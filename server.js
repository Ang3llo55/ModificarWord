const express = require('express');
const path = require('path');
const fs = require('fs');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');

const app = express();
const port = process.env.PORT || 3000; // Usa el puerto de Heroku/otro o 3000 localmente

// --- Configuración ---
const TEMPLATE_FOLDER = path.join(__dirname, 'word_templates');
const DEFAULT_TEMPLATE = 'mi_plantilla_1.docx'; // Puedes cambiar esto o hacerlo dinámico

// --- Middleware ---
// Para parsear datos de formularios URL-encoded
app.use(express.urlencoded({ extended: true }));
// Para servir archivos estáticos (CSS, JS del cliente) desde la carpeta 'public'
app.use(express.static(path.join(__dirname, 'public')));
// Configurar EJS como motor de plantillas
app.set('view engine', 'ejs');
app.set('views', path.join(__dirname, 'views')); // Especifica el directorio de vistas

// --- Rutas ---

// Ruta principal - Muestra el formulario
app.get('/', (req, res) => {
    // Pasar el nombre de la plantilla por defecto a la vista
    res.render('index', { defaultTemplate: DEFAULT_TEMPLATE });
});

// Ruta para generar y descargar el documento
app.post('/generate', (req, res) => {
    const templatePath = path.join(TEMPLATE_FOLDER, DEFAULT_TEMPLATE);

    try {
        // 1. Leer el archivo de plantilla .docx
        const content = fs.readFileSync(templatePath, 'binary');

        // 2. Cargar el contenido con PizZip
        const zip = new PizZip(content);

        // 3. Crear una instancia de Docxtemplater
        //    Configuramos los delimitadores para que coincidan con {{variable}}
        const doc = new Docxtemplater(zip, {
            paragraphLoop: true,
            linebreaks: true,
            delimiters: { start: '{{', end: '}}' } // ¡Importante!
        });

        // 4. Obtener los datos del formulario (req.body)
        const formData = req.body;
        console.log("Datos recibidos:", formData); // Para depuración

        // 5. Pasar los datos a la plantilla
        //    docxtemplater espera un objeto donde las claves coinciden
        //    con los placeholders (sin los delimitadores)
        doc.setData(formData);

        // 6. Realizar el reemplazo (renderizar el documento)
        doc.render(); // Esto puede lanzar errores si hay problemas con la plantilla/datos

        // 7. Generar el buffer del documento final
        const outputBuffer = doc.getZip().generate({
            type: 'nodebuffer',
            // compression: DEFLATE funciona bien para docx
            compression: "DEFLATE",
        });

        // 8. Preparar el nombre del archivo de salida
        const outputFilename = `documento_generado_${formData.numero_pedido || Date.now()}.docx`;

        // 9. Establecer las cabeceras para la descarga
        res.setHeader('Content-Disposition', `attachment; filename="${outputFilename}"`);
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');

        // 10. Enviar el buffer como respuesta
        res.send(outputBuffer);

    } catch (error) {
        console.error("Error generando el documento:", error);
        // Manejo básico de errores (podrías mostrar una página de error)
        if (error.code === 'ENOENT') {
            res.status(404).send(`Error: No se encontró la plantilla en ${templatePath}`);
        } else if (error.properties && error.properties.errors) {
            // Errores específicos de Docxtemplater (ej. variable no encontrada)
            console.error("Errores de Docxtemplater:", error.properties.errors);
            res.status(400).send(`Error en la plantilla: ${error.message}. Detalles: ${JSON.stringify(error.properties.errors)}`);
        }
         else {
            res.status(500).send(`Error interno del servidor al generar el documento: ${error.message}`);
        }
    }
});

// --- Iniciar Servidor ---
app.listen(port, () => {
    console.log(`Servidor escuchando en http://localhost:${port}`);
});