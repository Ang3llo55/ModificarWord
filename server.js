const express = require('express');
const path = require('path');
const fs = require('fs');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');
const { createClient } = require('@supabase/supabase-js'); // Importa el cliente de Supabase

const app = express();
const port = process.env.PORT || 3000;

// --- Configuración ---
const TEMPLATE_FOLDER = path.join(__dirname, 'word_templates');
const DEFAULT_TEMPLATE = 'mi_plantilla_1.docx'; // Asegúrate que esta plantilla exista

// --- Configuración de Supabase ---
const supabaseUrl = "https://xhvsqqyqkbaugyslcdig.supabase.co";
const supabaseKey = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InhodnNxcXlxa2JhdWd5c2xjZGlnIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NDQ0ODMxMjcsImV4cCI6MjA2MDA1OTEyN30.iEetUawDtUa2bAtsGCy9mqfPDKVdg5KWWeI-MACkxOQ";

if (!supabaseUrl || !supabaseKey) {
    console.error("Error: Faltan las variables de entorno SUPABASE_URL o SUPABASE_ANON_KEY.");
    process.exit(1); // Detiene la aplicación si faltan las claves
}
const supabase = createClient(supabaseUrl, supabaseKey);

// --- Middleware ---
app.use(express.urlencoded({ extended: true }));
app.use(express.static(path.join(__dirname, 'public')));
app.set('view engine', 'ejs');
app.set('views', path.join(__dirname, 'views'));

// --- Rutas ---

// Ruta principal - Muestra el formulario (ahora solo pide la cédula)
app.get('/', (req, res) => {
    res.render('index', { defaultTemplate: DEFAULT_TEMPLATE, error: null }); // Pasa null como error inicial
});

// Ruta para generar y descargar el documento (AHORA ASÍNCRONA por la llamada a DB)
app.post('/generate', async (req, res) => { // <--- Marcada como async
    const cedula = req.body.cedula; // Obtiene la cédula del formulario
    const templatePath = path.join(TEMPLATE_FOLDER, DEFAULT_TEMPLATE);

    if (!cedula) {
        // Si no se envió cédula, renderiza la misma página con un error
        return res.status(400).render('index', {
             defaultTemplate: DEFAULT_TEMPLATE,
             error: 'Por favor, introduce una cédula.'
        });
    }

    try {
        // 1. Buscar datos en Supabase usando la cédula
        //    AJUSTA 'clientes' al nombre de tu tabla y 'cedula' al nombre de tu columna de ID
        const { data, error: dbError } = await supabase
            .from('Base de datos') // <- Reemplaza 'clientes' con el nombre de tu tabla
            .select('cedula,nombres_y_apellidos,fecha_inicio,cargo_empresa,sede,salario') // Selecciona todas las columnas necesarias para la plantilla
            .eq('cedula', cedula) // <- Reemplaza 'cedula' si tu columna se llama diferente
            .single(); // Espera encontrar un único registro (o null)

        // Manejar errores de la base de datos
        if (dbError) {
            console.error("Error de Supabase:", dbError);
            // No devuelvas el error de Supabase directamente al usuario por seguridad
            return res.status(500).render('index', {
                defaultTemplate: DEFAULT_TEMPLATE,
                error: 'Error al consultar la base de datos.'
            });
        }

        // Manejar caso donde no se encuentra la cédula
        if (!data) {
            return res.status(404).render('index', {
                defaultTemplate: DEFAULT_TEMPLATE,
                error: `No se encontraron datos para la cédula: ${cedula}`
            });
        }

        // 'data' ahora contiene el objeto con los datos del cliente, ej:
        // { cedula: '123', nombre_cliente: 'Juan Perez', numero_pedido: 'P001', ... }
        const replacements = data; // Usamos directamente los datos de Supabase

        // 2. Leer el archivo de plantilla .docx
        if (!fs.existsSync(templatePath)) {
             console.error(`Error: No se encontró la plantilla en ${templatePath}`);
             return res.status(404).render('index', {
                 defaultTemplate: DEFAULT_TEMPLATE,
                 error: 'Archivo de plantilla no encontrado en el servidor.'
             });
        }
        const content = fs.readFileSync(templatePath, 'binary');

        // 3. Cargar el contenido con PizZip
        const zip = new PizZip(content);

        // 4. Crear una instancia de Docxtemplater (con delimitadores {{ }})
        const doc = new Docxtemplater(zip, {
            paragraphLoop: true,
            linebreaks: true,
            delimiters: { start: '{{', end: '}}' } // Asegúrate que coincida con tu plantilla
        });

        // 5. Pasar los datos (obtenidos de Supabase) a la plantilla
        doc.setData(replacements);

        // 6. Realizar el reemplazo
        doc.render(); // Puede lanzar errores si faltan variables en 'replacements' que están en la plantilla

        // 7. Generar el buffer del documento final
        const outputBuffer = doc.getZip().generate({
            type: 'nodebuffer',
            compression: "DEFLATE",
        });

        // 8. Preparar el nombre del archivo de salida (usa la cédula o un dato relevante)
        const outputFilename = `documento_${replacements.cedula || Date.now()}.docx`;

        // 9. Establecer las cabeceras para la descarga
        res.setHeader('Content-Disposition', `attachment; filename="${outputFilename}"`);
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');

        // 10. Enviar el buffer como respuesta
        res.send(outputBuffer);

    } catch (error) { // Captura errores generales (renderizado docxtemplater, etc.)
        console.error("Error generando el documento:", error);

        // Manejo específico de errores de Docxtemplater si es posible
        if (error.properties && error.properties.errors) {
             console.error("Errores de Docxtemplater:", error.properties.errors);
             return res.status(400).render('index', {
                 defaultTemplate: DEFAULT_TEMPLATE,
                 error: `Error en la plantilla: ${error.message}. Verifica que todas las variables {{variable}} existan en la base de datos para esta cédula.`
             });
        } else {
            // Error genérico
            return res.status(500).render('index', {
                defaultTemplate: DEFAULT_TEMPLATE,
                error: `Error interno del servidor al generar el documento.`
            });
        }
    }
});

// --- Iniciar Servidor ---
app.listen(port, () => {
    console.log(`Servidor escuchando en http://localhost:${port}`);
});