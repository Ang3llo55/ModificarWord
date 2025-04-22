// 2. Importar Módulos Necesarios
const express = require('express'); // Framework web
const path = require('path'); // Para manejar rutas de archivos
const fs = require('fs'); // Para leer archivos (las plantillas .docx)
const PizZip = require('pizzip'); // Para leer/manipular el contenido del .docx como un zip
const Docxtemplater = require('docxtemplater'); // La librería principal para reemplazar texto en .docx
const { createClient } = require('@supabase/supabase-js'); // El cliente oficial de Supabase

// --- 3. Inicialización del Cliente de Supabase ---
const supabaseUrl = "https://xhvsqqyqkbaugyslcdig.supabase.co"; // Obtiene la URL desde .env
const supabaseKey = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InhodnNxcXlxa2JhdWd5c2xjZGlnIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NDQ0ODMxMjcsImV4cCI6MjA2MDA1OTEyN30.iEetUawDtUa2bAtsGCy9mqfPDKVdg5KWWeI-MACkxOQ"; // Obtiene la clave pública desde .env

// Verificación de seguridad: si faltan las credenciales, detiene el servidor
if (!supabaseUrl || !supabaseKey) {
    console.error("Error: Faltan SUPABASE_URL o SUPABASE_ANON_KEY en el archivo .env.");
    process.exit(1); // Termina la ejecución del script
}
// Crea la instancia del cliente Supabase para interactuar con tu base de datos
const supabase = createClient(supabaseUrl, supabaseKey);

// Función para formatear la fecha
function formatearFecha(fecha) {
    const opciones = { day: 'numeric', month: 'long', year: 'numeric' };
    return fecha.toLocaleDateString('es-ES', opciones);
}

// Función para convertir un número a su representación en letras
function numeroALetras(numero) {
    const unidades = ['cero', 'uno', 'dos', 'tres', 'cuatro', 'cinco', 'seis', 'siete', 'ocho', 'nueve'];
    const decenas = ['diez', 'once', 'doce', 'trece', 'catorce', 'quince', 'dieciséis', 'diecisiete', 'dieciocho', 'diecinueve'];
    const decenasMultiplos = ['veinte', 'treinta', 'cuarenta', 'cincuenta', 'sesenta', 'setenta', 'ochenta', 'noventa'];
    const centenas = ['ciento', 'doscientos', 'trescientos', 'cuatrocientos', 'quinientos', 'seiscientos', 'setecientos', 'ochocientos', 'novecientos'];

    if (numero < 10) {
        return unidades[numero];
    } else if (numero < 20) {
        return decenas[numero - 10];
    } else if (numero < 100) {
        const decena = Math.floor(numero / 10);
        const unidad = numero % 10;
        return decenasMultiplos[decena - 2] + (unidad === 0 ? '' : ' y ' + unidades[unidad]);
    } else if (numero < 1000) {
        const centena = Math.floor(numero / 100);
        const resto = numero % 100;
        return centenas[centena - 1] + (resto === 0 ? '' : ' ' + numeroALetras(resto));
    } else {
        return 'Error: Número fuera del rango permitido';
    }
}

// Función para formatear la fecha en el formato específico
function formatearFechaEspecifica(fecha) {
    const opciones = { day: '2-digit', month: 'long', year: 'numeric' };
    const partesFecha = fecha.toLocaleDateString('es-ES', opciones).split(' ');
    const dia = partesFecha[0];
    const mes = partesFecha[2];
    const año = partesFecha[4];
    const diaEnLetras = numeroALetras(parseInt(dia, 10));
    return `${diaEnLetras} (${dia}) días del mes de ${mes} del ${año}`;
}

// --- 4. Configuración de la Aplicación Express ---
const app = express(); // Crea la aplicación Express
const port = process.env.PORT || 3000; // Puerto de escucha (configurable o 3000 por defecto)

// --- 5. Configuración Específica del Proyecto ---
const TEMPLATE_FOLDER = path.join(__dirname, 'word_templates'); // Ruta a la carpeta de plantillas .docx
const ACTIVE_TEMPLATE = 'plantilla_activo.docx';    // Nombre de la plantilla para activos
const INACTIVE_TEMPLATE = 'plantilla_inactivo.docx'; // Nombre de la plantilla para inactivos

// --- 6. Middleware (Funciones que se ejecutan en cada petición) ---
// Parsea datos de formularios enviados como URL-encoded (lo normal desde HTML)
app.use(express.urlencoded({ extended: true }));
// Sirve archivos estáticos (CSS, JS del cliente, imágenes) desde la carpeta 'public'
app.use(express.static(path.join(__dirname, 'public')));
// Configura EJS como el motor para renderizar las vistas HTML
app.set('view engine', 'ejs');
app.set('views', path.join(__dirname, 'views')); // Indica dónde están las vistas (.ejs)

// --- 7. Rutas de la Aplicación ---

// Ruta Principal (GET /): Muestra el formulario inicial
app.get('/', (req, res) => {
    res.render('index'); // Renderiza el archivo views/index.ejs
});

// Ruta de Generación (POST /generate): Procesa la cédula y genera el documento
// *** Se marca como 'async' para poder usar 'await' con la consulta a Supabase ***
app.post('/generate', async (req, res) => {
    // Obtiene la cédula enviada desde el formulario HTML
    const cedula = req.body.cedula;

    // Validación básica: Si no se envió cédula, devuelve un error
    if (!cedula) {
        return res.status(400).send('Error: Por favor, introduce una cédula.');
    }

    console.log(`Buscando empleado con cédula: ${cedula}`); // Mensaje para depuración

    try {
        // --- A. Consultar Datos en Supabase ---
        // Ajusta 'empleados' al nombre real de tu tabla
        // Ajusta 'cedula' al nombre real de tu columna de cédula
        const { data: employeeData, error: dbError } = await supabase
            .from('Base de datos') // Nombre de tu tabla en Supabase
            .select('cedula,nombres_y_apellidos,fecha_inicio,fecha_final,cargo_empresa,sede,salario,esta_activo')       // Selecciona todas las columnas (o especifica las necesarias: 'nombre, cargo, esta_activo, ...')
            .eq('cedula', cedula) // Filtra: busca donde la columna 'cedula' sea igual a la cédula recibida
            .maybeSingle();    // Espera 0 o 1 resultado. Devuelve 'null' si no lo encuentra (sin error).

        // Manejo de Error de Base de Datos
        if (dbError) {
            console.error("Error de Supabase:", dbError);
            return res.status(500).send(`Error al consultar la base de datos: ${dbError.message}`);
        }

        // Manejo de Empleado No Encontrado
        if (!employeeData) {
            console.log(`Empleado con cédula ${cedula} no encontrado.`);
            return res.status(404).send(`Error: No se encontró ningún empleado con la cédula ${cedula}.`);
        }

        console.log("Empleado encontrado:", employeeData); // Muestra los datos encontrados

        // --- B. Determinar Estado y Elegir Plantilla ---
        // Ajusta 'esta_activo' al nombre real de tu columna booleana de estado
        const isActive = employeeData.esta_activo; // Obtiene el valor booleano (true/false)
        const templateName = isActive ? ACTIVE_TEMPLATE : INACTIVE_TEMPLATE; // Elige el nombre del archivo .docx
        const templatePath = path.join(TEMPLATE_FOLDER, templateName); // Construye la ruta completa a la plantilla

        console.log(`Empleado ${isActive ? 'activo' : 'inactivo'}. Usando plantilla: ${templateName}`);

        // Verifica si el archivo de plantilla realmente existe antes de intentar leerlo
        if (!fs.existsSync(templatePath)) {
             console.error(`Error: Archivo de plantilla no encontrado en ${templatePath}`);
             return res.status(404).send(`Error: Archivo de plantilla '${templateName}' no encontrado en el servidor.`);
        }

        // --- C. Preparar Datos para Docxtemplater ---
        // Los datos de 'employeeData' ya vienen como un objeto {columna: valor}.
        // Si necesitas formatear algo (ej. fechas), este es un buen lugar.
        // Ejemplo: Formatear fecha_ingreso a formato local (si viene como ISO)
        if (employeeData.fecha_ingreso) {
            employeeData.fecha_ingreso_formateada = new Date(employeeData.fecha_ingreso).toLocaleDateString('es-ES');
            // En tu plantilla usarías {{fecha_ingreso_formateada}}
        }
        // Asegurarse de que 'fecha_final' tenga un valor si está inactivo
        if (!isActive && employeeData.fecha_final) {
            // Podrías formatear 'fecha_final' aquí si es necesario
            // employeeData.fecha_final_formateada = ...
        } else if (!isActive) {
            // Si está inactivo pero no hay fecha_final, asigna un valor por defecto
            employeeData.fecha_final = employeeData.fecha_final || "No especificada";
        }

        // --- D. Generar el Documento Word ---
        // 1. Lee el contenido binario de la plantilla seleccionada
        const content = fs.readFileSync(templatePath, 'binary');
        // 2. Carga ese contenido en PizZip
        const zip = new PizZip(content);
        // 3. Crea la instancia de Docxtemplater
        const doc = new Docxtemplater(zip, {
            paragraphLoop: true, // Permite bucles en párrafos
            linebreaks: true,    // Interpreta saltos de línea (\n) como saltos de línea en Word
            delimiters: { start: '{{', end: '}}' }, // Usa {{variable}} como delimitadores
            // ¡Importante! Maneja el caso de variables nulas o indefinidas en los datos
            // Si una variable {{ejemplo}} está en la plantilla pero employeeData.ejemplo no existe o es null,
            // en lugar de dar error, insertará una cadena vacía "".
             nullGetter: function(part) {
                if (!part.module) {
                     return ""; // Devuelve cadena vacía para variables simples no encontradas
                }
                return undefined; // Deja que Docxtemplater maneje sus módulos internos
             }
        });

        // Obtener la fecha actual
        const fechaActual = new Date();

        // Formatear la fecha
        const fechaFormateada = formatearFecha(fechaActual);

        // Formatear la fecha en el formato específico
        const fechaEspecifica = formatearFechaEspecifica(fechaActual);

        // 5. Pasar los datos (obtenidos de Supabase) a la plantilla
        const replacementsWithDate = {
            ...employeeData,
            fechaFormateada: fechaFormateada,
            fechaEspecifica: fechaEspecifica
        };

        // 4. Pasa los datos del empleado a la plantilla
        doc.setData(replacementsWithDate);
        // 5. Realiza los reemplazos (renderiza el documento en memoria)
        doc.render();
        // 6. Obtiene el documento final como un buffer de Node.js
        const outputBuffer = doc.getZip().generate({
            type: 'nodebuffer',
            compression: "DEFLATE", // Compresión estándar para .docx
        });

        // --- E. Enviar el Documento al Cliente ---
        // Nombre sugerido para el archivo descargado
        const outputFilename = `Certificado_${employeeData.nombre_completo || cedula}_${isActive ? 'Activo' : 'Inactivo'}.docx`;

        // Configura las cabeceras HTTP para indicar que es una descarga
        res.setHeader('Content-Disposition', `attachment; filename="${outputFilename}"`);
        // Indica el tipo de contenido correcto para un archivo .docx
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
        // Envía el buffer del archivo .docx como respuesta
        res.send(outputBuffer);

    } catch (error) { // Bloque 'catch' para capturar cualquier error inesperado
        console.error("Error procesando la solicitud:", error);
        // Manejo de errores específicos de Docxtemplater (ej. error de sintaxis en plantilla)
         if (error.properties && error.properties.errors) {
            console.error("Errores de Docxtemplater:", error.properties.errors);
            res.status(500).send(`Error al generar el documento (Error de plantilla): ${error.message}.`);
        } else {
            // Error genérico del servidor
            res.status(500).send(`Error interno del servidor al procesar la solicitud: ${error.message}`);
        }
    }
});

// --- 8. Iniciar el Servidor ---
app.listen(port, () => {
    console.log(`Servidor escuchando en http://localhost:${port}`);
});