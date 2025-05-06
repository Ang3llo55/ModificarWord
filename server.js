// 2. Importar Módulos Necesarios
const express = require('express');
const path = require('path');
const fs = require('fs');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');
const { createClient } = require('@supabase/supabase-js');
const ILovePDFApi = require('@ilovepdf/ilovepdf-nodejs'); // Import iLovePDF library
const { v4: uuidv4 } = require('uuid'); // Import UUID for unique filenames

// --- 3. Inicialización del Cliente de Supabase ---
const supabaseUrl = "https://xhvsqqyqkbaugyslcdig.supabase.co";
const supabaseKey = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InhodnNxcXlxa2JhdWd5c2xjZGlnIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NDQ0ODMxMjcsImV4cCI6MjA2MDA1OTEyN30.iEetUawDtUa2bAtsGCy9mqfPDKVdg5KWWeI-MACkxOQ";

if (!supabaseUrl || !supabaseKey) {
    console.error("Error: Faltan SUPABASE_URL o SUPABASE_ANON_KEY.");
    process.exit(1);
}
const supabase = createClient(supabaseUrl, supabaseKey);

// --- NEW: iLovePDF Initialization ---
const iLovePdfPublicKey = "project_public_797132303a83ec8cada9802bbba3af43_4ePhO5e99f6f19f4a71511f6923577496aeb5";
const iLovePdfSecretKey = "secret_key_9845f12dea8853738fb3b0529cdd1344_mls1d5c3c0a0dcb3925f8f734e831259b374e";

if (!iLovePdfPublicKey || iLovePdfPublicKey === "project_public_797132303a83ec8cada9802bbba3af43_4ePhO5e99f6f19f4a71511f6923577496aeb5" || !iLovePdfSecretKey || iLovePdfSecretKey === "secret_key_9845f12dea8853738fb3b0529cdd1344_mls1d5c3c0a0dcb3925f8f734e831259b374e") {
    console.error("Error: Faltan ILOVEPDF_PUBLIC_KEY o ILOVEPDF_SECRET_KEY. Obtén las claves desde https://developer.ilovepdf.com/");
    // Consider exiting in production: process.exit(1);
    // For development, we might allow it to continue but conversion will fail.
    console.warn("Advertencia: Las claves de iLovePDF no están configuradas. La conversión a PDF fallará.");
}
const iLovePdfApi = new ILovePDFApi(iLovePdfPublicKey, iLovePdfSecretKey);

// --- NEW: Supabase Storage Bucket Configuration ---
const TEMP_BUCKET_NAME = 'temp-docx'; // *** IMPORTANT: Create this PUBLIC bucket in your Supabase project ***

// --- Helper Functions (formatearFecha, numeroALetras, formatearFechaEspecifica) ---
// ... (Keep your existing helper functions here) ...
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

    // Asegurarse de que el número sea un entero
    numero = parseInt(numero, 10);
    if (isNaN(numero)) return 'Error: entrada no es un número';

    if (numero < 0) return 'menos ' + numeroALetras(Math.abs(numero));

    if (numero < 10) {
        return unidades[numero];
    } else if (numero < 20) {
        return decenas[numero - 10];
    } else if (numero < 100) {
        const decena = Math.floor(numero / 10);
        const unidad = numero % 10;
        if (numero === 20) return 'veinte'; // Caso especial para 20 exacto
        if (numero > 20 && numero < 30) return 'veinti' + unidades[unidad]; // Caso para veintiuno, veintidós, etc.
        return decenasMultiplos[decena - 2] + (unidad === 0 ? '' : ' y ' + unidades[unidad]);
    } else if (numero < 1000) {
        const centena = Math.floor(numero / 100);
        const resto = numero % 100;
        if (numero === 100) return 'cien'; // Caso especial para 100 exacto
        return centenas[centena - 1] + (resto === 0 ? '' : ' ' + numeroALetras(resto));
    } else if (numero < 1000000) { // Añadir soporte para miles
        const miles = Math.floor(numero / 1000);
        const resto = numero % 1000;
        const milesEnLetras = (miles === 1) ? 'mil' : numeroALetras(miles) + ' mil';
        return milesEnLetras + (resto === 0 ? '' : ' ' + numeroALetras(resto));
    }
    else {
        // Limitar para evitar complejidad excesiva o errores con números muy grandes
        console.warn("Número fuera del rango común para conversión a letras:", numero);
        return numero.toString(); // Devolver el número como string si es muy grande
    }
}

// Función para formatear la fecha en el formato específico
function formatearFechaEspecifica(fecha) {
    const opcionesDia = { day: 'numeric' };
    const opcionesMes = { month: 'long' };
    const opcionesAnio = { year: 'numeric' };

    const diaNum = parseInt(fecha.toLocaleDateString('es-ES', opcionesDia), 10);
    const mes = fecha.toLocaleDateString('es-ES', opcionesMes);
    const anio = fecha.toLocaleDateString('es-ES', opcionesAnio);

    // Validar que diaNum sea un número antes de llamar a numeroALetras
    if (isNaN(diaNum)) {
        console.error("Error al obtener el día numérico de la fecha:", fecha);
        return `Error al formatear fecha`;
    }

    const diaEnLetras = numeroALetras(diaNum);
    // Formato: "veinticinco (25) días del mes de julio del 2024"
    return `${diaEnLetras} (${diaNum}) días del mes de ${mes} del ${anio}`;
}

// --- 4. Configuración de la Aplicación Express ---
const app = express();
const port = process.env.PORT || 3000;

// --- 5. Configuración Específica del Proyecto ---
const TEMPLATE_FOLDER = path.join(__dirname, 'word_templates');
const ACTIVE_TEMPLATE = 'plantilla_activo.docx';
const INACTIVE_TEMPLATE = 'plantilla_inactivo.docx';

// --- 6. Middleware ---
app.use(express.urlencoded({ extended: true }));
app.use(express.static(path.join(__dirname, 'public')));
app.set('view engine', 'ejs');
app.set('views', path.join(__dirname, 'views'));

// --- 7. Rutas de la Aplicación ---

app.get('/', (req, res) => {
    res.render('index');
});

app.post('/generate', async (req, res) => {
    const cedula = req.body.cedula;
    let tempSupabaseFilePath = null; // Variable to store the path for cleanup

    if (!cedula) {
        return res.status(400).send('Error: Por favor, introduce una cédula.');
    }

    console.log(`Buscando empleado con cédula: ${cedula}`);

    try {
        // --- A. Consultar Datos en Supabase ---
        const { data: employeeData, error: dbError } = await supabase
            .from('Base de datos')
            .select('cedula,nombres_y_apellidos,fecha_inicio,fecha_final,cargo_empresa,sede,salario,esta_activo')
            .eq('cedula', cedula)
            .maybeSingle();

        if (dbError) {
            console.error("Error de Supabase:", dbError);
            return res.status(500).send(`Error al consultar la base de datos: ${dbError.message}`);
        }

        if (!employeeData) {
            console.log(`Empleado con cédula ${cedula} no encontrado.`);
            return res.status(404).send(`Error: No se encontró ningún empleado con la cédula ${cedula}.`);
        }

        console.log("Empleado encontrado:", employeeData);

        // --- B. Determinar Estado y Elegir Plantilla ---
        const isActive = employeeData.esta_activo;
        const templateName = isActive ? ACTIVE_TEMPLATE : INACTIVE_TEMPLATE;
        const templatePath = path.join(TEMPLATE_FOLDER, templateName);

        console.log(`Empleado ${isActive ? 'activo' : 'inactivo'}. Usando plantilla: ${templateName}`);

        if (!fs.existsSync(templatePath)) {
             console.error(`Error: Archivo de plantilla no encontrado en ${templatePath}`);
             return res.status(404).send(`Error: Archivo de plantilla '${templateName}' no encontrado en el servidor.`);
        }

        // --- C. Preparar Datos para Docxtemplater ---
        const fechaActual = new Date();
        const fechaFormateada = formatearFecha(fechaActual);
        const fechaEspecifica = formatearFechaEspecifica(fechaActual);

        // Formatear fechas del empleado si existen
        if (employeeData.fecha_inicio) {
            try {
                employeeData.fecha_inicio_formateada = formatearFecha(new Date(employeeData.fecha_inicio));
            } catch (e) {
                console.warn("No se pudo formatear fecha_inicio:", employeeData.fecha_inicio, e);
                employeeData.fecha_inicio_formateada = employeeData.fecha_inicio; // Usar original si falla
            }
        } else {
             employeeData.fecha_inicio_formateada = "No especificada";
        }

        if (!isActive) {
            if (employeeData.fecha_final) {
                 try {
                    employeeData.fecha_final_formateada = formatearFecha(new Date(employeeData.fecha_final));
                 } catch (e) {
                    console.warn("No se pudo formatear fecha_final:", employeeData.fecha_final, e);
                    employeeData.fecha_final_formateada = employeeData.fecha_final; // Usar original si falla
                 }
            } else {
                employeeData.fecha_final_formateada = "No especificada";
            }
        }
        // Si necesitas usar la fecha final en la plantilla de activos (quizás como 'hasta la fecha')
        // puedes añadirla aquí o manejarla en la plantilla con condicionales si Docxtemplater lo soporta.

        // Datos a reemplazar en la plantilla
        const replacements = {
            ...employeeData, // Todos los campos de Supabase
            fechaFormateada: fechaFormateada, // Fecha actual general
            fechaEspecifica: fechaEspecifica, // Fecha actual detallada
            // Usa los nombres de campo formateados si los creaste:
            fecha_inicio: employeeData.fecha_inicio_formateada, // Sobrescribe el original con el formateado
            fecha_final: employeeData.fecha_final_formateada || '', // Sobrescribe con formateado o vacío si es activo
            // Asegúrate de que los nombres coincidan EXACTAMENTE con las etiquetas {{variable}} en tus .docx
            // Ejemplo si en tu .docx tienes {{nombre_completo}} en lugar de {{nombres_y_apellidos}}
             nombre_completo: employeeData.nombres_y_apellidos,
             cargo: employeeData.cargo_empresa,
             // ... cualquier otro mapeo necesario entre nombres de Supabase y nombres en la plantilla
        };


        // --- D. Generar el Documento Word (Buffer) ---
        const content = fs.readFileSync(templatePath, 'binary');
        const zip = new PizZip(content);
        const doc = new Docxtemplater(zip, {
            paragraphLoop: true,
            linebreaks: true,
            delimiters: { start: '{{', end: '}}' },
            nullGetter: function(part) {
                // Si la variable no se encuentra en `replacements`, inserta una cadena vacía
                // Esto evita errores si una etiqueta existe en la plantilla pero no en los datos
                 if (!part.module) {
                     console.warn(`Advertencia: Etiqueta '{{${part.value}}}' encontrada en plantilla pero sin valor en los datos. Se reemplazará por vacío.`);
                     return "";
                 }
                 return undefined;
             }
        });

        // Renderizar el documento con los datos preparados
        doc.render(replacements);

        const outputDocxBuffer = doc.getZip().generate({
            type: 'nodebuffer',
            compression: "DEFLATE",
        });

        console.log("Documento DOCX generado en memoria.");

        // --- E. Subir DOCX a Supabase Storage ---
        const uniqueFilename = `${uuidv4()}_${cedula}.docx`; // Crear nombre único
        tempSupabaseFilePath = `${uniqueFilename}`; // Guardar solo el nombre para borrar luego

        console.log(`Subiendo DOCX a Supabase Storage: ${TEMP_BUCKET_NAME}/${tempSupabaseFilePath}`);

        const { error: uploadError } = await supabase.storage
            .from(TEMP_BUCKET_NAME)
            .upload(tempSupabaseFilePath, outputDocxBuffer, {
                contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                upsert: false // No sobrescribir si existe (improbable con UUID)
            });

        if (uploadError) {
            console.error("Error subiendo DOCX a Supabase Storage:", uploadError);
            throw new Error(`Error al subir archivo temporal: ${uploadError.message}`);
            // No necesitas retornar aquí, el catch general lo manejará
        }

        console.log("DOCX subido correctamente a Supabase Storage.");

        // --- F. Obtener URL Pública del DOCX ---
        const { data: urlData } = supabase.storage
            .from(TEMP_BUCKET_NAME)
            .getPublicUrl(tempSupabaseFilePath);

        if (!urlData || !urlData.publicUrl) {
             console.error("Error obteniendo URL pública de Supabase Storage para:", tempSupabaseFilePath);
             throw new Error('No se pudo obtener la URL pública del archivo temporal.');
        }
        const docxPublicUrl = urlData.publicUrl;
        console.log(`URL Pública del DOCX: ${docxPublicUrl}`);

        // --- G. Convertir DOCX a PDF usando iLovePDF ---
        console.log("Iniciando tarea de conversión a PDF con iLovePDF...");
        const task = iLovePdfApi.newTask('officepdf');
        await task.addFile(docxPublicUrl); // Añadir el archivo usando la URL pública
        await task.process({ // Opciones de procesamiento (puedes añadir más si es necesario)
             output_filename: `Certificado_${employeeData.nombres_y_apellidos || cedula}_${isActive ? 'Activo' : 'Inactivo'}`,
        });

        console.log("Tarea de iLovePDF procesada. Descargando PDF...");
        const pdfBuffer = await task.download(); // Descarga el PDF resultante como buffer
        console.log("PDF descargado de iLovePDF.");

        // --- H. Enviar el PDF al Cliente ---
        const outputPdfFilename = `Certificado_${employeeData.nombres_y_apellidos || cedula}_${isActive ? 'Activo' : 'Inactivo'}.pdf`;

        res.setHeader('Content-Disposition', `attachment; filename="${outputPdfFilename}"`);
        res.setHeader('Content-Type', 'application/pdf');
        res.send(pdfBuffer); // Enviar el buffer del PDF

        console.log(`PDF generado y enviado al cliente como ${outputPdfFilename}`);

    } catch (error) {
        console.error("Error procesando la solicitud:", error);
        // Manejo de errores específicos si es necesario
        if (error.properties && error.properties.errors) { // Error de Docxtemplater
            console.error("Errores de Docxtemplater:", error.properties.errors);
            res.status(500).send(`Error al generar el documento base (Error de plantilla): ${error.message}.`);
        } else if (error.response && error.response.data) { // Posible error de iLovePDF API
            console.error("Error de la API iLovePDF:", error.response.data);
            res.status(500).send(`Error durante la conversión a PDF: ${error.response.data.error?.message || error.message}`);
        }
         else { // Otro error (Supabase, red, etc.)
            res.status(500).send(`Error interno del servidor: ${error.message}`);
        }
    } finally {
        // --- I. Limpieza: Eliminar el archivo DOCX temporal de Supabase Storage ---
        if (tempSupabaseFilePath) {
            console.log(`Limpiando archivo temporal de Supabase Storage: ${TEMP_BUCKET_NAME}/${tempSupabaseFilePath}`);
            const { error: deleteError } = await supabase.storage
                .from(TEMP_BUCKET_NAME)
                .remove([tempSupabaseFilePath]); // .remove() espera un array de paths

            if (deleteError) {
                // Loguear el error pero no impedir que la respuesta (si ya se envió) llegue al usuario
                console.error("Error eliminando archivo temporal de Supabase Storage:", deleteError);
            } else {
                console.log("Archivo temporal eliminado correctamente.");
            }
        }
    }
});

// --- 8. Iniciar el Servidor ---
app.listen(port, () => {
    console.log(`Servidor escuchando en http://localhost:${port}`);
});