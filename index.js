require("dotenv").config();

const { createClient } = require("@supabase/supabase-js");
const ExcelJS = require("exceljs");
const fs = require("fs");
const path = require("path");

// ================================
// CONFIGURACIÓN
// ================================

const supabase = createClient(
    process.env.SUPABASE_URL,
    process.env.SUPABASE_SERVICE_KEY
);

const OUTPUT_FOLDER = path.join(__dirname, "public");
const OUTPUT_FILE = path.join(OUTPUT_FOLDER, "empresas.xlsx");

// Mapeo de columnas
const COLUMNAS = {
    trade_name: "RAZON_SOCIAL",
    nit: "NIT",
    fecha_matricula: "FECHA_MATRICULA",
    ciiu_code: "CIIU",
    category: "CLASIFICACION (NUEVA)",
    city: "CIUDAD",
    exports_usd: "EXPORTACIONES USD 2024"
};

// ================================
// OBTENER DATOS
// ================================

async function obtenerEmpresas() {

    console.log("Consultando Supabase...");

    const empresas = [];

    const PAGE_SIZE = 1000;

    let desde = 0;

    while (true) {

        const { data, error } = await supabase
            .from("companies")
            .select("*")
            .range(desde, desde + PAGE_SIZE - 1);

        if (error) throw error;

        empresas.push(...data);

        console.log(`Leídas ${empresas.length} empresas...`);

        if (data.length < PAGE_SIZE) break;

        desde += PAGE_SIZE;

    }

    console.log(`Total empresas: ${empresas.length}`);

    return empresas;

}

// ================================
// OBTENER TODOS LOS AÑOS
// ================================

function obtenerAnios(empresas) {

    const anios = new Set();

    empresas.forEach(empresa => {

        if (!empresa.sales_by_year)
            return;

        const ventas = empresa.sales_by_year;

        Object.keys(ventas).forEach(anio => {
            anios.add(anio);
        });

    });

    return [...anios].sort();

}

// ================================
// CREAR EXCEL
// ================================

async function generarExcel(empresas, anios) {

    console.log("Generando Excel...");

    const workbook = new ExcelJS.Workbook();

    const worksheet = workbook.addWorksheet("Empresas");

    //--------------------------------------------------
    // Encabezados
    //--------------------------------------------------

    const encabezados = [

        COLUMNAS.trade_name,
        COLUMNAS.nit,
        COLUMNAS.fecha_matricula,
        COLUMNAS.ciiu_code,
        COLUMNAS.category,
        COLUMNAS.city,
        COLUMNAS.exports_usd

    ];

    anios.forEach(anio => {

        encabezados.push(
            `INGRESOS_ACTIVIDAD_ORDINARIA_${anio}`
        );

    });

    worksheet.addRow(encabezados);

    //--------------------------------------------------
    // Estilo encabezado
    //--------------------------------------------------

    worksheet.getRow(1).font = {
        bold: true
    };

    //--------------------------------------------------
    // Filas
    //--------------------------------------------------

    empresas.forEach(empresa => {

        const fila = [

            empresa.trade_name,
            empresa.nit,
            empresa.fecha_matricula,
            empresa.ciiu_code,
            empresa.category,
            empresa.city,
            empresa.exports_usd

        ];

        anios.forEach(anio => {

            if (
                empresa.sales_by_year &&
                empresa.sales_by_year[anio] !== undefined
            ) {

                fila.push(
                    empresa.sales_by_year[anio]
                );

            } else {

                fila.push("");

            }

        });

        worksheet.addRow(fila);

    });

    //--------------------------------------------------
    // Ajustar ancho columnas
    //--------------------------------------------------

    worksheet.columns.forEach(column => {

        let max = 20;

        column.eachCell({ includeEmpty: true }, cell => {

            const length = cell.value
                ? cell.value.toString().length
                : 10;

            if (length > max)
                max = length;

        });

        column.width = max + 2;

    });

    //--------------------------------------------------
    // Crear carpeta si no existe
    //--------------------------------------------------

    if (!fs.existsSync(OUTPUT_FOLDER)) {

        fs.mkdirSync(OUTPUT_FOLDER);

    }

    //--------------------------------------------------
    // Guardar
    //--------------------------------------------------

    await workbook.xlsx.writeFile(OUTPUT_FILE);

    console.log("Excel generado correctamente");

    console.log(OUTPUT_FILE);

}

// ================================
// EXPORTAR FUNCIÓN
// ================================

async function generarExcelDesdeSupabase() {

    console.log("--------------------------------------");
    console.log("GENERANDO EXCEL");
    console.log("--------------------------------------");

    const empresas = await obtenerEmpresas();

    const anios = obtenerAnios(empresas);

    console.log("Años encontrados:", anios);

    await generarExcel(empresas, anios);

    console.log("--------------------------------------");
    console.log("EXCEL ACTUALIZADO");
    console.log("--------------------------------------");

}

module.exports = {
    generarExcelDesdeSupabase
};
