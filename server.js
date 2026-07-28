require("dotenv").config();

const express = require("express");
const path = require("path");
const cron = require("node-cron");


const {
    generarExcelDesdeSupabase
} = require("./index");

const app = express();

const PORT = process.env.PORT || 3001;

const PUBLIC = path.join(__dirname, "public");

app.use(express.static(PUBLIC));

app.get("/", (req, res) => {

    res.send("Servidor de Excel funcionando");

});

let ejecutando = false;

async function actualizarExcel() {

    if (ejecutando) {

        console.log("Ya existe una actualización en curso.");

        return;

    }

    ejecutando = true;

    try {

        console.log("");
        console.log("==================================");
        console.log(new Date().toLocaleString());
        console.log("Actualizando Excel...");
        console.log("==================================");

        await generarExcelDesdeSupabase();

        console.log("Actualización finalizada.");

    } catch (err) {

        console.error("Error actualizando Excel");

        console.error(err);

    }

    ejecutando = false;

}

// Primera generación al iniciar
(async () => {

    await actualizarExcel();

})();

// Cada hora
cron.schedule("0 6 * * 1", async () => {

    console.log("Iniciando actualización programada...");

    await actualizarExcel();

});

app.listen(PORT, () => {

    console.log(`Servidor iniciado en puerto ${PORT}`);

    console.log(
        `Excel disponible en http://localhost:${PORT}/empresas.xlsx`
    );

});
