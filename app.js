require('dotenv').config();
const express = require('express');
const path = require('path');
const bodyParser = require('body-parser');
const ExcelJS = require('exceljs');
const { sql, getPool } = require('./db');

const app = express();
const PORT = process.env.PORT || 3000;

app.use(bodyParser.json());
app.use(express.static(path.join(__dirname, 'public')));

// Búsqueda de colaboradores
app.get('/api/colaboradores', async (req, res) => {
    const q = req.query.search;
    const pool = await getPool();
    const result = await pool.request()
        .input('q', sql.NVarChar, `%${q}%`)
        .query('SELECT id_usuario, nombre, area, cargo, correo FROM Usuarios WHERE nombre LIKE @q');
    res.json(result.recordset);
});

// Generar y enviar Excel
app.post('/generate-excel', async (req, res) => {
    const data = req.body;
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(path.join(__dirname, 'public', 'templates', 'template.xlsx'));
    const sheet = workbook.getWorksheet(1);

    // Rellenar celdas
    sheet.getCell('D7').value = data.nombreEquipo;
    // Marcar tipo de equipo
    const tipos = {
        pc: 'C10', celular: 'C12', monitor: 'C14', uniformes: 'C16',
        notebook: 'H10', raton: 'H12', mochila: 'H14', otros: 'H16'
    };
    sheet.getCell(tipos[data.tipoEquipo]).value = 'X';

    sheet.getCell('B20').value = data.colaborador.nombre;
    sheet.getCell('B22').value = data.colaborador.area;
    sheet.getCell('B24').value = data.colaborador.cargo;
    sheet.getCell('B26').value = data.colaborador.correo;

    sheet.getCell('I20').value = new Date().toLocaleDateString('es-CL');

    const usos = { oficina: 'I24', terreno: 'I26' };
    sheet.getCell(usos[data.usoEquipo]).value = 'X';

    sheet.getCell('B33').value = data.marca;
    sheet.getCell('B35').value = data.modelo;
    sheet.getCell('B37').value = data.serie;
    sheet.getCell('B39').value = data.hostname;
    sheet.getCell('B41').value = data.imei;
    sheet.getCell('B43').value = data.otrosDetalle;
    sheet.getCell('F34').value = data.estadoEquipo;
    sheet.getCell('F38').value = data.comentarios;
    sheet.getCell('B48').value = data.elementosAdicionales;

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename=acta_equipo.xlsx');

    await workbook.xlsx.write(res);
    res.end();
});

// Ruta para mostrar formulario de devolución
app.get('/devolucion', (req, res) => {
    res.sendFile(path.join(__dirname, 'public', 'devolucion.html'));
});

// Generar y enviar Excel para devolución
app.post('/generate-excel-devolucion', async (req, res) => {
    const data = req.body;
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(
        path.join(__dirname, 'public', 'templates', 'templateD.xlsx')
    );
    const sheet = workbook.getWorksheet(1);

    // Rellenar celdas (igual a entrega, sin "utilización de equipo")
    sheet.getCell('D7').value = data.nombreEquipo;
    const tipos = {
        pc: 'C10', celular: 'C12', monitor: 'C14', uniformes: 'C16',
        notebook: 'H10', raton: 'H12', mochila: 'H14', otros: 'H16'
    };
    sheet.getCell(tipos[data.tipoEquipo]).value = 'X';

    sheet.getCell('B20').value = data.colaborador.nombre;
    sheet.getCell('B22').value = data.colaborador.area;
    sheet.getCell('B24').value = data.colaborador.cargo;
    sheet.getCell('B26').value = data.colaborador.correo;

    sheet.getCell('I20').value = new Date().toLocaleDateString('es-CL');

    // Campos técnicos y comentarios
    sheet.getCell('B33').value = data.marca;
    sheet.getCell('B35').value = data.modelo;
    sheet.getCell('B37').value = data.serie;
    sheet.getCell('B39').value = data.hostname;
    sheet.getCell('B41').value = data.imei;
    sheet.getCell('B43').value = data.otrosDetalle;
    sheet.getCell('F34').value = data.estadoEquipo;
    sheet.getCell('F38').value = data.comentarios;
    sheet.getCell('B48').value = data.elementosAdicionales;

    res.setHeader(
        'Content-Type',
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    );
    res.setHeader(
        'Content-Disposition',
        'attachment; filename=acta_devolucion.xlsx'
    );

    await workbook.xlsx.write(res);
    res.end();
});

// Ruta para mostrar formulario de devolución
app.get('/devolucion', (req, res) => {
    res.sendFile(path.join(__dirname, 'public', 'devolucion.html'));
});

// Generar y enviar Excel para devolución
app.post('/generate-excel-devolucion', async (req, res) => {
    const data = req.body;
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(
        path.join(__dirname, 'public', 'templates', 'templateD.xlsx')
    );
    const sheet = workbook.getWorksheet(1);

    // Rellenar celdas (igual a entrega, sin "utilización de equipo")
    sheet.getCell('D7').value = data.nombreEquipo;
    const tipos = {
        pc: 'C10', celular: 'C12', monitor: 'C14', uniformes: 'C16',
        notebook: 'H10', raton: 'H12', mochila: 'H14', otros: 'H16'
    };
    sheet.getCell(tipos[data.tipoEquipo]).value = 'X';

    sheet.getCell('B20').value = data.colaborador.nombre;
    sheet.getCell('B22').value = data.colaborador.area;
    sheet.getCell('B24').value = data.colaborador.cargo;
    sheet.getCell('B26').value = data.colaborador.correo;

    sheet.getCell('I20').value = new Date().toLocaleDateString('es-CL');

    // Campos técnicos y comentarios
    sheet.getCell('B33').value = data.marca;
    sheet.getCell('B35').value = data.modelo;
    sheet.getCell('B37').value = data.serie;
    sheet.getCell('B39').value = data.hostname;
    sheet.getCell('B41').value = data.imei;
    sheet.getCell('B43').value = data.otrosDetalle;
    sheet.getCell('F34').value = data.estadoEquipo;
    sheet.getCell('F38').value = data.comentarios;
    sheet.getCell('B48').value = data.elementosAdicionales;

    res.setHeader(
        'Content-Type',
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    );
    res.setHeader(
        'Content-Disposition',
        'attachment; filename=acta_devolucion.xlsx'
    );

    await workbook.xlsx.write(res);
    res.end();
});

// Ruta principal: menú
app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// Ruta para mostrar formulario de entrega
app.get('/entrega', (req, res) => {
    res.sendFile(path.join(__dirname, 'public', 'entrega.html'));
});

app.listen(PORT, () => console.log(`Servidor corriendo en http://localhost:${PORT}`));