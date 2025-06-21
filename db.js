require('dotenv').config();
const sql = require('mssql');

const config = {
    user: process.env.DB_USER,
    password: process.env.DB_PASSWORD,
    server: process.env.DB_SERVER,
    database: process.env.DB_DATABASE,
    options: {
        enableArithAbort: true,
        encrypt: true,                 // exige TLS
        trustServerCertificate: true   // confía en certificados autofirmados
    }
};

let pool;
async function getPool() {
    if (!pool) pool = await sql.connect(config);
    return pool;
}

module.exports = { sql, getPool };