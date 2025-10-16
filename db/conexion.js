const sqlite3 = require('sqlite3').verbose();
const path = require('path');

//const sql = require('mssql');
//require('dotenv').config();

const dbPath = path.join(__dirname, '../ZKTimeNet.db');
const db = new sqlite3.Database(dbPath, (err) => {
    if (err) {
        console.error('❌ Error de conexión a SQLite:', err);
    } else {
        console.log('✅ Conectado a SQLite');
    }
});

module.exports = db;
/*
const config = {
    user: process.env.DB_USER,
    password: process.env.DB_PASSWORD,
    server: process.env.DB_SERVER,
    database: process.env.DB_NAME,
    options: {
        encrypt: false,
        trustServerCertificate: true
    }
};

const poolPromise = new sql.ConnectionPool(config)
    .connect()
    .then(pool => {
        console.log('✅ Conectado a SQL Server');
        return pool;
    })
    .catch(err => console.log('❌ Error de conexión: ', err));

module.exports = {
    sql, poolPromise
};*/
