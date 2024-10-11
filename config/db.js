const mysql = require('mysql2');
require('dotenv').config();

// Crear conexión con MySQL
const connection = mysql.createConnection({
  host: process.env.DB_HOST,
  user: process.env.DB_USER,
  password: process.env.DB_PASSWORD,
  database: process.env.DB_NAME
});

connection.connect((err) => {
  if (err) {
    console.error('Error al conectarse a MySQL:', err);
    process.exit(1);
  } else {
    console.log('Conectado a MySQL');
  }
});

module.exports = connection;
