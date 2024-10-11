const XLSX = require('xlsx');
const cron = require('node-cron');
const connection = require('../config/db');
const fs = require('fs');
const path = require('path');

// Ruta del archivo Excel original y de respaldo
const originalFilePath = 'C:/Users/sharp/Desktop/muestra.xlsx'; // Cambia esta ruta por la de tu archivo
const backupFilePath = 'C:/Users/sharp/Desktop/archivos_originales/muestra_backup.xlsx'; // Ruta donde se guardará el archivo original

// Función para copiar el archivo original y agregar la columna "Estado" solo en el respaldo
function copyOriginalWithStatus() {
  const workbook = XLSX.readFile(originalFilePath);
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const data = XLSX.utils.sheet_to_json(sheet, { header: 1 });
  
  // Agregar el encabezado de "Estado" en el archivo de respaldo
  if (data.length > 0) {
    data[0].push('Estado'); // Agregar encabezado "Estado"
    for (let i = 1; i < data.length; i++) {
      data[i].push('Archivo procesado'); // Agregar estado "Archivo procesado" a cada fila
    }
  }

  // Convertir los datos actualizados de nuevo a una hoja
  const updatedSheet = XLSX.utils.aoa_to_sheet(data);
  workbook.Sheets[workbook.SheetNames[0]] = updatedSheet;

  // Guardar el archivo de respaldo con la columna de estado
  XLSX.writeFile(workbook, backupFilePath);
  console.log('Archivo original copiado y columna de estado agregada en:', backupFilePath);
}

// Función para leer y eliminar una fila del Excel (sin la columna "Estado")
function readAndDeleteRow() {
  const workbook = XLSX.readFile(originalFilePath);
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const data = XLSX.utils.sheet_to_json(sheet, { header: 1 });

  // Comprobar si hay más de una fila (las cabeceras no cuentan)
  if (data.length > 1) {
    const row = data[1];   // Toma la segunda fila (la primera es cabecera)
    
    // Eliminar la fila leída (sin crear un nuevo libro)
    data.splice(1, 1);

    // Actualizar el archivo Excel sin reescribir el libro completo
    const updatedSheet = XLSX.utils.json_to_sheet(data, { skipHeader: true });
    workbook.Sheets[workbook.SheetNames[0]] = updatedSheet;
    XLSX.writeFile(workbook, originalFilePath);

    console.log('Fila leída y eliminada del archivo Excel:', row);
    return row;
  } else {
    console.log("No hay más filas en el archivo Excel.");
    return null;
  }
}

// Función para insertar la fila en MySQL (sin la columna "Estado")
function insertRowIntoDatabase(row) {
  if (row) {
    const query = 'INSERT INTO articulos (nombre, descripcion, marca, precio) VALUES (?, ?, ?, ?)';
    connection.query(query, [row[0], row[1], row[2], row[3]], (err, results) => {
      if (err) {
        console.error('Error al insertar la fila en la base de datos:', err);
      } else {
        console.log('Fila insertada correctamente en la base de datos');
      }
    });
  }
}

// Función para eliminar el archivo Excel una vez que esté vacío
function deleteExcelFile() {
  fs.unlink(originalFilePath, (err) => {
    if (err) {
      console.error('Error al intentar eliminar el archivo:', err);
    } else {
      console.log('Archivo Excel eliminado exitosamente.');
      console.log('Proceso finalizado.'); // Mensaje de finalización
      process.exit(0);
    }
  });
}

// Ejecutar la función inmediatamente al iniciar
copyOriginalWithStatus(); // Copia el archivo original y agrega la columna de estado solo en el respaldo
const initialRow = readAndDeleteRow();
if (initialRow) {
  insertRowIntoDatabase(initialRow); // Insertar solo las columnas de datos en la base de datos
} else {
  deleteExcelFile();
}

// Cron para ejecutar cada 2 minutos
const job = cron.schedule('*/2 * * * *', () => {
  console.log('Ejecutando la tarea programada...');
  const row = readAndDeleteRow();
  if (row) {
    insertRowIntoDatabase(row); // Insertar solo las columnas de datos en la base de datos
  } else {
    // Si no hay más filas, eliminamos el archivo Excel
    deleteExcelFile();
    job.stop(); // Detenemos el cronjob para evitar ejecuciones futuras
  }
});
