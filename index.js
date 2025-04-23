// Importa las dependencias necesarias
const express = require('express');
const multer = require('multer');
const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');
const QRCode = require('qrcode');
const ExcelJS = require('exceljs');
const cors = require('cors');

// Crea una instancia de express
const app = express();

// Define el puerto en el que el servidor escuchará
const PORT = process.env.PORT || 3000;

// Middleware para parsear los datos JSON en las peticiones
app.use(express.json());

app.use(cors());

// Configura Multer para la carga de archivos
const upload = multer({ dest: 'uploads/' });

// Ruta para la carga de archivos y generación de QR
app.post('/api/generarQR', upload.single('archivo'), async (req, res) => {
  try {
    const filePath = req.file.path; // Ruta del archivo subido
    const ext = path.extname(req.file.originalname); // Extensión del archivo

    // Cargar el archivo Excel o CSV
    let workbook;
    if (ext === '.csv') {
      const data = fs.readFileSync(filePath, 'utf8');
      const sheet = XLSX.utils.csv_to_sheet(data);
      workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, sheet, 'Sheet1');
    } else {
      workbook = XLSX.readFile(filePath); // Si es Excel (.xlsx, .xls)
    }

    const sheet = workbook.Sheets[workbook.SheetNames[0]]; // Obtén la primera hoja
    const json = XLSX.utils.sheet_to_json(sheet); // Convierte la hoja a JSON

    // Crea un nuevo workbook para almacenar los QR generados
    const finalWorkbook = new ExcelJS.Workbook();
    const ws = finalWorkbook.addWorksheet('Personal');

    // Configura las columnas del Excel
    ws.columns = Object.keys(json[0]).map(key => ({ header: key, key }));
    ws.columns.push({ header: 'QR', key: 'qr' });

    // Genera los códigos QR para cada persona en el archivo
    for (let i = 0; i < json.length; i++) {
      const persona = json[i];
      const qrData = `ID: ${persona.ID}\nNombre: ${persona.Nombre}\nCargo: ${persona.Cargo}` +
        (persona.Documento ? `\nDocumento: ${persona.Documento}` : '');

      const qrPath = `uploads/qr_${persona.ID}.png`; // Ruta para guardar el QR
      await QRCode.toFile(qrPath, qrData); // Genera el QR y guárdalo en el disco

      // Agrega la persona a la hoja del nuevo Excel
      const row = ws.addRow(persona);

      // Agrega la imagen del QR a la celda correspondiente
      const imageId = finalWorkbook.addImage({
        filename: qrPath,
        extension: 'png',
      });
      ws.addImage(imageId, {
        tl: { col: ws.columns.length - 1, row: i + 1 },
        ext: { width: 100, height: 100 },
      });
    }

    // Guarda el archivo Excel final con los QR
    const outputPath = `uploads/output_${Date.now()}.xlsx`;
    await finalWorkbook.xlsx.writeFile(outputPath);

    // Devuelve el archivo generado al usuario
    res.download(outputPath);
  } catch (err) {
    console.error(err);
    res.status(500).send('Error generando QR.');
  }
});

// Ruta raíz de la API (opcional)
app.get('/', (req, res) => {
  res.send('¡Bienvenido a la API de generación de QRs!');
});

// Inicia el servidor
app.listen(PORT, () => {
  console.log(`Servidor corriendo en http://localhost:${PORT}`);
});
