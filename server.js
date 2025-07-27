const express = require('express');
const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');
const { Document, Packer, Paragraph, TextRun } = require('docx');

const app = express();
const PORT = 3000;
const DATA_FILE = path.join(__dirname, 'data.json');

app.use(express.static(__dirname));
app.use(express.json());

app.post('/pesan', (req, res) => {
  const dataBaru = req.body;
  let dataLama = [];
  try {
    dataLama = JSON.parse(fs.readFileSync(DATA_FILE, 'utf8'));
  } catch {}
  dataLama.push(dataBaru);
  fs.writeFileSync(DATA_FILE, JSON.stringify(dataLama, null, 2));
  res.json({ success: true });
});

app.get('/data', (req, res) => {
  try {
    const data = JSON.parse(fs.readFileSync(DATA_FILE, 'utf8'));
    res.json(data);
  } catch {
    res.json([]);
  }
});

app.get('/export/excel', async (req, res) => {
  const data = JSON.parse(fs.readFileSync(DATA_FILE, 'utf8') || '[]');
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Pesanan');
  ws.addRow(['Nama', 'Pesanan', 'Catatan', 'Waktu']);
  data.forEach(p => ws.addRow([p.nama, p.pesanan, p.catatan, p.waktu]));
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', 'attachment; filename="pesanan.xlsx"');
  await wb.xlsx.write(res);
  res.end();
});

app.get('/export/word', async (req, res) => {
  const data = JSON.parse(fs.readFileSync(DATA_FILE, 'utf8') || '[]');
  const doc = new Document({
    sections: [{
      children: [
        new Paragraph({ children: [new TextRun({ text: 'Daftar Pesanan', bold: true, size: 32 })] }),
        ...data.map(p =>
          new Paragraph({
            children: [
              new TextRun({ text: `${p.nama} - ${p.pesanan}`, bold: true }),
              new TextRun({ text: `\nCatatan: ${p.catatan || '-'}` }),
              new TextRun({ text: `\nWaktu: ${p.waktu}\n` }),
            ]
          })
        )
      ]
    }]
  });
  const buffer = await Packer.toBuffer(doc);
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
  res.setHeader('Content-Disposition', 'attachment; filename="pesanan.docx"');
  res.send(buffer);
});

app.listen(PORT, () => {
  console.log(`Server berjalan di http://localhost:${PORT}`);
});