const express = require('express');
const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');

const app = express();
const PORT = process.env.PORT || 3000;

const DATA_DIR = path.join(__dirname, 'data');
const PARTS_DIR = path.join(DATA_DIR, 'partituras');
const LEADS_FILE = path.join(DATA_DIR, 'leads.xlsx');

if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });
if (!fs.existsSync(PARTS_DIR)) fs.mkdirSync(PARTS_DIR, { recursive: true });

app.use(express.json({ limit: '10mb' }));
app.use(express.static(path.join(__dirname, 'public')));

app.get('/', (req,res) => {
  res.sendFile(path.join(__dirname, 'public', 'landing.html'));
});

app.post('/api/lead', async (req, res) => {
  try {
    const { nome, whatsapp, instrumento, cidade, origem, data } = req.body;
    let wb, ws;
    if (fs.existsSync(LEADS_FILE)) {
      wb = new ExcelJS.Workbook();
      await wb.xlsx.readFile(LEADS_FILE);
      ws = wb.getWorksheet('Leads');
    } else {
      wb = new ExcelJS.Workbook();
      ws = wb.addWorksheet('Leads');
      ws.addRow(['Data','Nome','WhatsApp','Instrumento','Cidade','Origem','Musica','Status']);
      ws.getRow(1).font = { bold: true };
    }
    ws.addRow([data||new Date().toLocaleString('pt-BR'), nome, whatsapp, instrumento, cidade, origem, '', 'ENTROU']);
    await wb.xlsx.writeFile(LEADS_FILE);
    res.json({ success: true });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.post('/api/partituras', async (req, res) => {
  try {
    fs.writeFileSync(path.join(PARTS_DIR, Date.now()+'.json'), JSON.stringify(req.body, null, 2));
    res.json({ success: true });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.listen(PORT, () => console.log('RODANDO na porta '+PORT));
