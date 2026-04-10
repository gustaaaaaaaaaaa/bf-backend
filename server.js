import express from "express";
import ExcelJS from "exceljs";
import fetch from "node-fetch";

const app = express();
app.use(express.json());

const FILE_PATH = "Orçamento - Finanças Pessoais - Ax - 2026.xlsm";

const GRAPH_URL =
  `https://graph.microsoft.com/v1.0/me/drive/root:/${FILE_PATH}:/content`;

async function getFile(token) {
  const res = await fetch(GRAPH_URL, {
    headers: {
      Authorization: `Bearer ${token}`
    }
  });
  return Buffer.from(await res.arrayBuffer());
}

async function saveFile(token, buffer) {
  await fetch(GRAPH_URL, {
    method: "PUT",
    headers: {
      Authorization: `Bearer ${token}`
    },
    body: buffer
  });
}

app.post("/despesas/adicionar", async (req, res) => {
  const { mesColuna, tipo, valor } = req.body;
  const token = process.env.MS_TOKEN;

  const buffer = await getFile(token);
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buffer);

  const sheet = wb.getWorksheet("Abril");

  let linha;
  for (let i = 131; i <= 144; i++) {
    if (sheet.getCell(`G${i}`).value === tipo) {
      linha = i;
      break;
    }
  }

  if (!linha) {
    return res.status(400).send("Tipo não encontrado");
  }

  const cell = sheet.getCell(`${mesColuna}${linha}`);
  cell.value = (cell.value || 0) + valor;

  const out = await wb.xlsx.writeBuffer();
  await saveFile(token, out);

  res.json({ ok: true });
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () =>
  console.log("Backend B&F rodando ✅")
);