import express from "express";
import ExcelJS from "exceljs";
import fetch from "node-fetch";

const app = express();
app.use(express.json());

const FILE_ID = process.env.EXCEL_FILE_ID;
const TOKEN = process.env.MS_TOKEN;

const GRAPH_URL =
  `https://graph.microsoft.com/v1.0/me/drive/items/${FILE_ID}/content`;

async function baixarPlanilha() {
  const res = await fetch(GRAPH_URL, {
    headers: {
      Authorization: `Bearer ${TOKEN}`
    }
  });

  if (!res.ok) {
    throw new Error("Erro ao baixar arquivo do OneDrive");
  }

  return Buffer.from(await res.arrayBuffer());
}

async function salvarPlanilha(buffer) {
  const res = await fetch(GRAPH_URL, {
    method: "PUT",
    headers: {
      Authorization: `Bearer ${TOKEN}`
    },
    body: buffer
  });

  if (!res.ok) {
    throw new Error("Erro ao salvar arquivo no OneDrive");
  }
}

app.post("/despesas/adicionar", async (req, res) => {
  try {
    const { mesColuna, tipo, valor } = req.body;

    const buffer = await baixarPlanilha();

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
      return res.status(400).json({ error: "Tipo não encontrado" });
    }

    const cell = sheet.getCell(`${mesColuna}${linha}`);
    cell.value = (cell.value || 0) + Number(valor);

    const out = await wb.xlsx.writeBuffer();
    await salvarPlanilha(out);

    res.json({ ok: true });

  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () =>
  console.log("✅ Backend B&F rodando")
);
