// label-print.js
const path = require("path");
const fs = require("fs");
const os = require("os");
const crypto = require("crypto");
const { execSync, execFile } = require("child_process");
const Fastify = require("fastify");
const { exec } = require("child_process");

// -----------------------
// Server config
// -----------------------
const PORT = 58124;
const HOST = "127.0.0.1";

const BASE_DIR = process.pkg ? path.dirname(process.execPath) : __dirname;

// -----------------------
// Certificados (mkcert)
// -----------------------
const CERT_DIR = path.join(BASE_DIR, "certs");
const CERT_FILE = path.join(CERT_DIR, "localhost.pem");
const KEY_FILE = path.join(CERT_DIR, "localhost-key.pem");
const MKCERT_PATH = path.join(BASE_DIR, "mkcert.exe");

if (!fs.existsSync(CERT_FILE) || !fs.existsSync(KEY_FILE)) {
  console.log("Instalando CA local de mkcert...");
  execSync(`"${MKCERT_PATH}" -install`, { stdio: "ignore" });

  console.log("Generando certificados HTTPS con mkcert...");

  if (!fs.existsSync(CERT_DIR)) fs.mkdirSync(CERT_DIR, { recursive: true });

  execSync(
    `"${MKCERT_PATH}" -cert-file "${CERT_FILE}" -key-file "${KEY_FILE}" localhost 127.0.0.1 ::1`,
    { stdio: "inherit" },
  );
}

const key = fs.readFileSync(KEY_FILE);
const cert = fs.readFileSync(CERT_FILE);

// -----------------------
// Fastify
// -----------------------
const app = Fastify({
  https: { key, cert },
  bodyLimit: 2 * 1024 * 1024,
});

// -----------------------
// CORS manual
// -----------------------
app.addHook("onRequest", (req, reply, done) => {
  reply.header("Access-Control-Allow-Origin", "*");
  reply.header("Access-Control-Allow-Methods", "GET,POST,OPTIONS");
  reply.header("Access-Control-Allow-Headers", "Content-Type");

  if (req.method === "OPTIONS") {
    reply.code(204).send();
    return;
  }

  done();
});

// -----------------------
// Config impresora etiquetas
// -----------------------
const CONFIG_PATH = path.join(BASE_DIR, "label-printer-config.json");

function loadConfig() {
  try {
    const raw = fs.readFileSync(CONFIG_PATH, "utf8");
    const cfg = JSON.parse(raw);
    return typeof cfg === "object" && cfg ? cfg : {};
  } catch {
    return {};
  }
}

function saveConfig(cfg) {
  fs.writeFileSync(CONFIG_PATH, JSON.stringify(cfg, null, 2), "utf8");
}

function getSelectedPrinterFromConfig() {
  const cfg = loadConfig();

  return typeof cfg.selectedPrinter === "string" && cfg.selectedPrinter.trim()
    ? cfg.selectedPrinter.trim()
    : null;
}

function setSelectedPrinterInConfig(printerNameOrNull) {
  const cfg = loadConfig();

  if (printerNameOrNull) cfg.selectedPrinter = printerNameOrNull;
  else delete cfg.selectedPrinter;

  saveConfig(cfg);
}

// -----------------------
// PowerShell RAW print helper
// -----------------------
const RAW_PRINT_PS1 = path.join(BASE_DIR, "raw-print-labels.ps1");

const RAW_PRINT_PS1_CONTENT = String.raw`
param(
  [Parameter(Mandatory=$true)][string]$PrinterName,
  [Parameter(Mandatory=$true)][string]$Path
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
using System.ComponentModel;

public static class RawPrinterHelper
{
    [StructLayout(LayoutKind.Sequential, CharSet=CharSet.Ansi)]
    public class DOCINFOA
    {
        [MarshalAs(UnmanagedType.LPStr)]
        public string pDocName;
        [MarshalAs(UnmanagedType.LPStr)]
        public string pOutputFile;
        [MarshalAs(UnmanagedType.LPStr)]
        public string pDataType;
    }

    [DllImport("winspool.Drv", EntryPoint="OpenPrinterA", SetLastError=true, CharSet=CharSet.Ansi)]
    public static extern bool OpenPrinter(string pPrinterName, out IntPtr phPrinter, IntPtr pDefault);

    [DllImport("winspool.Drv", SetLastError=true)]
    public static extern bool ClosePrinter(IntPtr hPrinter);

    [DllImport("winspool.Drv", EntryPoint="StartDocPrinterA", SetLastError=true, CharSet=CharSet.Ansi)]
    public static extern int StartDocPrinter(IntPtr hPrinter, int level, [In] DOCINFOA di);

    [DllImport("winspool.Drv", SetLastError=true)]
    public static extern bool EndDocPrinter(IntPtr hPrinter);

    [DllImport("winspool.Drv", SetLastError=true)]
    public static extern bool StartPagePrinter(IntPtr hPrinter);

    [DllImport("winspool.Drv", SetLastError=true)]
    public static extern bool EndPagePrinter(IntPtr hPrinter);

    [DllImport("winspool.Drv", SetLastError=true)]
    public static extern bool WritePrinter(IntPtr hPrinter, IntPtr pBytes, int dwCount, out int dwWritten);

    public static int SendBytes(string printerName, byte[] bytes)
    {
        IntPtr hPrinter;
        if (!OpenPrinter(printerName, out hPrinter, IntPtr.Zero))
            throw new Win32Exception(Marshal.GetLastWin32Error(), "OpenPrinter failed");

        try
        {
            var di = new DOCINFOA();
            di.pDocName = "Etiquetas TSPL";
            di.pDataType = "RAW";

            int jobId = StartDocPrinter(hPrinter, 1, di);
            if (jobId == 0)
                throw new Win32Exception(Marshal.GetLastWin32Error(), "StartDocPrinter failed");

            try
            {
                if (!StartPagePrinter(hPrinter))
                    throw new Win32Exception(Marshal.GetLastWin32Error(), "StartPagePrinter failed");

                try
                {
                    int written = 0;
                    IntPtr unmanaged = Marshal.AllocCoTaskMem(bytes.Length);
                    try
                    {
                        Marshal.Copy(bytes, 0, unmanaged, bytes.Length);

                        if (!WritePrinter(hPrinter, unmanaged, bytes.Length, out written))
                            throw new Win32Exception(Marshal.GetLastWin32Error(), "WritePrinter failed");

                        if (written != bytes.Length)
                            throw new Exception("Partial write: " + written + " of " + bytes.Length);
                    }
                    finally
                    {
                        Marshal.FreeCoTaskMem(unmanaged);
                    }
                }
                finally
                {
                    EndPagePrinter(hPrinter);
                }
            }
            finally
            {
                EndDocPrinter(hPrinter);
            }

            return jobId;
        }
        finally
        {
            ClosePrinter(hPrinter);
        }
    }
}
"@ -Language CSharp

$bytes = [System.IO.File]::ReadAllBytes($Path)
$jobId = [RawPrinterHelper]::SendBytes($PrinterName, $bytes)
Write-Output $jobId
`;

function ensureRawPrintScript() {
  const needsWrite =
    !fs.existsSync(RAW_PRINT_PS1) ||
    fs.readFileSync(RAW_PRINT_PS1, "utf8") !== RAW_PRINT_PS1_CONTENT;

  if (needsWrite) {
    fs.writeFileSync(RAW_PRINT_PS1, RAW_PRINT_PS1_CONTENT, "utf8");
  }
}

ensureRawPrintScript();

function runPowerShell(args) {
  return new Promise((resolve, reject) => {
    execFile(
      "powershell.exe",
      ["-NoProfile", "-ExecutionPolicy", "Bypass", ...args],
      { windowsHide: true, maxBuffer: 10 * 1024 * 1024 },
      (err, stdout, stderr) => {
        if (err) {
          const msg =
            (stderr && stderr.trim()) || err.message || "PowerShell falló";
          return reject(new Error(msg));
        }

        resolve((stdout || "").trim());
      },
    );
  });
}

// -----------------------
// Impresoras
// -----------------------
async function listPrinters() {
  const cmd =
    "Get-CimInstance Win32_Printer | Select-Object Name, Default | ConvertTo-Json -Compress";

  const out = await runPowerShell(["-Command", cmd]);
  const parsed = out ? JSON.parse(out) : [];
  const list = Array.isArray(parsed) ? parsed : [parsed];

  const printers = list.map((p) => p.Name).filter(Boolean);
  const defObj = list.find((p) => p.Default) || null;

  return {
    printers,
    defaultPrinter: defObj ? defObj.Name : null,
  };
}

async function getDefaultPrinterName() {
  const cmd =
    "(Get-CimInstance Win32_Printer | Where-Object { $_.Default -eq $true } | Select-Object -First 1 -ExpandProperty Name)";

  const name = await runPowerShell(["-Command", cmd]);

  if (!name) {
    throw new Error("No hay impresora predeterminada. Configúrala en Windows.");
  }

  return name;
}

async function resolveTargetPrinter() {
  const selected = getSelectedPrinterFromConfig();

  if (selected) return selected;

  return await getDefaultPrinterName();
}

async function resolveTargetPrinterWithOptionalOverride(printerName) {
  const name = String(printerName || "").trim();

  if (!name) return await resolveTargetPrinter();

  const { printers } = await listPrinters();

  if (!printers.includes(name)) {
    throw new Error(`La impresora "${name}" no existe en esta PC.`);
  }

  return name;
}

// -----------------------
// TSPL Label Builder
// -----------------------
function cleanTsplValue(value) {
  return String(value ?? "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/ñ/gi, "n")
    .replace(/["\r\n]/g, "-") // allow | for barcode payloads
    .replace(/[^\x20-\x7E]/g, "")
    .trim();
}

function stripKnownPrefix(value, prefixes = []) {
  const v = cleanTsplValue(value);

  for (const p of prefixes) {
    const re = new RegExp(`^${p}-`, "i");
    if (re.test(v)) return v.replace(re, "");
  }

  return v;
}

function bigintToBase36(value) {
  if (value === 0n) return "0";

  const alphabet = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ";
  let n = value;
  let out = "";

  while (n > 0n) {
    const digit = Number(n % 36n);
    out = alphabet[digit] + out;
    n = n / 36n;
  }

  return out;
}

function compactId(value, prefixes = []) {
  const id = stripKnownPrefix(value, prefixes);

  if (/^\d+$/.test(id)) {
    return bigintToBase36(BigInt(id));
  }

  return id.toUpperCase();
}

function withPrefix(value, prefix) {
  const v = cleanTsplValue(value);

  if (!v) return "";
  if (new RegExp(`^${prefix}-`, "i").test(v)) return v;

  return `${prefix}-${v}`;
}

function truncateText(value, max) {
  const v = cleanTsplValue(value);
  return v.length <= max ? v : v.slice(0, max - 1) + "*";
}

function copiesOf(label) {
  return Math.max(1, Math.min(999, Number(label.copies || 1) | 0));
}

function getLabelSize(label) {
  if (label.kind === "sku") return "LARGE_2X1";
  return "SMALL_1X0_5";
}

function getBarcodePayload(label) {
  if (label.kind === "component") {
    const componentId = compactId(label.componentCode, ["C"]);
    const serialId = label.serialNumber
      ? compactId(label.serialNumber, ["U"])
      : "";

    return serialId ? `C|${componentId}|${serialId}` : `C|${componentId}`;
  }

  if (label.kind === "sku") {
    const skuId = compactId(label.sku, ["SKU"]);
    const serialId = label.serialNumber
      ? compactId(label.serialNumber, ["U"])
      : "";

    return serialId ? `S|${skuId}|${serialId}` : `S|${skuId}`;
  }

  if (label.kind === "warranty") {
    const warrantyId = compactId(label.warrantyNumber, ["W"]);
    const employeeId = compactId(label.employeeCode, ["E"]);

    return `W|${warrantyId}|${employeeId}`;
  }

  throw new Error("Tipo de etiqueta no soportado.");
}

function validateLabel(label) {
  if (!label || typeof label !== "object") {
    throw new Error("Etiqueta inválida.");
  }

  if (label.kind === "component") {
    if (!cleanTsplValue(label.componentCode)) {
      throw new Error("La etiqueta de componente requiere componentCode.");
    }

    return;
  }

  if (label.kind === "sku") {
    if (!cleanTsplValue(label.sku)) {
      throw new Error("La etiqueta SKU requiere sku.");
    }

    return;
  }

  if (label.kind === "warranty") {
    if (!cleanTsplValue(label.warrantyNumber)) {
      throw new Error("La etiqueta de garantía requiere warrantyNumber.");
    }

    if (!cleanTsplValue(label.employeeCode)) {
      throw new Error("La etiqueta de garantía requiere employeeCode.");
    }

    return;
  }

  throw new Error(`Tipo de etiqueta no soportado: ${label.kind}`);
}

class TscLabelBuilder {
  constructor() {
    this.lines = [];
  }

  cmd(line) {
    this.lines.push(line + "\r\n");
  }

  quote(value) {
    return `"${cleanTsplValue(value)}"`;
  }

  header(size) {
    if (size === "SMALL_1X0_5") {
      this.cmd("SIZE 25.4 mm,12.7 mm");
    } else if (size === "LARGE_2X1") {
      this.cmd("SIZE 50.8 mm,25.4 mm");
    } else {
      throw new Error("Tamaño de etiqueta inválido.");
    }

    this.cmd("GAP 2 mm,0");
    this.cmd("DIRECTION 1");
    this.cmd("REFERENCE 24,18");
    this.cmd("SET RIBBON OFF");
    this.cmd("SPEED 3");
    this.cmd("DENSITY 8");
    this.cmd("CLS");
  }

  text(x, y, font, value, xMul = 1, yMul = 1) {
    this.cmd(
      `TEXT ${x},${y},${this.quote(font)},0,${xMul},${yMul},${this.quote(
        value,
      )}`,
    );
  }

  code128(x, y, height, payload, options = {}) {
    const narrow = options.narrow ?? 1;
    const wide = options.wide ?? 2;
    const readable = options.readable ?? 0;

    this.cmd(
      `BARCODE ${x},${y},"128",${height},${readable},0,${narrow},${wide},${this.quote(
        payload,
      )}`,
    );
  }

  addLabel(label) {
    validateLabel(label);

    const size = getLabelSize(label);
    const payload = getBarcodePayload(label);
    const copies = copiesOf(label);

    this.header(size);

    if (label.kind === "component") {
      const componentText = truncateText(
        withPrefix(label.componentCode, "C"),
        18,
      );

      const serialText = label.serialNumber
        ? truncateText(withPrefix(label.serialNumber, "U"), 18)
        : null;

      this.text(4, 2, "1", componentText);

      if (serialText) {
        this.text(4, 16, "1", serialText);
        this.code128(4, 36, 48, payload, { narrow: 1, wide: 2 });
      } else {
        this.code128(4, 28, 56, payload, { narrow: 1, wide: 2 });
      }
    }

    if (label.kind === "sku") {
      const skuText = truncateText(withPrefix(label.sku, "SKU"), 28);

      const serialText = label.serialNumber
        ? truncateText(withPrefix(label.serialNumber, "U"), 28)
        : null;

      this.text(12, 8, "3", skuText);

      if (serialText) {
        this.text(12, 42, "2", serialText);
        this.code128(12, 76, 88, payload, { narrow: 2, wide: 4 });
      } else {
        this.code128(12, 62, 104, payload, { narrow: 3, wide: 6 });
      }
    }

    if (label.kind === "warranty") {
      const warrantyText = truncateText(
        withPrefix(label.warrantyNumber, "W"),
        16,
      );

      const employeeText = truncateText(withPrefix(label.employeeCode, "E"), 8);

      this.text(4, 2, "1", warrantyText);
      this.text(4, 16, "1", employeeText);
      this.code128(4, 36, 48, payload, { narrow: 1, wide: 2 });
    }

    this.cmd(`PRINT 1,${copies}`);

    return this;
  }

  addBatch(labels) {
    if (!Array.isArray(labels) || labels.length === 0) {
      throw new Error("No se enviaron etiquetas.");
    }

    const sizes = new Set(labels.map(getLabelSize));

    if (sizes.size > 1) {
      throw new Error(
        "No mezcles etiquetas 1x0.5 y 2x1 en el mismo lote. La impresora solo puede tener un rollo físico cargado a la vez.",
      );
    }

    for (const label of labels) {
      this.addLabel(label);
    }

    return this;
  }

  buildRaw() {
    return this.lines.join("");
  }

  buildBuffer() {
    return Buffer.from(this.buildRaw(), "ascii");
  }
}

// -----------------------
// Helpers
// -----------------------
function makeTempFilePath(prefix, ext) {
  const id =
    typeof crypto.randomUUID === "function"
      ? crypto.randomUUID()
      : crypto.randomBytes(16).toString("hex");

  return path.join(os.tmpdir(), `${prefix}-${id}.${ext}`);
}

// -----------------------
// UI
// -----------------------
app.get("/ui", async (_req, reply) => {
  const html = `<!doctype html>
<html>
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>Selección de impresora de etiquetas</title>
  <style>
    body { font-family: system-ui, -apple-system, Segoe UI, Roboto, Arial; margin: 24px; }
    .card { max-width: 560px; border: 1px solid #ddd; border-radius: 12px; padding: 16px; }
    select, button { width: 100%; padding: 10px; font-size: 16px; margin-top: 10px; }
    .row { margin-top: 12px; }
    .muted { color: #666; font-size: 14px; }
    .ok { color: #0a7; }
    .err { color: #c00; white-space: pre-wrap; }
    code { background: #f4f4f4; padding: 2px 5px; border-radius: 5px; }
  </style>
</head>
<body>
  <div class="card">
    <h2>Elegir impresora de etiquetas</h2>
    <div class="muted">
      Este servidor solo imprime etiquetas. Endpoint:
      <code>https://localhost:${PORT}/print-labels</code>
    </div>

    <div class="row">
      <div class="muted" id="status">Cargando…</div>
      <select id="printerSelect"></select>
      <button id="saveBtn">Guardar selección</button>
      <button id="clearBtn">Usar la predeterminada de Windows</button>
    </div>

    <div class="row muted">
      Seleccionada: <span id="selectedLabel">(ninguna)</span><br/>
      Predeterminada (Windows): <span id="defaultLabel">(desconocida)</span>
    </div>

    <div class="row" id="msg"></div>
  </div>

<script>
  const $ = (id) => document.getElementById(id);
  const msg = (text, cls) => { $('msg').className = cls; $('msg').textContent = text; };

  async function refresh() {
    msg('', '');
    $('status').textContent = 'Cargando impresoras…';

    const pRes = await fetch('/api/printers');
    const p = await pRes.json();
    if (!pRes.ok) throw new Error(p.error || 'No se pudieron obtener las impresoras.');

    const sRes = await fetch('/api/selected-printer');
    const s = await sRes.json();
    if (!sRes.ok) throw new Error(s.error || 'No se pudo leer la configuración.');

    const sel = $('printerSelect');
    sel.innerHTML = '';

    (p.printers || []).forEach(name => {
      const opt = document.createElement('option');
      opt.value = name;
      opt.textContent = name;
      sel.appendChild(opt);
    });

    $('defaultLabel').textContent = p.default || '(ninguna)';
    $('selectedLabel').textContent = s.selectedPrinter || '(ninguna)';

    if (s.selectedPrinter) sel.value = s.selectedPrinter;
    else if (p.default) sel.value = p.default;

    $('status').textContent = 'Listo.';
  }

  $('saveBtn').addEventListener('click', async () => {
    msg('', '');

    const printerName = $('printerSelect').value;

    const r = await fetch('/api/selected-printer', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ printerName })
    });

    const j = await r.json();

    if (!r.ok) return msg(j.error || 'No se pudo guardar.', 'err');

    msg('Guardado.', 'ok');
    await refresh();
  });

  $('clearBtn').addEventListener('click', async () => {
    msg('', '');

    const r = await fetch('/api/selected-printer', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ printerName: null })
    });

    const j = await r.json();

    if (!r.ok) return msg(j.error || 'No se pudo actualizar.', 'err');

    msg('Ahora se usa la predeterminada de Windows.', 'ok');
    await refresh();
  });

  refresh().catch(e => msg(String(e), 'err'));
</script>
</body>
</html>`;

  reply.type("text/html; charset=utf-8").send(html);
});

// -----------------------
// APIs UI
// -----------------------
app.get("/api/printers", async (_req, reply) => {
  try {
    const { printers, defaultPrinter } = await listPrinters();

    reply.send({
      printers,
      default: defaultPrinter,
    });
  } catch (e) {
    reply.code(500).send({
      error: "No se pudieron listar las impresoras.",
    });
  }
});

app.get("/api/selected-printer", async (_req, reply) => {
  reply.send({
    selectedPrinter: getSelectedPrinterFromConfig(),
  });
});

app.post("/api/selected-printer", async (req, reply) => {
  try {
    const { printerName } = req.body || {};

    if (printerName === null) {
      setSelectedPrinterInConfig(null);

      return reply.send({
        success: true,
        selectedPrinter: null,
      });
    }

    const name = String(printerName || "").trim();

    if (!name) {
      return reply.code(400).send({
        error: "Debes indicar una impresora o null para limpiar.",
      });
    }

    const { printers } = await listPrinters();

    if (!printers.includes(name)) {
      return reply.code(400).send({
        error: "Esa impresora no existe en esta PC.",
      });
    }

    setSelectedPrinterInConfig(name);

    reply.send({
      success: true,
      selectedPrinter: name,
    });
  } catch (e) {
    reply.code(500).send({
      error: "No se pudo guardar la impresora seleccionada.",
    });
  }
});

// -----------------------
// Print labels
// -----------------------
app.post("/print-labels", async (req, reply) => {
  let tmpPath = null;

  try {
    const { labels, printerName } = req.body || {};

    const buf = new TscLabelBuilder().addBatch(labels).buildBuffer();
    const target = await resolveTargetPrinterWithOptionalOverride(printerName);

    tmpPath = makeTempFilePath("labels", "tspl");
    fs.writeFileSync(tmpPath, buf);

    const jobIdStr = await runPowerShell([
      "-File",
      RAW_PRINT_PS1,
      "-PrinterName",
      target,
      "-Path",
      tmpPath,
    ]);

    const jobId = Number(jobIdStr);

    reply.send({
      success: true,
      printer: target,
      labels: labels.length,
      bytes: buf.length,
      jobId: Number.isFinite(jobId) ? jobId : null,
    });
  } catch (err) {
    console.error("Error de impresión de etiquetas:", err);

    reply.code(500).send({
      error:
        err instanceof Error
          ? err.message
          : "Error al imprimir etiquetas. Revisa la impresora y el controlador.",
    });
  } finally {
    if (tmpPath) {
      try {
        fs.unlinkSync(tmpPath);
      } catch {}
    }
  }
});

// -----------------------
// Start
// -----------------------
async function start() {
  try {
    await app.listen({ port: PORT, host: HOST });

    console.log(`Servidor de etiquetas HTTPS: https://localhost:${PORT}`);
    console.log(`UI para elegir impresora: https://localhost:${PORT}/ui`);
    console.log(
      `Endpoint etiquetas: POST https://localhost:${PORT}/print-labels`,
    );

    try {
      console.log(
        "Impresora predeterminada (Windows):",
        await getDefaultPrinterName(),
      );

      console.log(
        "Impresora de etiquetas seleccionada:",
        getSelectedPrinterFromConfig() || "(ninguna)",
      );
    } catch (e) {
      console.log("Impresora predeterminada:", e.message);
    }
  } catch (err) {
    console.error("Error iniciando servidor de etiquetas:", err);
    process.exit(1);
  }
}

console.log("");
console.log("======================================");
console.log("  SERVIDOR DE ETIQUETAS - LAPTOP OUTLET");
console.log("======================================");
console.log("");
console.log("⚠️  NO CERRAR ESTA VENTANA");
console.log("");
console.log("Si se cierra, las etiquetas no imprimirán.");
console.log("");
console.log(`Servidor: https://localhost:${PORT}`);
console.log("");

setTimeout(() => {
  exec(`start "" "https://localhost:${PORT}/ui"`);
}, 1000);

start();
