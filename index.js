// index.js
const path = require("path");
const fs = require("fs");
const os = require("os");
const crypto = require("crypto");
const { execSync, execFile } = require("child_process");
const Fastify = require("fastify");
const cors = require("@fastify/cors");

// -----------------------
// Certificados (mkcert)
// -----------------------
const CERT_DIR = path.resolve(__dirname, "certs");
const CERT_FILE = path.join(CERT_DIR, "localhost+1.pem");
const KEY_FILE = path.join(CERT_DIR, "localhost+1-key.pem");
const MKCERT_PATH = path.resolve(__dirname, "mkcert.exe"); // ajusta si es necesario

if (!fs.existsSync(CERT_FILE) || !fs.existsSync(KEY_FILE)) {
  console.log("Generando certificados HTTPS con mkcert...");
  if (!fs.existsSync(CERT_DIR)) fs.mkdirSync(CERT_DIR, { recursive: true });
  execSync(
    `${MKCERT_PATH} -cert-file "${CERT_FILE}" -key-file "${KEY_FILE}" localhost 127.0.0.1 ::1`,
    { stdio: "inherit" }
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
app.register(cors, { origin: "*" });

// -----------------------
// Config (persistencia)
// -----------------------
const CONFIG_PATH = path.join(__dirname, "printer-config.json");

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
// Helper PowerShell (Winspool WritePrinter)
// -----------------------
const RAW_PRINT_PS1 = path.join(__dirname, "raw-print.ps1");

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
            di.pDocName = "Recibo";
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

  if (needsWrite) fs.writeFileSync(RAW_PRINT_PS1, RAW_PRINT_PS1_CONTENT, "utf8");
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
          const msg = (stderr && stderr.trim()) || err.message || "PowerShell falló";
          return reject(new Error(msg));
        }
        resolve((stdout || "").trim());
      }
    );
  });
}

// -----------------------
// Impresoras (sin libs de Node nativas)
// -----------------------
async function listPrinters() {
  const cmd =
    "Get-CimInstance Win32_Printer | Select-Object Name, Default | ConvertTo-Json -Compress";
  const out = await runPowerShell(["-Command", cmd]);
  const parsed = out ? JSON.parse(out) : [];
  const list = Array.isArray(parsed) ? parsed : [parsed];
  const printers = list.map((p) => p.Name).filter(Boolean);
  const defObj = list.find((p) => p.Default) || null;
  return { printers, defaultPrinter: defObj ? defObj.Name : null };
}

async function getDefaultPrinterName() {
  const cmd =
    '(Get-CimInstance Win32_Printer | Where-Object { $_.Default -eq $true } | Select-Object -First 1 -ExpandProperty Name)';
  const name = await runPowerShell(["-Command", cmd]);
  if (!name) throw new Error("No hay impresora predeterminada. Configúrala en Windows.");
  return name;
}

async function resolveTargetPrinter() {
  const selected = getSelectedPrinterFromConfig();
  if (selected) return selected;
  return await getDefaultPrinterName();
}

// -----------------------
// Base64 -> bytes (RAW)
// -----------------------
function bufferFromBase64(dataB64) {
  if (typeof dataB64 !== "string" || !dataB64.trim()) {
    throw new Error("Datos inválidos: se esperaba un string en base64.");
  }
  let b64 = dataB64.trim();
  const m = b64.match(/^data:.*;base64,(.*)$/i);
  if (m) b64 = m[1];
  return Buffer.from(b64, "base64");
}

// -----------------------
// UI (español)
// -----------------------
app.get("/ui", async (_req, reply) => {
  const html = `<!doctype html>
<html>
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>Selección de impresora</title>
  <style>
    body { font-family: system-ui, -apple-system, Segoe UI, Roboto, Arial; margin: 24px; }
    .card { max-width: 560px; border: 1px solid #ddd; border-radius: 12px; padding: 16px; }
    select, button { width: 100%; padding: 10px; font-size: 16px; margin-top: 10px; }
    .row { margin-top: 12px; }
    .muted { color: #666; font-size: 14px; }
    .ok { color: #0a7; }
    .err { color: #c00; white-space: pre-wrap; }
  </style>
</head>
<body>
  <div class="card">
    <h2>Elegir impresora</h2>
    <div class="muted">Esto solo afecta la impresión en esta PC.</div>

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
// APIs para UI (español)
// -----------------------
app.get("/api/printers", async (_req, reply) => {
  try {
    const { printers, defaultPrinter } = await listPrinters();
    reply.send({ printers, default: defaultPrinter });
  } catch (e) {
    reply.code(500).send({ error: "No se pudieron listar las impresoras." });
  }
});

app.get("/api/selected-printer", async (_req, reply) => {
  reply.send({ selectedPrinter: getSelectedPrinterFromConfig() });
});

app.post("/api/selected-printer", async (req, reply) => {
  try {
    const { printerName } = req.body || {};

    if (printerName === null) {
      setSelectedPrinterInConfig(null);
      return reply.send({ success: true, selectedPrinter: null });
    }

    const name = String(printerName || "").trim();
    if (!name) return reply.code(400).send({ error: "Debes indicar una impresora (o null para limpiar)." });

    const { printers } = await listPrinters();
    if (!printers.includes(name)) {
      return reply.code(400).send({ error: "Esa impresora no existe en esta PC." });
    }

    setSelectedPrinterInConfig(name);
    reply.send({ success: true, selectedPrinter: name });
  } catch (e) {
    reply.code(500).send({ error: "No se pudo guardar la impresora seleccionada." });
  }
});

// -----------------------
// Imprimir (solo impresora seleccionada o predeterminada)
// -----------------------
app.post("/print", async (req, reply) => {
  let tmpPath = null;
  try {
    const { data } = req.body || {};
    if (!data) return reply.code(400).send({ error: "No se proporcionaron datos para imprimir." });

    const target = await resolveTargetPrinter();
    const buf = bufferFromBase64(data);

    tmpPath = path.join(
      os.tmpdir(),
      `recibo-${crypto.randomUUID?.() || crypto.randomBytes(16).toString("hex")}.bin`
    );
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
      bytes: buf.length,
      jobId: Number.isFinite(jobId) ? jobId : null,
    });
  } catch (err) {
    console.error("Error de impresión:", err);
    reply.code(500).send({ error: "Error al imprimir. Revisa la impresora y el controlador." });
  } finally {
    if (tmpPath) {
      try { fs.unlinkSync(tmpPath); } catch {}
    }
  }
});

// -----------------------
// Escuchar (solo local por defecto)
// -----------------------
app.listen({ port: 9100, host: "127.0.0.1" }, async (err) => {
  if (err) throw err;
  console.log("Servidor de impresión HTTPS: https://localhost:9100");
  console.log("UI para elegir impresora: https://localhost:9100/ui");
  try {
    console.log("Impresora predeterminada (Windows):", await getDefaultPrinterName());
    console.log("Impresora seleccionada:", getSelectedPrinterFromConfig() || "(ninguna)");
  } catch (e) {
    console.log("Impresora predeterminada:", e.message);
  }
});
