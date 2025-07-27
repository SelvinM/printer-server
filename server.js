const path = require("path");
const fs = require("fs");
const { execSync } = require("child_process");
const Fastify = require("fastify");
const { SerialPort } = require("serialport");
const cors = require("@fastify/cors");

const CERT_DIR = path.resolve(__dirname, "certs");
const CERT_FILE = path.join(CERT_DIR, "localhost+1.pem");
const KEY_FILE = path.join(CERT_DIR, "localhost+1-key.pem");
const MKCERT_PATH = path.resolve(__dirname, "mkcert.exe"); // adjust if needed

// Generate certs if missing
if (!fs.existsSync(CERT_FILE) || !fs.existsSync(KEY_FILE)) {
  console.log("Generating HTTPS certificates with mkcert...");
  if (!fs.existsSync(CERT_DIR)) fs.mkdirSync(CERT_DIR);
  execSync(
    `${MKCERT_PATH} -cert-file "${CERT_FILE}" -key-file "${KEY_FILE}" localhost 127.0.0.1 ::1`,
    { stdio: "inherit" }
  );
}

// Load certs
const key = fs.readFileSync(KEY_FILE);
const cert = fs.readFileSync(CERT_FILE);

// Fastify HTTPS
const app = Fastify({ https: { key, cert } });
app.register(cors, { origin: "*" });

const PRINTER_PORT = "COM4";
const PRINTER_BAUD = 9600;
const port = new SerialPort({ path: PRINTER_PORT, baudRate: PRINTER_BAUD });

function sanitizeForAscii(str) {
  return str
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "") // remove combining diacritics
    .replace(/[áàâäãå]/gi, "a")
    .replace(/[éèêë]/gi, "e")
    .replace(/[íìîï]/gi, "i")
    .replace(/[óòôöõ]/gi, "o")
    .replace(/[úùûü]/gi, "u")
    .replace(/ñ/gi, "n")
    .replace(/\r\n?/g, "\n");
}

app.post("/print", async (req, reply) => {
  try {
    const { data } = req.body;
    if (!data) return reply.code(400).send({ error: "No data provided" });

    let rawStr = Buffer.from(data, "base64").toString("utf8");
    rawStr = sanitizeForAscii(rawStr); // <-- ensure safe ASCII
    const raw = Buffer.from(rawStr, "ascii");

    port.write(raw);
    reply.send({ success: true });
  } catch (err) {
    console.error("Print error:", err);
    reply.code(500).send({ error: "Print failed" });
  }
});

app.listen({ port: 9100, host: "0.0.0.0" }, (err) => {
  if (err) throw err;
  console.log("HTTPS Printer server running at https://localhost:9100");
});
