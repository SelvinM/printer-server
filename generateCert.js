const Fastify = require("fastify");
const fs = require("fs");
const cors = require("@fastify/cors");
const { SerialPort } = require("serialport");

const app = Fastify({
  https: {
    key: fs.readFileSync("./certs/key.pem"),
    cert: fs.readFileSync("./certs/cert.pem"),
  },
});

app.register(cors, {
  origin: "https://erp.laptopoutlet.hn", // restrict to your ERP domain
});

const port = new SerialPort({ path: "COM4", baudRate: 9600 });

app.post("/print", async (req, reply) => {
  try {
    const { data } = req.body;
    if (!data) return reply.code(400).send({ error: "No data provided" });

    const buffer = Buffer.from(data, "base64");
    port.write(buffer);

    reply.send({ success: true });
  } catch (err) {
    console.error(err);
    reply.code(500).send({ error: "Print failed" });
  }
});

app.listen({ port: 9100, host: "0.0.0.0" }, (err) => {
  if (err) throw err;
  console.log("HTTPS Printer server running at https://localhost:9100");
});
