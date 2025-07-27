const Fastify = require("fastify");
const cors = require("@fastify/cors");
const { SerialPort } = require("serialport");

const app = Fastify();

app.register(cors, {
  origin: "*", // allow all for now (can restrict to your ERP domain later)
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
  console.log("Printer server running at http://localhost:9100");
});
