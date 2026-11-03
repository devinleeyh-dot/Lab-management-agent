// WhatsApp bridge for DEVELOPMENT (unofficial, QR login).
//  - Forwards messages from approved senders & owner to agent.py
//  - Lets agent.py send WhatsApp messages (PO requests to the lab owner, updates to requesters)
//  - If the lab owner IS the linked WhatsApp account, any message typed in the
//    "Message yourself" chat is sent to agent.py as an owner command.
const { Client, LocalAuth } = require("whatsapp-web.js");
const qrcode = require("qrcode-terminal");
const http = require("http");

const AGENT_URL = "http://127.0.0.1:5000/incoming";
const USERS_URL = "http://127.0.0.1:5000/users";
const UNKNOWN_URL = "http://127.0.0.1:5000/unknown";
const SEND_PORT = 5001; // agent.py calls this to send messages

// The lab owner is the same person as the linked WhatsApp account (+6589033626).
const SELF_OWNER = { phone: "6589033626", name: "YH Lee (SG)", role: "owner" };

// Your own account appears under two ids in the "Message yourself" chat:
// the phone number and a long WhatsApp id (seen in the log as "to = ...@lid").
const MY_IDS = ["6589033626", "256762238238751"];

const startTime = Math.floor(Date.now() / 1000);

const client = new Client({
  authStrategy: new LocalAuth(),
  puppeteer: { args: ["--no-sandbox"] },
});

client.on("qr", (qr) => {
  console.log("Scan this QR code in WhatsApp > Linked devices:");
  qrcode.generate(qr, { small: true });
});
client.on("authenticated", () => console.log("Authenticated (login OK)"));
client.on("auth_failure", (m) => console.log("Auth failure:", m));
client.on("ready", () => console.log("WhatsApp bridge ready - waiting for messages..."));

// Fetch approved requesters dynamically from agent.py
async function getRequestersMap() {
  try {
    const res = await fetch(USERS_URL);
    if (!res.ok) return {};
    const data = await res.json();
    const map = {};
    if (data && Array.isArray(data.requesters)) {
      for (const r of data.requesters) {
        if (r.id) {
          map[r.id] = { phone: r.phone, name: r.name, role: "requester" };
        }
      }
    }
    return map;
  } catch (e) {
    console.error("Failed to fetch users from agent:", e.message);
    return {};
  }
}

// Report unknown sender to agent.py so it gets logged in the Unknown sheet
async function reportUnknown(idUser, text) {
  try {
    await fetch(UNKNOWN_URL, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ id: idUser, text: text }),
    });
  } catch (e) {
    // ignore network errors if agent is restarting
  }
}

async function forwardToAgent(msg, user) {
  console.log("[1] Text received. name =", user.name, "| role =", user.role, "| text =", msg.body);
  const res = await fetch(AGENT_URL, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ sender: `${user.name} (${user.phone})`, role: user.role, text: msg.body }),
  });
  const data = await res.json();
  console.log("[2] Agent replied:", (data.reply || "").split("\n")[0]);
  if (data.reply) {
    await msg.reply(data.reply);
  }
}

// Messages from other people
client.on("message", async (msg) => {
  try {
    if (msg.fromMe) return;
    if (msg.timestamp < startTime) return; // old message
    if (msg.from === "status@broadcast") return; // Status updates
    if (!msg.from.endsWith("@c.us") && !msg.from.endsWith("@lid")) return; // groups/channels
    if (!msg.body || !msg.body.trim()) return; // empty / non-text events

    const idUser = msg.from.split("@")[0];
    const usersMap = await getRequestersMap();
    const user = usersMap[idUser];

    if (!user) {
      console.log(
        "[x] Ignored: unknown sender. id =", idUser, "| text =", msg.body,
        "\n    Sender logged to agent's 'Unknown' sheet. Link them via agent startup option 2."
      );
      await reportUnknown(idUser, msg.body);
      return;
    }
    await forwardToAgent(msg, user);
  } catch (e) {
    console.error("Bridge error:", e.message);
  }
});

// Messages typed by the linked account itself ("Message yourself" chat only)
client.on("message_create", async (msg) => {
  try {
    if (!msg.fromMe) return;
    if (msg.timestamp < startTime) return;
    if (!msg.body || !msg.body.trim()) return;

    // PREVENT LOOP: Ignore bot's own replies so it doesn't loop forever
    const text = msg.body.trim();
    if (
      text.startsWith("Inventory:") ||
      text.startsWith("⚠️ Low-Stock Dashboard") ||
      text.startsWith("Unknown owner command") ||
      text.startsWith("Approved.") ||
      text.startsWith("Successfully received") ||
      text.startsWith("Ticket ") ||
      text.startsWith("Request received:")
    ) {
      return;
    }

    const me = (client.info && client.info.wid && client.info.wid._serialized) || "";
    const toUser = String(msg.to || "").split("@")[0];
    const meUser = me.split("@")[0];
    console.log("[self] Command seen:", msg.body, "| from =", msg.from, "| to =", msg.to, "| me =", me);

    // The self chat is the chat addressed to your own account (phone id or long id).
    let isSelfChat = msg.to === msg.from || msg.to === me || MY_IDS.includes(toUser) || toUser === meUser;
    if (!isSelfChat) {
      try { // fallback: ask WhatsApp whether the chat is "me"
        const chat = await msg.getChat();
        const ct = await chat.getContact();
        isSelfChat = !!(ct && ct.isMe);
        if (isSelfChat) console.log("[self] Self chat detected via contact. Add", toUser, "to MY_IDS to skip this check.");
      } catch (e) {}
    }
    if (!isSelfChat) return console.log("[x] Ignored: command was typed in another chat, not 'Message yourself'");

    // Forward valid user commands to agent.py
    await forwardToAgent(msg, SELF_OWNER);
  } catch (e) {
    console.error("Bridge error (self chat):", e.message);
  }
});

// Small local server so agent.py can send WhatsApp messages by phone number.
http
  .createServer((req, res) => {
    if (req.method !== "POST" || req.url !== "/send") {
      res.writeHead(404);
      return res.end();
    }
    let body = "";
    req.on("data", (c) => (body += c));
    req.on("end", async () => {
      try {
        const { phone, text } = JSON.parse(body);
        const wid = await client.getNumberId(String(phone));
        if (!wid) {
          res.writeHead(404);
          return res.end(JSON.stringify({ ok: false, error: "number not on WhatsApp" }));
        }
        await client.sendMessage(wid._serialized, text);
        console.log("[send] Message sent to", phone);
        res.writeHead(200);
        res.end(JSON.stringify({ ok: true }));
      } catch (e) {
        console.error("[send] failed:", e.message);
        res.writeHead(500);
        res.end(JSON.stringify({ ok: false, error: e.message }));
      }
    });
  })
  .listen(SEND_PORT, "127.0.0.1", () => console.log("Send endpoint on port", SEND_PORT));

client.initialize();