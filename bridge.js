// WhatsApp bridge (unofficial, QR login) - cloud-ready version.
//  - Forwards messages from approved senders & owner to agent.py
//  - Lets agent.py send WhatsApp messages (PO requests to the lab owner, updates to requesters)
//  - Messages typed in the linked account's "Message yourself" chat are owner commands.
//  - Reports login status / QR code to agent.py so you can scan it from the /admin web page.
const fs = require("fs");
const path = require("path");
const http = require("http");
const { Client, LocalAuth } = require("whatsapp-web.js");
const qrcode = require("qrcode-terminal");

const AGENT_PORT = process.env.PORT || 5000; // agent.py listens on the host-assigned PORT
const AGENT = `http://127.0.0.1:${AGENT_PORT}`;
const SEND_PORT = Number(process.env.BRIDGE_PORT || 5001); // agent.py calls this to send messages
const AUTH_DIR = process.env.WWEBJS_AUTH_DIR || path.join(__dirname, ".wwebjs_auth");

// The lab owner is the same person as the linked WhatsApp account.
const OWNER_PHONE = (process.env.OWNER_PHONE || "6589033626").replace(/\D/g, "");
const SELF_OWNER = { phone: OWNER_PHONE, name: process.env.BUYER_NAME || "YH Lee (SG)", role: "owner" };

// Your own account appears under two ids in the "Message yourself" chat:
// the phone number and a long WhatsApp id (seen in the log as "to = ...@lid").
const MY_IDS = [OWNER_PHONE, ...(process.env.OWNER_LID || "256762238238751").split(",")]
  .map((s) => s.replace(/\D/g, ""))
  .filter(Boolean);

const startTime = Math.floor(Date.now() / 1000);

// After a container restart Chromium can leave lock files behind and refuse to start.
function clearChromiumLocks(dir) {
  if (!fs.existsSync(dir)) return;
  for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
    const p = path.join(dir, entry.name);
    if (entry.isDirectory()) clearChromiumLocks(p);
    else if (/^Singleton(Lock|Socket|Cookie)$/.test(entry.name)) {
      try { fs.rmSync(p, { force: true }); } catch (e) {}
    }
  }
}
clearChromiumLocks(AUTH_DIR);

async function reportState(status, qr = null) {
  try {
    await fetch(`${AGENT}/bridge-state`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ status, qr }),
    });
  } catch (e) {
    // agent may still be starting
  }
}

const puppeteer = {
  args: ["--no-sandbox", "--disable-setuid-sandbox", "--disable-dev-shm-usage", "--disable-gpu"],
};
if (process.env.PUPPETEER_EXECUTABLE_PATH) puppeteer.executablePath = process.env.PUPPETEER_EXECUTABLE_PATH;

// Optional: fresh logins against the newest WhatsApp Web build can crash with "Execution context was
// destroyed". Setting WA_WEB_VERSION (e.g. 2.3000.1052477439) pins a build; the matching <version>.html
// must exist in WA_WEB_CACHE_DIR. Leave unset to use the live WhatsApp Web (default).
const clientOptions = {
  authStrategy: new LocalAuth({ dataPath: AUTH_DIR }),
  puppeteer,
  // The library's default identity is Chrome 101, which WhatsApp Web no longer fully loads for.
  userAgent: process.env.WA_USER_AGENT || "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/156.0.0.0 Safari/537.36",
  // First load on a small cloud machine can exceed the library's 30 s default.
  authTimeoutMs: 180000,
  webVersionCache: { type: "none" },
};
if (process.env.WA_WEB_VERSION) {
  clientOptions.webVersion = process.env.WA_WEB_VERSION;
  clientOptions.webVersionCache = {
    type: "local",
    path: process.env.WA_WEB_CACHE_DIR || path.join(__dirname, "wa-web-cache"),
    strict: true,
  };
  console.log("Using pinned WhatsApp Web version", process.env.WA_WEB_VERSION);
}
const client = new Client(clientOptions);

client.on("qr", (qr) => {
  console.log("Scan this QR code in WhatsApp > Linked devices (or open /admin?key=... in a browser):");
  qrcode.generate(qr, { small: true });
  reportState("waiting for QR scan", qr);
});
client.on("authenticated", () => {
  console.log("Authenticated (login OK)");
  reportState("authenticated");
});
client.on("auth_failure", (m) => {
  console.log("Auth failure:", m);
  reportState("auth failure - restart to get a new QR");
});
client.on("ready", () => {
  console.log("WhatsApp bridge ready - waiting for messages...");
  reportState("ready");
});
client.on("disconnected", (reason) => {
  console.log("WhatsApp disconnected:", reason);
  reportState("disconnected");
  process.exit(1); // agent.py restarts the bridge
});

// Fetch approved requesters dynamically from agent.py
async function getRequestersMap() {
  try {
    const res = await fetch(`${AGENT}/users`);
    if (!res.ok) return {};
    const data = await res.json();
    const map = {};
    if (data && Array.isArray(data.requesters)) {
      for (const r of data.requesters) {
        if (r.id) map[r.id] = { phone: r.phone, name: r.name, role: "requester" };
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
    await fetch(`${AGENT}/unknown`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ id: idUser, text: text }),
    });
  } catch (e) {}
}

async function forwardToAgent(msg, user) {
  console.log("[1] Text received. name =", user.name, "| role =", user.role, "| text =", msg.body);
  const res = await fetch(`${AGENT}/incoming`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ sender: `${user.name} (${user.phone})`, role: user.role, text: msg.body }),
  });
  const data = await res.json();
  console.log("[2] Agent replied:", (data.reply || "").split("\n")[0]);
  if (data.reply) await msg.reply(data.reply);
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
      console.log("[x] Ignored: unknown sender. id =", idUser, "| text =", msg.body,
        "\n    Logged to the 'Unknown' sheet. Copy the id into the Requesters sheet (WhatsAppID) to allow them.");
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

    // PREVENT LOOP: ignore the bot's own replies
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

    let isSelfChat = msg.to === msg.from || msg.to === me || MY_IDS.includes(toUser) || toUser === meUser;
    if (!isSelfChat) {
      try { // fallback: ask WhatsApp whether the chat is "me"
        const chat = await msg.getChat();
        const ct = await chat.getContact();
        isSelfChat = !!(ct && ct.isMe);
        if (isSelfChat) console.log("[self] Self chat detected via contact. Add", toUser, "to OWNER_LID to skip this check.");
      } catch (e) {}
    }
    if (!isSelfChat) return console.log("[x] Ignored: command was typed in another chat, not 'Message yourself'");

    await forwardToAgent(msg, SELF_OWNER);
  } catch (e) {
    console.error("Bridge error (self chat):", e.message);
  }
});

// Small local server so agent.py can send WhatsApp messages by phone number (not exposed publicly).
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

// Any startup crash (e.g. "Execution context was destroyed" while WhatsApp Web reloads)
// exits the process; agent.py restarts the bridge after 10 s, which normally succeeds.
process.on("unhandledRejection", (e) => {
  console.error("Bridge crashed:", (e && e.message) || e);
  process.exit(1);
});

reportState("starting");
client.initialize().catch((e) => {
  console.error("WhatsApp start failed, will retry:", e.message);
  process.exit(1);
});
