const http = require("http");
const fs = require("fs");
const path = require("path");

const root = __dirname;
// 默认监听所有网卡，方便同一局域网的设备访问；仍可用 HOST=127.0.0.1 限制为本机。
const host = process.env.HOST || "0.0.0.0";

function loadEnvFile() {
  const file = path.join(root, ".env");
  if (!fs.existsSync(file)) return;
  for (const line of fs.readFileSync(file, "utf8").split(/\r?\n/)) {
    const match = line.match(/^\s*([A-Z0-9_]+)\s*=\s*(.*?)\s*$/);
    if (match && !process.env[match[1]]) process.env[match[1]] = match[2].replace(/^['"]|['"]$/g, "");
  }
}
loadEnvFile();

const port = Number(process.env.PORT || 3000);
const types = { ".html": "text/html; charset=utf-8", ".js": "text/javascript; charset=utf-8", ".css": "text/css; charset=utf-8", ".json": "application/json; charset=utf-8", ".svg": "image/svg+xml", ".png": "image/png", ".ico": "image/x-icon" };
const limits = new Map();
const RATE_WINDOW_MS = 60_000;
const RATE_LIMITS = { chat: Number(process.env.CHAT_RATE_LIMIT || 20), speech: Number(process.env.SPEECH_RATE_LIMIT || 10), voices: 20, status: 60 };

function securityHeaders(contentType = "application/json; charset=utf-8") {
  return {
    "Content-Type": contentType,
    "Cache-Control": "no-store",
    "X-Content-Type-Options": "nosniff",
    "X-Frame-Options": "DENY",
    "Referrer-Policy": "no-referrer",
    "Permissions-Policy": "geolocation=(), microphone=(), camera=()",
  };
}

function json(res, status, body) {
  res.writeHead(status, securityHeaders());
  res.end(JSON.stringify(body));
}

function limited(req, res, bucket) {
  const ip = req.socket.remoteAddress || "local";
  const key = `${bucket}:${ip}`;
  const now = Date.now();
  const record = limits.get(key) || { started: now, count: 0 };
  if (now - record.started >= RATE_WINDOW_MS) Object.assign(record, { started: now, count: 0 });
  record.count += 1;
  limits.set(key, record);
  if (record.count <= RATE_LIMITS[bucket]) return false;
  res.writeHead(429, { ...securityHeaders(), "Retry-After": String(Math.ceil((RATE_WINDOW_MS - (now - record.started)) / 1000)) });
  res.end(JSON.stringify({ error: "请求过于频繁，请稍后再试。" }));
  return true;
}
setInterval(() => {
  const cutoff = Date.now() - RATE_WINDOW_MS * 2;
  for (const [key, value] of limits) if (value.started < cutoff) limits.delete(key);
}, RATE_WINDOW_MS).unref();

function readJson(req, maxBytes = 32 * 1024) {
  return new Promise((resolve, reject) => {
    let size = 0;
    const chunks = [];
    req.on("data", (chunk) => {
      size += chunk.length;
      if (size > maxBytes) { reject(new Error("请求内容过大")); req.destroy(); return; }
      chunks.push(chunk);
    });
    req.on("end", () => {
      try { resolve(JSON.parse(Buffer.concat(chunks).toString("utf8"))); }
      catch { reject(new Error("请求格式无效")); }
    });
    req.on("error", reject);
  });
}

function cleanText(value, max = 80) {
  return String(value || "").replace(/[\u0000-\u001f]/g, " ").trim().slice(0, max);
}

function validateLearningContext(context) {
  if (!context || typeof context !== "object" || Array.isArray(context)) throw new Error("学习信息格式无效");
  const result = {
    teacherName: cleanText(context.teacherName, 80), teacherNote: cleanText(context.teacherNote, 140), grade: cleanText(context.grade, 30), subject: cleanText(context.subject, 30), style: cleanText(context.style, 30), currentTask: cleanText(context.currentTask, 180),
    speechStyle: cleanText(context.speechStyle, 30), language: cleanText(context.language, 30), textbook: cleanText(context.textbook, 40),
    goal: cleanText(context.goal, 40), minutes: Number(context.minutes), days: Number(context.days), lowPressure: Boolean(context.lowPressure), companion: cleanText(context.companion, 12),
  };
  if (!result.grade || !result.subject || !result.language || !Number.isFinite(result.minutes) || result.minutes < 5 || result.minutes > 180 || !Number.isFinite(result.days) || result.days < 1 || result.days > 7) throw new Error("学习信息不完整");
  return result;
}

function validateMessages(messages) {
  if (!Array.isArray(messages) || messages.length < 1 || messages.length > 12) throw new Error("聊天上下文无效");
  const safe = messages.filter((message) => message && ["user", "assistant"].includes(message.role)).map((message) => ({ role: message.role, content: cleanText(message.content, 1200) })).filter((message) => message.content);
  if (!safe.length || safe.filter((message) => message.role === "user").length < 1) throw new Error("请先输入学习问题");
  return safe.slice(-10);
}

function serverPrompt(context) {
  const companion = context.companion === "学习伙伴" ? "学习伙伴" : "老师";
  return `你是造师台中的温和${companion}${context.teacherName ? `「${context.teacherName}」` : ""}，负责${context.subject}学习支持。学生当前阶段：${context.grade}；教材：${context.textbook || "通用教材"}；目标：${context.goal || "循序学习"}；每次约${context.minutes}分钟、每周${context.days}天。${context.currentTask ? ` 当前学习任务：${context.currentTask}。` : ""}请使用${context.language}，以${context.style || "耐心启发"}和${context.speechStyle || "清晰亲切"}的方式交流。${context.teacherNote ? ` 你的专属陪学设定是：${context.teacherNote}。` : ""}\n\n回答质量要求：先判断用户是在提问概念、求解题目、练习、复盘，还是需要情绪支持。对学习问题，先用一句话给出核心结论或思路，再用 2～4 个短步骤解释；有计算或推理时写出关键依据和过程，不能只给最终答案。尽量给一个贴近学生阶段的例子，并指出一个常见误区。信息不足时只问一个最关键的澄清问题，不要连续追问。对不确定、可能因教材版本不同而变化的内容，要明确说明范围或不确定性，不要编造事实。\n\n使用易扫读的格式：可用 **核心思路**、**一步一步来**、**小练习** 作为短标签和项目符号；普通问题控制在约 250 字内，复杂问题可更长但不要堆砌。每次学习性回复的结尾给出一个可选的“下一步”，例如一道 1～3 分钟的小练习、让学生复述一句话，或询问是否需要换一种讲法。学生只想要答案时，仍给答案并用一句话说明关键理由。\n\n不要要求或猜测学校、班级、住址、教师姓名、联系方式、身份证件、密码、验证码等个人或敏感信息。若学生说“抗拒学习、害怕被批评、很疲惫、想缓一缓”，先共情，允许暂停或把任务缩小到5分钟；不要强迫其解释原因。若出现自伤、自杀、被欺凌、严重恐惧上学或立即危险，先表达关心和认真对待，鼓励尽快联系可信赖的成人（家长、监护人、学校心理老师）或当地紧急/专业支持；如有即时危险，应立即联系当地紧急服务。不要替代专业医疗、心理或危机支持。${context.lowPressure ? " 当前为低压力陪学模式：避免积分、排名、打卡和催促措辞；提供可跳过的5分钟微任务。" : ""}`;
}

function agentConfig() {
  const apiKey = process.env.AGENT_API_KEY || process.env.DEEPSEEK_API_KEY;
  const baseUrl = (process.env.AGENT_BASE_URL || "https://api.deepseek.com").replace(/\/$/, "");
  const model = process.env.AGENT_MODEL || "deepseek-chat";
  const url = new URL(baseUrl);
  if (!apiKey || url.protocol !== "https:") throw new Error("agent unavailable");
  return { apiKey, baseUrl: url.toString().replace(/\/$/, ""), model };
}

async function askConfiguredAgent(context, messages) {
  const config = agentConfig();
  const response = await fetch(`${config.baseUrl}/chat/completions`, {
    method: "POST", headers: { "Content-Type": "application/json", Authorization: `Bearer ${config.apiKey}` },
    body: JSON.stringify({ model: config.model, messages: [{ role: "system", content: serverPrompt(context) }, ...messages], temperature: 0.35, max_tokens: 1000, stream: false }),
  });
  const data = await response.json().catch(() => ({}));
  if (!response.ok) throw new Error(`agent ${response.status}`);
  const content = cleanText(data.choices?.[0]?.message?.content, 5000);
  if (!content) throw new Error("empty agent reply");
  return content;
}

function fallbackLearningReply(context, messages) {
  const last = [...messages].reverse().find((message) => message && message.role === "user");
  const question = cleanText(last?.content, 240) || "这个知识点";
  const companion = context.companion === "学习伙伴" ? "学习伙伴" : "老师";
  return `我是${context.teacherName || "你的"}${companion}。外部智能服务暂时不可用，但我仍可以先陪你完成这一小步。\n\n关于「${question}」，请先说说你最不明白的是概念、题目步骤，还是记忆方法？我们从一个最小的问题开始。`;
}

function solveArithmetic(expression) {
  if (typeof expression !== "string" || expression.length > 80 || !/^[\d+\-*/().]+$/.test(expression) || !/[+\-*/]/.test(expression)) return null;
  let index = 0;
  const skip = () => { while (/\s/.test(expression[index] || "")) index += 1; };
  const number = () => {
    skip();
    const match = expression.slice(index).match(/^(?:\d+(?:\.\d*)?|\.\d+)/);
    if (!match) throw new Error("number expected");
    index += match[0].length;
    return Number(match[0]);
  };
  const factor = () => {
    skip();
    if (expression[index] === "+") { index += 1; return factor(); }
    if (expression[index] === "-") { index += 1; return -factor(); }
    if (expression[index] === "(") {
      index += 1;
      const value = sum();
      skip();
      if (expression[index] !== ")") throw new Error("closing bracket expected");
      index += 1;
      return value;
    }
    return number();
  };
  const product = () => {
    let value = factor();
    while (true) {
      skip();
      const operator = expression[index];
      if (operator !== "*" && operator !== "/") return value;
      index += 1;
      const next = factor();
      if (operator === "/" && next === 0) throw new Error("division by zero");
      value = operator === "*" ? value * next : value / next;
      if (!Number.isFinite(value) || Math.abs(value) > 1e15) throw new Error("result out of range");
    }
  };
  const sum = () => {
    let value = product();
    while (true) {
      skip();
      const operator = expression[index];
      if (operator !== "+" && operator !== "-") return value;
      index += 1;
      const next = product();
      value = operator === "+" ? value + next : value - next;
      if (!Number.isFinite(value) || Math.abs(value) > 1e15) throw new Error("result out of range");
    }
  };
  try {
    const value = sum();
    skip();
    return index === expression.length && Number.isFinite(value) ? value : null;
  } catch {
    return null;
  }
}

function arithmeticFromQuestion(question, subject) {
  const raw = String(question || "").replace(/＝/g, "=").trim();
  const hasMathCue = /(?:计算|算一下|帮我算|请算|等于|多少|得多少|结果|答案)/.test(raw);
  const possibleExpression = raw.replace(/[×xX]/g, "*").replace(/÷/g, "/").split("=")[0].replace(/[？?！!。。，,\s]/g, "");
  const looksLikeArithmetic = /^[\d+\-*/().]+$/.test(possibleExpression) && /[+\-*/]/.test(possibleExpression);
  if (!hasMathCue && !looksLikeArithmetic) return null;
  if (!hasMathCue && /^20\d{2}[-/.]\d{1,2}[-/.]\d{1,2}$/.test(raw)) return null;
  const expression = raw.split("=")[0]
    .replace(/[×xX]/g, "*").replace(/÷/g, "/")
    .replace(/(?:请|帮我|帮忙|计算一下|计算|算一下|请算|求解|求|等于多少|是多少|得多少|结果|答案|[？?！!。。，,])/g, "")
    .replace(/\s+/g, "");
  const result = solveArithmetic(expression);
  return result === null ? null : { expression, result };
}

function formatArithmeticNumber(value) {
  if (Number.isInteger(value)) return String(value);
  return String(Number(value.toFixed(10)));
}

function arithmeticReason(expression, result) {
  const simple = expression.match(/^(\d+(?:\.\d+)?)\s*([+\-*/])\s*(\d+(?:\.\d+)?)$/);
  if (!simple) return `按“先算括号，再算乘除，最后算加减”的顺序，算得 ${formatArithmeticNumber(result)}。`;
  const [, left, operator, right] = simple;
  const verb = { "+": "加上", "-": "减去", "*": "乘以", "/": "除以" }[operator];
  return `${left} ${verb} ${right}，得到 ${formatArithmeticNumber(result)}。`;
}

function buildLocalLearningReply(context, messages) {
  const last = [...messages].reverse().find((message) => message && message.role === "user");
  const question = cleanText(last?.content, 240) || "今天的学习";
  const teacher = context.teacherName ? `${context.teacherName}${context.companion === "学习伙伴" ? "同学" : "老师"}` : "学习伙伴";
  const subject = context.subject || "学习";
  const goal = context.goal || "打牢基础";
  const currentTask = cleanText(context.currentTask, 120) || `${subject}的${goal}`;
  const minutes = Math.min(90, Math.max(10, Math.round(Number(context.minutes) || 30)));
  const warmup = Math.max(3, Math.min(5, Math.round(minutes * 0.15)));
  const review = Math.max(3, Math.min(6, Math.round(minutes * 0.2)));
  const practice = Math.max(5, minutes - warmup - review);
  const subjectWarmup = { 数学: "回想一条公式或一道例题的解题步骤", 英语: "朗读 5 个单词或 2 句课文", 语文: "读一小段课文，圈出一个关键词", 科学: "写下今天最想弄懂的一个现象", 编程: "回顾一个已学过的命令或代码例子" }[subject] || "回想昨天学到的一件小事";
  const lower = question.toLowerCase();
  const arithmetic = arithmeticFromQuestion(question, subject);
  if (arithmetic) {
    const display = arithmetic.expression.replace(/\*/g, "×").replace(/\//g, "÷");
    return `${teacher}来解这道题：\n\n**答案：** ${display} = ${formatArithmeticNumber(arithmetic.result)}。\n\n**怎么算：** ${arithmeticReason(arithmetic.expression, arithmetic.result)}\n\n下一步：再发一道同类型题，我会继续帮你检查步骤。`;
  }
  if (/(抗拒|害怕|疲惫|累了|想缓|不想学|难过)/.test(question)) return `${teacher}在。你不用马上解释原因，也不必硬撑。\n\n今天先选一个最轻的版本：\n1. 喝口水、伸伸手，休息 2 分钟。\n2. 只做 5 分钟：${subjectWarmup}。\n3. 做完就可以暂停，或者把一句感受告诉我。\n\n你愿意先做第 1 步，还是直接开始这 5 分钟？`;
  if (/(计划|安排|今天.*学|今日.*学|学习.*计划)/.test(lower)) return `${teacher}给你排一份 ${minutes} 分钟的 ${subject} 小计划：\n\n1. 热身 ${warmup} 分钟：${subjectWarmup}。\n2. 专注 ${practice} 分钟：围绕「${currentTask}」看一个例子，再完成 2 道基础练习。\n3. 复盘 ${review} 分钟：写下“我学会了什么”和“我还想问什么”各一句。\n\n现在先做第一步：${subjectWarmup}。做完后告诉我“热身完成”，我再带你做下一步。`;
  if (/(出题|练习|测验|小题)/.test(lower)) return `${teacher}给你一个小练习：\n\n请用一句话说出「${currentTask}」里最重要的规则或方法。\n\n如果一时想不起来，也可以先回答：它更像是“概念”“计算步骤”还是“记忆内容”？我会按你的回答出下一题。`;
  if (/(什么是|解释|讲解|不明白|概念|为什么)/.test(lower)) return `${teacher}可以讲解「${question}」。先告诉我：你卡在“词语意思”“题目步骤”还是“为什么要这样做”？\n\n你只要把题目、课本句子或知识点发来，我会按“先说结论 → 举一个例子 → 给一题小练习”的方式讲，不会一下塞很多内容。`;
  return `${teacher}收到「${question}」。我们先把它变成一个小目标：\n\n1. 找出你要学的一个知识点或一道题。\n2. 用 ${Math.min(10, minutes)} 分钟看例子或课本。\n3. 用一句话复述你理解的内容。\n\n把具体题目、课文句子或知识点发给我，我会直接陪你一步一步完成。`;
}

async function createElevenLabsSpeech(text, selectedVoiceId = "") {
  const apiKey = process.env.ELEVENLABS_API_KEY;
  const voiceId = selectedVoiceId || process.env.ELEVENLABS_VOICE_ID;
  if (!apiKey || !/^[A-Za-z0-9_-]{5,128}$/.test(voiceId || "")) throw new Error("speech unavailable");
  const response = await fetch(`https://api.elevenlabs.io/v1/text-to-speech/${encodeURIComponent(voiceId)}?output_format=mp3_44100_128`, {
    method: "POST", headers: { "Content-Type": "application/json", "xi-api-key": apiKey },
    body: JSON.stringify({ text: text.slice(0, 1200), model_id: process.env.ELEVENLABS_TTS_MODEL || "eleven_multilingual_v2", language_code: "zh", voice_settings: { stability: 0.42, similarity_boost: 0.78, style: 0.28, use_speaker_boost: true } }),
  });
  if (!response.ok) throw new Error(`speech ${response.status}`);
  return Buffer.from(await response.arrayBuffer());
}

const server = http.createServer(async (req, res) => {
  let pathname;
  try { pathname = decodeURIComponent((req.url || "/").split("?")[0]); }
  catch { return json(res, 400, { error: "请求地址无效" }); }

  if (req.method === "POST" && pathname === "/api/chat") {
    if (limited(req, res, "chat")) return;
    try {
      const body = await readJson(req);
      const context = validateLearningContext(body.learningContext);
      const messages = validateMessages(body.messages);
      try { return json(res, 200, { reply: await askConfiguredAgent(context, messages) }); }
      catch (error) { console.warn("chat agent unavailable, using local fallback:", error.message); return json(res, 200, { reply: buildLocalLearningReply(context, messages), fallback: true }); }
    } catch (error) { console.warn("chat request rejected:", error.message); return json(res, 400, { error: "聊天请求格式无效，请重新打开页面后再试。" }); }
  }
  if (req.method === "POST" && pathname === "/api/speech") {
    if (limited(req, res, "speech")) return;
    try { const body = await readJson(req); const text = cleanText(body.text, 1200); if (!text) throw new Error("empty text"); const audio = await createElevenLabsSpeech(text, cleanText(body.voiceId, 128)); res.writeHead(200, securityHeaders("audio/mpeg")); return res.end(audio); }
    catch (error) { console.warn("speech request rejected:", error.message); return json(res, 400, { error: "语音服务暂时不可用。" }); }
  }
  if (req.method === "GET" && pathname === "/api/voices") {
    if (limited(req, res, "voices")) return;
    const apiKey = process.env.ELEVENLABS_API_KEY;
    if (!apiKey) return json(res, 200, { enabled: false, voices: [] });
    try { const response = await fetch("https://api.elevenlabs.io/v1/voices", { headers: { "xi-api-key": apiKey } }); const data = await response.json().catch(() => ({})); if (!response.ok) throw new Error("voices unavailable"); const voices = (data.voices || []).map((voice) => ({ id: cleanText(voice.voice_id, 128), name: cleanText(voice.name, 80), labels: { accent: cleanText(voice.labels?.accent, 40) } })).filter((voice) => voice.id && voice.name); return json(res, 200, { enabled: true, voices }); }
    catch (error) { console.warn("voice list rejected:", error.message); return json(res, 200, { enabled: false, voices: [] }); }
  }
  if (req.method === "GET" && pathname === "/api/status") { if (limited(req, res, "status")) return; return json(res, 200, { agent: Boolean(process.env.AGENT_API_KEY || process.env.DEEPSEEK_API_KEY), speech: Boolean(process.env.ELEVENLABS_API_KEY) }); }
  if (req.method !== "GET" && req.method !== "HEAD") { res.writeHead(405, { Allow: "GET, HEAD, POST" }); return res.end(); }
  const requested = pathname === "/" ? "/index.html" : pathname;
  const file = path.resolve(root, `.${requested}`);
  if (!file.startsWith(root + path.sep) && file !== path.join(root, "index.html")) { res.writeHead(403); return res.end("Forbidden"); }
  fs.readFile(file, (error, data) => { if (error) { res.writeHead(error.code === "ENOENT" ? 404 : 500, securityHeaders("text/plain; charset=utf-8")); return res.end(error.code === "ENOENT" ? "Not found" : "Server error"); } res.writeHead(200, securityHeaders(types[path.extname(file)] || "application/octet-stream")); if (req.method === "HEAD") return res.end(); res.end(data); });
});
server.listen(port, host, () => console.log(`造师台已运行：http://localhost:${port}（监听 ${host}）`));
