const pptxgen = require('pptxgenjs');

const pptx = new pptxgen();
pptx.layout = 'LAYOUT_WIDE';
pptx.author = '造师台项目组';
pptx.company = '造师台';
pptx.subject = '造师台项目介绍';
pptx.title = '造师台｜你的专属 AI 学习空间';
pptx.lang = 'zh-CN';
pptx.theme = {
  headFontFace: 'Microsoft YaHei',
  bodyFontFace: 'Microsoft YaHei',
  lang: 'zh-CN',
};
pptx.defineLayout({ name: 'CUSTOM_WIDE', width: 13.333, height: 7.5 });
pptx.layout = 'CUSTOM_WIDE';
pptx.defineSlideMaster({
  title: 'MASTER',
  background: { color: 'F7F9FD' },
  objects: [
    { rect: { x: 0, y: 0, w: 13.333, h: 0.08, fill: { color: '315EFB' }, line: { color: '315EFB' } } },
    { text: { text: '造师台 · AI LEARNING STUDIO', options: { x: 0.55, y: 7.08, w: 4.5, h: 0.18, fontFace: 'Microsoft YaHei', fontSize: 8, color: '8090A9', margin: 0 } } },
    { text: { text: '项目介绍', options: { x: 11.8, y: 7.08, w: 0.95, h: 0.18, fontFace: 'Microsoft YaHei', fontSize: 8, color: '8090A9', align: 'right', margin: 0 } } },
  ],
  slideNumber: { x: 12.8, y: 7.08, color: '8090A9', fontFace: 'Microsoft YaHei', fontSize: 8 },
});

const C = { blue: '315EFB', indigo: '6654D8', navy: '26385C', text: '33415C', muted: '73819A', line: 'DFE6F3', pale: 'EEF3FF', mint: 'EAF9F4', green: '20A782', orange: 'F49A47', white: 'FFFFFF', bg: 'F7F9FD', red: 'C75B6B' };
const FONT = 'Microsoft YaHei';

function addTitle(slide, title, subtitle = '') {
  slide.addText(title, { x: 0.62, y: 0.43, w: 8.9, h: 0.45, fontFace: FONT, fontSize: 25, bold: true, color: C.navy, margin: 0 });
  if (subtitle) slide.addText(subtitle, { x: 0.64, y: 0.96, w: 10.8, h: 0.27, fontFace: FONT, fontSize: 10.5, color: C.muted, margin: 0 });
}

function addPill(slide, text, x, y, w, color = C.blue, fill = C.pale) {
  slide.addShape(pptx.ShapeType.roundRect, { x, y, w, h: 0.32, rectRadius: 0.06, fill: { color: fill }, line: { color: fill } });
  slide.addText(text, { x, y: y + 0.06, w, h: 0.16, fontFace: FONT, fontSize: 8.5, bold: true, color, align: 'center', margin: 0 });
}

function addCard(slide, { x, y, w, h, title, body, accent = C.blue, icon = '' }) {
  slide.addShape(pptx.ShapeType.roundRect, { x, y, w, h, rectRadius: 0.08, fill: { color: C.white }, line: { color: C.line, width: 1 } });
  slide.addShape(pptx.ShapeType.rect, { x, y, w: 0.07, h, fill: { color: accent }, line: { color: accent } });
  if (icon) {
    slide.addShape(pptx.ShapeType.ellipse, { x: x + 0.25, y: y + 0.25, w: 0.42, h: 0.42, fill: { color: accent, transparency: 86 }, line: { color: accent, transparency: 100 } });
    slide.addText(icon, { x: x + 0.25, y: y + 0.31, w: 0.42, h: 0.18, fontFace: FONT, fontSize: 11, color: accent, align: 'center', margin: 0 });
  }
  slide.addText(title, { x: x + (icon ? 0.8 : 0.25), y: y + 0.25, w: w - (icon ? 1.0 : 0.5), h: 0.25, fontFace: FONT, fontSize: 13, bold: true, color: C.text, margin: 0 });
  slide.addText(body, { x: x + 0.25, y: y + 0.72, w: w - 0.5, h: h - 0.9, fontFace: FONT, fontSize: 10, color: C.muted, breakLine: false, valign: 'top', margin: 0.02, paraSpaceAfterPt: 7, fit: 'shrink' });
}

function addBulletList(slide, items, x, y, w, h, color = C.text, size = 13) {
  const runs = [];
  items.forEach((item, index) => {
    runs.push({ text: item, options: { bullet: { indent: 15 }, hanging: 3, breakLine: index !== items.length - 1 } });
  });
  slide.addText(runs, { x, y, w, h, fontFace: FONT, fontSize: size, color, breakLine: false, paraSpaceAfterPt: 12, breakLine: false, margin: 0.02, valign: 'mid', fit: 'shrink' });
}

function addArrow(slide, x, y, w, color = 'AAB8D4') {
  slide.addShape(pptx.ShapeType.chevron, { x, y, w, h: 0.35, fill: { color }, line: { color } });
}

// 1. Cover
{
  const s = pptx.addSlide('MASTER');
  s.background = { color: 'F4F7FF' };
  s.addShape(pptx.ShapeType.arc, { x: 8.65, y: -1.3, w: 5.2, h: 5.2, adjustPoint: 0.25, line: { color: 'DCE5FF', width: 20, transparency: 35 }, rotate: 25 });
  s.addShape(pptx.ShapeType.ellipse, { x: 10.5, y: 4.65, w: 2.2, h: 2.2, fill: { color: 'DFF6EE', transparency: 8 }, line: { color: 'DFF6EE' } });
  addPill(s, 'AI LEARNING STUDIO', 0.72, 1.25, 1.85, C.blue, 'E5ECFF');
  s.addText('造师台', { x: 0.7, y: 1.82, w: 4.8, h: 0.82, fontFace: FONT, fontSize: 42, bold: true, color: C.navy, margin: 0 });
  s.addText('让每一段学习，都有懂你的老师', { x: 0.73, y: 2.73, w: 6.5, h: 0.45, fontFace: FONT, fontSize: 23, color: C.blue, bold: true, margin: 0 });
  s.addText('一个面向学生的个性化 AI 学习空间\n从专属老师、学习方案到安全陪学，帮助学习更清晰、更温和地发生。', { x: 0.74, y: 3.48, w: 6.1, h: 0.8, fontFace: FONT, fontSize: 14, color: C.muted, breakLine: false, margin: 0, breakLine: false });
  s.addShape(pptx.ShapeType.roundRect, { x: 0.72, y: 5.25, w: 2.1, h: 0.62, rectRadius: 0.08, fill: { color: C.blue }, line: { color: C.blue } });
  s.addText('项目展示 · 2026', { x: 0.72, y: 5.46, w: 2.1, h: 0.18, fontFace: FONT, fontSize: 10, color: C.white, bold: true, align: 'center', margin: 0 });
  // Decorative dashboard
  s.addShape(pptx.ShapeType.roundRect, { x: 7.45, y: 1.1, w: 4.8, h: 4.95, rectRadius: 0.14, fill: { color: C.white }, line: { color: 'D9E2F4', width: 1.2 }, shadow: { type: 'outer', color: '7F91B6', opacity: 0.16, blur: 3, angle: 45, distance: 2 } });
  s.addShape(pptx.ShapeType.roundRect, { x: 7.72, y: 1.4, w: 1.16, h: 0.36, rectRadius: 0.06, fill: { color: C.pale }, line: { color: C.pale } });
  s.addText('本周学习', { x: 7.86, y: 1.51, w: 0.86, h: 0.13, fontFace: FONT, fontSize: 8, bold: true, color: C.blue, margin: 0 });
  s.addText('从一小段学习开始', { x: 7.75, y: 2.05, w: 2.7, h: 0.3, fontFace: FONT, fontSize: 15, bold: true, color: C.text, margin: 0 });
  s.addText('专属老师会按你的节奏陪学', { x: 7.75, y: 2.43, w: 2.7, h: 0.22, fontFace: FONT, fontSize: 9, color: C.muted, margin: 0 });
  s.addShape(pptx.ShapeType.roundRect, { x: 7.75, y: 2.95, w: 3.95, h: 1.02, rectRadius: 0.08, fill: { color: 'F4F7FF' }, line: { color: 'E1E8F6' } });
  s.addText('林知远老师', { x: 8.0, y: 3.2, w: 1.2, h: 0.2, fontFace: FONT, fontSize: 11, bold: true, color: C.text, margin: 0 });
  s.addText('数学 · 温和启发 · 在线陪学', { x: 8.0, y: 3.5, w: 2.6, h: 0.17, fontFace: FONT, fontSize: 8.5, color: C.muted, margin: 0 });
  s.addShape(pptx.ShapeType.roundRect, { x: 10.48, y: 3.18, w: 0.95, h: 0.34, rectRadius: 0.06, fill: { color: C.blue }, line: { color: C.blue } });
  s.addText('开始学习', { x: 10.48, y: 3.29, w: 0.95, h: 0.12, fontFace: FONT, fontSize: 7.5, color: C.white, bold: true, align: 'center', margin: 0 });
  s.addText('●  今日任务  2 / 5 已完成', { x: 7.75, y: 4.45, w: 2.65, h: 0.2, fontFace: FONT, fontSize: 10, color: C.green, bold: true, margin: 0 });
  s.addShape(pptx.ShapeType.roundRect, { x: 7.75, y: 4.83, w: 3.75, h: 0.1, rectRadius: 0.05, fill: { color: 'E4EAF7' }, line: { color: 'E4EAF7' } });
  s.addShape(pptx.ShapeType.roundRect, { x: 7.75, y: 4.83, w: 1.5, h: 0.1, rectRadius: 0.05, fill: { color: '5AC8A7' }, line: { color: '5AC8A7' } });
}

// 2. Problem
{
  const s = pptx.addSlide('MASTER');
  addTitle(s, '为什么需要造师台？', '把“找不到合适学习方式”的困扰，转化为可持续的个性化陪学体验。');
  addCard(s, { x: 0.68, y: 1.55, w: 3.85, h: 2.05, title: '学习路径不够清晰', body: '学生常常不知道从哪里开始、当前学到哪一步，也难以看到短期进展。', accent: C.blue, icon: '01' });
  addCard(s, { x: 4.75, y: 1.55, w: 3.85, h: 2.05, title: '学习支持不够贴合', body: '统一节奏与统一表达，难以同时照顾学科、基础、教材和个人偏好。', accent: C.indigo, icon: '02' });
  addCard(s, { x: 8.82, y: 1.55, w: 3.85, h: 2.05, title: '压力与隐私被忽略', body: '疲惫、抗拒或害怕被批评时，学生需要被理解，也需要安全边界。', accent: C.green, icon: '03' });
  s.addShape(pptx.ShapeType.roundRect, { x: 0.68, y: 4.25, w: 12.0, h: 1.55, rectRadius: 0.1, fill: { color: 'EEF3FF' }, line: { color: 'D9E4FF' } });
  s.addText('造师台的回答', { x: 1.0, y: 4.6, w: 1.8, h: 0.28, fontFace: FONT, fontSize: 17, bold: true, color: C.blue, margin: 0 });
  s.addText('通过“专属老师 + 定制方案 + 学习路径 + 安全对话”，让每位学习者拥有可开始、可坚持、可回顾的个人学习空间。', { x: 2.8, y: 4.52, w: 8.95, h: 0.47, fontFace: FONT, fontSize: 15, color: C.text, margin: 0, valign: 'mid', fit: 'shrink' });
}

// 3. Product flow
{
  const s = pptx.addSlide('MASTER');
  addTitle(s, '产品核心：一条清晰、可执行的学习主线', '从第一次进入，到完成一节课程，用户始终知道下一步该做什么。');
  const steps = [
    ['1', '选择老师', '学科、风格、陪学设定'],
    ['2', '定制计划', '年级、教材、频率与目标'],
    ['3', '开始对话', '讲解、提问、练习、复盘'],
    ['4', '自动更新', '任务、路径、成长记录'],
  ];
  steps.forEach((step, i) => {
    const x = 0.76 + i * 3.1;
    s.addShape(pptx.ShapeType.roundRect, { x, y: 2.12, w: 2.52, h: 2.34, rectRadius: 0.1, fill: { color: C.white }, line: { color: C.line } });
    s.addShape(pptx.ShapeType.ellipse, { x: x + 0.2, y: 2.34, w: 0.5, h: 0.5, fill: { color: i === 2 ? C.indigo : C.blue }, line: { color: i === 2 ? C.indigo : C.blue } });
    s.addText(step[0], { x: x + 0.2, y: 2.49, w: 0.5, h: 0.14, fontFace: FONT, fontSize: 10, bold: true, color: C.white, align: 'center', margin: 0 });
    s.addText(step[1], { x: x + 0.2, y: 3.08, w: 2.0, h: 0.26, fontFace: FONT, fontSize: 15, bold: true, color: C.text, margin: 0 });
    s.addText(step[2], { x: x + 0.2, y: 3.5, w: 2.05, h: 0.42, fontFace: FONT, fontSize: 10, color: C.muted, margin: 0, fit: 'shrink' });
    if (i < 3) addArrow(s, x + 2.62, 3.08, 0.31);
  });
  s.addShape(pptx.ShapeType.roundRect, { x: 1.26, y: 5.27, w: 10.85, h: 0.58, rectRadius: 0.08, fill: { color: 'EFFAF6' }, line: { color: 'D3EDE3' } });
  s.addText('学习任务会随课程完成自动点亮；学习路径同步显示“当前阶段 / 已完成 / 下一步”。', { x: 1.26, y: 5.46, w: 10.85, h: 0.17, fontFace: FONT, fontSize: 11, color: '27846B', bold: true, align: 'center', margin: 0 });
}

// 4. Feature matrix
{
  const s = pptx.addSlide('MASTER');
  addTitle(s, '核心功能模块', '覆盖创建、学习、陪伴、记录与个性化设置的完整闭环。');
  const cards = [
    ['专属老师', '创建、选择、搜索、排序与折叠管理；不同老师拥有独立对话与学习设定。', C.blue, '人'],
    ['学习方案', '设置年级、教材、每日时长、每周频率、复盘方式、提醒与朗读风格。', C.indigo, '计'],
    ['学习对话', '支持智能回复、历史加载、复制、朗读、停止回答、草稿恢复与快捷提问。', C.green, '聊'],
    ['学习成长', '本周任务自动点亮、专注计时、签到、成就体系与学习报告。', C.orange, '长'],
    ['低压力陪学', '隐藏积分与连胜压力，允许暂停或跳过，提供 5 分钟微任务。', '3F9B88', '柔'],
    ['本机数据管理', '按昵称隔离资料；支持本机导入、导出、清除与隐私退出。', '7887A8', '盾'],
  ];
  cards.forEach((card, i) => addCard(s, { x: 0.7 + (i % 3) * 4.15, y: 1.52 + Math.floor(i / 3) * 2.25, w: 3.85, h: 1.82, title: card[0], body: card[1], accent: card[2], icon: card[3] }));
}

// 5. AI conversation
{
  const s = pptx.addSlide('MASTER');
  addTitle(s, 'AI 对话：像一位真实、温和的学习伙伴', '服务端组合学习上下文与安全提示词，向模型发起受控请求。');
  s.addShape(pptx.ShapeType.roundRect, { x: 0.78, y: 1.55, w: 5.15, h: 4.6, rectRadius: 0.12, fill: { color: C.white }, line: { color: C.line } });
  s.addText('对话体验设计', { x: 1.05, y: 1.88, w: 2, h: 0.28, fontFace: FONT, fontSize: 16, bold: true, color: C.text, margin: 0 });
  addBulletList(s, ['根据老师风格、学科、年级、教材、目标与陪学模式调整表达。', '优先给自然回应、短段落、例子与可选练习，避免一次塞入过多步骤。', '支持“停止回答”、自动/手动朗读、复制内容、回到最新消息。', '每位老师独立保存聊天记录与未发送草稿。'], 1.05, 2.45, 4.45, 2.75, C.text, 11.5);
  s.addShape(pptx.ShapeType.roundRect, { x: 6.45, y: 1.55, w: 5.92, h: 4.6, rectRadius: 0.12, fill: { color: 'F9FBFF' }, line: { color: 'DBE5F7' } });
  s.addText('一次请求的学习上下文', { x: 6.78, y: 1.88, w: 2.8, h: 0.28, fontFace: FONT, fontSize: 16, bold: true, color: C.text, margin: 0 });
  const tags = ['老师名称与设定', '学科 / 年级', '教材 / 学习目标', '学习时长与频率', '语言与表达风格', '低压力陪学状态'];
  tags.forEach((tag, i) => addPill(s, tag, 6.8 + (i % 2) * 2.65, 2.5 + Math.floor(i / 2) * 0.66, 2.3, i % 2 ? C.indigo : C.blue, i % 2 ? 'F0EDFF' : C.pale));
  s.addShape(pptx.ShapeType.roundRect, { x: 6.8, y: 4.75, w: 4.9, h: 0.72, rectRadius: 0.08, fill: { color: 'EAF9F4' }, line: { color: 'D1EEE3' } });
  s.addText('输出：短小、准确、可执行的个性化学习支持', { x: 6.99, y: 5.02, w: 4.52, h: 0.16, fontFace: FONT, fontSize: 10.5, bold: true, color: '27846B', align: 'center', margin: 0 });
}

// 6. Support and safety
{
  const s = pptx.addSlide('MASTER');
  addTitle(s, '不只追求完成，更关注学生的感受与安全', '将“低压力陪学”和情绪支持内置为产品能力，而不是额外负担。');
  s.addShape(pptx.ShapeType.roundRect, { x: 0.72, y: 1.55, w: 5.75, h: 4.62, rectRadius: 0.12, fill: { color: 'F0FBF7' }, line: { color: 'D0ECE2' } });
  s.addText('低压力陪学模式', { x: 1.05, y: 1.9, w: 2.4, h: 0.3, fontFace: FONT, fontSize: 18, bold: true, color: '21846B', margin: 0 });
  addBulletList(s, ['隐藏积分、连胜和完成压力，减少比较与催促。', '老师可切换为更温和的“学习伙伴”。', '允许暂停、跳过，并提供可立即开始的 5 分钟微任务。', '页面视觉和提示语同步转为更舒缓的节奏。'], 1.05, 2.45, 4.9, 2.65, '396B62', 12);
  s.addShape(pptx.ShapeType.roundRect, { x: 6.85, y: 1.55, w: 5.75, h: 4.62, rectRadius: 0.12, fill: { color: 'FFF9F1' }, line: { color: 'F1DEC3' } });
  s.addText('情绪与安全支持', { x: 7.18, y: 1.9, w: 2.4, h: 0.3, fontFace: FONT, fontSize: 18, bold: true, color: 'A46B2C', margin: 0 });
  addBulletList(s, ['为“抗拒学习、很疲惫、害怕被批评、想缓一缓”提供直接入口。', 'AI 优先共情，不强迫说明原因，也不以积分或排名施压。', '遇到自伤、欺凌、严重恐惧上学等风险时，提示联系可信赖成人或专业支持。', '不替代专业医疗、心理或危机干预。'], 7.18, 2.45, 4.88, 2.65, '765D3B', 12);
}

// 7. Architecture
{
  const s = pptx.addSlide('MASTER');
  addTitle(s, '系统架构：前端轻量，敏感能力放在服务端', 'API Key 不进入浏览器；前端通过本机服务访问聊天、语音与状态接口。');
  const nodes = [
    { x: 0.72, title: '浏览器前端', detail: 'index.html\n界面、状态、localStorage', color: C.blue },
    { x: 4.15, title: 'Node.js 服务', detail: 'server.js\n校验、限流、安全响应头', color: C.indigo },
    { x: 7.58, title: 'AI / 语音服务', detail: '聊天模型 API\nElevenLabs 语音 API', color: C.green },
    { x: 11.01, title: '本机数据', detail: '按昵称隔离\n导入 / 导出 / 清除', color: C.orange },
  ];
  nodes.forEach((node, i) => {
    s.addShape(pptx.ShapeType.roundRect, { x: node.x, y: 2.27, w: 1.62, h: 1.52, rectRadius: 0.09, fill: { color: C.white }, line: { color: node.color, width: 1.3 } });
    s.addShape(pptx.ShapeType.rect, { x: node.x, y: 2.27, w: 1.62, h: 0.1, fill: { color: node.color }, line: { color: node.color } });
    s.addText(node.title, { x: node.x + 0.12, y: 2.63, w: 1.38, h: 0.28, fontFace: FONT, fontSize: 12, bold: true, color: C.text, align: 'center', margin: 0, fit: 'shrink' });
    s.addText(node.detail, { x: node.x + 0.12, y: 3.1, w: 1.38, h: 0.38, fontFace: FONT, fontSize: 8.6, color: C.muted, align: 'center', margin: 0, fit: 'shrink' });
    if (i < 3) addArrow(s, node.x + 1.85, 2.85, 0.42);
  });
  s.addShape(pptx.ShapeType.roundRect, { x: 1.18, y: 4.7, w: 11.0, h: 0.9, rectRadius: 0.09, fill: { color: 'EEF3FF' }, line: { color: 'D8E3FF' } });
  s.addText('关键原则：  前端只发送结构化学习信息  ·  服务端验证参数  ·  环境变量保存密钥  ·  单 IP 限流  ·  响应安全头', { x: 1.38, y: 5.03, w: 10.6, h: 0.19, fontFace: FONT, fontSize: 11.2, bold: true, color: C.blue, align: 'center', margin: 0, fit: 'shrink' });
}

// 8. Security and privacy
{
  const s = pptx.addSlide('MASTER');
  addTitle(s, '安全与隐私设计', '对学生而言，安全不是说明书，而应当是产品默认行为。');
  const entries = [
    ['敏感信息保护', '聊天框明确提示：不输入密码、验证码、证件、银行卡、住址等敏感内容；前端会拦截可疑输入。', C.red],
    ['提示词注入防护', '自定义老师与任务内容按纯文本处理；服务端清洗字段、限制长度，并固定安全系统提示。', C.indigo],
    ['数据最小化', '不要求学生透露学校、班级、教师姓名等信息；聊天与设置保存在本机浏览器。', C.green],
    ['可控的数据权利', '按昵称隔离本机资料，支持导出备份、导入恢复、清除当前档案与隐私退出。', C.blue],
  ];
  entries.forEach((item, i) => addCard(s, { x: 0.85 + (i % 2) * 6.0, y: 1.58 + Math.floor(i / 2) * 2.15, w: 5.63, h: 1.72, title: item[0], body: item[1], accent: item[2], icon: '✓' }));
}

// 9. UX
{
  const s = pptx.addSlide('MASTER');
  addTitle(s, '体验优化：让复杂能力保持简单', '围绕首次使用、长页面、移动端和辅助访问，持续降低使用门槛。');
  const points = [
    ['快速开始', '默认推荐老师与标准方案，减少首次登录前的配置成本。'],
    ['今日学习控制台', '根据是否选老师、任务完成情况，直接展示“下一步”。'],
    ['清晰学习路径', '用“当前阶段 / 已完成 / 下一步”替代静态步骤展示。'],
    ['响应式布局', '覆盖电脑、平板与手机，避免老师卡片、工具栏和按钮重叠。'],
    ['键盘与动效', '支持 Ctrl + Enter 发送、Alt + L 打开聊天、明显焦点样式和减少动画设置。'],
  ];
  points.forEach((point, i) => {
    const y = 1.5 + i * 0.91;
    s.addShape(pptx.ShapeType.ellipse, { x: 0.9, y: y + 0.05, w: 0.37, h: 0.37, fill: { color: i === 1 ? C.green : C.blue }, line: { color: i === 1 ? C.green : C.blue } });
    s.addText(String(i + 1), { x: 0.9, y: y + 0.15, w: 0.37, h: 0.1, fontFace: FONT, fontSize: 8, bold: true, color: C.white, align: 'center', margin: 0 });
    s.addText(point[0], { x: 1.48, y, w: 1.85, h: 0.25, fontFace: FONT, fontSize: 14, bold: true, color: C.text, margin: 0 });
    s.addText(point[1], { x: 3.38, y: y + 0.02, w: 7.95, h: 0.32, fontFace: FONT, fontSize: 11, color: C.muted, margin: 0, fit: 'shrink' });
    if (i < points.length - 1) s.addShape(pptx.ShapeType.line, { x: 1.47, y: y + 0.61, w: 9.95, h: 0, line: { color: 'E1E7F2', width: 1 } });
  });
}

// 10. LAN deployment
{
  const s = pptx.addSlide('MASTER');
  addTitle(s, '局域网部署：同一网络即可共同使用', '通过本机 Node.js 服务提供页面和 AI 接口，其他设备无需配置 API Key。');
  const deploy = [
    ['1', '启动服务', 'npm start\n服务监听 0.0.0.0:3000'],
    ['2', '放行网络', 'Windows 防火墙允许\nTCP 3000 入站访问'],
    ['3', '分享地址', '同一 Wi‑Fi 访问\nhttp://主机IP:3000'],
    ['4', '安全使用', '主机持续运行\n密钥始终只保存在主机 .env'],
  ];
  deploy.forEach((item, i) => {
    const x = 0.75 + i * 3.1;
    s.addShape(pptx.ShapeType.roundRect, { x, y: 2.0, w: 2.5, h: 2.45, rectRadius: 0.1, fill: { color: C.white }, line: { color: C.line } });
    s.addShape(pptx.ShapeType.ellipse, { x: x + 0.95, y: 2.28, w: 0.6, h: 0.6, fill: { color: i === 2 ? C.green : C.blue }, line: { color: i === 2 ? C.green : C.blue } });
    s.addText(item[0], { x: x + 0.95, y: 2.49, w: 0.6, h: 0.12, fontFace: FONT, fontSize: 10, bold: true, color: C.white, align: 'center', margin: 0 });
    s.addText(item[1], { x: x + 0.25, y: 3.15, w: 2.0, h: 0.24, fontFace: FONT, fontSize: 14, bold: true, color: C.text, align: 'center', margin: 0 });
    s.addText(item[2], { x: x + 0.25, y: 3.62, w: 2.0, h: 0.38, fontFace: FONT, fontSize: 9.7, color: C.muted, align: 'center', margin: 0, fit: 'shrink' });
  });
  s.addShape(pptx.ShapeType.roundRect, { x: 1.37, y: 5.32, w: 10.6, h: 0.55, rectRadius: 0.08, fill: { color: 'FFF7E9' }, line: { color: 'F0D9A9' } });
  s.addText('注意：应分享访问地址而不是单独发送 index.html 文件；后者无法连接运行在主机上的 AI 服务。', { x: 1.37, y: 5.51, w: 10.6, h: 0.15, fontFace: FONT, fontSize: 10.5, bold: true, color: '91672A', align: 'center', margin: 0 });
}

// 11. Roadmap
{
  const s = pptx.addSlide('MASTER');
  addTitle(s, '下一步规划', '在稳固的学习闭环基础上，逐步扩展结构化内容与协作能力。');
  const roadmap = [
    ['近期', '内容更结构化', '课程库、教材章节、在线测验、错题本与复习建议。', C.blue],
    ['中期', '学习洞察更可见', '学习热力图、进步趋势、知识掌握雷达与阶段报告。', C.indigo],
    ['长期', '协作学习更丰富', '家长模式、教师任务发布、学习小组与多模态提问。', C.green],
  ];
  roadmap.forEach((item, i) => {
    const x = 0.82 + i * 4.15;
    s.addShape(pptx.ShapeType.roundRect, { x, y: 1.82, w: 3.72, h: 3.5, rectRadius: 0.12, fill: { color: C.white }, line: { color: C.line } });
    addPill(s, item[0], x + 0.28, 2.13, 0.72, item[3], item[3] === C.green ? 'EAF9F4' : item[3] === C.indigo ? 'F0EDFF' : 'EAF0FF');
    s.addText(item[1], { x: x + 0.28, y: 2.74, w: 3.0, h: 0.34, fontFace: FONT, fontSize: 18, bold: true, color: C.text, margin: 0 });
    s.addText(item[2], { x: x + 0.28, y: 3.42, w: 3.02, h: 0.78, fontFace: FONT, fontSize: 11.3, color: C.muted, margin: 0, breakLine: false, fit: 'shrink' });
    s.addShape(pptx.ShapeType.line, { x: x + 0.28, y: 4.75, w: 2.95, h: 0, line: { color: item[3], width: 2.5 } });
  });
}

// 12. Closing
{
  const s = pptx.addSlide('MASTER');
  s.background = { color: 'F2F5FF' };
  s.addShape(pptx.ShapeType.ellipse, { x: 9.85, y: 0.75, w: 3.2, h: 3.2, fill: { color: 'DDE6FF', transparency: 16 }, line: { color: 'DDE6FF' } });
  s.addShape(pptx.ShapeType.ellipse, { x: 10.95, y: 4.38, w: 1.8, h: 1.8, fill: { color: 'DDF5EC', transparency: 6 }, line: { color: 'DDF5EC' } });
  addPill(s, 'PROJECT SUMMARY', 0.78, 1.28, 1.55, C.blue, 'E3EAFF');
  s.addText('造师台', { x: 0.76, y: 1.95, w: 4.0, h: 0.6, fontFace: FONT, fontSize: 34, bold: true, color: C.navy, margin: 0 });
  s.addText('不仅是一个 AI 聊天网页，\n更是一套尊重节奏、守住边界的学习陪伴体验。', { x: 0.78, y: 2.85, w: 7.2, h: 0.86, fontFace: FONT, fontSize: 20, bold: true, color: C.blue, margin: 0, breakLine: false });
  s.addText('专属老师 · 清晰路径 · 温和陪学 · 本机安全', { x: 0.8, y: 4.35, w: 5.8, h: 0.28, fontFace: FONT, fontSize: 13, color: C.muted, margin: 0 });
  s.addShape(pptx.ShapeType.roundRect, { x: 0.78, y: 5.32, w: 2.05, h: 0.58, rectRadius: 0.08, fill: { color: C.blue }, line: { color: C.blue } });
  s.addText('谢谢观看', { x: 0.78, y: 5.52, w: 2.05, h: 0.16, fontFace: FONT, fontSize: 12, bold: true, color: C.white, align: 'center', margin: 0 });
}

pptx.writeFile({ fileName: '造师台-项目介绍.pptx' });
