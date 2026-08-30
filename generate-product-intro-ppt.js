const pptxgen = require('pptxgenjs');

const pptx = new pptxgen();
pptx.defineLayout({ name: 'WIDE', width: 13.333, height: 7.5 });
pptx.layout = 'WIDE';
pptx.author = '造师台项目组';
pptx.company = '造师台';
pptx.subject = '产品介绍演示稿（配套三分钟字幕）';
pptx.title = '造师台产品介绍';
pptx.lang = 'zh-CN';
pptx.theme = { headFontFace: 'Microsoft YaHei', bodyFontFace: 'Microsoft YaHei', lang: 'zh-CN' };

const C = {
  bg: 'F7F9FF', white: 'FFFFFF', navy: '172A55', text: '304361', muted: '6E7E99',
  blue: '4364F5', purple: '7657DE', green: '28AA87', mint: 'E9F9F3', pale: 'EBEFFF',
  orange: 'F39845', peach: 'FFF0E3', line: 'DDE5F2', dark: '202E73'
};
const font = 'Microsoft YaHei';

function slide() {
  const s = pptx.addSlide();
  s.background = { color: C.bg };
  return s;
}
function addHeader(s, index, heading, sub, time) {
  s.addText(`造师台 · 产品介绍`, { x: 0.62, y: 0.32, w: 2.2, h: 0.18, fontFace: font, fontSize: 9, bold: true, color: C.blue, margin: 0 });
  s.addText(heading, { x: 0.62, y: 0.68, w: 8.9, h: 0.42, fontFace: font, fontSize: 25, bold: true, color: C.navy, margin: 0 });
  s.addText(sub, { x: 0.64, y: 1.18, w: 10.6, h: 0.19, fontFace: font, fontSize: 10.5, color: C.muted, margin: 0, fit: 'shrink' });
  s.addShape(pptx.ShapeType.roundRect, { x: 11.52, y: 0.59, w: 1.2, h: 0.39, rectRadius: 0.08, fill: { color: C.pale }, line: { color: C.pale } });
  s.addText(time, { x: 11.52, y: 0.72, w: 1.2, h: 0.1, fontFace: font, fontSize: 8.5, bold: true, color: C.blue, align: 'center', margin: 0 });
  s.addShape(pptx.ShapeType.line, { x: 0.62, y: 6.89, w: 12.1, h: 0, line: { color: C.line, width: 1 } });
  s.addText('造师台 · AI LEARNING STUDIO', { x: 0.62, y: 7.06, w: 2.6, h: 0.12, fontFace: font, fontSize: 8, color: '8492AA', margin: 0 });
  s.addText(String(index), { x: 12.38, y: 7.06, w: 0.25, h: 0.12, fontFace: font, fontSize: 8, color: '8492AA', align: 'right', margin: 0 });
}
function addCard(s, x, y, w, h, title, body, accent = C.blue, icon = '') {
  s.addShape(pptx.ShapeType.roundRect, { x, y, w, h, rectRadius: 0.08, fill: { color: C.white }, line: { color: C.line } });
  s.addShape(pptx.ShapeType.roundRect, { x, y, w: 0.07, h, rectRadius: 0.03, fill: { color: accent }, line: { color: accent } });
  if (icon) s.addText(icon, { x: x + 0.28, y: y + 0.28, w: 0.38, h: 0.3, fontFace: font, fontSize: 21, margin: 0, align: 'center' });
  s.addText(title, { x: x + (icon ? 0.78 : 0.28), y: y + 0.32, w: w - (icon ? 1.0 : 0.55), h: 0.22, fontFace: font, fontSize: 14, bold: true, color: C.text, margin: 0, fit: 'shrink' });
  s.addText(body, { x: x + 0.28, y: y + 0.82, w: w - 0.56, h: h - 1.02, fontFace: font, fontSize: 11, color: C.muted, margin: 0, breakLine: false, fit: 'shrink', valign: 'top' });
}
function addNotes(s, lines) {
  s.addNotes(lines.join('\n'));
}
function bulletText(s, items, x, y, w, h, size = 14) {
  const runs = items.map((text, i) => ({ text, options: { bullet: { indent: 16 }, hanging: 3, breakLine: i !== items.length - 1 } }));
  s.addText(runs, { x, y, w, h, fontFace: font, fontSize: size, color: C.text, margin: 0, paraSpaceAfterPt: 12, breakLine: false, fit: 'shrink' });
}

// 1 · 00:00–00:16
{
  const s = slide();
  s.addShape(pptx.ShapeType.ellipse, { x: 9.35, y: -0.86, w: 4.9, h: 4.9, fill: { color: 'DFE6FF', transparency: 15 }, line: { color: 'DFE6FF', transparency: 100 } });
  s.addShape(pptx.ShapeType.ellipse, { x: 10.7, y: 4.55, w: 2.0, h: 2.0, fill: { color: 'DDF8EF', transparency: 8 }, line: { color: 'DDF8EF', transparency: 100 } });
  s.addShape(pptx.ShapeType.roundRect, { x: 0.76, y: 1.02, w: 1.45, h: 0.42, rectRadius: 0.12, fill: { color: C.blue }, line: { color: C.blue } });
  s.addText('AI 学习平台', { x: 0.76, y: 1.16, w: 1.45, h: 0.11, fontFace: font, fontSize: 9.5, bold: true, color: C.white, align: 'center', margin: 0 });
  s.addText('造师台', { x: 0.75, y: 1.85, w: 3.8, h: 0.62, fontFace: font, fontSize: 42, bold: true, color: C.navy, margin: 0 });
  s.addText('让每一段学习，都有懂你的老师', { x: 0.78, y: 2.7, w: 6.3, h: 0.36, fontFace: font, fontSize: 23, bold: true, color: C.blue, margin: 0 });
  s.addText('当学习缺少清晰方向、合适节奏与及时回应，努力也容易被焦虑消耗。', { x: 0.8, y: 3.48, w: 5.8, h: 0.35, fontFace: font, fontSize: 14, color: C.muted, margin: 0, fit: 'shrink' });
  s.addShape(pptx.ShapeType.roundRect, { x: 7.6, y: 1.25, w: 4.35, h: 4.7, rectRadius: 0.14, fill: { color: C.white }, line: { color: 'D9E2F4' } });
  s.addText('👩🏻‍🏫', { x: 8.08, y: 1.65, w: 0.7, h: 0.6, fontFace: font, fontSize: 44, margin: 0 });
  s.addText('今天，从一小段学习开始', { x: 8.05, y: 2.55, w: 3.2, h: 0.3, fontFace: font, fontSize: 17, bold: true, color: C.text, margin: 0 });
  s.addText('专属老师 · 清晰任务 · 温和陪伴', { x: 8.05, y: 3.03, w: 3.2, h: 0.18, fontFace: font, fontSize: 10.5, color: C.muted, margin: 0 });
  s.addShape(pptx.ShapeType.roundRect, { x: 8.05, y: 3.75, w: 3.45, h: 0.68, rectRadius: 0.1, fill: { color: C.pale }, line: { color: 'DDE5FF' } });
  s.addText('本周任务  ·  2 / 5', { x: 8.3, y: 4.02, w: 2.2, h: 0.12, fontFace: font, fontSize: 12, bold: true, color: C.green, margin: 0 });
  s.addText('产品介绍演示 · 3 分 41 秒', { x: 0.8, y: 5.65, w: 3.1, h: 0.16, fontFace: font, fontSize: 10.5, bold: true, color: C.purple, margin: 0 });
  addNotes(s, ['大家好，这里是造师台。', '我们希望让每一段学习，都有一位真正懂你的老师陪在身边。', '很多时候，学习并不是因为不够努力，而是缺少清晰的方向、合适的节奏和及时的回应。']);
}

// 2 · audiences and pain points
{
  const s = slide(); addHeader(s, 2, '我们服务哪些人？', '核心用户是学生与自主学习者，同时支持家长参与，并为学校和培训机构预留服务能力。', '00:16–00:33');
  addCard(s, 0.8, 1.75, 3.7, 3.55, '学生与自主学习者', '面对任务堆积或陌生知识点时，需要一个能把目标拆小、能随时回应、不会催促自己的学习伙伴。', C.blue, '🎒');
  addCard(s, 4.82, 1.75, 3.7, 3.55, '关注成长的家长', '希望了解学习是否在稳步推进，也重视孩子的情绪状态、学习节奏和数字隐私。', C.green, '🏡');
  addCard(s, 8.84, 1.75, 3.7, 3.55, '学校与培训机构（后续）', '可作为学习任务发布、学习进度汇总与个性化陪学能力的补充工具。', C.purple, '🏫');
  s.addShape(pptx.ShapeType.roundRect, { x: 1.45, y: 5.75, w: 10.4, h: 0.52, rectRadius: 0.1, fill: { color: C.peach }, line: { color: 'F8DFC1' } });
  s.addText('共同痛点：不知道从哪里开始、讲解不够贴合、压力影响坚持、学习记录分散。', { x: 1.45, y: 5.93, w: 10.4, h: 0.12, fontFace: font, fontSize: 12, bold: true, color: 'A65E22', align: 'center', margin: 0 });
  addNotes(s, ['面对堆积的任务和陌生的知识点，学生容易不知道从哪里开始，也容易在压力中失去信心。', '我们的核心用户是中小学生与自主学习者，也服务关注孩子成长的家长，并面向学校和培训机构持续拓展。']);
}

// 3 · product definition
{
  const s = slide(); addHeader(s, 3, '不只是聊天工具，而是学习闭环', '老师、计划、任务、复习与记录被连接在同一处，帮助用户从开始走向掌握。', '00:33–00:42');
  const items = [
    ['专属老师', '设定形象、学科与沟通方式', C.blue, '👩🏻‍🏫'],
    ['学习计划', '将目标拆成可执行的今日任务', C.green, '🗺️'],
    ['任务对话', '围绕当前任务解释、举例与引导', C.orange, '💬'],
    ['成长记录', '笔记、复习、成就与学习记录沉淀', C.purple, '🌱']
  ];
  items.forEach((item, i) => addCard(s, 0.82 + i * 3.08, 2.0, 2.7, 2.65, item[0], item[1], item[2], item[3]));
  s.addShape(pptx.ShapeType.roundRect, { x: 1.17, y: 5.45, w: 10.9, h: 0.65, rectRadius: 0.1, fill: { color: C.dark }, line: { color: C.dark } });
  s.addText('造师台 = 懂你的老师 × 可执行的学习路径 × 持续的成长反馈', { x: 1.17, y: 5.7, w: 10.9, h: 0.16, fontFace: font, fontSize: 15, bold: true, color: C.white, align: 'center', margin: 0 });
  addNotes(s, ['造师台不是只回答问题的聊天工具，它是一个将老师、计划、任务、复习与记录连接起来的学习空间。']);
}

// 4 · teacher customisation
{
  const s = slide(); addHeader(s, 4, '一位适合自己的专属老师', '让学习者从关系感开始建立学习意愿，再以合适的讲解风格持续陪伴。', '00:42–01:15');
  s.addShape(pptx.ShapeType.roundRect, { x: 0.86, y: 1.65, w: 4.35, h: 4.55, rectRadius: 0.15, fill: { color: 'EEF1FF' }, line: { color: 'DCE3FF' } });
  s.addShape(pptx.ShapeType.ellipse, { x: 2.18, y: 2.12, w: 1.65, h: 1.65, fill: { color: 'FFE8D2' }, line: { color: 'FFE8D2' } });
  s.addText('🧑🏻‍🏫', { x: 2.37, y: 2.37, w: 1.3, h: 1.1, fontFace: font, fontSize: 72, margin: 0, align: 'center' });
  s.addText('林知远老师', { x: 1.25, y: 4.2, w: 3.55, h: 0.26, fontFace: font, fontSize: 18, bold: true, color: C.navy, align: 'center', margin: 0 });
  s.addText('数学 · 温和启发型 · 在线陪学', { x: 1.25, y: 4.64, w: 3.55, h: 0.16, fontFace: font, fontSize: 10.5, color: C.muted, align: 'center', margin: 0 });
  s.addShape(pptx.ShapeType.roundRect, { x: 1.58, y: 5.18, w: 2.9, h: 0.44, rectRadius: 0.08, fill: { color: C.blue }, line: { color: C.blue } });
  s.addText('开始一起学习', { x: 1.58, y: 5.33, w: 2.9, h: 0.1, fontFace: font, fontSize: 10, bold: true, color: C.white, align: 'center', margin: 0 });
  addCard(s, 5.68, 1.65, 3.05, 1.35, '可选择形象', '可爱人型或动物形象，让老师更有亲近感。', C.blue, '🎨');
  addCard(s, 9.08, 1.65, 3.05, 1.35, '可设定语气', '清晰直接、温和陪伴、活力引导等方式可选。', C.green, '🗣️');
  addCard(s, 5.68, 3.42, 3.05, 1.35, '可聚焦学科', '让老师围绕当前的学习主题、教材与目标陪学。', C.orange, '📚');
  addCard(s, 9.08, 3.42, 3.05, 1.35, '可持续对话', '选择老师后，学习任务和对话体验会持续围绕该老师展开。', C.purple, '✨');
  addNotes(s, ['打开平台后，你可以快速开始，也可以创建一位属于自己的专属老师。', '老师可以选择可爱的人型或动物形象，并根据自己的喜好设定名字、学科和沟通方式。', '有人喜欢清晰直接的讲解，也有人更需要温和、耐心、一步一步的陪伴。', '在造师台中，老师不只是一个头像。选择老师后，学习页面和对话会围绕这位老师持续展开。']);
}

// 5 · learning plan
{
  const s = slide(); addHeader(s, 5, '把大目标拆成今天就能完成的一步', '根据每日学习时长与每周频次生成计划；首页持续呈现当前阶段和下一步。', '01:15–01:32');
  const steps = [['学习目标', '例如：本周掌握分数基础'], ['每周计划', '每周 5 次，每次 20 分钟'], ['今日任务', '完成一个知识点讲解与练习'], ['完成反馈', '自动点亮下一项本周任务']];
  steps.forEach((item, i) => {
    const x = 0.72 + i * 3.14;
    s.addShape(pptx.ShapeType.roundRect, { x, y: 2.18, w: 2.55, h: 2.18, rectRadius: 0.1, fill: { color: C.white }, line: { color: C.line } });
    s.addShape(pptx.ShapeType.ellipse, { x: x + 0.22, y: 2.44, w: 0.45, h: 0.45, fill: { color: i === 3 ? C.green : C.blue }, line: { color: i === 3 ? C.green : C.blue } });
    s.addText(String(i + 1), { x: x + 0.22, y: 2.59, w: 0.45, h: 0.1, fontFace: font, fontSize: 9, bold: true, color: C.white, align: 'center', margin: 0 });
    s.addText(item[0], { x: x + 0.25, y: 3.18, w: 1.95, h: 0.2, fontFace: font, fontSize: 14, bold: true, color: C.text, margin: 0 });
    s.addText(item[1], { x: x + 0.25, y: 3.64, w: 2.0, h: 0.33, fontFace: font, fontSize: 10.5, color: C.muted, margin: 0, fit: 'shrink' });
    if (i < 3) s.addShape(pptx.ShapeType.chevron, { x: x + 2.72, y: 3.05, w: 0.25, h: 0.3, fill: { color: 'A6B5D6' }, line: { color: 'A6B5D6' } });
  });
  s.addShape(pptx.ShapeType.roundRect, { x: 1.3, y: 5.15, w: 10.7, h: 0.62, rectRadius: 0.1, fill: { color: C.mint }, line: { color: 'D7F1E7' } });
  s.addText('清晰路径的价值：减少“我该做什么”的犹豫，把注意力放在真正的学习行动。', { x: 1.3, y: 5.38, w: 10.7, h: 0.15, fontFace: font, fontSize: 13, bold: true, color: C.green, align: 'center', margin: 0 });
  addNotes(s, ['接下来，你可以按每天的学习时长和每周频次，生成适合自己的学习计划。', '大的学习目标会被拆成较轻松的今日任务，让“我要学习”变成现在就能完成的一小步。', '首页会显示当前学习阶段、本周任务和下一步目标。完成课程后，后续任务也会自动点亮。']);
}

// 6 · task-aware chat
{
  const s = slide(); addHeader(s, 6, '围绕当前任务的智能对话', '不只给答案：先讲核心思路，再给步骤、例子、常见误区和小行动。', '01:32–01:50');
  s.addShape(pptx.ShapeType.roundRect, { x: 0.82, y: 1.72, w: 5.15, h: 4.65, rectRadius: 0.12, fill: { color: C.white }, line: { color: C.line } });
  s.addText('👩🏻‍🏫 林知远老师', { x: 1.14, y: 2.03, w: 2.4, h: 0.2, fontFace: font, fontSize: 13, bold: true, color: C.navy, margin: 0 });
  s.addShape(pptx.ShapeType.roundRect, { x: 1.15, y: 2.55, w: 3.95, h: 0.76, rectRadius: 0.1, fill: { color: 'EEF2FF' }, line: { color: 'E0E6FC' } });
  s.addText('分数相加时，为什么要先通分？', { x: 1.38, y: 2.84, w: 3.4, h: 0.12, fontFace: font, fontSize: 11.5, color: C.text, margin: 0 });
  s.addShape(pptx.ShapeType.roundRect, { x: 1.67, y: 3.6, w: 3.75, h: 1.3, rectRadius: 0.1, fill: { color: C.mint }, line: { color: 'D7F1E7' } });
  s.addText('核心思路：先让分数的单位相同。\n例如 1/2 = 2/4，所以 1/2 + 1/4 = 3/4。\n下一步：试着完成 2 道同类题。', { x: 1.94, y: 3.92, w: 3.18, h: 0.65, fontFace: font, fontSize: 10.5, color: C.text, margin: 0, breakLine: false, fit: 'shrink' });
  const points = [['任务相关', '结合正在学习的内容，而非泛泛而谈'], ['结构清晰', '核心思路 → 步骤 → 例子 → 易错点'], ['行动导向', '最后给出可完成的下一步，而非停在解释中']];
  points.forEach((p, i) => addCard(s, 6.55, 1.74 + i * 1.45, 5.55, 1.05, p[0], p[1], [C.blue, C.purple, C.green][i]));
  addNotes(s, ['在学习过程中，随时可以打开和老师的对话。提问时，AI 会结合当前任务给出更有针对性的帮助。', '它会先说明核心思路，再给出可操作的步骤、例子和常见误区，避免只给出一个让人看不懂的结论。']);
}

// 7 · low pressure
{
  const s = slide(); addHeader(s, 7, '低压力陪学：先理解，再行动', '学习者可以暂停或跳过，也可以通过五分钟微任务重新回到学习节奏。', '01:50–02:15');
  s.addShape(pptx.ShapeType.roundRect, { x: 0.82, y: 1.76, w: 5.1, h: 4.55, rectRadius: 0.12, fill: { color: 'F3F9FF' }, line: { color: 'D8E9F7' } });
  s.addText('“今天有点累，不想学。”', { x: 1.22, y: 2.2, w: 3.7, h: 0.22, fontFace: font, fontSize: 15, bold: true, color: C.text, margin: 0 });
  s.addShape(pptx.ShapeType.roundRect, { x: 1.25, y: 2.85, w: 3.9, h: 1.2, rectRadius: 0.12, fill: { color: C.white }, line: { color: 'E0E8F4' } });
  s.addText('没关系，先缓一缓也可以。\n你想暂停，还是只做一个 5 分钟的小任务？', { x: 1.52, y: 3.2, w: 3.3, h: 0.44, fontFace: font, fontSize: 11.5, color: C.text, margin: 0, breakLine: false, fit: 'shrink' });
  s.addShape(pptx.ShapeType.roundRect, { x: 1.3, y: 4.66, w: 1.62, h: 0.48, rectRadius: 0.09, fill: { color: C.pale }, line: { color: C.pale } });
  s.addText('暂停一下', { x: 1.3, y: 4.84, w: 1.62, h: 0.1, fontFace: font, fontSize: 9.5, bold: true, color: C.blue, align: 'center', margin: 0 });
  s.addShape(pptx.ShapeType.roundRect, { x: 3.18, y: 4.66, w: 1.62, h: 0.48, rectRadius: 0.09, fill: { color: C.green }, line: { color: C.green } });
  s.addText('5 分钟微任务', { x: 3.18, y: 4.84, w: 1.62, h: 0.1, fontFace: font, fontSize: 9.5, bold: true, color: C.white, align: 'center', margin: 0 });
  addCard(s, 6.55, 1.8, 5.5, 1.12, '允许暂停与跳过', '用户可以调整节奏，不把一次无法完成当作失败。', C.blue, '⏸️');
  addCard(s, 6.55, 3.2, 5.5, 1.12, '提供情绪入口', '“抗拒学习、害怕被批评、很疲惫、想缓一缓”等入口，AI 先共情、不强迫说明原因。', C.orange, '💛');
  addCard(s, 6.55, 4.6, 5.5, 1.12, '保留安全边界', '遇到自伤、欺凌或严重恐惧上学等情况，引导联系可信赖的成人或专业支持。', C.purple, '🛟');
  addNotes(s, ['如果暂时不想回答很多问题，也没有关系。你可以选择低压力陪学模式，允许暂停、跳过，或先完成五分钟微任务。', '当你觉得疲惫、害怕被批评，或者只是想缓一缓，平台会先给出理解和陪伴，而不是催促你必须立刻完成。']);
}

// 8 · notes and review
{
  const s = slide(); addHeader(s, 8, '把学习过程沉淀为可复习的内容', '本地笔记记录思考；错题与薄弱知识点按合适间隔再次出现。', '02:15–02:33');
  s.addShape(pptx.ShapeType.roundRect, { x: 0.95, y: 1.8, w: 4.85, h: 4.25, rectRadius: 0.12, fill: { color: C.white }, line: { color: C.line } });
  s.addText('📝  学习笔记', { x: 1.3, y: 2.16, w: 2.1, h: 0.24, fontFace: font, fontSize: 16, bold: true, color: C.navy, margin: 0 });
  s.addShape(pptx.ShapeType.roundRect, { x: 1.3, y: 2.7, w: 4.1, h: 2.15, rectRadius: 0.06, fill: { color: 'FFFDF6' }, line: { color: 'F2E8CC' } });
  s.addText('今天理解了：\n分数相加前要让单位相同，\n也就是先通分。\n\n下次复习：把异分母加法再练两题。', { x: 1.62, y: 3.04, w: 3.55, h: 1.3, fontFace: font, fontSize: 12, color: C.text, margin: 0, breakLine: false, fit: 'shrink' });
  s.addShape(pptx.ShapeType.roundRect, { x: 6.5, y: 1.8, w: 5.75, h: 4.25, rectRadius: 0.12, fill: { color: C.white }, line: { color: C.line } });
  s.addText('🔁  间隔重复复习', { x: 6.86, y: 2.16, w: 2.5, h: 0.24, fontFace: font, fontSize: 16, bold: true, color: C.navy, margin: 0 });
  const days = ['今天', '第 1 天', '第 3 天', '第 7 天', '第 14 天'];
  days.forEach((d, i) => { const x = 6.9 + i * 1.02; s.addShape(pptx.ShapeType.ellipse, { x, y: 3.25, w: 0.56, h: 0.56, fill: { color: i === 0 ? C.green : C.pale }, line: { color: i === 0 ? C.green : C.pale } }); s.addText(i === 0 ? '✓' : String(i + 1), { x, y: 3.44, w: 0.56, h: 0.1, fontFace: font, fontSize: 9, bold: true, color: i === 0 ? C.white : C.blue, align: 'center', margin: 0 }); s.addText(d, { x: x - 0.2, y: 4.04, w: 0.95, h: 0.12, fontFace: font, fontSize: 8.5, color: C.muted, align: 'center', margin: 0 }); if (i < 4) s.addShape(pptx.ShapeType.line, { x: x + 0.56, y: 3.53, w: 0.45, h: 0, line: { color: 'BFCBE3', width: 1.3 } }); });
  s.addText('把错题和薄弱知识点在更合适的时间带回学习中。', { x: 6.85, y: 4.82, w: 4.9, h: 0.16, fontFace: font, fontSize: 12, color: C.green, bold: true, margin: 0, align: 'center' });
  addNotes(s, ['学习完成后，可以把重点写进学习笔记。笔记支持本地保存，帮助你留下每一次思考和复盘。', '错题和需要复习的知识点，也能够按照间隔重复的节奏提醒复习，让复习出现在更合适的时间。']);
}

// 9 · growth
{
  const s = slide(); addHeader(s, 9, '让“我正在进步”变得看得见', '成长不只是一串分数。任务、笔记、复习与学习天数共同记录每一次坚持。', '02:33–02:50');
  const stats = [['本周任务', '4 / 5', C.blue], ['学习笔记', '31 次', C.green], ['连续学习', '7 天', C.purple], ['本年学习之星', '12 次', C.orange]];
  stats.forEach((item, i) => { const x = 0.84 + i * 3.08; s.addShape(pptx.ShapeType.roundRect, { x, y: 1.92, w: 2.6, h: 1.62, rectRadius: 0.12, fill: { color: C.white }, line: { color: C.line } }); s.addText(item[0], { x: x + 0.24, y: 2.25, w: 2.1, h: 0.15, fontFace: font, fontSize: 10.5, color: C.muted, margin: 0, align: 'center' }); s.addText(item[1], { x: x + 0.24, y: 2.66, w: 2.1, h: 0.32, fontFace: font, fontSize: 22, bold: true, color: item[2], margin: 0, align: 'center' }); });
  s.addShape(pptx.ShapeType.roundRect, { x: 1.25, y: 4.25, w: 4.55, h: 1.35, rectRadius: 0.1, fill: { color: 'FFF9EE' }, line: { color: 'F5E4BD' } });
  s.addText('📝  学习之星', { x: 1.62, y: 4.63, w: 2.2, h: 0.22, fontFace: font, fontSize: 15, bold: true, color: '9A6A2A', margin: 0 });
  s.addText('本月保存 31 次学习笔记获得，每月刷新。', { x: 1.62, y: 5.05, w: 3.4, h: 0.13, fontFace: font, fontSize: 10.5, color: '8B795B', margin: 0 });
  s.addShape(pptx.ShapeType.roundRect, { x: 7.0, y: 4.25, w: 4.55, h: 1.35, rectRadius: 0.1, fill: { color: 'F2EEFF' }, line: { color: 'E3D9FF' } });
  s.addText('🏅  年度之星', { x: 7.37, y: 4.63, w: 2.2, h: 0.22, fontFace: font, fontSize: 15, bold: true, color: C.purple, margin: 0 });
  s.addText('本年获得 12 次“学习之星”后解锁，每年刷新。', { x: 7.37, y: 5.05, w: 3.5, h: 0.13, fontFace: font, fontSize: 10.5, color: '74688F', margin: 0 });
  addNotes(s, ['成长不只是一串分数。学习天数、笔记、完成任务和成就，都会帮助你看见自己的坚持。', '例如，连续记录学习笔记可以获得学习之星，长期积累还会解锁年度之星等成长徽章。', '这些设计不是为了制造比较，而是想把“我正在进步”这件事变得更具体、更可感知。']);
}

// 10 · privacy
{
  const s = slide(); addHeader(s, 10, '学习数据更私密，也更可控', '不要求学生透露学校、教师姓名等敏感信息；提供本机数据管理与对话保护。', '02:50–03:16');
  const items = [['敏感信息提醒', '聊天输入区提示：不输入学校、教师姓名等个人敏感信息。', C.orange, '🔒'], ['按昵称隔离', '不同昵称的本机学习数据分开保存，减少混用。', C.blue, '👤'], ['本机数据管理', '支持导出、导入和清除本机数据，用户拥有更多控制权。', C.green, '💾'], ['安全服务端', '接口限流、校验结构化学习信息，服务密钥不进入前端。', C.purple, '🛡️']];
  items.forEach((item, i) => addCard(s, 0.86 + (i % 2) * 6.15, 1.8 + Math.floor(i / 2) * 2.05, 5.6, 1.6, item[0], item[1], item[2], item[3]));
  addNotes(s, ['在隐私方面，造师台提醒用户不要输入学校、老师姓名等敏感信息，并提供本机数据导出、导入和清除功能。', '个人学习记录按昵称隔离，聊天记录也会控制数量，让学习资料尽可能由自己掌握。']);
}

// 11 · loop
{
  const s = slide(); addHeader(s, 11, '从开始到掌握的学习闭环', '选择老师、生成计划、完成学习、复习和复盘：让每个环节都能自然衔接。', '03:16–03:33');
  const steps = [['选择老师', '找到合适的陪学方式', '👩🏻‍🏫', C.blue], ['生成计划', '明确今天与本周的任务', '🗺️', C.green], ['完成学习', '在任务对话中理解与练习', '💬', C.orange], ['复习与复盘', '笔记、错题、间隔复习与成长记录', '🌱', C.purple]];
  steps.forEach((item, i) => {
    const x = 0.7 + i * 3.16;
    s.addShape(pptx.ShapeType.roundRect, { x, y: 2.0, w: 2.68, h: 2.85, rectRadius: 0.12, fill: { color: C.white }, line: { color: C.line } });
    s.addShape(pptx.ShapeType.ellipse, { x: x + 0.9, y: 2.35, w: 0.88, h: 0.88, fill: { color: i === 0 ? 'E9EDFF' : i === 1 ? 'E2FAF4' : i === 2 ? 'FFF0E3' : 'F1EEFF' }, line: { color: 'FFFFFF', transparency: 100 } });
    s.addText(item[2], { x: x + 0.9, y: 2.55, w: 0.88, h: 0.3, fontFace: font, fontSize: 24, align: 'center', margin: 0 });
    s.addText(item[0], { x: x + 0.22, y: 3.58, w: 2.24, h: 0.21, fontFace: font, fontSize: 14, bold: true, color: C.text, margin: 0, align: 'center' });
    s.addText(item[1], { x: x + 0.25, y: 4.05, w: 2.18, h: 0.33, fontFace: font, fontSize: 10.3, color: C.muted, margin: 0, align: 'center', fit: 'shrink' });
    if (i < 3) s.addShape(pptx.ShapeType.chevron, { x: x + 2.75, y: 3.2, w: 0.26, h: 0.35, fill: { color: 'AAB9D8' }, line: { color: 'AAB9D8' } });
  });
  s.addText('好的学习工具不该增加压力，而应帮助每个人找到适合自己的节奏。', { x: 1.4, y: 5.72, w: 10.5, h: 0.22, fontFace: font, fontSize: 15, bold: true, color: C.green, align: 'center', margin: 0 });
  addNotes(s, ['从选择老师，到制定计划；从完成任务，到复习和复盘，造师台把看似分散的学习行为连接成一个完整的闭环。', '我们相信，好的学习工具不应增加压力，而应该帮每个人找到合适的节奏，慢慢建立能力和信心。']);
}

// 12 · close
{
  const s = slide();
  s.addShape(pptx.ShapeType.ellipse, { x: -1.2, y: 3.8, w: 4.2, h: 4.2, fill: { color: 'DDF5EC', transparency: 8 }, line: { color: 'DDF5EC', transparency: 100 } });
  s.addShape(pptx.ShapeType.ellipse, { x: 9.6, y: -1.1, w: 4.8, h: 4.8, fill: { color: 'E0E5FF', transparency: 6 }, line: { color: 'E0E5FF', transparency: 100 } });
  s.addShape(pptx.ShapeType.roundRect, { x: 0.76, y: 1.22, w: 1.45, h: 0.42, rectRadius: 0.12, fill: { color: C.blue }, line: { color: C.blue } });
  s.addText('造师台', { x: 0.76, y: 1.36, w: 1.45, h: 0.11, fontFace: font, fontSize: 9.5, bold: true, color: C.white, align: 'center', margin: 0 });
  s.addText('让每一段学习\n都有懂你的老师', { x: 0.75, y: 2.05, w: 6.3, h: 1.05, fontFace: font, fontSize: 37, bold: true, color: C.navy, margin: 0, breakLine: false });
  s.addText('从一次轻松的开始出发，在持续行动中建立能力和信心。', { x: 0.8, y: 3.6, w: 5.7, h: 0.22, fontFace: font, fontSize: 15, color: C.muted, margin: 0 });
  s.addShape(pptx.ShapeType.roundRect, { x: 0.8, y: 4.56, w: 2.1, h: 0.57, rectRadius: 0.12, fill: { color: C.green }, line: { color: C.green } });
  s.addText('立即体验', { x: 0.8, y: 4.77, w: 2.1, h: 0.12, fontFace: font, fontSize: 11, bold: true, color: C.white, align: 'center', margin: 0 });
  s.addShape(pptx.ShapeType.roundRect, { x: 8.35, y: 1.6, w: 3.4, h: 3.95, rectRadius: 0.16, fill: { color: C.white }, line: { color: 'D8E2F4' } });
  s.addText('🌟', { x: 9.48, y: 2.05, w: 1.1, h: 0.65, fontFace: font, fontSize: 44, align: 'center', margin: 0 });
  s.addText('选择老师\n生成计划\n完成并复盘', { x: 8.85, y: 3.0, w: 2.4, h: 1.0, fontFace: font, fontSize: 19, bold: true, color: C.text, align: 'center', margin: 0, breakLine: false });
  s.addText('谢谢观看', { x: 9.15, y: 4.74, w: 1.8, h: 0.18, fontFace: font, fontSize: 13, bold: true, color: C.blue, align: 'center', margin: 0 });
  s.addText('12', { x: 12.38, y: 7.06, w: 0.25, h: 0.12, fontFace: font, fontSize: 8, color: '8492AA', align: 'right', margin: 0 });
  addNotes(s, ['造师台，让每一段学习都有懂你的老师。', '现在，就从一次轻松的开始出发。']);
}

pptx.writeFile({ fileName: '造师台-产品介绍（3分钟字幕版）.pptx' });
