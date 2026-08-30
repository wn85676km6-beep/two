const pptxgen = require('pptxgenjs');

const pptx = new pptxgen();
pptx.defineLayout({ name: 'WIDE', width: 13.333, height: 7.5 });
pptx.layout = 'WIDE';
pptx.author = '造师台项目组';
pptx.company = '造师台';
pptx.title = '造师台路演PPT';
pptx.lang = 'zh-CN';
pptx.theme = { headFontFace: 'Microsoft YaHei', bodyFontFace: 'Microsoft YaHei', lang: 'zh-CN' };

const C = { blue: '315EFB', purple: '7158DD', navy: '26385C', text: '33415C', muted: '73819A', line: 'DFE6F3', bg: 'F7F9FD', white: 'FFFFFF', mint: 'EAF9F4', green: '20A782', orange: 'F49A47', pale: 'EDF2FF' };
const font = 'Microsoft YaHei';

function title(slide, heading, sub = '') {
  slide.addText(heading, { x: 0.68, y: 0.42, w: 8.7, h: 0.42, fontFace: font, fontSize: 25, bold: true, color: C.navy, margin: 0 });
  if (sub) slide.addText(sub, { x: 0.7, y: 0.94, w: 10.8, h: 0.2, fontFace: font, fontSize: 10.5, color: C.muted, margin: 0 });
}
function footer(slide, page) {
  slide.addShape(pptx.ShapeType.line, { x: 0.65, y: 6.95, w: 12.05, h: 0, line: { color: 'E6EBF5', width: 1 } });
  slide.addText('造师台 · AI LEARNING STUDIO', { x: 0.68, y: 7.08, w: 3.2, h: 0.12, fontFace: font, fontSize: 8, color: '8795AB', margin: 0 });
  slide.addText(String(page), { x: 12.3, y: 7.08, w: 0.35, h: 0.12, fontFace: font, fontSize: 8, color: '8795AB', align: 'right', margin: 0 });
}
function card(slide, x, y, w, h, heading, body, color = C.blue) {
  slide.addShape(pptx.ShapeType.roundRect, { x, y, w, h, rectRadius: 0.08, fill: { color: C.white }, line: { color: C.line } });
  slide.addShape(pptx.ShapeType.rect, { x, y, w: 0.07, h, fill: { color }, line: { color } });
  slide.addText(heading, { x: x + 0.25, y: y + 0.25, w: w - 0.5, h: 0.24, fontFace: font, fontSize: 14, bold: true, color: C.text, margin: 0, fit: 'shrink' });
  slide.addText(body, { x: x + 0.25, y: y + 0.73, w: w - 0.48, h: h - 0.92, fontFace: font, fontSize: 10.5, color: C.muted, margin: 0, breakLine: false, fit: 'shrink', valign: 'top' });
}
function pill(slide, text, x, y, w, color = C.blue, fill = C.pale) {
  slide.addShape(pptx.ShapeType.roundRect, { x, y, w, h: 0.32, rectRadius: 0.08, fill: { color: fill }, line: { color: fill } });
  slide.addText(text, { x, y: y + 0.085, w, h: 0.12, fontFace: font, fontSize: 8.5, bold: true, color, align: 'center', margin: 0 });
}
function bullets(slide, items, x, y, w, h, size = 12) {
  slide.addText(items.map((text, i) => ({ text, options: { bullet: { indent: 15 }, hanging: 3, breakLine: i < items.length - 1 } })), { x, y, w, h, fontFace: font, fontSize: size, color: C.text, margin: 0, paraSpaceAfterPt: 10, fit: 'shrink' });
}
function slide() { const s = pptx.addSlide(); s.background = { color: C.bg }; return s; }

// 1 Cover
{
  const s = slide();
  s.addShape(pptx.ShapeType.ellipse, { x: 9.4, y: -1.0, w: 4.6, h: 4.6, fill: { color: 'DDE6FF', transparency: 16 }, line: { color: 'DDE6FF' } });
  s.addShape(pptx.ShapeType.ellipse, { x: 10.8, y: 4.5, w: 2.0, h: 2.0, fill: { color: 'DDF5EC', transparency: 7 }, line: { color: 'DDF5EC' } });
  pill(s, 'AI LEARNING STUDIO', 0.78, 1.22, 1.72);
  s.addText('造师台', { x: 0.76, y: 1.82, w: 4.0, h: 0.65, fontFace: font, fontSize: 39, bold: true, color: C.navy, margin: 0 });
  s.addText('让每一段学习，都有懂你的老师', { x: 0.78, y: 2.75, w: 6.3, h: 0.42, fontFace: font, fontSize: 22, bold: true, color: C.blue, margin: 0 });
  s.addText('面向学生的个性化 AI 学习空间\n专属老师、学习计划与温和陪学，一站完成。', { x: 0.8, y: 3.52, w: 5.8, h: 0.72, fontFace: font, fontSize: 14, color: C.muted, margin: 0, breakLine: false });
  s.addShape(pptx.ShapeType.roundRect, { x: 0.8, y: 5.25, w: 2.08, h: 0.56, rectRadius: 0.08, fill: { color: C.blue }, line: { color: C.blue } });
  s.addText('项目路演 · 2026', { x: 0.8, y: 5.46, w: 2.08, h: 0.13, fontFace: font, fontSize: 10, bold: true, color: C.white, align: 'center', margin: 0 });
  s.addShape(pptx.ShapeType.roundRect, { x: 7.55, y: 1.24, w: 4.5, h: 4.75, rectRadius: 0.12, fill: { color: C.white }, line: { color: 'D9E2F4' }, shadow: { type: 'outer', color: '7183AA', opacity: 0.15, blur: 3, angle: 45, distance: 2 } });
  s.addText('TODAY · 学习控制台', { x: 7.85, y: 1.65, w: 2.2, h: 0.16, fontFace: font, fontSize: 9, bold: true, color: C.blue, margin: 0 });
  s.addText('下一步，从一小段学习开始', { x: 7.85, y: 2.05, w: 3.35, h: 0.27, fontFace: font, fontSize: 15, bold: true, color: C.text, margin: 0 });
  s.addText('林知远老师 · 数学 · 温和启发', { x: 7.85, y: 2.5, w: 3.1, h: 0.16, fontFace: font, fontSize: 9, color: C.muted, margin: 0 });
  s.addShape(pptx.ShapeType.roundRect, { x: 7.85, y: 3.0, w: 3.85, h: 0.78, rectRadius: 0.07, fill: { color: 'F2F6FF' }, line: { color: 'E0E8F8' } });
  s.addText('本周任务  2 / 5', { x: 8.08, y: 3.28, w: 2.0, h: 0.18, fontFace: font, fontSize: 12, bold: true, color: C.green, margin: 0 });
  s.addShape(pptx.ShapeType.roundRect, { x: 7.85, y: 4.35, w: 3.85, h: 0.1, rectRadius: 0.05, fill: { color: 'E4EAF7' }, line: { color: 'E4EAF7' } });
  s.addShape(pptx.ShapeType.roundRect, { x: 7.85, y: 4.35, w: 1.54, h: 0.1, rectRadius: 0.05, fill: { color: '57C4A4' }, line: { color: '57C4A4' } });
  footer(s, 1);
}

// 2 Product intro
{
  const s = slide(); title(s, '产品介绍', '造师台把 AI 对话能力变成一套“可开始、可坚持、可回顾”的学习体验。');
  card(s, 0.75, 1.55, 3.76, 3.4, '专属 AI 老师', '自定义学科、教学风格、说话方式与陪学设定；不同老师拥有独立对话与任务。', C.blue);
  card(s, 4.78, 1.55, 3.76, 3.4, '个性化学习闭环', '从学习方案、课程对话、每周任务，到专注、复盘、成就与学习报告，形成连续反馈。', C.purple);
  card(s, 8.81, 1.55, 3.76, 3.4, '温和且安全的陪伴', '低压力模式、情绪支持、隐私提示、本机数据管理与服务端密钥保护。', C.green);
  s.addShape(pptx.ShapeType.roundRect, { x: 1.3, y: 5.55, w: 10.7, h: 0.56, rectRadius: 0.07, fill: { color: 'EDF2FF' }, line: { color: 'DCE5FF' } });
  s.addText('产品定位：面向学生和自主学习者的 AI 学习空间，而非单纯问答机器人。', { x: 1.3, y: 5.75, w: 10.7, h: 0.13, fontFace: font, fontSize: 11.5, bold: true, color: C.blue, align: 'center', margin: 0 }); footer(s, 2);
}

// 3 User pain points
{
  const s = slide(); title(s, '用户痛点', '学习者的核心障碍通常不是“没有内容”，而是缺少适合自己的开始方式和持续支持。');
  const items = [['不知道从哪里开始', '计划抽象、目标过大，第一次使用就需要完成很多设置。'], ['讲解不适合自己', '统一内容无法兼顾学科、基础、教材与表达偏好。'], ['压力影响坚持', '对打卡、积分、批评的焦虑，可能让学习变成负担。'], ['记录分散且不安全', '聊天、笔记、错题和复习计划割裂，隐私边界也不清晰。']];
  items.forEach((item, i) => card(s, 0.82 + (i % 2) * 6.05, 1.55 + Math.floor(i / 2) * 2.05, 5.65, 1.6, item[0], item[1], [C.blue, C.purple, C.orange, C.green][i]));
  footer(s, 3);
}

// 4 Solution
{
  const s = slide(); title(s, '解决方案', '以“老师 + 计划 + 对话 + 反馈”串起每一次学习。');
  const steps = [['1', '选择老师', '学科、风格、陪学设定'], ['2', '生成计划', '年级、教材、目标与节奏'], ['3', '开始学习', '讲解、提问、练习与朗读'], ['4', '自动反馈', '任务点亮、笔记、复习与报告']];
  steps.forEach((item, i) => { const x = 0.76 + i * 3.12; s.addShape(pptx.ShapeType.roundRect, { x, y: 2.0, w: 2.55, h: 2.5, rectRadius: 0.1, fill: { color: C.white }, line: { color: C.line } }); s.addShape(pptx.ShapeType.ellipse, { x: x + 0.2, y: 2.25, w: 0.52, h: 0.52, fill: { color: i === 2 ? C.purple : C.blue }, line: { color: i === 2 ? C.purple : C.blue } }); s.addText(item[0], { x: x + 0.2, y: 2.42, w: 0.52, h: 0.12, fontFace: font, fontSize: 9, bold: true, color: C.white, align: 'center', margin: 0 }); s.addText(item[1], { x: x + 0.22, y: 3.05, w: 2.0, h: 0.22, fontFace: font, fontSize: 14, bold: true, color: C.text, margin: 0 }); s.addText(item[2], { x: x + 0.22, y: 3.5, w: 2.0, h: 0.35, fontFace: font, fontSize: 10, color: C.muted, margin: 0, fit: 'shrink' }); if (i < 3) s.addShape(pptx.ShapeType.chevron, { x: x + 2.65, y: 3.0, w: 0.28, h: 0.33, fill: { color: 'AAB8D4' }, line: { color: 'AAB8D4' } }); });
  s.addText('核心差异：系统不仅回答问题，还会根据学习状态安排下一步，并以低压力方式帮助学生坚持。', { x: 1.0, y: 5.42, w: 11.2, h: 0.25, fontFace: font, fontSize: 13, bold: true, color: C.green, align: 'center', margin: 0 }); footer(s, 4);
}

// 5 Product demo
{
  const s = slide(); title(s, '产品演示', '一次典型使用流程：从首次进入到完成一节学习课。');
  s.addShape(pptx.ShapeType.roundRect, { x: 0.86, y: 1.5, w: 5.2, h: 4.7, rectRadius: 0.12, fill: { color: C.white }, line: { color: C.line } });
  s.addText('01 · 立即体验', { x: 1.18, y: 1.88, w: 1.5, h: 0.18, fontFace: font, fontSize: 12, bold: true, color: C.blue, margin: 0 });
  bullets(s, ['点击“快速开始”，自动选用推荐老师与标准学习方案。', '进入“今日学习控制台”，直接看到本周任务和下一步。', '点击“继续学习”，进入 AI 老师对话。'], 1.16, 2.35, 4.4, 1.55, 11.5);
  s.addShape(pptx.ShapeType.roundRect, { x: 1.2, y: 4.45, w: 4.45, h: 0.85, rectRadius: 0.07, fill: { color: C.pale }, line: { color: 'DCE5FF' } });
  s.addText('继续和林知远老师学习 →', { x: 1.2, y: 4.75, w: 4.45, h: 0.15, fontFace: font, fontSize: 11, bold: true, color: C.blue, align: 'center', margin: 0 });
  s.addShape(pptx.ShapeType.roundRect, { x: 6.55, y: 1.5, w: 5.9, h: 4.7, rectRadius: 0.12, fill: { color: 'FBFCFF' }, line: { color: 'DBE4F4' } });
  s.addText('02 · 学习后的自动反馈', { x: 6.9, y: 1.88, w: 2.5, h: 0.18, fontFace: font, fontSize: 12, bold: true, color: C.purple, margin: 0 });
  bullets(s, ['课程完成后自动点亮下一项周任务，避免手动“刷进度”。', '可把薄弱知识点加入错题本，按 1、3、7、14 天间隔复习。', '用本机笔记记录易错点；每日挑战提供 5 分钟可完成的微任务。'], 6.9, 2.35, 4.95, 1.55, 11.5);
  s.addShape(pptx.ShapeType.roundRect, { x: 6.9, y: 4.45, w: 4.95, h: 0.85, rectRadius: 0.07, fill: { color: 'EFFAF6' }, line: { color: 'D5EEE4' } });
  s.addText('学习工具箱 · 知识点 / 间隔复习 / 笔记 / 每日挑战', { x: 6.9, y: 4.75, w: 4.95, h: 0.15, fontFace: font, fontSize: 10.5, bold: true, color: C.green, align: 'center', margin: 0 }); footer(s, 5);
}

// 6 User research
{
  const s = slide(); title(s, '用户调研结果（MVP 模板）', '当前项目尚未沉淀可公开的量化问卷；本页用于展示已验证方向与下一轮调研指标，请在路演前替换为真实数据。');
  card(s, 0.8, 1.56, 3.72, 2.05, '方向 1：更低的首次门槛', '重点验证：快速开始是否让新用户更快进入第一节学习；记录首次完成率与首次学习耗时。', C.blue);
  card(s, 4.8, 1.56, 3.72, 2.05, '方向 2：更温和的陪伴', '重点验证：低压力模式、5 分钟微任务和共情表达，是否提升疲惫/抗拒场景的留存。', C.green);
  card(s, 8.8, 1.56, 3.72, 2.05, '方向 3：更连续的复习', '重点验证：错题本与间隔复习提醒是否提升知识点复习完成率。', C.purple);
  s.addShape(pptx.ShapeType.roundRect, { x: 0.8, y: 4.25, w: 11.72, h: 1.28, rectRadius: 0.08, fill: { color: 'FFF8ED' }, line: { color: 'F0DAB2' } });
  s.addText('建议补充的真实指标', { x: 1.1, y: 4.55, w: 1.7, h: 0.2, fontFace: font, fontSize: 13, bold: true, color: '9A6A2A', margin: 0 });
  s.addText('访谈样本数 · 首次学习完成率 · 次日/7 日留存 · 单次有效学习时长 · 低压力模式满意度 · 隐私信任度', { x: 2.75, y: 4.55, w: 8.9, h: 0.2, fontFace: font, fontSize: 11.5, color: '755C3A', margin: 0, fit: 'shrink' }); footer(s, 6);
}

// 7 Business model
{
  const s = slide(); title(s, '商业模式', '以免费体验建立学习习惯，以更深的个性化能力形成增值。');
  card(s, 0.78, 1.55, 3.78, 3.62, '免费版', '基础老师创建\n标准学习方案\n本机笔记、任务、错题与学习记录\n局域网体验能力', C.blue);
  card(s, 4.78, 1.55, 3.78, 3.62, '高级版', '更多 AI 老师和高级语音\n深度学习报告与个性化复习\n多学科知识点卡片与强化计划', C.purple);
  card(s, 8.78, 1.55, 3.78, 3.62, '家庭 / 机构版', '家庭学习看板与时长边界\n教师批量发布任务和班级汇总\n学校、培训机构定制化部署', C.green);
  s.addText('增长路径：学生个人体验 → 家庭共享 → 学校 / 培训机构合作。', { x: 1.25, y: 5.72, w: 10.7, h: 0.2, fontFace: font, fontSize: 13, bold: true, color: C.blue, align: 'center', margin: 0 }); footer(s, 7);
}

// 8 Roadmap + team
{
  const s = slide(); title(s, '产品规划与团队介绍', '先打磨学习闭环，再扩展协作、数据洞察和机构服务。');
  s.addShape(pptx.ShapeType.roundRect, { x: 0.78, y: 1.5, w: 7.15, h: 4.85, rectRadius: 0.1, fill: { color: C.white }, line: { color: C.line } });
  s.addText('产品路线图', { x: 1.08, y: 1.85, w: 1.5, h: 0.24, fontFace: font, fontSize: 17, bold: true, color: C.text, margin: 0 });
  [['近期','丰富知识点卡片、错题本、学习笔记、每日挑战与学习日历'], ['中期','家长模式、教师任务发布、学习数据洞察与多模态提问'], ['长期','学习小组、机构 SaaS、个性化学习图谱与跨端体验']].forEach((item, i) => { const y = 2.55 + i * 1.0; pill(s, item[0], 1.08, y, 0.68, [C.blue, C.purple, C.green][i], ['EAF0FF', 'F1EEFF', 'EAF9F4'][i]); s.addText(item[1], { x: 1.98, y: y + 0.06, w: 5.35, h: 0.17, fontFace: font, fontSize: 11, color: C.muted, margin: 0, fit: 'shrink' }); });
  s.addShape(pptx.ShapeType.roundRect, { x: 8.28, y: 1.5, w: 4.27, h: 4.85, rectRadius: 0.1, fill: { color: 'F4F7FF' }, line: { color: 'DDE6FA' } });
  s.addText('团队介绍（可替换）', { x: 8.62, y: 1.85, w: 2.6, h: 0.24, fontFace: font, fontSize: 17, bold: true, color: C.text, margin: 0 });
  s.addText('产品负责人\n用户研究、学习体验设计与商业验证\n\n技术负责人\nAI 对话、服务端安全与跨端体验\n\n教育顾问\n课程设计、学习评估与儿童/学生安全边界', { x: 8.62, y: 2.5, w: 3.35, h: 2.5, fontFace: font, fontSize: 11.5, color: C.muted, margin: 0, breakLine: false, fit: 'shrink' });
  s.addText('谢谢观看', { x: 8.62, y: 5.55, w: 2.1, h: 0.2, fontFace: font, fontSize: 13, bold: true, color: C.blue, margin: 0 }); footer(s, 8);
}

pptx.writeFile({ fileName: '造师台-路演PPT.pptx' });
