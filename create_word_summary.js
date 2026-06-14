const fs = require('fs');
const path = require('path');
const zlib = require('zlib');

const out = 'C:/Users/User/Desktop/清北目标与赚钱思考.docx';

const paragraphs = [
  ['title', '高中生身份下关于赚钱与清北目标的对话整理'],
  ['normal', '本文档整理了刚才的聊天内容，重点保留与你当前“高中生身份”、赚钱想法、清华北大目标相关的核心讨论。'],
  ['heading', '1. 身份切换'],
  ['quote', '用户：我现在作为高中生的身份再和你聊天'],
  ['normal', '回复要点：后续交流按高中生身份来处理，复杂概念会拆成简单步骤；涉及论文、专利、通信科研时，会区分“高中生学习理解”和“科研/专利写作”两个层级。'],
  ['heading', '2. 关于“怎么赚大钱”的真实判断'],
  ['quote', '用户：你觉得咋能赚大钱'],
  ['normal', '核心结论：高中阶段想赚大钱，最现实的路线不是马上找暴利项目，而是尽早积累可以复利的能力：技术、表达、商业理解和信用。'],
  ['normal', '不建议走的方向包括：炒币、彩票、刷单、灰产、平台漏洞、所谓稳赚项目、流量课和过早重资产生意。这些更可能让你被别人赚钱，而不是你赚钱。'],
  ['normal', '建议押注的能力组合：AI 工具能力 + 编程能力 + 某个行业知识。结合你关注通信、6G、低空通信、无线网络数字孪生等方向，未来可以考虑科研资料自动整理工具、专利检索和分析工具、论文图表生成工具、通信仿真可视化工具、科研人员 AI 助手等。'],
  ['normal', '关键商业思维：不要只做一次性接单，要做可以复用和规模化的小产品。一次投入，多次销售，才有放大空间。'],
  ['heading', '3. 推荐的长期路线'],
  ['normal', '结合“高中生 + 通信科研相关背景”，建议路线是：'],
  ['bullet', '高中阶段：AI + 编程 + 科研写作基础。'],
  ['bullet', '大学前后：做科研、专利、论文辅助工具。'],
  ['bullet', '长期：做面向科研人员、学生、工程师的产品或服务。'],
  ['normal', '阶段性目标：30 天内做出一个能帮别人解决问题的小工具或资料产品，例如高中物理错题整理模板、AI 论文阅读笔记模板、通信基础知识小网站、专利交底书生成模板、PDF 文献整理成表格的小脚本等。'],
  ['heading', '4. 清华北大目标优先级'],
  ['quote', '用户：可是我现在的目标是考上清华北大'],
  ['normal', '核心结论：如果目标是清华北大，现在“赚钱”应降级为副线，主线必须是高考、竞赛、强基或相关升学路径。高中阶段真正能让未来赚钱能力放大的最大杠杆，是先把清北这个入口拿到。'],
  ['normal', '真实判断：现在不应该把主要精力放在赚钱上，因为高中时间窗口极短，清北录取机会的机会成本极高。当前最值钱的是分数、学科能力、竞赛/科研履历和稳定心态。'],
  ['heading', '5. 清北路径需要先判断'],
  ['normal', '清北路线大致包括：'],
  ['bullet', '裸分高考。'],
  ['bullet', '强基计划。'],
  ['bullet', '竞赛保送、破格或相关降分路径。'],
  ['bullet', '专项计划或综合评价类路径，具体取决于省份和政策。'],
  ['bullet', '港校或海外作为备选，如果你愿意。'],
  ['normal', '关键问题不是“怎么努力”，而是：我离清北差多少？差在哪一科？靠什么路径补？'],
  ['heading', '6. 学习策略要点'],
  ['bullet', '分数优先级通常是：数学 > 物理/化学/理综方向 > 英语 > 语文。不是语文不重要，而是数学和理科更容易拉开顶尖差距。'],
  ['bullet', '不要平均用力。要找出拖后腿科目、能冲满分科目、最容易提分模块和最容易失误题型，然后集中爆破。'],
  ['bullet', '建立错题系统，不是简单抄错题，而是记录错误原因、下次第一步怎么做、同类题模型是什么。'],
  ['bullet', '真题和高质量模拟题优先。一套卷子真正吃透，价值大于刷三套不复盘。'],
  ['heading', '7. 赚钱副线如何处理'],
  ['normal', '如果保留赚钱副线，每周最多 2 小时，并且必须服务于主线学习。可以做高中物理错题模板、数学模型总结、英语作文素材库、高考复盘系统、学习计划自动生成表等。这些既能提高自己，也可能未来变现。'],
  ['heading', '8. 建议建立“清北目标作战文档”'],
  ['normal', '建议新建一个清北目标作战文档，先定规则再执行。示例结构如下：'],
  ['code', '# 清北目标作战规则\n\n## 总目标\n目标：考入清华/北大。\n时间节点：____年高考。\n目标路径：裸分 / 强基 / 竞赛 / 其他。\n\n## 当前基线\n省份：\n年级：\n选科：\n最近排名：\n当前总分：\n目标总分：\n差距：\n\n## 科目策略\n数学：\n物理：\n化学/生物/政治/历史/地理：\n语文：\n英语：\n\n## 每周规则\n- 每周复盘一次成绩和错题\n- 每天固定数学训练\n- 每天固定英语输入\n- 每周至少一次完整限时训练\n- 错题必须分类，不只抄题\n\n## 红线\n- 不为了副业牺牲主线学习\n- 不沉迷低质量资料\n- 不用计划代替执行\n- 不刷题不复盘'],
  ['heading', '9. 下一步需要补充的信息'],
  ['normal', '为了制定具体清北备考方案，需要填写：'],
  ['code', '年级：\n省份：\n选科：\n最近一次大考总分：\n年级排名/全市或全省排名：\n数学分数：\n语文分数：\n英语分数：\n物理/化学/其他分数：\n最强科目：\n最弱科目：\n是否学竞赛：\n目标专业：\n每天可自主学习时间：'],
  ['heading', '10. 总结'],
  ['normal', '当前主线：清华北大。赚钱不是取消，而是降级为服务学习的副线。最优策略是先用分数和能力拿到更高平台，再把 AI、编程、通信科研理解和产品化能力放大成未来的收入能力。'],
];

const styleMap = {
  title: '<w:pStyle w:val="Title"/>',
  heading: '<w:pStyle w:val="Heading1"/>',
  quote: '<w:pStyle w:val="Quote"/>',
  bullet: '<w:pStyle w:val="ListBullet"/>',
  code: '<w:pStyle w:val="Code"/>',
};

function xmlEscape(s) {
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

function makeParagraph(kind, text) {
  const ppr = styleMap[kind] ? `<w:pPr>${styleMap[kind]}</w:pPr>` : '';
  const lines = String(text).split('\n');
  const runs = [];
  for (let i = 0; i < lines.length; i++) {
    if (i > 0) runs.push('<w:r><w:br/></w:r>');
    const preserve = kind === 'code' || /^\s|\s$/.test(lines[i]) ? ' xml:space="preserve"' : '';
    runs.push(`<w:r><w:t${preserve}>${xmlEscape(lines[i])}</w:t></w:r>`);
  }
  return `<w:p>${ppr}${runs.join('')}</w:p>`;
}

const body = paragraphs.map(([kind, text]) => makeParagraph(kind, text)).join('');

const files = {
  '[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>`,
  '_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`,
  'word/document.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    ${body}
    <w:sectPr>
      <w:pgSz w:w="11906" w:h="16838"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/>
    </w:sectPr>
  </w:body>
</w:document>`,
  'word/styles.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:rPr><w:rFonts w:ascii="Microsoft YaHei" w:eastAsia="Microsoft YaHei"/><w:sz w:val="22"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:rPr><w:b/><w:rFonts w:ascii="Microsoft YaHei" w:eastAsia="Microsoft YaHei"/><w:sz w:val="36"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:rPr><w:b/><w:rFonts w:ascii="Microsoft YaHei" w:eastAsia="Microsoft YaHei"/><w:sz w:val="28"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="Quote"><w:name w:val="Quote"/><w:basedOn w:val="Normal"/><w:pPr><w:ind w:left="420"/></w:pPr><w:rPr><w:i/><w:color w:val="666666"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="ListBullet"><w:name w:val="List Bullet"/><w:basedOn w:val="Normal"/><w:pPr><w:ind w:left="420" w:hanging="240"/></w:pPr></w:style>
  <w:style w:type="paragraph" w:styleId="Code"><w:name w:val="Code"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:before="120" w:after="120"/></w:pPr><w:rPr><w:rFonts w:ascii="Consolas" w:eastAsia="Microsoft YaHei"/><w:sz w:val="20"/></w:rPr></w:style>
</w:styles>`,
  'word/_rels/document.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`,
};

function crc32(buf) {
  let table = crc32.table;
  if (!table) {
    table = crc32.table = new Uint32Array(256);
    for (let i = 0; i < 256; i++) {
      let c = i;
      for (let k = 0; k < 8; k++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
      table[i] = c >>> 0;
    }
  }
  let c = 0xFFFFFFFF;
  for (const b of buf) c = table[(c ^ b) & 0xFF] ^ (c >>> 8);
  return (c ^ 0xFFFFFFFF) >>> 0;
}

function dosDateTime(date = new Date()) {
  const time = (date.getHours() << 11) | (date.getMinutes() << 5) | Math.floor(date.getSeconds() / 2);
  const dosDate = ((date.getFullYear() - 1980) << 9) | ((date.getMonth() + 1) << 5) | date.getDate();
  return { time, date: dosDate };
}

function u16(n) { const b = Buffer.alloc(2); b.writeUInt16LE(n); return b; }
function u32(n) { const b = Buffer.alloc(4); b.writeUInt32LE(n >>> 0); return b; }

function createZip(entries) {
  const chunks = [];
  const central = [];
  let offset = 0;
  const dt = dosDateTime();
  for (const [name, content] of Object.entries(entries)) {
    const nameBuf = Buffer.from(name, 'utf8');
    const data = Buffer.from(content, 'utf8');
    const comp = zlib.deflateRawSync(data);
    const crc = crc32(data);
    const local = Buffer.concat([
      u32(0x04034b50), u16(20), u16(0x0800), u16(8), u16(dt.time), u16(dt.date),
      u32(crc), u32(comp.length), u32(data.length), u16(nameBuf.length), u16(0), nameBuf, comp
    ]);
    chunks.push(local);
    central.push(Buffer.concat([
      u32(0x02014b50), u16(20), u16(20), u16(0x0800), u16(8), u16(dt.time), u16(dt.date),
      u32(crc), u32(comp.length), u32(data.length), u16(nameBuf.length), u16(0), u16(0),
      u16(0), u16(0), u32(0), u32(offset), nameBuf
    ]));
    offset += local.length;
  }
  const centralOffset = offset;
  const centralBuf = Buffer.concat(central);
  const end = Buffer.concat([
    u32(0x06054b50), u16(0), u16(0), u16(central.length), u16(central.length),
    u32(centralBuf.length), u32(centralOffset), u16(0)
  ]);
  return Buffer.concat([...chunks, centralBuf, end]);
}

fs.writeFileSync(out, createZip(files));
console.log(out);
