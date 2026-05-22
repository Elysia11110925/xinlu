const pptxgen = require('pptxgenjs');
const pptx = new pptxgen();

/* ═══════════════ 全局主题 ═══════════════ */
const C = {
  orange:   'e67e22',
  dark:     '2c3e50',
  gray:     '6b7b8d',
  white:    'ffffff',
  lightBg:  'fdf2e9',
  black:    '1a1a1a',
};

pptx.defineLayout({ name:'CUSTOM', width:13.33, height:7.5 });
pptx.layout = 'CUSTOM';

function titleBar(slide, num, title) {
  slide.addText(`0${num}`, { x:0.6, y:0.4, w:0.8, h:0.6, fontSize:11, color:C.orange, fontFace:'Arial', bold:true });
  slide.addText(title, { x:1.4, y:0.4, w:10, h:0.7, fontSize:26, color:C.dark, fontFace:'Microsoft YaHei', bold:true });
  slide.addShape(pptx.ShapeType.rect, { x:0.6, y:1.1, w:1.5, h:3, fill:{color:C.orange} });
}

/* ══ Slide 1 — 封面 ══ */
{
  const s = pptx.addSlide();
  s.background = { fill: C.dark };
  s.addShape(pptx.ShapeType.rect, { x:0, y:0, w:13.33, h:0.08, fill:{color:C.orange} });
  s.addShape(pptx.ShapeType.rect, { x:0, y:7.42, w:13.33, h:0.08, fill:{color:C.orange} });
  s.addText('心路志愿服务实践团', { x:0, y:1.8, w:'100%', h:1.2, fontSize:48, color:C.white, fontFace:'Microsoft YaHei', bold:true, align:'center' });
  s.addText('以心换心 · 路路相通 · 用热爱点亮希望', { x:0, y:3.0, w:'100%', h:0.6, fontSize:18, color:C.orange, fontFace:'Microsoft YaHei', align:'center' });
  s.addText('北京科技大学 · 计通学院', { x:0, y:4.8, w:'100%', h:0.5, fontSize:16, color:'bbbbbb', fontFace:'Microsoft YaHei', align:'center' });
  s.addText('2026 暑期社会实践宣传', { x:0, y:5.4, w:'100%', h:0.5, fontSize:13, color:'888888', fontFace:'Microsoft YaHei', align:'center' });
}

/* ══ Slide 2 — 团队简介 ══ */
{
  const s = pptx.addSlide();
  titleBar(s, 1, '团队简介');
  s.addText([
    {text:'北京科技大学心路志愿服务实践团', options:{bold:true, fontSize:18, color:C.dark}},
    {text:'是一支致力于\n教育、调研、社会服务的队伍，以支教、陪伴、募捐义卖、文艺汇演等多种活动形式关爱留守儿童和务工子女。', options:{fontSize:14, color:C.gray}},
  ], { x:2.8, y:1.6, w:9.5, h:1.6, fontFace:'Microsoft YaHei', lineSpacingMultiple:1.8 });
  s.addText([
    {text:'作为计通学院的老牌暑期实践团', options:{bold:true}},
    {text:'，心路积极响应国家和学校号召，已将帮扶对象从务工子女逐步扩展至全国教育资源不均衡地区的儿童们。心路人走遍大江南北，足迹遍布祖国各地，以爱心驱散阴霾，用陪伴消弭隔阂。', options:{}},
  ], { x:2.8, y:3.2, w:9.5, h:1.6, fontSize:14, color:C.gray, fontFace:'Microsoft YaHei', lineSpacingMultiple:1.8 });

  s.addText([
    {text:'2009年至今\n', options:{bold:true, fontSize:14, color:C.orange}},
    {text:'长期志愿项目十七年\n短期实践八年\n活动 600+ 次\n服务 4000+ 人次\n志愿时长 20000+ 小时', options:{fontSize:12, color:C.gray}},
  ], { x:2.8, y:5.0, w:9.5, h:2.2, fontFace:'Microsoft YaHei', lineSpacingMultiple:1.6 });
}

/* ══ Slide 3 — 实践内容 ══ */
{
  const s = pptx.addSlide();
  titleBar(s, 2, '实践内容');

  const items = [
    { icon:'📚', title:'支教助学', desc:'深入教育资源不均衡地区\n开展课业辅导与知识拓展\n帮助儿童夯实学业基础' },
    { icon:'🤝', title:'陪伴关怀', desc:'以日常陪伴与心灵沟通\n为核心\n关注儿童心理健康' },
    { icon:'🎭', title:'募捐义卖\n文艺汇演', desc:'组织爱心募捐、义卖活动\n和文艺汇演\n汇聚社会力量' },
    { icon:'💡', title:'科技创新', desc:'发挥专业优势\n研发智能点读系统\n以科技赋能教育公平' },
  ];

  items.forEach((it, i) => {
    const x = 0.6 + i * 3.1;
    s.addShape(pptx.ShapeType.roundRect, { x, y:1.6, w:2.8, h:5.2, fill:{color:C.lightBg}, rectRadius:0.1 });
    s.addText(it.icon, { x, y:1.9, w:2.8, h:0.8, fontSize:30, align:'center' });
    s.addText(it.title, { x, y:2.8, w:2.8, h:1.0, fontSize:18, color:C.dark, fontFace:'Microsoft YaHei', bold:true, align:'center' });
    s.addText(it.desc, { x:x+0.3, y:4.0, w:2.2, h:2.5, fontSize:13, color:C.gray, fontFace:'Microsoft YaHei', lineSpacingMultiple:1.8 });
  });
}

/* ══ Slide 4 — 成果数据 ══ */
{
  const s = pptx.addSlide();
  titleBar(s, 3, '成果数据');

  const stats = [
    { num:'600+',  label:'活动次数' },
    { num:'4000+', label:'累计服务人次' },
    { num:'20000+', label:'志愿服务时长\n（小时）' },
    { num:'20余项',  label:'荣誉奖项' },
  ];

  stats.forEach((st, i) => {
    const x = 1.2 + i * 2.9;
    s.addShape(pptx.ShapeType.rect, { x, y:1.8, w:2.4, h:3.6, fill:{color:C.white}, line:{color:'e8e8e8', width:1} });
    s.addText(st.num, { x, y:2.0, w:2.4, h:1.2, fontSize:38, color:C.orange, fontFace:'Arial', bold:true, align:'center' });
    s.addText(st.label, { x, y:3.4, w:2.4, h:1.2, fontSize:13, color:C.gray, fontFace:'Microsoft YaHei', align:'center' });
  });

  s.addText([{text:'近百位', options:{bold:true, fontSize:16, color:C.orange}}, {text:'社会实践先进个人', options:{fontSize:14, color:C.gray}}], { x:0, y:5.8, w:'100%', h:0.6, fontSize:14, fontFace:'Microsoft YaHei', align:'center' });
}

/* ══ Slide 5 — 荣誉奖项 ══ */
{
  const s = pptx.addSlide();
  titleBar(s, 4, '荣誉奖项');

  const awards = [
    '北京市志愿服务项目大赛 银奖',
    '北京市重点志愿项目',
    '团中央"七彩假期"重点示范团队',
    '校级重点志愿服务项目',
    '2024"青年服务国家"首都大学生\n暑期社会实践优秀团队',
    '校级社会实践金奖',
    '……等二十余项荣誉',
  ];

  s.addText(awards.map((a, i) => ({
    text:`${i === awards.length-1 ? '' : '▸ '}${a}\n`,
    options:{fontSize:16, color: i===awards.length-1 ? C.orange : C.dark, bold:i===awards.length-1, lineSpacingMultiple:1.8},
  })), { x:3.0, y:1.8, w:9, h:5, fontFace:'Microsoft YaHei', lineSpacingMultiple:2.0 });
}

/* ══ Slide 6 — 科技创新 ══ */
{
  const s = pptx.addSlide();
  titleBar(s, 5, '科技创新成果');

  const techs = [
    { name:'智能识别点读系统 iReader', desc:'国家级发明专利\n随时满足儿童阅读需求' },
    { name:'孩童专注度分析系统', desc:'科学分析学习专注度\n提高学习效率' },
    { name:'"心路-用心陪伴"小程序', desc:'独立开发，共享优质\n教育资源' },
    { name:'盲人阅读眼镜', desc:'自主研制\n以科技驱散黑暗阴霾' },
  ];

  techs.forEach((t, i) => {
    const y = 1.6 + i * 1.4;
    s.addShape(pptx.ShapeType.roundRect, { x:0.6, y, w:12, h:1.2, fill:{color: i%2===0?C.lightBg:C.white}, rectRadius:0.06 });
    s.addText(t.name, { x:1.0, y:y+0.1, w:5, h:1.0, fontSize:16, color:C.dark, fontFace:'Microsoft YaHei', bold:true });
    s.addText(t.desc, { x:6.5, y:y+0.1, w:5.5, h:1.0, fontSize:13, color:C.gray, fontFace:'Microsoft YaHei', lineSpacingMultiple:1.4 });
  });

  s.addText('发挥计通学院专业优势，聚力科技创新，以科技赋能教育公平', { x:0, y:7.0, w:'100%', h:0.4, fontSize:11, color:C.gray, fontFace:'Microsoft YaHei', align:'center' });
}

/* ══ Slide 7 — 团队构成 ══ */
{
  const s = pptx.addSlide();
  titleBar(s, 6, '团队构成（10人）');

  const depts = [
    '队长', '外联组', '摄影组', '设计组',
    '文案组 ×2', '宣传组 ×2', '程序开发组 ×2',
  ];

  depts.forEach((d, i) => {
    const col = i % 4;
    const row = Math.floor(i / 4);
    const x = 0.6 + col * 3.1;
    const y = 1.6 + row * 2.6;

    s.addShape(pptx.ShapeType.roundRect, { x, y, w:2.8, h:2.3, fill:{color:C.lightBg}, rectRadius:0.1 });
    s.addText(d, { x, y:y+0.5, w:2.8, h:1.2, fontSize:18, color:C.dark, fontFace:'Microsoft YaHei', bold:true, align:'center' });
  });
}

/* ══ Slide 8 — 心路之旅（时间线式结语） ══ */
{
  const s = pptx.addSlide();
  s.background = { fill: C.dark };
  s.addShape(pptx.ShapeType.rect, { x:0, y:0, w:13.33, h:0.08, fill:{color:C.orange} });
  s.addShape(pptx.ShapeType.rect, { x:0, y:7.42, w:13.33, h:0.08, fill:{color:C.orange} });

  s.addText('心路串联岁月\n奔赴烂漫纯真的童话秘境', { x:0, y:1.2, w:'100%', h:1.6, fontSize:32, color:C.white, fontFace:'Microsoft YaHei', bold:true, align:'center' });
  s.addText('以爱意守护童真，用陪伴书写专属成长篇章\n这是属于孩子的心灵旅程，更是我们携手同行、双向奔赴的心路之旅', { x:0, y:3.0, w:'100%', h:1.0, fontSize:15, color:'bbbbbb', fontFace:'Microsoft YaHei', align:'center', lineSpacingMultiple:1.8 });

  s.addText('回望 2025，心路满载温柔与硕果\n奔赴 2026，初心不改，温暖续航', { x:0, y:4.5, w:'100%', h:1.0, fontSize:16, color:C.orange, fontFace:'Microsoft YaHei', align:'center', lineSpacingMultiple:1.8 });

  s.addText('让我们继续以心相伴，逐光前行\n共赴全新的心路盛宴，书写璀璨新章', { x:0, y:5.8, w:'100%', h:1.0, fontSize:20, color:C.white, fontFace:'Microsoft YaHei', bold:true, align:'center', lineSpacingMultiple:1.8 });
}

/* ═══════════════ 输出 ═══════════════ */
pptx.writeFile({ fileName:'心路实践团宣传.pptx' }).then(() => {
  console.log('PPT 已生成: 心路实践团宣传.pptx');
}).catch(err => {
  console.error(err);
});
