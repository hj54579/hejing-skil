// 锈鉴智检 PPT 重建脚本
// 设计系统：藏青 + 青绿点缀；白底内容页；封面/结尾深色（三明治结构）
const PptxGenJS = require('pptxgenjs');
const pres = new PptxGenJS();
pres.layout = 'LAYOUT_16x9'; // 10 x 5.625 in

const F = 'Microsoft YaHei';
const SW = 10, SH = 5.625, MX = 0.5, CW = 9;
const C = {
  navy: '12283D', navy2: '16324A', navyLine: '2A4A66',
  teal: '0FA3B1', tealDark: '0B7A86', tealLt: 'E3F4F6', tealBg: 'EAF6F7', tealLine: 'CFE8EA',
  red: 'C0392B', redLt: 'FBECEA', redLine: 'F2C9C4',
  orange: 'F09F13',
  ink: '1F2A37', gray: '5B6B7A', line: 'E3E9EF', white: 'FFFFFF',
  dSub: 'AAC4D6', dSub2: '8FA8BC', dTxt: 'C9D8E4', dCard: 'F7FAFC'
};

// ---------- 基础工具 ----------
function T(s, text, o = {}) {
  s.addText(text, {
    x: o.x, y: o.y, w: o.w, h: o.h,
    fontFace: o.font || F,
    fontSize: o.fs || 11,
    color: o.color || C.ink,
    bold: !!o.bold,
    italic: !!o.italic,
    align: o.align || 'left',
    valign: o.valign || 'top',
    margin: 0,
    lineSpacingMultiple: o.ls == null ? 1.12 : o.ls,
    fit: 'shrink',
    charSpacing: o.cs || 0
  });
  return s;
}
function B(s, x, y, w, h, fill, line, radius) {
  return s.addShape(pres.ShapeType.roundRect, {
    x, y, w, h,
    fill: { color: fill },
    line: line ? (typeof line === 'string' ? { color: line, width: 1 } : line) : { type: 'none' },
    rectRadius: radius == null ? 0.055 : radius,
    shadow: { type: 'none' }
  });
}
function SQ(s, x, y, w, h, fill) {
  return s.addShape(pres.ShapeType.rect, { x, y, w, h, fill: { color: fill }, line: { type: 'none' }, shadow: { type: 'none' } });
}
function CI(s, x, y, d, fill, line) {
  return s.addShape(pres.ShapeType.ellipse, {
    x, y, w: d, h: d,
    fill: { color: fill },
    line: line ? { color: line, width: 1.1 } : { type: 'none' },
    shadow: { type: 'none' }
  });
}
function rt(text, options) { return options ? { text, options } : { text }; }

function header(s, tag, title, page, dark) {
  T(s, tag, { x: MX, y: 0.30, w: 6, h: 0.24, fs: 10, color: C.teal, bold: true });
  T(s, title, { x: MX, y: 0.52, w: CW, h: 0.60, fs: 21, color: dark ? C.white : C.ink, bold: true });
  T(s, String(page).padStart(2, '0'), { x: 8.85, y: 0.28, w: 0.65, h: 0.26, fs: 9.5, color: dark ? '5E7C92' : C.gray, align: 'right' });
}
function band(s, y, text, o = {}) {
  const h = o.h || 0.42, fill = o.fill || C.navy;
  B(s, MX, y, CW, h, fill);
  SQ(s, MX + 0.16, y + h / 2 - 0.055, 0.11, 0.11, o.dot || C.teal);
  T(s, text, { x: MX + 0.4, y, w: CW - 0.7, h, fs: o.fs || 12, color: o.color || C.white, bold: o.bold !== false, valign: 'middle' });
}

// ================= S1 封面（左对齐 · 右侧留白 · 三栏底部信息） =================
(function () {
  const s = pres.addSlide(); s.background = { color: '161D30' };
  // 顶部：赛事标识 + 日期
  CI(s, 0.50, 0.545, 0.12, C.orange);
  T(s, '中国国际大学生创新大赛（2024）拟推荐项目', { x: 0.72, y: 0.47, w: 7, h: 0.30, fs: 12, color: C.white });
  T(s, '2024.10', { x: 8.4, y: 0.47, w: 1.1, h: 0.30, fs: 11, color: '94A3B8', align: 'right' });
  // 赛道标签（橙色路标）
  T(s, '学生赛道（本科组）', { x: 0.5, y: 1.04, w: 6, h: 0.34, fs: 13.5, color: C.orange, bold: true });
  // 主标题 + 橙色短横线
  T(s, '「锈鉴智检」', { x: 0.5, y: 1.56, w: 8.8, h: 0.95, fs: 54, color: C.white, bold: true });
  SQ(s, 6.30, 2.34, 1.25, 0.055, C.orange);
  // 副标题
  T(s, '低成本高精度钢材质检专家', { x: 0.5, y: 2.66, w: 9, h: 0.55, fs: 27, color: C.white, bold: true });
  // Slogan
  T(s, '让高质量质检普惠每一家中小工厂', { x: 0.5, y: 3.44, w: 9, h: 0.4, fs: 15, color: 'B0B8C4' });
  // 分隔线
  SQ(s, 0.5, 4.46, 9, 0.012, '2E4A63');
  // 底部信息栏（三列）
  const cols = [
    { x: 0.5, w: 3.1, label: '学院 · 专业', lines: ['广西科技师范学院', '先进制造工程学院 · 智能制造工程'] },
    { x: 3.8, w: 3.6, label: '团队 · 负责人', lines: ['锈鉴智检创新研发团队', '负责人：谢宇｜成员：曹晓倩、伍锦琦、谢耀宽、何军'] },
    { x: 7.6, w: 2.4, label: '指导教师', lines: ['陈旭红、赖道辉'] }
  ];
  cols.forEach(c => {
    T(s, c.label, { x: c.x, y: 4.68, w: c.w, h: 0.24, fs: 9.5, color: '94A3B8' });
    T(s, c.lines.map(t => ({ text: t, options: { breakLine: true } })),
      { x: c.x, y: 4.98, w: c.w, h: 0.50, fs: 10.3, color: 'D5E2EC', ls: 1.25 });
  });
})();

// ================= S2 目录 =================
(function () {
  const s = pres.addSlide(); s.background = { color: C.white };
  T(s, 'CONTENTS', { x: MX, y: 0.30, w: 6, h: 0.24, fs: 10, color: C.teal, bold: true });
  T(s, '目录', { x: MX, y: 0.50, w: 4, h: 0.62, fs: 26, color: C.ink, bold: true });
  T(s, '02', { x: 8.85, y: 0.28, w: 0.65, h: 0.26, fs: 9.5, color: C.gray, align: 'right' });
  const items = [
    ['01', '项目概述', '一页看懂"锈鉴智检"核心优势与千元级突破。'],
    ['02', '市场痛点', '被忽视的97%中小钢企，面临高成本与高漏检。'],
    ['03', '产品与技术', '光学成像源头创新 × 轻量化AI算法三级联动。'],
    ['04', '市场分析', '60亿存量蓝海，轻量化普惠赛道增速35%居首。'],
    ['05', '商业模式', '"硬件+软件+服务+租赁"四位一体盈利闭环。'],
    ['06', '投资与财务', '三级阶梯资金投入，投资回收期仅2-2.5年。'],
    ['07', '生产与风险', '轻资产稳健落地，三级质控与全流程风险管控。'],
    ['08', '团队与规划', '智能制造学子领航，高校双导师全程护航。']
  ];
  const xs = [0.5, 2.8, 5.1, 7.4], ys = [1.42, 3.28], w = 2.1, h = 1.70;
  items.forEach((it, i) => {
    const x = xs[i % 4], y = ys[Math.floor(i / 4)];
    if (i === 0) { B(s, x, y, w, h, C.tealBg, C.tealLine); SQ(s, x, y + 0.18, 0.045, 1.34, C.teal); }
    else B(s, x, y, w, h, C.white, C.line);
    T(s, it[0], { x: x + 0.18, y: y + 0.16, w: 1.2, h: 0.34, fs: 19, color: C.teal, bold: true });
    T(s, it[1], { x: x + 0.18, y: y + 0.62, w: 1.75, h: 0.32, fs: 13, color: C.ink, bold: true });
    T(s, it[2], { x: x + 0.18, y: y + 1.00, w: 1.75, h: 0.62, fs: 9.3, color: C.gray, ls: 1.2 });
  });
})();

// ================= S3 项目概述 =================
(function () {
  const s = pres.addSlide(); s.background = { color: C.white };
  header(s, 'EXECUTIVE SUMMARY', '项目概述｜千元级成本 · 万元级售价 · 十万元级效果', 3);
  B(s, 0.5, 1.30, 9, 0.94, C.navy);
  T(s, '面向国内海量中小钢材企业，打造轻量化外挂式智能质检系统——无需改造产线、即装即用，用 AI 视觉替代人工肉眼，精准识别钢材表面早期锈蚀与端面微裂纹，让中小工厂用得起、用得好。',
    { x: 0.8, y: 1.30, w: 8.4, h: 0.94, fs: 12, color: C.white, valign: 'middle', ls: 1.3 });
  const kpis = [
    ['≥ 95%', '锈蚀识别准确率', '可精准捕获0.1mm肉眼难辨的早期点锈与微锈，在车间粉尘、震动等复杂工况下保持高稳定输出。'],
    ['≤ 4%', '伪缺陷误检率', '依托12维几何与反光特征以及SVM二次甄别，完美剔除油污、灰尘、水印干扰，检出即为真缺陷。'],
    ['≤ 8000元', '单工位硬件成本', '采用民用工业级器件集成与全国产化选型，硬件造价极低，彻底打破大型高端设备的价格壁垒。'],
    ['仅 30%', '综合部署成本占比', '整套系统落地价格仅为传统大型定制方案的三分之一，极大降低中小微企智能化改造门槛。']
  ];
  const xs = [0.5, 2.8, 5.1, 7.4], w = 2.1;
  kpis.forEach((k, i) => {
    const x = xs[i];
    B(s, x, 2.42, w, 2.05, C.white, C.line);
    SQ(s, x, 2.42 + 0.22, 0.05, 0.42, C.teal);
    T(s, k[0], { x: x + 0.18, y: 2.42 + 0.24, w: 1.8, h: 0.5, fs: 24, color: C.teal, bold: true });
    T(s, k[1], { x: x + 0.18, y: 2.42 + 0.82, w: 1.8, h: 0.32, fs: 11.5, color: C.ink, bold: true });
    T(s, k[2], { x: x + 0.18, y: 2.42 + 1.18, w: 1.76, h: 0.82, fs: 8.6, color: C.gray, ls: 1.18 });
  });
  band(s, 4.62, '核心价值——一套设备替代多道人工质检工序，降本增效即刻可见。', { h: 0.50, fill: C.tealLt, color: C.ink, fs: 12.5 });
})();

// ================= S4 四大核心功能 =================
(function () {
  const s = pres.addSlide(); s.background = { color: C.white };
  header(s, 'CORE FEATURES', '四大核心功能｜一套设备替代多道人工工序', 4);
  const cards = [
    ['①', '高精度早期锈蚀识别分级', '捕捉 0.1mm 肉眼难辨的早期点锈、微锈；严格对标 GB/T 8923.1-2011 国标，按锈蚀面积占比与分布密度自动完成 A/B/C/D 等级判定；判定结果全程留痕，供需对账有据可依，消除等级争议。'],
    ['②', '工业伪缺陷智能滤除', '针对车间常见的油污、灰尘、水印、划痕等干扰，依托反光率差异与 12 维几何特征二次甄别；误检率控制在 4% 以内，检出缺陷无需人工复核，真正"减人"而非"人机双保险"。'],
    ['③', '端面微裂纹同步检测', '同一工位同步完成端面质检，"边缘提取初筛 + 深度学习复核"双重把关；可识别低至 0.05mm 的微裂纹并评估延伸趋势，杜绝隐性缺陷流入下道工序引发批量报废。'],
    ['④', '质检数据可视化管理', '每根钢材的检测图像、等级、缺陷位置自动归档，按批次生成多维质量报表；支持历史追溯与趋势分析，帮工厂把"质检数据"变成"工艺改进依据"，建立标准化数字质控闭环。']
  ];
  const xs = [0.5, 5.125], ys = [1.36, 3.06], w = 4.375, h = 1.60;
  cards.forEach((c, i) => {
    const x = xs[i % 2], y = ys[Math.floor(i / 2)];
    B(s, x, y, w, h, C.white, C.line);
    CI(s, x + 0.22, y + 0.20, 0.40, C.tealLt);
    T(s, c[0], { x: x + 0.22, y: y + 0.28, w: 0.40, h: 0.28, fs: 13, color: C.tealDark, bold: true, align: 'center' });
    T(s, c[1], { x: x + 0.75, y: y + 0.18, w: 3.45, h: 0.34, fs: 13.5, color: C.ink, bold: true });
    T(s, c[2], { x: x + 0.25, y: y + 0.72, w: 3.875, h: 0.82, fs: 9.2, color: C.gray, ls: 1.15 });
  });
})();

// ================= S5 政策东风 × 市场刚需 =================
(function () {
  const s = pres.addSlide(); s.background = { color: C.white };
  header(s, 'MARKET & POLICY', '政策东风 × 市场刚需｜智能化升级正当时', 5);
  T(s, '政策红利驱动', { x: 0.5, y: 1.24, w: 5, h: 0.30, fs: 12.5, color: C.ink, bold: true });
  T(s, '庞大市场体量', { x: 5.9, y: 1.24, w: 3.6, h: 0.30, fs: 12.5, color: C.ink, bold: true });
  B(s, 0.5, 1.58, 5.1, 2.84, C.dCard, C.line);
  B(s, 5.9, 1.58, 3.6, 2.84, C.dCard, C.line);
  const policies = [
    ['国家纲领引导', '《质量强国建设纲要》《"十四五"智能制造发展规划》明确普及智能在线检测，发展轻量化、模块化智能质检装备，智能质检上升为国家战略。'],
    ['广西财政补助', '制造业智能化改造项目，按设备及软件实际投入的20%给予专项财政补助，直接降低企业采购成本。'],
    ['柳州数字化券', '企业采购智能检测设备可抵扣30%采购费用，政策窗口期内采购意愿显著增强。']
  ];
  policies.forEach((p, i) => {
    const y = 1.86 + i * 0.88;
    CI(s, 0.78, y + 0.10, 0.30, C.tealLt);
    T(s, String(i + 1), { x: 0.78, y: y + 0.16, w: 0.30, h: 0.22, fs: 10.5, color: C.tealDark, bold: true, align: 'center' });
    T(s, p[0], { x: 1.22, y: y + 0.02, w: 4.2, h: 0.30, fs: 12, color: C.ink, bold: true });
    T(s, p[1], { x: 1.22, y: y + 0.34, w: 4.25, h: 0.52, fs: 9.3, color: C.gray, ls: 1.15 });
  });
  const stats = [
    ['10亿吨+', '2024年全国粗钢年产量 / 海量检测基数'],
    ['10万家+', '钢材加工贸易相关企业 / 中小微企业占比超90%'],
    ['近60亿元', '2024钢材智能质检市场 / 细分赛道空间广阔'],
    ['35%/年', '轻量化质检方案增速 / 全行业增长最快领域']
  ];
  stats.forEach((st, i) => {
    const y = 1.62 + i * 0.68;
    SQ(s, 6.05, y + 0.10, 0.04, 0.46, C.teal);
    T(s, st[0], { x: 6.22, y: y + 0.02, w: 1.42, h: 0.62, fs: 17, color: C.teal, bold: true, valign: 'middle' });
    T(s, st[1], { x: 7.72, y: y + 0.06, w: 1.72, h: 0.56, fs: 9.3, color: C.ink, valign: 'middle', ls: 1.15 });
  });
  band(s, 4.72, [rt('政策真金白银 + 海量中小钢企 = ', { bold: true, color: C.white }), rt('智能质检需求加速释放', { bold: true, color: C.teal })], { h: 0.38, fs: 11.5 });
})();

// ================= S6 痛点① =================
(function () {
  const s = pres.addSlide(); s.background = { color: C.white };
  header(s, 'PAIN POINT 01', '痛点①｜人工质检：漏检高 · 标准乱 · 成本贵 · 难追溯', 6);
  const cards = [
    [[rt('漏检率高达 '), rt('30%-40%', { bold: true, color: C.red })], '早期微锈、点锈与端面裂纹极其细微，肉眼难以辨识；质检员长时间目视作业极易视觉疲劳，漏检率随工时持续攀升；残次品流入后续生产，轻则返工、重则整批报废。'],
    ['标准不一争议多', '锈蚀等级判定高度依赖质检员个人经验与当天状态，同一批钢材不同人判出不同等级；供需双方常因等级认定分歧产生退货纠纷，企业质量口碑受损。'],
    ['人力成本持续攀升', '一名专职质检员年用工成本高达8-12万元；质检岗位枯燥、流动性大，中小工厂"招人难、留人更难"；订单旺季临时加人，新手误判率更高，形成恶性循环。'],
    ['数据零散无法追溯', '纸质记录易丢失、难统计，出现质量问题无法按批次追溯根源；管理层看不到质量趋势，工艺改进只能凭经验拍脑袋，精细化管理无从谈起。']
  ];
  const xs = [0.5, 5.125], ys = [1.36, 3.06], w = 4.375, h = 1.60;
  cards.forEach((cr, i) => {
    const x = xs[i % 2], y = ys[Math.floor(i / 2)];
    B(s, x, y, w, h, C.white, C.line);
    CI(s, x + 0.22, y + 0.20, 0.40, C.redLt, C.redLine);
    T(s, String(i + 1), { x: x + 0.22, y: y + 0.28, w: 0.40, h: 0.28, fs: 13, color: C.red, bold: true, align: 'center' });
    T(s, cr[0], { x: x + 0.75, y: y + 0.18, w: 3.45, h: 0.34, fs: 13.5, color: C.ink, bold: true });
    T(s, cr[1], { x: x + 0.25, y: y + 0.72, w: 3.875, h: 0.82, fs: 9.2, color: C.gray, ls: 1.15 });
  });
  band(s, 4.74, '× 四大短板叠加 = 中小钢企质检数字化全面卡壳', { h: 0.38, fill: C.redLt, color: C.red, fs: 12, dot: C.red });
})();

// ================= S7 痛点②③ =================
(function () {
  const s = pres.addSlide(); s.background = { color: C.white };
  header(s, 'PAIN POINTS 02 & 03', '痛点②③｜传统设备"买不起、用不好"，97%市场仍空白', 7);
  B(s, 0.5, 1.28, 5.5, 3.38, C.dCard, C.line);
  SQ(s, 0.72, 1.46, 0.07, 0.28, C.teal);
  T(s, '传统智能质检"三大门槛"', { x: 0.9, y: 1.42, w: 4.8, h: 0.34, fs: 12.5, color: C.ink, bold: true });
  const gates = [
    ['造价昂贵', '大型钢厂定制化方案单工位5-10万元，后期运维另计，中小厂一年利润买不起一套设备。'],
    ['部署笨重', '需改造产线、停产调试数天，与中小厂"多批次、小批量"的柔性生产节奏严重冲突。'],
    ['干扰误检', '传统单光源无法区分油污、水印与真实锈蚀，误报频繁，检出后仍需人工二次复核，省了人又没完全省。']
  ];
  gates.forEach((g, i) => {
    const y = 1.88 + i * 0.86;
    CI(s, 0.75, y + 0.06, 0.30, C.redLt, C.redLine);
    T(s, String(i + 1), { x: 0.75, y: y + 0.12, w: 0.30, h: 0.22, fs: 10.5, color: C.red, bold: true, align: 'center' });
    T(s, g[0], { x: 1.20, y: y + 0.02, w: 4.6, h: 0.30, fs: 11.5, color: C.ink, bold: true });
    T(s, g[1], { x: 1.20, y: y + 0.34, w: 4.62, h: 0.50, fs: 9.4, color: C.gray, ls: 1.15 });
  });
  B(s, 6.2, 1.28, 3.3, 3.38, C.navy);
  T(s, '< 3%', { x: 6.2, y: 1.56, w: 3.3, h: 0.75, fs: 44, color: C.teal, bold: true, align: 'center' });
  T(s, '中小钢企智能质检渗透率', { x: 6.2, y: 2.36, w: 3.3, h: 0.30, fs: 11, color: C.dSub, align: 'center' });
  SQ(s, 7.35, 2.80, 1.0, 0.015, C.navyLine);
  T(s, '超90%中小微钢企仍停留在纯人工质检阶段。"高端设备买不起、复杂设备用不好、低价设备精度差"是普遍困境——他们不是不想升级，是市场上没有为他们做的产品。',
    { x: 6.45, y: 2.98, w: 2.8, h: 1.50, fs: 9.2, color: C.dTxt, ls: 1.28, align: 'center' });
  band(s, 4.72, [rt('供需严重错配 = ', { bold: true, color: C.white }), rt('我们的精准切入机会', { bold: true, color: C.teal })], { h: 0.40, fs: 12 });
})();

// ================= S8 我们的答案 =================
(function () {
  const s = pres.addSlide(); s.background = { color: C.white };
  header(s, 'SOLUTION OVERVIEW', '我们的答案｜轻量化外挂式智能质检系统', 8);
  B(s, 0.5, 1.28, 9, 0.66, C.tealLt);
  T(s, [rt('打破行业"只优化算法"的固有思路，从', {}), rt('光学成像源头创新', { bold: true, color: C.tealDark }), rt('，结合轻量化 AI 算法，实现', {}), rt('低成本、高精度、易落地', { bold: true, color: C.tealDark }), rt('。', {})],
    { x: 0.8, y: 1.28, w: 8.4, h: 0.66, fs: 11.5, color: C.ink, valign: 'middle', ls: 1.25 });
  const feats = [
    ['0', '零改造', '跨装于现有输送带上方，不改动任何原有生产设备与产线布局，220V常规供电即插即用。'],
    ['2h', '快部署', '龙门式可调支架+软件参数化调试，2小时完成全套安装、标定与上线，全程不停产。'],
    ['∞', '广适配', '光源角度、亮度软件可调，一套设备覆盖板材、型材等多种钢材品类与不同车间光照环境。']
  ];
  const xs = [0.5, 3.567, 6.633], w = 2.867;
  feats.forEach((f, i) => {
    const x = xs[i];
    B(s, x, 2.08, w, 2.12, C.white, C.line);
    CI(s, x + (w - 0.60) / 2, 2.08 + 0.20, 0.60, C.teal);
    T(s, f[0], { x: x + (w - 0.60) / 2, y: 2.08 + 0.33, w: 0.60, h: 0.36, fs: 15, color: C.white, bold: true, align: 'center' });
    T(s, f[1], { x: x + 0.2, y: 2.08 + 0.92, w: w - 0.4, h: 0.34, fs: 13.5, color: C.ink, bold: true, align: 'center' });
    T(s, f[2], { x: x + 0.2, y: 2.08 + 1.30, w: w - 0.4, h: 0.76, fs: 9.3, color: C.gray, ls: 1.18, align: 'center' });
  });
  band(s, 4.55, [rt('龙门式外挂支架 + 软件参数化调试', { bold: true, color: C.teal }), rt(' —— 板材、型材等多品类钢材全覆盖，全程不停产', { bold: true, color: C.white })], { h: 0.52, fs: 11.5 });
})();

// ================= S9 核心技术① =================
(function () {
  const s = pres.addSlide(); s.background = { color: C.white };
  header(s, 'CORE TECH 01', '核心技术①｜多角度强光辅助成像——从光学源头创新', 9);
  B(s, 0.5, 1.26, 5.2, 3.34, C.dCard, C.line);
  SQ(s, 0.72, 1.44, 0.07, 0.28, C.teal);
  T(s, '技术原理', { x: 0.9, y: 1.38, w: 3, h: 0.34, fs: 13, color: C.ink, bold: true });
  const principles = [
    ['对称补光', '搭载4-6组LED条形光源，以15°-45°倾角对钢材表面差异化补光。'],
    ['反光差异', '粗糙锈蚀区产生漫反射呈暗斑，光滑基材呈均匀镜面反射，油污水印反光特征又与锈蚀不同。'],
    ['特征显形', '对比多角度图像，将肉眼难辨的0.1mm早期点锈转化为高对比度图像特征。']
  ];
  principles.forEach((p, i) => {
    const y = 1.90 + i * 0.88;
    SQ(s, 0.75, y + 0.12, 0.06, 0.42, C.teal);
    T(s, p[0], { x: 0.96, y: y, w: 4.5, h: 0.28, fs: 11.5, color: C.ink, bold: true });
    T(s, p[1], { x: 0.96, y: y + 0.30, w: 4.55, h: 0.54, fs: 9.3, color: C.gray, ls: 1.15 });
  });
  B(s, 5.9, 1.26, 3.6, 3.34, C.dCard, C.line);
  SQ(s, 6.12, 1.44, 0.07, 0.28, C.teal);
  T(s, '四大优势', { x: 6.3, y: 1.38, w: 3, h: 0.34, fs: 13, color: C.ink, bold: true });
  const advs = [
    ['检测下限极低', '灵敏度较传统方案提升3倍以上，微米级早期锈蚀无处遁形。'],
    ['抗干扰能力强', '从光学层区分锈蚀与油污/水印/灰尘，大幅降低后端算法压力。'],
    ['场景适配广泛', '软件调节光源参数，适配热轧板、冷轧板、型材等多种表面。'],
    ['硬件成本低廉', '民用工业级LED模组，无需昂贵特种光学设备。']
  ];
  advs.forEach((a, i) => {
    const y = 1.86 + i * 0.66;
    SQ(s, 6.15, y + 0.08, 0.06, 0.30, C.teal);
    T(s, a[0], { x: 6.34, y: y, w: 3.0, h: 0.26, fs: 11, color: C.ink, bold: true });
    T(s, a[1], { x: 6.34, y: y + 0.26, w: 3.0, h: 0.38, fs: 9.2, color: C.gray, ls: 1.1 });
  });
  band(s, 4.72, [
    rt('灵敏度 ', { bold: true, color: C.white }), rt('×3', { bold: true, color: C.teal }), rt('　|　', { color: '5E7C92' }),
    rt('抗干扰 ', { bold: true, color: C.white }), rt('↑', { bold: true, color: C.teal }), rt('　|　', { color: '5E7C92' }),
    rt('光学成本 ', { bold: true, color: C.white }), rt('↓', { bold: true, color: C.teal })
  ], { h: 0.38, fs: 12 });
})();

// ================= S10 核心技术② =================
(function () {
  const s = pres.addSlide(); s.background = { color: C.white };
  header(s, 'CORE TECH 02', '核心技术②｜"增强 → 分割 → 滤伪"三级联动算法', 10);
  const stages = [
    ['STAGE 01', '第一级 前置图像增强', '去噪矫正 · 多帧融合 · 直方图均衡化', '解决车间强光反光、画面模糊、对比度低问题，为检测提供高质量图像输入；运算量小，边缘端实时运行。'],
    ['STAGE 02', '第二级 改进 YOLOv8-seg', '小目标检测头 · CBAM 注意力 · Ghost 卷积', '针对点锈、微裂纹等极小目标专项优化；参数量减少 35%、推理速度提升 40%，低成本网关即可流畅运行。'],
    ['STAGE 03', '第三级 伪缺陷智能滤除', '12 维特征提取 · SVM 二次甄别', '提取面积、灰度、反光率、纹理熵等特征，结合自建工业伪缺陷样本库精准剔除油污划痕干扰，误检率 ≤ 4%。']
  ];
  const xs = [0.5, 3.567, 6.633], w = 2.867;
  stages.forEach((st, i) => {
    const x = xs[i];
    B(s, x, 1.28, w, 3.10, C.white, C.line);
    B(s, x + 0.2, 1.28 + 0.18, 1.24, 0.30, C.navy, null, 0.15);
    T(s, st[0], { x: x + 0.2, y: 1.28 + 0.18, w: 1.24, h: 0.30, fs: 9.5, color: C.white, bold: true, align: 'center', valign: 'middle' });
    T(s, st[1], { x: x + 0.2, y: 1.28 + 0.58, w: w - 0.4, h: 0.62, fs: 12, color: C.ink, bold: true });
    T(s, st[2], { x: x + 0.2, y: 1.28 + 1.24, w: w - 0.4, h: 0.55, fs: 9.3, color: C.tealDark, bold: true, ls: 1.15 });
    T(s, st[3], { x: x + 0.2, y: 1.28 + 1.86, w: w - 0.4, h: 1.05, fs: 9.2, color: C.gray, ls: 1.18 });
  });
  T(s, '→', { x: 3.435, y: 2.55, w: 0.13, h: 0.5, fs: 15, color: C.teal, bold: true, align: 'center' });
  T(s, '→', { x: 6.505, y: 2.55, w: 0.13, h: 0.5, fs: 15, color: C.teal, bold: true, align: 'center' });
  band(s, 4.56, [rt('全流程检测延迟 ', { bold: true, color: C.white }), rt('≤ 100ms', { bold: true, color: C.teal }), rt('，完美跟上高速产线节奏，兼顾高精度与计算速度。', { bold: true, color: C.white })], { h: 0.52, fs: 12 });
})();

// ================= S11 不止于锈 =================
(function () {
  const s = pres.addSlide(); s.background = { color: C.white };
  header(s, 'EXTENDED CAPABILITIES', '不止于锈｜国标等级判定 + 端面微裂纹检测', 11);
  const cards = [
    ['锈蚀等级智能判定', '严格对标GB/T 8923.1-2011国家标准，基于像素级分割结果精准计算锈蚀面积占比、分布密度与形态特征；自动完成A/B/C/D四级判定，输出带依据的判定报告；支持企业按客户要求自定义判定阈值，对接不同采购标准。',
      ['GB/T 8923.1-2011 国标', 'A/B/C/D 四级判定', '阈值可自定义']],
    ['端面微裂纹检测', '采用"边缘提取初筛 + 深度学习复核"融合方案：先由传统视觉算法过滤氧化皮、加工划痕等干扰，再由轻量化深度学习模型精准判定；可识别低至0.05mm的微裂纹，同步估算裂纹长度与延伸趋势，提前拦截断裂风险。',
      ['0.05mm 微裂纹', '初筛 + 深度学习', '断裂风险提前拦截']]
  ];
  const xs = [0.5, 5.125], w = 4.375;
  cards.forEach((c, i) => {
    const x = xs[i], y = 1.28, h = 3.28;
    B(s, x, y, w, h, C.dCard, C.line);
    SQ(s, x + 0.24, y + 0.28, 0.07, 0.28, C.teal);
    T(s, c[0], { x: x + 0.44, y: y + 0.22, w: 3.6, h: 0.36, fs: 13.5, color: C.ink, bold: true });
    T(s, c[1], { x: x + 0.28, y: y + 0.78, w: w - 0.56, h: 1.35, fs: 9.8, color: C.gray, ls: 1.35 });
    const chips = c[2];
    chips.forEach((ch, j) => {
      const cx = x + 0.28 + j * 1.32;
      B(s, cx, y + h - 0.55, 1.22, 0.34, C.tealBg, C.tealLine, 0.17);
      T(s, ch, { x: cx, y: y + h - 0.55, w: 1.22, h: 0.34, fs: 8.6, color: C.tealDark, bold: true, align: 'center', valign: 'middle' });
    });
  });
})();

// ================= S12 系统组成 × 快速部署 =================
(function () {
  const s = pres.addSlide(); s.background = { color: C.white };
  header(s, 'SYSTEM ARCHITECTURE', '系统组成 × 快速部署｜三层架构 · 2 小时上线', 12);
  B(s, 0.5, 1.26, 4.6, 3.32, C.dCard, C.line);
  SQ(s, 0.72, 1.44, 0.07, 0.28, C.teal);
  T(s, '三层系统架构', { x: 0.9, y: 1.38, w: 3, h: 0.34, fs: 13, color: C.ink, bold: true });
  const layers = [
    ['硬件层（感知载体）', '多角度可调 LED 光源模组 · 高清工业面阵相机 · 国产化边缘计算网关 · 声光报警模块。'],
    ['软件层（智能核心）', '高精度算法引擎 · 全量数据管理系统 · 可视化操作平台，检测、判定、存储、报表一站式。'],
    ['应用服务层（价值延伸）', '现场部署调试 · 人员培训 · 算法迭代升级 · 定制开发 · 质量数据分析咨询。']
  ];
  layers.forEach((l, i) => {
    const y = 1.84 + i * 0.84;
    B(s, 0.72, y, 4.16, 0.76, C.white, C.line);
    T(s, l[0], { x: 0.9, y: y + 0.09, w: 3.8, h: 0.28, fs: 11, color: C.ink, bold: true });
    T(s, l[1], { x: 0.9, y: y + 0.38, w: 3.82, h: 0.36, fs: 9.0, color: C.gray, ls: 1.1 });
  });
  B(s, 5.3, 1.26, 4.7, 3.32, C.dCard, C.line);
  SQ(s, 5.52, 1.44, 0.07, 0.28, C.teal);
  T(s, '2 小时快部署 · 七步自动化流程', { x: 5.7, y: 1.38, w: 4.2, h: 0.34, fs: 12.5, color: C.ink, bold: true });
  T(s, '龙门架跨装输送带上方 1.2–1.5 米，220V 供电、全程不停产', { x: 5.55, y: 1.74, w: 4.2, h: 0.26, fs: 8.8, color: C.gray });
  const steps = ['传感器感应钢材到位', '多角度分时采图', '图像预处理增强', 'YOLO 像素级分割', '滤伪缺陷 + 国标定级', '合格放行 / 声光预警', '数据存档可追溯'];
  steps.forEach((st, i) => {
    const y = 2.02 + i * 0.36;
    CI(s, 5.55, y + 0.01, 0.24, C.teal);
    T(s, String(i + 1), { x: 5.55, y: y + 0.05, w: 0.24, h: 0.18, fs: 8.5, color: C.white, bold: true, align: 'center' });
    T(s, st, { x: 5.9, y, w: 3.9, h: 0.30, fs: 9.8, color: C.ink });
  });
  band(s, 4.72, [rt('全程检测延迟 ', { bold: true, color: C.white }), rt('≤ 100ms', { bold: true, color: C.teal }), rt('，零门槛融入现有流水线', { bold: true, color: C.white })], { h: 0.38, fs: 11.5 });
})();

// ================= S13 为中小工厂而生 =================
(function () {
  const s = pres.addSlide(); s.background = { color: C.white };
  header(s, 'FEASIBILITY & USABILITY', '为中小工厂而生｜装得上 · 用得起 · 管得好', 13);
  const cards = [
    ['装', '装得上', ['模块化外挂设计，不改变产线布局', '单工位占地 < 0.5㎡，普通电工即可完成接线', '机械旋钮粗调 + 软件细调，无需专业工程师驻场']],
    ['用', '用得起', ['单工位硬件成本 ≤ 8000 元', '整套落地价仅为传统方案 1/3', '叠加广西 20% 补助 + 柳州 30% 券，投入再降近半']],
    ['管', '管得好', ['批次报表 + 质量趋势图自动生成', '远程运维 + 在线算法升级，故障远程诊断', '数据本地存储，工厂数据资产完全自主']]
  ];
  const xs = [0.5, 3.567, 6.633], w = 2.867;
  cards.forEach((c, i) => {
    const x = xs[i], y = 1.28, h = 3.18;
    B(s, x, y, w, h, C.white, C.line);
    CI(s, x + (w - 0.58) / 2, y + 0.18, 0.58, C.teal);
    T(s, c[0], { x: x + (w - 0.58) / 2, y: y + 0.32, w: 0.58, h: 0.34, fs: 13, color: C.white, bold: true, align: 'center' });
    T(s, c[1], { x: x + 0.2, y: y + 0.88, w: w - 0.4, h: 0.36, fs: 14, color: C.ink, bold: true, align: 'center' });
    T(s, c[2].map(t => rt('· ' + t, { breakLine: true, bullet: { code: '2713' } }).options ? { text: '· ' + t, options: { breakLine: true } } : { text: '· ' + t, options: { breakLine: true } }),
      { x: x + 0.2, y: y + 1.36, w: w - 0.4, h: 1.70, fs: 9.4, color: C.gray, ls: 1.3 });
  });
  band(s, 4.62, [rt('广西 20% 技改补助 + 柳州 30% 数字化券 ', { bold: true, color: C.white }), rt('→', { bold: true, color: C.teal }), rt(' 客户实际投入再降近半', { bold: true, color: C.white })], { h: 0.44, fs: 11.5 });
})();

// ================= S14 市场空间 =================
(function () {
  const s = pres.addSlide(); s.background = { color: C.white };
  header(s, 'MARKET OPPORTUNITY', '市场空间｜60 亿存量 → 200 亿未来，轻量化赛道增速第一', 14);
  T(s, '市场容量跨越增长', { x: 0.5, y: 1.24, w: 3.5, h: 0.30, fs: 12, color: C.ink, bold: true });
  T(s, '市场结构三维度拆解', { x: 4.5, y: 1.24, w: 5, h: 0.30, fs: 12, color: C.ink, bold: true });
  const sizes = [
    ['近 60 亿元', '2024 年国内钢材智能质检细分市场规模。'],
    ['突破 200 亿元', '预计 2030 年市场规模，六年实现 3 倍+ 暴涨。'],
    ['超 400 亿元', '中小钢企设备更新 + 新增需求（超 200 万套），未来五年最核心增量。']
  ];
  sizes.forEach((c, i) => {
    const y = 1.62 + i * 0.98;
    B(s, 0.5, y, 3.8, 0.90, C.white, C.line);
    SQ(s, 0.5, y + 0.14, 0.05, 0.62, C.teal);
    T(s, c[0], { x: 0.72, y: y + 0.12, w: 3.3, h: 0.44, fs: 21, color: C.teal, bold: true });
    T(s, c[1], { x: 0.72, y: y + 0.58, w: 3.4, h: 0.30, fs: 9.3, color: C.gray });
  });
  const dims = [
    ['按产品类型分', [rt('高端定制 40%（难下沉）｜中端 35%（精度有限）｜'), rt('轻量化普惠 25%，增速 35% 最快——我们在这', { bold: true, color: C.tealDark })]],
    ['按技术路径分', [rt('传统机器视觉 55%（增速放缓）｜'), rt('AI 智能视觉 45%，正快速替代成为绝对主流', { bold: true, color: C.tealDark })]],
    ['按应用场景分', [rt('大型钢厂在线质检 40%（趋饱和）｜'), rt('中小厂来料质检 60%，需求分散旺盛、竞争薄弱——我们聚焦', { bold: true, color: C.tealDark })]]
  ];
  dims.forEach((dm, i) => {
    const y = 1.62 + i * 0.98;
    B(s, 4.5, y, 5.5, 0.90, C.white, C.line);
    T(s, dm[0], { x: 4.7, y: y + 0.10, w: 5.0, h: 0.28, fs: 10.5, color: C.ink, bold: true });
    T(s, dm[1], { x: 4.7, y: y + 0.40, w: 5.1, h: 0.46, fs: 9.3, color: C.gray, ls: 1.15 });
  });
  band(s, 4.72, [rt('大型钢企改造趋近饱和，占行业 ', { bold: true, color: C.ink }), rt('55%', { bold: true, color: C.tealDark }), rt(' 体量的中小微钢企仍处蓝海爆发前期', { bold: true, color: C.ink })], { h: 0.40, fill: C.tealLt, fs: 11.5, dot: C.teal });
})();

// ================= S15 竞争格局 =================
(function () {
  const s = pres.addSlide(); s.background = { color: C.white };
  header(s, 'COMPETITIVE LANDSCAPE', '竞争格局｜三类玩家各有短板，专属方案存在空白', 15);
  const cols = [[0.5, 2.0], [2.5, 3.4], [5.9, 3.6]];
  const heads = ['竞争主体', '主要优势', '明显短板'];
  B(s, 0.5, 1.30, 9, 0.46, C.navy);
  cols.forEach((c, i) => {
    T(s, heads[i], { x: c[0] + 0.15, y: 1.30, w: c[1] - 0.3, h: 0.46, fs: 11.5, color: C.white, bold: true, valign: 'middle' });
  });
  const rows = [
    [[[rt('头部视觉厂商', { bold: true })], '（海康威视、大恒图像）'], '品牌影响力强、技术成熟度高、产品线布局全面。', '聚焦中高端定制高价方案，单工位5-10万，忽视中小厂轻量化需求。'],
    [[[rt('垂直AI创业公司', { bold: true })], ''], '算法针对性强、检测精度高、技术迭代速度快。', '硬件成本高昂、部署过程复杂，仅服务大型企业，不具备基层下沉能力。'],
    [[[rt('地方小型设备厂商', { bold: true })], ''], '产品售价低廉、本地化响应极其迅速。', '技术积累薄弱、检测精度差、系统稳定性不足，完全无算法持续迭代能力。']
  ];
  rows.forEach((r, i) => {
    const y = 1.80 + i * 0.88;
    B(s, 0.5, y, 9, 0.80, i % 2 === 0 ? C.white : C.dCard, C.line);
    // 竞争主体
    T(s, r[0][0], { x: 0.65, y: y + 0.16, w: 1.75, h: 0.30, fs: 11.5, color: C.ink });
    T(s, r[0][1], { x: 0.65, y: y + 0.46, w: 1.75, h: 0.26, fs: 8.8, color: C.gray });
    // 优势 / 短板
    T(s, r[1], { x: 2.68, y: y + 0.18, w: 3.0, h: 0.48, fs: 9.6, color: C.gray, ls: 1.15 });
    T(s, r[2], { x: 6.08, y: y + 0.18, w: 3.2, h: 0.48, fs: 9.6, color: C.red, ls: 1.15 });
  });
  B(s, 0.5, 4.42, 9, 0.66, C.navy);
  T(s, [rt('我们的位置：锈鉴智检——定位 ', { bold: true, color: C.white }), rt('"高精度、低成本、快部署"', { bold: true, color: C.teal }), rt(' 平衡点', { bold: true, color: C.white })],
    { x: 0.8, y: 4.48, w: 8.4, h: 0.24, fs: 11.5 });
  T(s, '"高精度、低成本、快部署"三者兼得——当前行业高度分散、尚无绝对龙头，中小钢企专属质检方案存在明显市场空白，先发者将迅速建立区域标杆壁垒。',
    { x: 0.8, y: 4.74, w: 8.4, h: 0.32, fs: 8.6, color: C.dSub });
})();

// ================= S16 四维壁垒 =================
(function () {
  const s = pres.addSlide(); s.background = { color: C.white };
  header(s, 'COMPETITIVE ADVANTAGES', '四维壁垒｜技术 · 成本 · 部署 · 服务', 16);
  const cards = [
    ['01', '技术壁垒', [rt('独创多角度强光辅助成像，从光学源头突破行业瓶颈；配合三级联动算法，高反光、多污渍复杂工况下精度与抗干扰能力行业领先，', {}), rt('0.1mm', { bold: true }), rt('点锈、', {}), rt('0.05mm', { bold: true }), rt('裂纹稳定检出。', {})]],
    ['02', '成本壁垒', [rt('光学创新降低算力依赖，无需高端GPU；全国产化民用器件选型，单工位硬件 ', {}), rt('≤ 8000元', { bold: true }), rt('，整套落地价仅为传统方案 ', {}), rt('1/3', { bold: true }), rt('，中小厂一年省下的人工成本即可覆盖采购。', {})]],
    ['03', '部署壁垒', [rt('模块化外挂设计，不改产线、不停产，', {}), rt('2小时', { bold: true }), rt('全流程落地；参数化软件调试适配多品类钢材，同行定制方案需要数周的部署周期，我们当天上线。', {})]],
    ['04', '服务壁垒', '本地化团队深耕广西，上门安装调试、一对一操作培训、远程极速运维；比头部厂商标准化工单响应更贴身，比小厂商更有技术纵深，客户粘性转化为复购与转介绍。']
  ];
  const xs = [0.5, 5.125], ys = [1.30, 2.96], w = 4.375, h = 1.56;
  cards.forEach((c, i) => {
    const x = xs[i % 2], y = ys[Math.floor(i / 2)];
    B(s, x, y, w, h, C.white, C.line);
    CI(s, x + 0.22, y + 0.18, 0.40, C.tealLt);
    T(s, c[0], { x: x + 0.22, y: y + 0.26, w: 0.40, h: 0.26, fs: 11, color: C.tealDark, bold: true, align: 'center' });
    T(s, c[1], { x: x + 0.72, y: y + 0.16, w: 3.4, h: 0.34, fs: 13, color: C.ink, bold: true });
    T(s, c[2], { x: x + 0.24, y: y + 0.70, w: 3.875, h: 0.80, fs: 9.3, color: C.gray, ls: 1.18 });
  });
  band(s, 4.66, [rt('传统方案单工位 5-10 万元', { bold: true, color: 'EDA7A2' }), rt('  vs  ', { color: '8FA8BC' }), rt('锈鉴智检 ≤ 2 万元级整体落地', { bold: true, color: C.teal })], { h: 0.44, fs: 12 });
})();

// ================= S17 先标杆、后规模 =================
(function () {
  const s = pres.addSlide(); s.background = { color: C.white };
  header(s, 'TARGET CUSTOMERS', '先标杆、后规模｜目标客户清晰可触达', 17);
  const cards = [
    ['客', '目标客户画像', ['广西及周边（柳州、玉林、贵港、佛山辐射圈）中小型钢材加工、贸易企业', '来料检验依赖人工、质量纠纷频发', '有数字化意愿但预算有限', '单厂 1-3 个质检工位，年质检人力成本 10-30 万元']],
    ['标', '标杆打造路径', ['试点期免费 / 低价部署 3-5 家本地标杆客户，全程陪跑迭代', '输出《质检数字化对比报告》：漏检率、人力成本、纠纷率前后对比', '形成案例手册与视频素材，作为规模化推广的核心弹药']],
    ['赢', '为什么能赢', ['政府补贴背书，降低决策门槛', '7 天免费试用，先见效果再付费', '设备钱 < 一年质检员工资，算账即下单', '本地化服务随叫随到，消除"买了没人管"顾虑']]
  ];
  const xs = [0.5, 3.567, 6.633], w = 2.867;
  cards.forEach((c, i) => {
    const x = xs[i], y = 1.28, h = 3.20;
    B(s, x, y, w, h, C.white, C.line);
    CI(s, x + (w - 0.56) / 2, y + 0.16, 0.56, C.teal);
    T(s, c[0], { x: x + (w - 0.56) / 2, y: y + 0.29, w: 0.56, h: 0.34, fs: 13, color: C.white, bold: true, align: 'center' });
    T(s, c[1], { x: x + 0.2, y: y + 0.84, w: w - 0.4, h: 0.36, fs: 13.5, color: C.ink, bold: true, align: 'center' });
    T(s, c[2].map(t => ({ text: '· ' + t, options: { breakLine: true } })),
      { x: x + 0.2, y: y + 1.30, w: w - 0.4, h: 1.78, fs: 9.3, color: C.gray, ls: 1.3 });
  });
  band(s, 4.64, [rt('标杆案例 ', { bold: true, color: C.white }), rt('→', { bold: true, color: C.teal }), rt(' 区域口碑 ', { bold: true, color: C.white }), rt('→', { bold: true, color: C.teal }), rt(' 渠道复制，滚雪球式扩张', { bold: true, color: C.white })], { h: 0.46, fs: 12 });
})();

// ================= S18 商业模式 =================
(function () {
  const s = pres.addSlide(); s.background = { color: C.white };
  header(s, 'BUSINESS MODEL', '商业模式｜四位一体，"一次销售 + 长期服务"盈利闭环', 18);
  const cards = [
    ['硬件销售', '毛利约 40%', '三档配置差异化定价（基础/标准/高级版），规模化采购+自主集成控成本，贡献基础收入与现金流。'],
    ['软件授权', '毛利超 80% · 利润主源', '基础功能随硬件附赠，高级算法模块按年授权订阅，老客户续费8折，形成持续高粘性收入。'],
    ['增值服务', '长期增长', '部署调试按次、年度维保按年、定制开发按项目、质量数据咨询按次收费，粘住客户全生命周期。'],
    ['设备租赁', '低门槛获客', '按月收租、租期灵活，租满设备归客户，打消中小厂一次性投入顾虑，快速抢占市场。']
  ];
  const xs = [0.5, 5.125], ys = [1.26, 2.60], w = 4.375, h = 1.24;
  cards.forEach((c, i) => {
    const x = xs[i % 2], y = ys[Math.floor(i / 2)];
    B(s, x, y, w, h, C.white, C.line);
    SQ(s, x, y + 0.16, 0.05, 0.30, C.teal);
    T(s, c[0], { x: x + 0.22, y: y + 0.12, w: 2.6, h: 0.34, fs: 13, color: C.ink, bold: true });
    B(s, x + w - 1.55, y + 0.13, 1.38, 0.30, C.tealLt, null, 0.15);
    T(s, c[1], { x: x + w - 1.55, y: y + 0.13, w: 1.38, h: 0.30, fs: 8.8, color: C.tealDark, bold: true, align: 'center', valign: 'middle' });
    T(s, c[2], { x: x + 0.22, y: y + 0.56, w: 3.95, h: 0.62, fs: 9.3, color: C.gray, ls: 1.18 });
  });
  B(s, 0.5, 4.06, 9, 0.62, C.navy);
  SQ(s, 0.66, 4.23, 0.11, 0.11, C.teal);
  T(s, '盈利闭环逻辑', { x: 0.9, y: 4.10, w: 1.5, h: 0.26, fs: 10.5, color: C.teal, bold: true });
  T(s, '高性价比硬件引流获客 → 技术与服务赢得信任 → 软件订阅与增值服务提升粘性 → 可持续盈利与规模化发展。',
    { x: 0.9, y: 4.38, w: 8.3, h: 0.26, fs: 9.2, color: C.dTxt });
})();

// ================= S19 营销策略 =================
(function () {
  const s = pres.addSlide(); s.background = { color: C.white };
  header(s, 'MARKETING STRATEGY', '营销策略｜产品 · 价格 · 渠道 · 推广', 19);
  const cards = [
    ['Product', '产品策略', '差异化定位"低成本、高精度、快部署"；基础/标准/高级三档矩阵覆盖不同预算客户；每6-12个月迭代升级，软件功能远程推送，硬件模块可加装，让客户"买后不落后"。'],
    ['Price', '价格策略', '渗透定价快速抢占细分市场；阶梯定价鼓励多工位批量采购；"硬件+软件+服务"组合优惠包；主动协助客户申报政府技改补贴与数字化券，降低实际支付价。'],
    ['Place', '渠道策略', '直销团队攻坚本地标杆客户；经销商下沉钢材市场与产业园区；与钢材贸易商、产线集成商捆绑销售分成；官网+短视频线上获客引流，四线并行。'],
    ['Promotion', '推广策略', '3-5家标杆案例手册与对比视频；行业展会现场实机演示"肉眼vs设备"检出对比；7天免费试用零风险体验；老客户推荐裂变奖励，口碑滚雪球。']
  ];
  const xs = [0.5, 5.125], ys = [1.30, 2.96], w = 4.375, h = 1.56;
  cards.forEach((c, i) => {
    const x = xs[i % 2], y = ys[Math.floor(i / 2)];
    B(s, x, y, w, h, C.white, C.line);
    T(s, [rt(c[0], { color: C.tealDark, bold: true, fontSize: 9.5 }), rt('  ' + c[1], { bold: true })],
      { x: x + 0.22, y: y + 0.14, w: 3.9, h: 0.34, fs: 13, color: C.ink });
    T(s, c[2], { x: x + 0.22, y: y + 0.62, w: 3.9, h: 0.88, fs: 9.2, color: C.gray, ls: 1.18 });
  });
  band(s, 4.68, [rt('补贴借力 + 试用先行 + 标杆背书 ', { bold: true, color: C.white }), rt('=', { bold: true, color: C.teal }), rt(' 低预算高效获客', { bold: true, color: C.white })], { h: 0.44, fs: 12 });
})();

// ================= S20 产品矩阵 =================
(function () {
  const s = pres.addSlide(); s.background = { color: C.white };
  header(s, 'PRODUCT MATRIX', '产品矩阵｜三档配置 · 三级迭代', 20);
  T(s, '产品组合阶梯', { x: 0.5, y: 1.24, w: 4.5, h: 0.30, fs: 12.5, color: C.ink, bold: true });
  T(s, '研发迭代规划', { x: 5.3, y: 1.24, w: 4.5, h: 0.30, fs: 12.5, color: C.ink, bold: true });
  const products = [
    ['基础版', '单相机 + 4 组光源，专注表面锈蚀检测——价格敏感的微型厂首选。'],
    ['标准版', '双相机 + 6 组光源，锈蚀 + 端面裂纹同步检测——主力走量款。'],
    ['高级版', '加装 3D 模组 + 云端质量管理平台——面向多产线规模客户。']
  ];
  products.forEach((p, i) => {
    const y = 1.66 + i * 1.0;
    B(s, 0.5, y, 4.6, 0.90, C.white, C.line);
    CI(s, 0.76, y + 0.26, 0.38, C.tealLt);
    T(s, String(i + 1), { x: 0.76, y: y + 0.33, w: 0.38, h: 0.26, fs: 11, color: C.tealDark, bold: true, align: 'center' });
    T(s, p[0], { x: 1.28, y: y + 0.14, w: 3.6, h: 0.32, fs: 12, color: C.ink, bold: true });
    T(s, p[1], { x: 1.28, y: y + 0.48, w: 3.7, h: 0.40, fs: 9.4, color: C.gray, ls: 1.15 });
  });
  const rd = [
    ['短期目标', '优化核心算法、新增型材适配、完善报表功能，夯实标杆客户体验。'],
    ['中期目标', 'AI 自学习迭代、拓展管材/不锈钢版本、对接主流 ERP 系统。'],
    ['长期目标', '3D/多光谱融合检测、工业质检大模型、多语言版本布局海外。']
  ];
  rd.forEach((p, i) => {
    const y = 1.66 + i * 1.0;
    B(s, 5.3, y, 4.7, 0.90, C.white, C.line);
    CI(s, 5.56, y + 0.26, 0.38, C.tealLt);
    T(s, String(i + 1), { x: 5.56, y: y + 0.33, w: 0.38, h: 0.26, fs: 11, color: C.tealDark, bold: true, align: 'center' });
    T(s, p[0], { x: 6.08, y: y + 0.14, w: 3.6, h: 0.32, fs: 12, color: C.ink, bold: true });
    T(s, p[1], { x: 6.08, y: y + 0.48, w: 3.8, h: 0.40, fs: 9.4, color: C.gray, ls: 1.15 });
  });
})();

// ================= S21 实施路线 =================
(function () {
  const s = pres.addSlide(); s.background = { color: C.white };
  header(s, 'ROADMAP', '实施路线｜六步走，从原型到国际化', 21);
  const stages = [
    ['01', '启动', '（1-3月）', '核心团队组建完成，算法原型验证，关键指标实验室达标。'],
    ['02', '开发', '（3-6月）', '深入工厂需求调研，软硬件联调，完成首台工程样机。'],
    ['03', '试点', '（6-12月）', '3-5家标杆客户免费部署，真实产线陪跑迭代，沉淀案例。'],
    ['04', '推广', '（12-18月）', '产品标准化定型，渠道体系搭建，启动区域规模化销售。'],
    ['05', '扩张', '（18-36月）', '跨区域复制推广，品类拓展至管材/不锈钢，客户破200家。'],
    ['06', '生态', '（长期）', '多金属质检平台化，布局东南亚市场，构建工业质检生态。']
  ];
  const xs = [0.5, 3.567, 6.633], ys = [1.28, 2.84], w = 2.867, h = 1.46;
  stages.forEach((st, i) => {
    const x = xs[i % 3], y = ys[Math.floor(i / 3)];
    const row0 = Math.floor(i / 3) === 0;
    B(s, x, y, w, h, row0 ? C.tealBg : C.white, row0 ? C.tealLine : C.line);
    T(s, [rt(st[0] + ' ', { bold: true, color: C.teal }), rt(st[1], { bold: true })], { x: x + 0.2, y: y + 0.14, w: w - 0.4, h: 0.30, fs: 13, color: C.ink });
    T(s, st[2], { x: x + 0.2, y: y + 0.48, w: w - 0.4, h: 0.24, fs: 9.3, color: C.tealDark, bold: true });
    T(s, st[3], { x: x + 0.2, y: y + 0.76, w: w - 0.4, h: 0.62, fs: 9.3, color: C.gray, ls: 1.18 });
  });
  band(s, 4.52, '每 6-12 个月形成一个可验证里程碑，稳健高效推进', { h: 0.42, fill: C.tealBg, color: C.tealDark, fs: 11.5, dot: C.teal });
})();

// ================= S22 投资分析 =================
(function () {
  const s = pres.addSlide(); s.background = { color: C.white };
  header(s, 'FINANCIAL ANALYSIS', '投资分析｜双向可行 + 阶梯投入 5 万 → 20 万 → 50 万', 22);
  const xs = [0.5, 3.567, 6.633], w = 2.867;
  B(s, xs[0], 1.26, w, 2.40, C.white, C.line);
  CI(s, xs[0] + 0.2, 1.26 + 0.16, 0.48, C.tealLt);
  T(s, '技', { x: xs[0] + 0.2, y: 1.26 + 0.27, w: 0.48, h: 0.30, fs: 12, color: C.tealDark, bold: true, align: 'center' });
  T(s, '技术可行', { x: xs[0] + 0.82, y: 1.26 + 0.14, w: 2.0, h: 0.32, fs: 13, color: C.ink, bold: true });
  T(s, '准确率 ≥ 95%、误检率 ≤ 4%，关键指标实验室验证达标；原型机完成组装联调；多角度成像 + 轻量化算法技术路线完全自主；模块化软硬件选型成熟，具备规模化生产条件。',
    { x: xs[0] + 0.22, y: 1.26 + 0.82, w: w - 0.44, h: 1.55, fs: 9.0, color: C.gray, ls: 1.22 });
  B(s, xs[1], 1.26, w, 2.40, C.white, C.line);
  CI(s, xs[1] + 0.2, 1.26 + 0.16, 0.48, C.tealLt);
  T(s, '市', { x: xs[1] + 0.2, y: 1.26 + 0.27, w: 0.48, h: 0.30, fs: 12, color: C.tealDark, bold: true, align: 'center' });
  T(s, '市场可行', { x: xs[1] + 0.82, y: 1.26 + 0.14, w: 2.0, h: 0.32, fs: 13, color: C.ink, bold: true });
  T(s, '60亿存量市场，中小企业渗透率不足 3%，蓝海明确；地方政策与补贴到位，客户采购意愿强；普惠赛道竞争薄弱，无强势对手；"设备钱 < 一年质检员工资"，商业账算得过来。',
    { x: xs[1] + 0.22, y: 1.26 + 0.82, w: w - 0.44, h: 1.55, fs: 9.0, color: C.gray, ls: 1.22 });
  const fin = [
    ['种子期', '5 万元', '算法研发、原型制作、标杆试点验证'],
    ['成长期', '20 万元', '扩大产能、渠道建设、技术升级'],
    ['成熟期', '50 万元', '品类拓展、生态平台与海内外布局']
  ];
  B(s, xs[2], 1.26, w, 2.40, C.white, C.line);
  CI(s, xs[2] + 0.2, 1.26 + 0.16, 0.48, C.tealLt);
  T(s, '资', { x: xs[2] + 0.2, y: 1.26 + 0.27, w: 0.48, h: 0.30, fs: 12, color: C.tealDark, bold: true, align: 'center' });
  T(s, '三阶段资金规划', { x: xs[2] + 0.82, y: 1.26 + 0.14, w: 2.05, h: 0.32, fs: 13, color: C.ink, bold: true });
  fin.forEach((f, i) => {
    const y = 1.26 + 0.82 + i * 0.54;
    T(s, [rt(f[0] + '　', { bold: true }), rt(f[1], { bold: true, color: C.tealDark })], { x: xs[2] + 0.22, y: y, w: w - 0.4, h: 0.24, fs: 10.5, color: C.ink });
    T(s, f[2], { x: xs[2] + 0.22, y: y + 0.25, w: w - 0.4, h: 0.26, fs: 8.8, color: C.gray });
  });
  B(s, 0.5, 3.86, 9, 0.98, C.navy);
  T(s, '种子期 5 万元资金精准拆分', { x: 0.8, y: 3.96, w: 5, h: 0.30, fs: 10.5, color: C.teal, bold: true });
  const splits = [['研发费用', '2.25 万 · 45%'], ['硬件原型', '1.25 万 · 25%'], ['试点市场', '0.75 万 · 15%'], ['人力办公', '0.75 万 · 15%']];
  const chipXs = [0.5, 2.8, 5.1, 7.4], chipW = 2.1;
  splits.forEach((sp, i) => {
    const x = chipXs[i], y = 4.34;
    B(s, x, y, chipW, 0.40, C.navy2, C.navyLine, 0.08);
    T(s, [rt(sp[0] + ' ', { color: C.white }), rt(sp[1], { bold: true, color: C.teal })], { x, y, w: chipW, h: 0.40, fs: 9.5, align: 'center', valign: 'middle' });
  });
})();

// ================= S23 回报预期 =================
(function () {
  const s = pres.addSlide(); s.background = { color: C.white };
  header(s, 'FINANCIAL RETURN', '回报预期｜回收期 2-2.5 年，五年营收 60 万 → 1040 万', 23);
  const kpis = [
    ['5.2 万元', '种子期净利润', '小规模验证商业与盈利模式顺畅跑通。'],
    ['> 80 万元', '成长期累计净利', '标准化批量复制，开启规模化盈利。'],
    ['> 110 万元', '成熟期年净利', '平台化全面放量，获取稳健长效收益。']
  ];
  const xs = [0.5, 3.567, 6.633], w = 2.867;
  kpis.forEach((k, i) => {
    const x = xs[i], y = 1.24, h = 1.24;
    B(s, x, y, w, h, C.white, C.line);
    T(s, k[0], { x: x + 0.22, y: y + 0.14, w: w - 0.4, h: 0.46, fs: 21, color: C.teal, bold: true });
    T(s, k[1], { x: x + 0.22, y: y + 0.62, w: w - 0.4, h: 0.28, fs: 10.8, color: C.ink, bold: true });
    T(s, k[2], { x: x + 0.22, y: y + 0.90, w: w - 0.4, h: 0.30, fs: 8.8, color: C.gray });
  });
  T(s, '五年营收演进路径（单位：万元）', { x: 0.5, y: 2.66, w: 5, h: 0.26, fs: 9.5, color: C.gray });
  const bars = [60, 180, 400, 700, 1040];
  const colors = ['9FC1D4', '6F9AB8', '40759B', '1E5078', '0E325A'];
  const bx = [1.15, 2.85, 4.55, 6.25, 7.95], bw = 0.7, baseY = 4.40, maxH = 1.04;
  SQ(s, 0.9, baseY, 8.4, 0.015, 'C9D6DE');
  bars.forEach((v, i) => {
    const h = Math.max(0.08, v / 1040 * maxH);
    const y = baseY - h;
    B(s, bx[i], y, bw, h, colors[i], null, 0.03);
    T(s, String(v), { x: bx[i] - 0.1, y: y - 0.26, w: bw + 0.2, h: 0.22, fs: 9.5, color: C.navy, bold: true, align: 'center' });
    T(s, 'Y' + (i + 1), { x: bx[i], y: baseY + 0.06, w: bw, h: 0.20, fs: 9, color: C.gray, align: 'center' });
  });
  band(s, 4.78, '投资回收期仅 2-2.5 年，五年实现营收十七倍跨越！', { h: 0.34, fill: C.tealLt, color: C.tealDark, fs: 10.5, dot: C.teal });
})();

// ================= S24 生产管理 =================
(function () {
  const s = pres.addSlide(); s.background = { color: C.white };
  header(s, 'PRODUCTION & QUALITY', '生产管理｜轻资产 · 强质控 · 全生命周期', 24);
  const xs = [0.5, 3.567, 6.633], w = 2.867, y = 1.28, h = 3.18;
  B(s, xs[0], y, w, h, C.white, C.line);
  CI(s, xs[0] + 0.2, y + 0.16, 0.48, C.teal);
  T(s, '轻', { x: xs[0] + 0.2, y: y + 0.27, w: 0.48, h: 0.30, fs: 12, color: C.white, bold: true, align: 'center' });
  T(s, '轻资产布局', { x: xs[0] + 0.82, y: y + 0.16, w: 2.0, h: 0.32, fs: 13, color: C.ink, bold: true });
  T(s, '以组装集成与测试为主，核心零部件外采，不自建重资产产线；初期依托学校实验室与孵化基地完成小批量生产；中期在柳州/玉林设区域组装中心，贴近客户、降低物流与响应成本。',
    { x: xs[0] + 0.22, y: y + 0.80, w: w - 0.44, h: 2.30, fs: 9.3, color: C.gray, ls: 1.32 });
  B(s, xs[1], y, w, h, C.white, C.line);
  CI(s, xs[1] + 0.2, y + 0.16, 0.48, C.teal);
  T(s, '质', { x: xs[1] + 0.2, y: y + 0.27, w: 0.48, h: 0.30, fs: 12, color: C.white, bold: true, align: 'center' });
  T(s, '三级质控管控', { x: xs[1] + 0.82, y: y + 0.16, w: 2.0, h: 0.32, fs: 13, color: C.ink, bold: true });
  const qcs = [
    ['来料检验', '核心器件入库全检。'],
    ['过程检验', '组装关键工序双人复核。'],
    ['出厂检验', [rt('每台设备完成 ', {}), rt('72 小时', { bold: true, color: C.tealDark }), rt('连续运行测试方可发货，确保到厂即稳定运行。', {})]]
  ];
  qcs.forEach((q, i) => {
    const yy = y + 0.80 + i * 0.82;
    SQ(s, xs[1] + 0.22, yy + 0.08, 0.05, 0.30, C.teal);
    T(s, q[0], { x: xs[1] + 0.38, y: yy + 0.02, w: 2.3, h: 0.26, fs: 10.5, color: C.ink, bold: true });
    T(s, q[1], { x: xs[1] + 0.38, y: yy + 0.28, w: w - 0.62, h: 0.52, fs: 9.0, color: C.gray, ls: 1.15 });
  });
  B(s, xs[2], y, w, h, C.white, C.line);
  CI(s, xs[2] + 0.2, y + 0.16, 0.48, C.teal);
  T(s, '全', { x: xs[2] + 0.2, y: y + 0.27, w: 0.48, h: 0.30, fs: 12, color: C.white, bold: true, align: 'center' });
  T(s, '全生命周期闭环', { x: xs[2] + 0.82, y: y + 0.16, w: 2.0, h: 0.32, fs: 13, color: C.ink, bold: true });
  T(s, '标准接口设计支持旧设备升级，以旧换新政策促进迭代复购；上门安装培训 + 定期主动回访，全程无忧；',
    { x: xs[2] + 0.22, y: y + 0.78, w: w - 0.44, h: 1.05, fs: 9.3, color: C.gray, ls: 1.3 });
  const chips = ['旧设备可升级', '上门安装培训', '以旧换新', '研发·服务闭环'];
  chips.forEach((c, i) => {
    const cx = xs[2] + 0.22 + (i % 2) * 1.32, cy = y + h - 0.86 + Math.floor(i / 2) * 0.44;
    B(s, cx, cy, 1.22, 0.36, C.tealBg, C.tealLine, 0.18);
    T(s, c, { x: cx, y: cy, w: 1.22, h: 0.36, fs: 8.6, color: C.tealDark, bold: true, align: 'center', valign: 'middle' });
  });
  band(s, 4.64, [rt('轻资产模式 ', { bold: true, color: C.white }), rt('=', { bold: true, color: C.teal }), rt(' 低固定投入 ', { bold: true, color: C.white }), rt('=', { bold: true, color: C.teal }), rt(' 强抗风险能力', { bold: true, color: C.white })], { h: 0.44, fs: 12 });
})();

// ================= S25 风险可控 =================
(function () {
  const s = pres.addSlide(); s.background = { color: C.white };
  header(s, 'RISK MANAGEMENT', '风险可控｜六大风险 · 六项对策', 25);
  const risks = [
    ['技术迭代风险', '每年营收15%-20%投入研发，6-12个月快速迭代，保持技术绝对领先地位。'],
    ['市场竞争风险', '深耕中小钢企细分场景，以标杆案例与服务口碑构筑坚实的区域品牌壁垒。'],
    ['政策变动风险', '紧跟政策导向调整申报策略，持续提升产品本体价值，降低对地方补贴依赖。'],
    ['需求变化风险', '采用模块化软硬件架构，快速响应客户定制需求，小步快跑验证新品类。'],
    ['供应链风险', '核心零部件维持2-3家供应商备选，签订长协锁价锁产能，常备安全库存。'],
    ['财务风险', '推行全面预算管理，坚持轻资产运营，利用硬件销售保现金流，严控账期。']
  ];
  const xs = [0.5, 5.125], ys = [1.28, 2.50, 3.72], w = 4.375, h = 1.12;
  risks.forEach((r, i) => {
    const x = xs[i % 2], y = ys[Math.floor(i / 2)];
    B(s, x, y, w, h, C.white, C.line);
    CI(s, x + 0.2, y + 0.18, 0.32, C.redLt, C.redLine);
    T(s, '▲', { x: x + 0.2, y: y + 0.24, w: 0.32, h: 0.20, fs: 9, color: C.red, bold: true, align: 'center' });
    T(s, r[0], { x: x + 0.62, y: y + 0.14, w: 3.5, h: 0.30, fs: 12, color: C.red, bold: true });
    T(s, r[1], { x: x + 0.62, y: y + 0.48, w: 3.55, h: 0.58, fs: 9.3, color: C.gray, ls: 1.18 });
  });
})();

// ================= S26 团队组织 =================
(function () {
  const s = pres.addSlide(); s.background = { color: C.white };
  header(s, 'TEAM & ADVISORS', '团队组织｜能力互补 · 结构合理 · 导师护航', 26);
  B(s, 0.5, 1.26, 9, 1.02, C.tealBg, C.tealLine);
  CI(s, 0.82, 1.26 + 0.17, 0.68, C.teal);
  T(s, '谢', { x: 0.82, y: 1.26 + 0.29, w: 0.68, h: 0.46, fs: 18, color: C.white, bold: true, align: 'center' });
  T(s, '谢宇（项目总负责人）', { x: 1.72, y: 1.26 + 0.16, w: 7.5, h: 0.34, fs: 13.5, color: C.ink, bold: true });
  T(s, '统筹战略规划、技术决策与团队管理｜精通 AutoCAD、UG NX 软件设计，多次斩获 AI 创新大赛校级一等奖。',
    { x: 1.72, y: 1.26 + 0.54, w: 7.5, h: 0.32, fs: 9.8, color: C.gray });
  const members = [
    ['伍', '伍锦琦', '技术研发 / 市场运营', '精通 Python 与深度学习算法，专注缺陷识别模型研发，逻辑严密、执行力强。'],
    ['谢', '谢耀宽', '硬件工程 / 项目执行', '机电工程基础扎实，擅长设备硬件集成与现场联调，工程实践经验突出。'],
    ['曹', '曹晓倩', '人力行政经理', '沟通协调能力优秀，文字功底扎实，精通团队行政组织与后勤综合保障。'],
    ['何', '何军', '财务总监', '严谨细致，精通项目成本核算、财务风险控制与资金规划精细化管理。']
  ];
  const xs = [0.5, 2.8, 5.1, 7.4], w = 2.1, y = 2.44, h = 2.0;
  members.forEach((m, i) => {
    const x = xs[i];
    B(s, x, y, w, h, C.white, C.line);
    CI(s, x + (w - 0.52) / 2, y + 0.14, 0.52, 'F0F6F8', C.tealLine);
    T(s, m[0], { x: x + (w - 0.52) / 2, y: y + 0.25, w: 0.52, h: 0.34, fs: 13, color: C.tealDark, bold: true, align: 'center' });
    T(s, m[1], { x: x + 0.1, y: y + 0.74, w: w - 0.2, h: 0.30, fs: 12, color: C.ink, bold: true, align: 'center' });
    T(s, m[2], { x: x + 0.1, y: y + 1.02, w: w - 0.2, h: 0.26, fs: 8.8, color: C.tealDark, align: 'center' });
    T(s, m[3], { x: x + 0.17, y: y + 1.32, w: w - 0.34, h: 0.64, fs: 8.6, color: C.gray, ls: 1.15, align: 'center' });
  });
  B(s, 0.5, 4.56, 9, 0.55, C.navy);
  SQ(s, 0.66, 4.72, 0.11, 0.11, C.teal);
  T(s, [rt('指导教师 / 顾问团队：', { bold: true, color: C.white }), rt('陈旭红、赖道辉', { bold: true, color: C.teal })],
    { x: 0.9, y: 4.60, w: 8.3, h: 0.26, fs: 12 });
  T(s, '高级工程师、南航工学硕士坐镇，工业项目经验丰富——为团队提供实验室场地、关键技术攻关及产业资源全面支持。',
    { x: 0.9, y: 4.88, w: 8.3, h: 0.20, fs: 8.8, color: C.dSub });
})();

// ================= S27 为什么是我们 =================
(function () {
  const s = pres.addSlide(); s.background = { color: C.white };
  header(s, 'TEAM DNA', '为什么是我们？四大团队基因', 27);
  const cards = [
    ['01', '目标一致 · 使命驱动', '团队成员同专业同班级，因"用智能技术赋能中小制造"的共同愿景组队；目标高度一致，分工不分家，关键时刻顶得上。'],
    ['02', '跨域互补 · 能力均衡', '五人覆盖算法研发、硬件工程、市场运营、行政管理、财务风控全链路；创业所需的每个关键环节都有专人负责，构建完整创业闭环。'],
    ['03', '创新实干 · 执行力强', '敢于突破传统成像思维，从光学源头另辟蹊径；原型机亲手组装调试，算法一行行代码打磨，脚踏实地推动产品快速迭代落地。'],
    ['04', '导师护航 · 资源支撑', '依托广西科技师范学院高校平台，双导师全程指导；获得实验室场地、技术攻关、政策对接与产业资源多维支持，站在巨人肩膀上创业。']
  ];
  const xs = [0.5, 5.125], ys = [1.30, 2.90], w = 4.375, h = 1.48;
  cards.forEach((c, i) => {
    const x = xs[i % 2], y = ys[Math.floor(i / 2)];
    B(s, x, y, w, h, C.white, C.line);
    T(s, c[0], { x: x + 0.22, y: y + 0.14, w: 0.65, h: 0.34, fs: 14, color: C.teal, bold: true });
    T(s, c[1], { x: x + 0.85, y: y + 0.16, w: 3.35, h: 0.34, fs: 13, color: C.ink, bold: true });
    T(s, c[2], { x: x + 0.22, y: y + 0.62, w: 3.875, h: 0.82, fs: 9.3, color: C.gray, ls: 1.18 });
  });
  band(s, 4.52, [rt('学生团队 + 双导师 ', { bold: true, color: C.white }), rt('=', { bold: true, color: C.teal }), rt(' 敢创新 · 能落地 · 可持续', { bold: true, color: C.white })], { h: 0.44, fs: 12.5 });
})();

// ================= S28 三阶段战略（结尾） =================
(function () {
  const s = pres.addSlide(); s.background = { color: C.navy };
  header(s, 'FUTURE STRATEGY & VISION', '三阶段战略｜从技术试点到生态引领', 28, true);
  const stages = [
    ['1-2 年', '试点期 · 打磨落地', '完成产品工业化迭代；落地 3-5 家标杆客户并沉淀案例；年营收 60 万元，验证商业闭环。'],
    ['3-4 年', '扩张期 · 品类拓展', '覆盖板材/型材/管材多品类；客户超 200 家；区域市占率突破 5%，成为区域知名品牌。'],
    ['5 年 +', '生态期 · 平台出海', '拓展多金属质检平台；布局东南亚市场；年营收突破千万级，成为中小制造企业质检数字化伙伴。']
  ];
  const xs = [0.5, 3.567, 6.633], w = 2.867, y = 1.26, h = 2.14;
  stages.forEach((st, i) => {
    const x = xs[i];
    B(s, x, y, w, h, C.navy2, i === 2 ? { color: C.teal, width: 1.4 } : C.navyLine);
    T(s, st[0], { x: x + 0.22, y: y + 0.18, w: w - 0.44, h: 0.36, fs: 16, color: C.teal, bold: true });
    T(s, st[1], { x: x + 0.22, y: y + 0.60, w: w - 0.44, h: 0.36, fs: 13.5, color: C.white, bold: true });
    T(s, st[2], { x: x + 0.22, y: y + 1.04, w: w - 0.44, h: 1.0, fs: 9.4, color: C.dTxt, ls: 1.28 });
  });
  T(s, '锈鉴智检——让每一根钢材，都被精准看见', { x: 0.5, y: 3.60, w: 9, h: 0.50, fs: 24, color: C.white, bold: true, align: 'center' });
  T(s, '以光学创新降低智能质检门槛，以AI算法守护制造质量底线', { x: 0.5, y: 4.18, w: 9, h: 0.26, fs: 10.5, color: '9FB4C6', align: 'center' });
  T(s, '广西科技师范学院 先进制造工程学院 · 锈鉴智检创新研发团队｜负责人：谢宇｜指导教师：陈旭红、赖道辉',
    { x: 0.5, y: 4.60, w: 9, h: 0.24, fs: 8.8, color: C.dSub2, align: 'center' });
  T(s, '恳请各位评委老师批评指正，谢谢！', { x: 0.5, y: 4.84, w: 9, h: 0.22, fs: 9.5, color: C.dSub, align: 'center' });
})();

pres.author = '锈鉴智检创新研发团队';
pres.title = '锈鉴智检——低成本高精度钢材质检专家';
pres.subject = '中国国际大学生创新大赛（2024）拟推荐项目';

pres.writeFile({ fileName: '/workspace/deck/锈鉴智检-低成本高精度钢材质检专家.pptx' }).then(() => {
  console.log('OK: 28 slides written');
});