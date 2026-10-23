// 何军 · 职规赛成长赛道答辩 PPT 重建脚本（15 页，绿色主题）
const PptxGenJS = require('pptxgenjs');
const QRCode = require('qrcode');

(async () => {
  const pres = new PptxGenJS();
  pres.layout = 'LAYOUT_WIDE'; // 13.333 x 7.5 in

  const F = 'Microsoft YaHei';
  const SW = 13.333, SH = 7.5, MX = 0.6, CW = 12.133;
  // 原版绿色主题
  const C = {
    deep: '0D1B1E', deep2: '12242A', deepLine: '2A4A46', chip: '0F3D2C',
    green: '0E7A4F', greenDark: '0B5B3A', greenBright: '2ED573',
    greenLt: 'E3F4EA', greenBg: 'EAF6EF', greenLine: 'C9E6D4',
    ink: '212B28', gray: '5A6661', line: 'E4E3E0', bg: 'FAF9F6', card: 'FFFFFF',
    white: 'FFFFFF', dSub: 'A9BDB4', dSub2: '7E948A', dTxt: 'CBD9D1',
    soft: 'F7F9F7', softLine: 'E5EBE7'
  };

  // ---------- 工具 ----------
  function T(s, text, o = {}) {
    s.addText(text, {
      x: o.x, y: o.y, w: o.w, h: o.h,
      fontFace: o.font || F, fontSize: o.fs || 11,
      color: o.color || C.ink, bold: !!o.bold,
      align: o.align || 'left', valign: o.valign || 'top',
      margin: 0, lineSpacingMultiple: o.ls == null ? 1.15 : o.ls,
      fit: 'shrink'
    });
  }
  function B(s, x, y, w, h, fill, line, radius) {
    return s.addShape(pres.ShapeType.roundRect, {
      x, y, w, h,
      fill: { color: fill },
      line: line ? (typeof line === 'string' ? { color: line, width: 1 } : line) : { type: 'none' },
      rectRadius: radius == null ? 0.05 : radius,
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
    T(s, tag, { x: MX, y: 0.42, w: 8, h: 0.26, fs: 11, color: dark ? C.greenBright : C.green, bold: true });
    T(s, title, { x: MX, y: 0.76, w: CW, h: 0.62, fs: 27, color: dark ? C.white : C.ink, bold: true });
    T(s, String(page).padStart(2, '0'), { x: 11.9, y: 0.40, w: 0.83, h: 0.28, fs: 10, color: dark ? '5E8B78' : C.gray, align: 'right' });
  }

  // 生成真实二维码
  const URL = 'https://hejun54579com.netlify.app';
  await QRCode.toFile('assets/qr_512x512.png', URL, {
    width: 512, margin: 1,
    color: { dark: '#0B4F35', light: '#FFFFFF' },
    errorCorrectionLevel: 'M'
  });

  // ================= S1 封面 =================
  (function () {
    const s = pres.addSlide(); s.background = { color: C.deep };
    CI(s, 0.60, 0.63, 0.13, C.greenBright);
    T(s, '全国大学生职业规划大赛 · 成长赛道', { x: 0.85, y: 0.52, w: 8, h: 0.32, fs: 13, color: C.white });
    T(s, '2026.10', { x: 11.3, y: 0.54, w: 1.43, h: 0.3, fs: 12, color: C.dSub2, align: 'right' });
    T(s, 'CAREER GROWTH DECK', { x: 0.6, y: 1.66, w: 8, h: 0.32, fs: 13, color: C.greenBright, bold: true });
    T(s, '从被调剂到主动选择', { x: 0.6, y: 2.18, w: 12, h: 0.95, fs: 44, color: C.white, bold: true });
    SQ(s, 6.55, 3.16, 1.7, 0.07, C.greenBright);
    T(s, '—— AI 与机器人交叉领域的技术型产品经理成长路', { x: 0.6, y: 3.52, w: 12, h: 0.5, fs: 21, color: C.dTxt, bold: true });
    B(s, 0.6, 5.30, 12.133, 1.35, C.deep2, C.deepLine);
    SQ(s, 0.6, 5.30, 0.075, 1.35, C.greenBright);
    T(s, '“我要做的不是和机器人抢饭碗的人，而是设计和驾驭机器人的人。”', { x: 1.05, y: 5.55, w: 8.1, h: 0.75, fs: 13.5, color: C.white, valign: 'middle' });
    SQ(s, 9.45, 5.52, 0.012, 0.9, C.deepLine);
    T(s, '何军', { x: 9.75, y: 5.58, w: 2.7, h: 0.4, fs: 16, color: C.white, bold: true });
    T(s, '机器人工程 261 班 · 18 岁', { x: 9.75, y: 6.02, w: 2.7, h: 0.3, fs: 11, color: C.dSub2 });
  })();

  // ================= S2 三个关键词 =================
  (function () {
    const s = pres.addSlide(); s.background = { color: C.bg };
    header(s, '01 / PERSONAL PROFILE', '三个关键词，快速认识何军', 2);
    const cards = [
      ['</>', '动手派', '开学利用课余约一周，独立上线个人网站全流程。', '约 1 周独立上线'],
      ['"', '沟通者', '高中优秀班干、学习委员；作为校园合伙人推广产品。', '沟通表达见长'],
      ['∞', '长期主义', '面向短板打补丁，本学期起每日背单词攻坚英语。', '每日背单词攻坚']
    ];
    const xs = [0.6, 4.727, 8.853], w = 3.877, y = 2.05, h = 4.55;
    cards.forEach((c, i) => {
      const x = xs[i];
      B(s, x, y, w, h, C.card, C.line);
      CI(s, x + (w - 0.8) / 2, y + 0.42, 0.8, C.greenLt);
      T(s, c[0], { x: x + (w - 0.8) / 2, y: y + 0.62, w: 0.8, h: 0.44, fs: 15, color: C.greenDark, bold: true, align: 'center' });
      T(s, '关键词 0' + (i + 1), { x: x, y: y + 1.48, w, h: 0.3, fs: 10.5, color: C.green, bold: true, align: 'center' });
      T(s, c[1], { x: x, y: y + 1.88, w, h: 0.44, fs: 20, color: C.ink, bold: true, align: 'center' });
      T(s, c[2], { x: x + 0.3, y: y + 2.52, w: w - 0.6, h: 1.35, fs: 11, color: C.gray, ls: 1.55, align: 'center' });
      B(s, x + (w - 1.75) / 2, y + h - 0.75, 1.75, 0.4, C.greenBg, C.greenLine, 0.2);
      T(s, c[3], { x: x + (w - 1.75) / 2, y: y + h - 0.75, w: 1.75, h: 0.4, fs: 9.5, color: C.greenDark, bold: true, align: 'center', valign: 'middle' });
    });
  })();

  // ================= S3 四站路线图 =================
  (function () {
    const s = pres.addSlide(); s.background = { color: C.bg };
    header(s, '02 / GROWTH ROADMAP', '四站路线图：目标是怎么长出来的', 3);
    SQ(s, 1.35, 2.78, 10.63, 0.04, 'D3E5DA');
    const st = [
      ['起点', '被调剂的专业', '录取机器人工程，重新认知专业内涵。'],
      ['探索', '三月深度调研', '对话AI、查岗位要求、跟踪产业政策。'],
      ['行动', '实战与积累', '建网站、练沟通、补短板、三年规划。'],
      ['调整', '复盘与备份', '三信号校准方向，设立AI服务Plan B。']
    ];
    const cxs = [1.7, 4.6, 7.5, 10.4];
    st.forEach((c, i) => {
      const cx = cxs[i];
      CI(s, cx - 0.31, 2.47, 0.62, C.green);
      T(s, String(i + 1), { x: cx - 0.31, y: 2.58, w: 0.62, h: 0.4, fs: 15, color: C.white, bold: true, align: 'center' });
      T(s, c[0], { x: cx - 1.5, y: 3.35, w: 3, h: 0.42, fs: 16, color: C.ink, bold: true, align: 'center' });
      T(s, c[1], { x: cx - 1.5, y: 3.85, w: 3, h: 0.3, fs: 11, color: C.greenDark, bold: true, align: 'center' });
      T(s, c[2], { x: cx - 1.7, y: 4.28, w: 3.4, h: 0.75, fs: 10.5, color: C.gray, ls: 1.3, align: 'center' });
      if (i < 3) T(s, '→', { x: cxs[i] + 0.95, y: 2.56, w: 0.5, h: 0.4, fs: 14, color: C.green, bold: true, align: 'center' });
    });
    T(s, [rt('我的目标，本身就是 ', {}), rt('“调”', { bold: true, color: C.green }), rt(' 出来的。', {})],
      { x: 0.6, y: 5.9, w: CW, h: 0.5, fs: 14, color: C.gray, align: 'center' });
  })();

  // ================= S4 起点：被调剂 =================
  (function () {
    const s = pres.addSlide(); s.background = { color: C.bg };
    header(s, '03 / STARTING POINT', '起点：我的专业，是“被调剂”来的', 4);
    B(s, 0.6, 2.0, 7.7, 4.2, C.soft, C.softLine);
    SQ(s, 0.6, 2.0, 0.07, 4.2, C.green);
    T(s, '“', { x: 0.95, y: 1.82, w: 0.9, h: 0.95, fs: 48, color: C.green, bold: true });
    T(s, [
      rt('高考录取时，机器人工程并非第一志愿。', { breakLine: true }),
      rt('第一反应是：“以后要去修机器人吗？”')
    ], { x: 1.1, y: 2.95, w: 6.95, h: 1.9, fs: 17, color: C.ink, bold: true, ls: 1.7 });
    B(s, 8.55, 2.0, 4.18, 4.2, C.deep);
    T(s, '我的决定', { x: 8.95, y: 2.85, w: 3.4, h: 0.4, fs: 15, color: C.greenBright, bold: true });
    SQ(s, 8.95, 3.42, 0.6, 0.04, C.greenBright);
    T(s, '先看清楚，再下结论。', { x: 8.95, y: 3.75, w: 3.45, h: 1.2, fs: 21, color: C.white, bold: true, ls: 1.5 });
  })();

  // ================= S5 三个月深度调研 =================
  (function () {
    const s = pres.addSlide(); s.background = { color: C.bg };
    header(s, '04 / RESEARCH METHODOLOGY', '行动：拒绝躺平，展开三个月深度调研', 5);
    const cards = [
      ['问', '问 AI', '从生活问答升级到与 DeepSeek、ChatGPT、Claude 深度协作。', 'AI 深度协作'],
      ['查', '查岗位', '精细分析“机器人产品经理”与“AI产品经理”招聘需求。', '岗位画像分析'],
      ['看', '看行业', '持续追踪具身智能政策导向与前沿产业新闻发展。', '产业动态追踪']
    ];
    const xs = [0.6, 4.727, 8.853], w = 3.877, y = 2.05, h = 4.55;
    cards.forEach((c, i) => {
      const x = xs[i];
      B(s, x, y, w, h, C.card, C.line);
      CI(s, x + (w - 0.8) / 2, y + 0.42, 0.8, C.green);
      T(s, c[0], { x: x + (w - 0.8) / 2, y: y + 0.62, w: 0.8, h: 0.44, fs: 15, color: C.white, bold: true, align: 'center' });
      T(s, c[1], { x: x, y: y + 1.55, w, h: 0.44, fs: 18, color: C.ink, bold: true, align: 'center' });
      T(s, c[2], { x: x + 0.3, y: y + 2.25, w: w - 0.6, h: 1.7, fs: 11, color: C.gray, ls: 1.55, align: 'center' });
      B(s, x + (w - 1.75) / 2, y + h - 0.75, 1.75, 0.4, C.greenBg, C.greenLine, 0.2);
      T(s, c[3], { x: x + (w - 1.75) / 2, y: y + h - 0.75, w: 1.75, h: 0.4, fs: 9.5, color: C.greenDark, bold: true, align: 'center', valign: 'middle' });
    });
  })();

  // ================= S6 三个发现 =================
  (function () {
    const s = pres.addSlide(); s.background = { color: C.bg };
    header(s, '05 / KEY FINDINGS', '三个发现，替我汇聚出职业方向', 6);
    const cards = [
      '具身智能列入国家六大未来产业，写入政府工作报告。',
      '岗位核心要求拆解为：“懂技术、能沟通、有项目”。',
      '个人优势（沟通表达与AI协作）与市场需求高度契合。'
    ];
    const xs = [0.6, 4.727, 8.853], w = 3.877, y = 2.05, h = 3.55;
    cards.forEach((t, i) => {
      const x = xs[i];
      B(s, x, y, w, h, C.card, C.line);
      CI(s, x + (w - 0.55) / 2, y + 0.3, 0.55, C.greenLt);
      T(s, '0' + (i + 1), { x: x + (w - 0.55) / 2, y: y + 0.42, w: 0.55, h: 0.34, fs: 12, color: C.greenDark, bold: true, align: 'center' });
      T(s, t, { x: x + 0.35, y: y + 1.2, w: w - 0.7, h: 2.1, fs: 11.5, color: C.ink, bold: true, ls: 1.65, align: 'center' });
    });
    B(s, 0.6, 5.75, CW, 0.85, C.deep);
    B(s, 0.9, 5.95, 1.7, 0.5, C.chip, null, 0.1);
    T(s, '明确目标', { x: 0.9, y: 5.95, w: 1.7, h: 0.5, fs: 11, color: C.greenBright, bold: true, align: 'center', valign: 'middle' });
    T(s, 'AI 与机器人交叉领域的技术型产品经理', { x: 2.85, y: 5.95, w: 9.5, h: 0.5, fs: 16, color: C.white, bold: true, valign: 'middle' });
  })();

  // ================= S7 目标岗位 =================
  (function () {
    const s = pres.addSlide(); s.background = { color: C.bg };
    header(s, '06 / TARGET POSITION', '技术型产品经理：把技术“翻译”成产品', 7);
    const infos = [
      ['目标企业', '宇树科技等具身智能公司 / 包含机器人业务的科技龙头'],
      ['意向地域', '深圳、杭州、上海等产业集聚城市']
    ];
    infos.forEach((inf, i) => {
      const x = i === 0 ? 0.6 : 6.73, w = 6.0;
      B(s, x, 1.95, w, 1.5, C.card, C.line);
      SQ(s, x, 1.95 + 0.3, 0.06, 0.9, C.green);
      T(s, inf[0], { x: x + 0.3, y: 1.95 + 0.28, w: w - 0.6, h: 0.36, fs: 12, color: C.green, bold: true });
      T(s, inf[1], { x: x + 0.3, y: 1.95 + 0.72, w: w - 0.6, h: 0.62, fs: 12.5, color: C.ink, bold: true, ls: 1.3 });
    });
    const elems = [
      ['要素 01', '懂技术', '熟练协同AI大模型，理解机器人底层硬件与控制原理。'],
      ['要素 02', '能沟通', '具备跨部门表达协作能力，精准对接技术与需求。'],
      ['要素 03', '有项目', '持续积累真实落地实践，用可量化的作品说话。']
    ];
    const xs = [0.6, 4.727, 8.853], w = 3.877, y = 3.75, h = 2.85;
    elems.forEach((el, i) => {
      const x = xs[i];
      B(s, x, y, w, h, C.card, C.line);
      B(s, x + 0.3, y + 0.3, 1.35, 0.42, C.deep, null, 0.12);
      T(s, el[0], { x: x + 0.3, y: y + 0.3, w: 1.35, h: 0.42, fs: 10, color: C.greenBright, bold: true, align: 'center', valign: 'middle' });
      T(s, el[1], { x: x + 0.3, y: y + 0.92, w: w - 0.6, h: 0.44, fs: 17, color: C.ink, bold: true });
      T(s, el[2], { x: x + 0.3, y: y + 1.5, w: w - 0.6, h: 1.1, fs: 10.5, color: C.gray, ls: 1.5 });
    });
  })();

  // ================= S8 行业分析 =================
  (function () {
    const s = pres.addSlide(); s.background = { color: C.bg };
    header(s, '07 / INDUSTRY ANALYSIS', '具身智能：从数字思考者到物理行动者', 8);
    B(s, 0.6, 1.95, CW, 0.9, C.deep);
    B(s, 0.9, 2.15, 1.2, 0.5, C.chip, null, 0.1);
    T(s, '结论', { x: 0.9, y: 2.15, w: 1.2, h: 0.5, fs: 11, color: C.greenBright, bold: true, align: 'center', valign: 'middle' });
    T(s, 'AI 与机器人的交叉，是确定性极强的长期向上赛道', { x: 2.35, y: 2.13, w: 10, h: 0.55, fs: 15.5, color: C.white, bold: true, valign: 'middle' });
    const cols = [
      ['政策端', '国家战略支柱', '列入国家六大未来产业方向，连续写进政府工作报告。'],
      ['学术端', '清华张涛教授观点', '机器人正“从数字世界的思考者，变成物理世界的行动者”。'],
      ['产业端', '技术落地破局', 'OpenAI刷新模型能力，宇树大模型接入机械臂无人抓取。']
    ];
    const xs = [0.6, 4.727, 8.853], w = 3.877, y = 3.2, h = 3.35;
    cols.forEach((c, i) => {
      const x = xs[i];
      B(s, x, y, w, h, C.card, C.line);
      CI(s, x + 0.3, y + 0.32, 0.5, C.greenLt);
      T(s, String(i + 1), { x: x + 0.3, y: y + 0.44, w: 0.5, h: 0.34, fs: 12, color: C.greenDark, bold: true, align: 'center' });
      T(s, c[0], { x: x + 0.95, y: y + 0.32, w: 2.6, h: 0.42, fs: 16, color: C.ink, bold: true });
      T(s, c[1], { x: x + 0.3, y: y + 1.0, w: w - 0.6, h: 0.36, fs: 11.5, color: C.greenDark, bold: true });
      T(s, c[2], { x: x + 0.3, y: y + 1.5, w: w - 0.6, h: 1.7, fs: 11, color: C.gray, ls: 1.55 });
    });
  })();

  // ================= S9 未来十年 · 三个世界 =================
  (function () {
    const s = pres.addSlide(); s.background = { color: C.bg };
    header(s, '08 / FUTURE VISION', '未来十年，我的产品在这样三个世界里', 9);
    const worlds = [
      ['assets/world1_740x1046.png', '黑灯工厂', '机械臂精密组装机器人'],
      ['assets/world2_740x1046.png', '救援与服务', '特种机器人走进危险与日常'],
      ['assets/world3_740x1046.png', '人形机器人', '高密度 human-robot 协作共事']
    ];
    const xs = [0.6, 4.88, 9.16], w = 3.57, y = 1.9, h = 5.0;
    worlds.forEach((wd, i) => {
      const x = xs[i];
      s.addImage({ path: wd[0], x, y, w, h });
      T(s, wd[1], { x: x + 0.25, y: y + h - 0.85, w: w - 0.5, h: 0.42, fs: 16, color: C.white, bold: true });
      T(s, wd[2], { x: x + 0.25, y: y + h - 0.42, w: w - 0.5, h: 0.34, fs: 10, color: 'B9CFC4' });
    });
  })();

  // ================= S10 行动①：个人网站 =================
  (function () {
    const s = pres.addSlide(); s.background = { color: C.bg };
    header(s, '09 / ACTION EVIDENCE 01', '行动①：开学课余约一周，独立上线个人网站', 10);
    // 左卡：流程闭环
    B(s, 0.6, 1.95, 7.3, 5.0, C.soft, C.softLine);
    T(s, '全流程实践闭环', { x: 1.0, y: 2.3, w: 6, h: 0.42, fs: 16, color: C.ink, bold: true });
    const chips1 = [['策划', 0.85], ['AI协同写代码', 1.75], ['Netlify部署', 1.55]];
    let cx = 1.0;
    chips1.forEach((ch, i) => {
      B(s, cx, 2.95, ch[1], 0.5, C.greenBg, C.greenLine, 0.25);
      T(s, ch[0], { x: cx, y: 2.95, w: ch[1], h: 0.5, fs: 11, color: C.greenDark, bold: true, align: 'center', valign: 'middle' });
      cx += ch[1];
      if (i < chips1.length - 1) { T(s, '→', { x: cx, y: 2.95, w: 0.35, h: 0.5, fs: 12, color: C.green, bold: true, align: 'center' }); cx += 0.35; }
    });
    const chips2 = [['绑定域名', 1.3], ['动态维护', 1.3]];
    cx = 1.0;
    chips2.forEach((ch, i) => {
      B(s, cx, 3.62, ch[1], 0.5, C.greenBg, C.greenLine, 0.25);
      T(s, ch[0], { x: cx, y: 3.62, w: ch[1], h: 0.5, fs: 11, color: C.greenDark, bold: true, align: 'center', valign: 'middle' });
      cx += ch[1];
      if (i < chips2.length - 1) { T(s, '→', { x: cx, y: 3.62, w: 0.35, h: 0.5, fs: 12, color: C.green, bold: true, align: 'center' }); cx += 0.35; }
    });
    B(s, 1.0, 4.42, 6.5, 1.15, C.greenBg, C.greenLine);
    T(s, '人机协同范式', { x: 1.25, y: 4.56, w: 3, h: 0.3, fs: 11.5, color: C.greenDark, bold: true });
    const para = [['我描述现象', 1.4], ['AI定位问题', 1.4], ['我验证修复', 1.4]];
    cx = 1.25;
    para.forEach((p, i) => {
      B(s, cx, 4.95, p[1], 0.42, C.card, C.greenLine, 0.21);
      T(s, p[0], { x: cx, y: 4.95, w: p[1], h: 0.42, fs: 10, color: C.ink, bold: true, align: 'center', valign: 'middle' });
      cx += p[1];
      if (i < para.length - 1) { T(s, '→', { x: cx, y: 4.95, w: 0.3, h: 0.42, fs: 12, color: C.green, bold: true, align: 'center' }); cx += 0.3; }
    });
    const stats = [['约 1 周上线', 1.5], ['独立完成', 1.15], ['持续更新至今', 1.6]];
    let sx = 1.0;
    stats.forEach(st => {
      B(s, sx, 5.95, st[1], 0.5, C.deep, null, 0.25);
      T(s, st[0], { x: sx, y: 5.95, w: st[1], h: 0.5, fs: 10.5, color: C.greenBright, bold: true, align: 'center', valign: 'middle' });
      sx += st[1] + 0.2;
    });
    // 右卡：二维码
    B(s, 8.15, 1.95, 4.58, 5.0, C.card, C.line);
    T(s, '现场可扫码体验网站', { x: 8.15, y: 2.4, w: 4.58, h: 0.4, fs: 13, color: C.ink, bold: true, align: 'center' });
    s.addImage({ path: 'assets/qr_512x512.png', x: 9.27, y: 3.0, w: 2.34, h: 2.34 });
    T(s, URL, { x: 8.15, y: 5.62, w: 4.58, h: 0.34, fs: 9.8, color: C.greenDark, bold: true, align: 'center' });
    SQ(s, 9.35, 6.14, 1.6, 0.015, C.line);
    T(s, '从策划、开发到部署全流程独立完成', { x: 8.15, y: 6.26, w: 4.58, h: 0.3, fs: 9.5, color: C.gray, align: 'center' });
  })();

  // ================= S11 行动②：对话人 =================
  (function () {
    const s = pres.addSlide(); s.background = { color: C.bg };
    header(s, '10 / ACTION EVIDENCE 02', '行动②：做网站练对话AI，卖电脑练对话人', 11);
    B(s, 0.6, 2.0, CW, 1.15, C.greenLt);
    T(s, '“把复杂配置参数，讲成同学听得懂的语言”', { x: 0.6, y: 2.0, w: CW, h: 1.15, fs: 16, color: C.ink, bold: true, align: 'center', valign: 'middle' });
    const cards = [
      ['高', '高中经历', '任劳动委员 · 获评优秀班干'],
      ['大', '大学职务', '机器人工程261班学习委员'],
      ['商', '商业实践', '星辰科技校园合伙人促成成交']
    ];
    const xs = [0.6, 4.727, 8.853], w = 3.877, y = 3.6, h = 2.9;
    cards.forEach((c, i) => {
      const x = xs[i];
      B(s, x, y, w, h, C.card, C.line);
      CI(s, x + (w - 0.62) / 2, y + 0.4, 0.62, C.greenLt);
      T(s, c[0], { x: x + (w - 0.62) / 2, y: y + 0.54, w: 0.62, h: 0.4, fs: 14, color: C.greenDark, bold: true, align: 'center' });
      T(s, c[1], { x: x, y: y + 1.3, w, h: 0.4, fs: 15, color: C.ink, bold: true, align: 'center' });
      T(s, c[2], { x: x + 0.35, y: y + 1.9, w: w - 0.7, h: 0.75, fs: 11, color: C.gray, ls: 1.45, align: 'center' });
    });
  })();

  // ================= S12 行动③：补短板 =================
  (function () {
    const s = pres.addSlide(); s.background = { color: C.bg };
    header(s, '11 / SELF IMPROVEMENT', '行动③：不藏短板——用工具给性格打补丁', 12);
    const cards = [
      ['短板 01：性子急', '代码误操作致网站空白', '打补丁 · 立规矩', 'AI代码跑完必人工复核，关键操作严格列清单执行。'],
      ['短板 02：英语弱', '高考英语 53 分底子薄', '打补丁 · 日攻坚', '本学期起每日坚持背单词，攻坚下学期四级首考。']
    ];
    const xs = [0.6, 6.78], w = 5.95, y = 2.0, h = 4.6;
    cards.forEach((c, i) => {
      const x = xs[i];
      B(s, x, y, w, h, C.soft, C.softLine);
      B(s, x + 0.3, y + 0.3, 2.3, 0.52, 'ECF2ED', 'D5E4DA', 0.26);
      T(s, c[0], { x: x + 0.3, y: y + 0.3, w: 2.3, h: 0.52, fs: 10.5, color: '4A5751', bold: true, align: 'center', valign: 'middle' });
      T(s, c[1], { x: x + 0.35, y: y + 1.05, w: w - 0.7, h: 0.4, fs: 12.5, color: C.ink, bold: true });
      T(s, '↓', { x: x + 0.4, y: y + 1.52, w: 0.4, h: 0.4, fs: 17, color: C.green, bold: true });
      B(s, x + 0.3, y + 2.0, w - 0.6, 2.3, '0E5B40');
      T(s, c[2], { x: x + 0.55, y: y + 2.22, w: w - 1.1, h: 0.36, fs: 11, color: C.greenBright, bold: true });
      T(s, c[3], { x: x + 0.55, y: y + 2.72, w: w - 1.1, h: 1.35, fs: 12, color: C.white, ls: 1.55 });
    });
  })();

  // ================= S13 三年计划 =================
  (function () {
    const s = pres.addSlide(); s.background = { color: C.bg };
    header(s, '12 / ACTION PLAN', '三年三步：技能作品 → 竞赛实战 → 考研深化', 13);
    const steps = ['技能作品', '竞赛实战', '考研深化'];
    let cx2 = (SW - (1.7 * 3 + 0.55 * 2)) / 2;
    steps.forEach((st, i) => {
      B(s, cx2, 2.0, 1.7, 0.55, C.deep, null, 0.12);
      T(s, st, { x: cx2, y: 2.0, w: 1.7, h: 0.55, fs: 13, color: C.white, bold: true, align: 'center', valign: 'middle' });
      cx2 += 1.7;
      if (i < steps.length - 1) { T(s, '→', { x: cx2, y: 2.0, w: 0.55, h: 0.55, fs: 15, color: C.green, bold: true, align: 'center' }); cx2 += 0.55; }
    });
    const cols = [
      ['大一 · 技能与作品', '打牢基础根基', ['持续迭代优化个人网站', '每日背单词备战英语四级', '整理 PM 岗位学习清单']],
      ['大二 · 竞赛与实战', '工程项目破局', ['数模竞赛（学长邀约入队）', '机器人与双创类竞赛', '争取进入机器人实验室']],
      ['大三 · 考研与深化', '学术学历提升', ['启动机器人/AI方向考研', '用竞赛成果支撑复试', '寻找前沿企业实习机会']]
    ];
    const xs = [0.6, 4.727, 8.853], w = 3.877, y = 2.95, h = 3.65;
    cols.forEach((c, i) => {
      const x = xs[i];
      B(s, x, y, w, h, C.card, C.line);
      CI(s, x + 0.3, y + 0.32, 0.5, C.green);
      T(s, String(i + 1), { x: x + 0.3, y: y + 0.44, w: 0.5, h: 0.34, fs: 12, color: C.white, bold: true, align: 'center' });
      T(s, c[0], { x: x + 0.95, y: y + 0.3, w: 2.75, h: 0.42, fs: 14.5, color: C.ink, bold: true });
      T(s, c[1], { x: x + 0.3, y: y + 0.8, w: w - 0.6, h: 0.32, fs: 10.5, color: C.greenDark, bold: true });
      SQ(s, x + 0.3, y + 1.22, w - 0.6, 0.015, C.line);
      T(s, c[2].map(t => ({ text: '· ' + t, options: { breakLine: true } })),
        { x: x + 0.3, y: y + 1.45, w: w - 0.6, h: 2.05, fs: 10.5, color: C.gray, ls: 1.6 });
    });
  })();

  // ================= S14 风险与 Plan B =================
  (function () {
    const s = pres.addSlide(); s.background = { color: C.bg };
    header(s, '13 / RISK CONTROL', '动态调整机制与 Plan B 备选方案', 14);
    B(s, 0.6, 2.0, 7.0, 4.55, C.soft, C.softLine);
    T(s, '复盘机制：三个触发信号', { x: 1.0, y: 2.3, w: 6, h: 0.42, fs: 15, color: C.ink, bold: true });
    const signals = [
      '课程成绩与四级考核是否达标',
      '竞赛项目是否有阶段性产出',
      '目标岗位招聘需求是否发生变化'
    ];
    signals.forEach((sig, i) => {
      const y = 3.05 + i * 1.0;
      CI(s, 1.0, y + 0.08, 0.4, C.greenLt);
      T(s, String(i + 1), { x: 1.0, y: y + 0.16, w: 0.4, h: 0.26, fs: 10.5, color: C.greenDark, bold: true, align: 'center' });
      T(s, [rt('信号 ' + (i + 1) + '：', { bold: true, color: C.greenDark }), rt(sig)],
        { x: 1.6, y, w: 5.75, h: 0.5, fs: 11.5, color: C.ink, valign: 'middle' });
    });
    T(s, [rt('信号异常 → ', { bold: true, color: C.greenDark }), rt('及时调整路径', { color: C.gray })],
      { x: 1.0, y: 6.15, w: 6, h: 0.32, fs: 10.5 });
    B(s, 7.85, 2.0, 5.28, 4.55, C.deep);
    B(s, 8.15, 2.35, 2.0, 0.5, C.chip, null, 0.1);
    T(s, 'Plan B 备份方案', { x: 8.15, y: 2.35, w: 2.0, h: 0.5, fs: 11, color: C.greenBright, bold: true, align: 'center', valign: 'middle' });
    T(s, 'AI 应用服务', { x: 8.15, y: 3.2, w: 4.6, h: 0.5, fs: 18, color: C.white, bold: true });
    T(s, '聚焦 AI 内容产品制作、规划与展示包装。基于已验证的沟通与 AI 协作能力，与主线互为备份。',
      { x: 8.15, y: 3.95, w: 4.65, h: 1.6, fs: 11, color: C.dTxt, ls: 1.6 });
  })();

  // ================= S15 结尾 =================
  (function () {
    const s = pres.addSlide(); s.background = { color: C.deep };
    SQ(s, 6.42, 1.62, 0.5, 0.045, C.greenBright);
    T(s, '从被调剂到主动选择，我才刚刚出发', { x: 1.0, y: 2.15, w: 11.33, h: 0.75, fs: 30, color: C.white, bold: true, align: 'center' });
    T(s, '“我要做的不是和机器人抢饭碗的人，而是设计和驾驭机器人的人。”', { x: 1.0, y: 3.25, w: 11.33, h: 0.5, fs: 14, color: C.dSub, align: 'center' });
    T(s, '谢谢聆听，请各位老师批评指正', { x: 1.0, y: 4.35, w: 11.33, h: 0.55, fs: 14.5, color: C.white, align: 'center' });
    T(s, [rt('何军 · 机器人工程 261 班', {}), rt('　|　', { color: '4E6A5F' }), rt('hejun54579com.netlify.app', { color: C.greenBright })],
      { x: 1.0, y: 5.75, w: 11.33, h: 0.4, fs: 10.5, color: C.dSub2, align: 'center' });
  })();

  pres.author = '何军';
  pres.title = '何军 · 职规赛成长赛道答辩';
  pres.subject = '全国大学生职业规划大赛 · 成长赛道';

  await pres.writeFile({ fileName: '/workspace/hejun_work/deck/何军_职规赛成长赛道答辩_优化.pptx' });
  console.log('OK: 15 slides written');
})();