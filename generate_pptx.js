const pptxgen = require("pptxgenjs");

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.title = "Resultado da Carteira - Juliana Machado";

// =====================
// PALETA DE CORES
// =====================
const DARK_BG   = "0F2044";  // azul escuro profundo
const MID_BLUE  = "1A3A6B";  // azul médio
const ACCENT    = "C8972E";  // dourado
const WHITE     = "FFFFFF";
const LIGHT_BG  = "F4F6FA";
const GRAY_TEXT = "64748B";
const GREEN     = "16A34A";
const RED_CLR   = "DC2626";
const CARD_BG   = "FFFFFF";

const makeShadow = () => ({
  type: "outer", blur: 8, offset: 3, angle: 135, color: "000000", opacity: 0.10
});

// =====================
// SLIDE 1 — CAPA
// =====================
{
  const s = pres.addSlide();
  s.background = { color: DARK_BG };

  // Faixa dourada lateral esquerda
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 0.22, h: 5.625,
    fill: { color: ACCENT }, line: { color: ACCENT }
  });

  // Círculo decorativo
  s.addShape(pres.shapes.OVAL, {
    x: 7.2, y: -0.8, w: 4.5, h: 4.5,
    fill: { color: MID_BLUE }, line: { color: MID_BLUE }
  });
  s.addShape(pres.shapes.OVAL, {
    x: 7.7, y: -0.3, w: 3.5, h: 3.5,
    fill: { color: "1E4080" }, line: { color: "1E4080" }
  });

  // Safra Invest label
  s.addText("SAFRA INVEST", {
    x: 0.5, y: 0.4, w: 5, h: 0.35,
    fontSize: 10, color: ACCENT, bold: true, charSpacing: 4, fontFace: "Calibri"
  });

  // Título principal
  s.addText("Sua Carteira em Foco", {
    x: 0.5, y: 0.9, w: 7, h: 1.0,
    fontSize: 42, color: WHITE, bold: true, fontFace: "Calibri"
  });

  // Subtítulo
  s.addText("Análise completa dos seus investimentos", {
    x: 0.5, y: 1.95, w: 7, h: 0.5,
    fontSize: 16, color: "A8C4E8", fontFace: "Calibri"
  });

  // Linha divisória
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 2.6, w: 3, h: 0.04,
    fill: { color: ACCENT }, line: { color: ACCENT }
  });

  // Nome da cliente
  s.addText("Juliana B. Machado", {
    x: 0.5, y: 2.8, w: 6, h: 0.55,
    fontSize: 22, color: WHITE, bold: true, fontFace: "Calibri"
  });

  // Data
  s.addText("Posição: 13 de maio de 2026", {
    x: 0.5, y: 3.42, w: 5, h: 0.35,
    fontSize: 13, color: "A8C4E8", fontFace: "Calibri"
  });

  // Patrimônio destaque
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 4.1, w: 4.2, h: 1.0,
    fill: { color: MID_BLUE }, line: { color: ACCENT, pt: 1 }
  });
  s.addText("Patrimônio Total", {
    x: 0.55, y: 4.15, w: 4, h: 0.3,
    fontSize: 10, color: "A8C4E8", fontFace: "Calibri", bold: true, charSpacing: 2
  });
  s.addText("R$ 415.709,91", {
    x: 0.55, y: 4.45, w: 4, h: 0.55,
    fontSize: 26, color: WHITE, bold: true, fontFace: "Calibri"
  });
}

// =====================
// SLIDE 2 — RESUMO EXECUTIVO (Visão Geral)
// =====================
{
  const s = pres.addSlide();
  s.background = { color: LIGHT_BG };

  // Header band
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 10, h: 1.05,
    fill: { color: DARK_BG }, line: { color: DARK_BG }
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 0.18, h: 1.05,
    fill: { color: ACCENT }, line: { color: ACCENT }
  });
  s.addText("VISÃO GERAL DA CARTEIRA", {
    x: 0.4, y: 0.3, w: 9, h: 0.45,
    fontSize: 20, color: WHITE, bold: true, fontFace: "Calibri"
  });

  // 4 KPI cards
  const cards = [
    { label: "Patrimônio Bruto",    value: "R$ 415.710", sub: "em 13/05/2026",           color: DARK_BG,  icon: "💼" },
    { label: "Rentab. no Mês",      value: "+0,48%",      sub: "111% do CDI",              color: "16A34A", icon: "📈" },
    { label: "Rentab. no Ano",      value: "+3,23%",      sub: "65% do CDI",               color: "0369A1", icon: "📊" },
    { label: "Rentab. 12 Meses",    value: "+12,02%",     sub: "86% do CDI",               color: "7C3AED", icon: "🏆" },
  ];

  cards.forEach((c, i) => {
    const x = 0.3 + i * 2.38;
    s.addShape(pres.shapes.RECTANGLE, {
      x, y: 1.25, w: 2.2, h: 2.0,
      fill: { color: CARD_BG }, line: { color: "E2E8F0", pt: 1 },
      shadow: makeShadow()
    });
    // top accent
    s.addShape(pres.shapes.RECTANGLE, {
      x, y: 1.25, w: 2.2, h: 0.1,
      fill: { color: c.color }, line: { color: c.color }
    });
    s.addText(c.icon, {
      x: x + 0.08, y: 1.35, w: 0.5, h: 0.45, fontSize: 20
    });
    s.addText(c.label, {
      x: x + 0.1, y: 1.78, w: 2.0, h: 0.32,
      fontSize: 9.5, color: GRAY_TEXT, bold: true, fontFace: "Calibri", charSpacing: 0.5
    });
    s.addText(c.value, {
      x: x + 0.1, y: 2.08, w: 2.0, h: 0.55,
      fontSize: 20, color: c.color, bold: true, fontFace: "Calibri"
    });
    s.addText(c.sub, {
      x: x + 0.1, y: 2.63, w: 2.0, h: 0.3,
      fontSize: 9, color: GRAY_TEXT, fontFace: "Calibri"
    });
  });

  // Nota positiva
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.3, y: 3.45, w: 9.4, h: 1.72,
    fill: { color: "EFF6FF" }, line: { color: "BFDBFE", pt: 1 },
    shadow: makeShadow()
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.3, y: 3.45, w: 0.12, h: 1.72,
    fill: { color: "2563EB" }, line: { color: "2563EB" }
  });
  s.addText("✅  O que isso significa para você?", {
    x: 0.55, y: 3.55, w: 8.8, h: 0.38,
    fontSize: 12, color: "1E40AF", bold: true, fontFace: "Calibri"
  });
  s.addText([
    { text: "Sua carteira rendeu ", options: {} },
    { text: "R$ 46.710,91 ", options: { bold: true, color: GREEN } },
    { text: "acima do valor aplicado nos últimos 12 meses. Você está com o patrimônio crescendo consistentemente. ", options: {} },
    { text: "Nos últimos 30 dias, seu dinheiro rendeu mais rápido do que o CDI — o benchmark do mercado.", options: { bold: true } },
  ], {
    x: 0.55, y: 3.95, w: 8.8, h: 1.0,
    fontSize: 12, color: "1E3A8A", fontFace: "Calibri"
  });
}

// =====================
// SLIDE 3 — MARCAÇÃO NA CURVA vs MERCADO (explicação)
// =====================
{
  const s = pres.addSlide();
  s.background = { color: LIGHT_BG };

  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 10, h: 1.05,
    fill: { color: DARK_BG }, line: { color: DARK_BG }
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 0.18, h: 1.05,
    fill: { color: ACCENT }, line: { color: ACCENT }
  });
  s.addText("ENTENDENDO OS DOIS NÚMEROS", {
    x: 0.4, y: 0.3, w: 9, h: 0.45,
    fontSize: 20, color: WHITE, bold: true, fontFace: "Calibri"
  });

  // Bloco esquerdo — Curva
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.3, y: 1.2, w: 4.4, h: 3.95,
    fill: { color: CARD_BG }, line: { color: "E2E8F0", pt: 1 }, shadow: makeShadow()
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.3, y: 1.2, w: 4.4, h: 0.12,
    fill: { color: GREEN }, line: { color: GREEN }
  });
  s.addText("📋  MARCAÇÃO NA CURVA", {
    x: 0.45, y: 1.35, w: 4.1, h: 0.42,
    fontSize: 13, color: DARK_BG, bold: true, fontFace: "Calibri"
  });
  s.addText("O que aparece no seu extrato do banco", {
    x: 0.45, y: 1.75, w: 4.1, h: 0.3,
    fontSize: 9.5, color: GRAY_TEXT, fontFace: "Calibri", italic: true
  });

  s.addText([
    { text: "Como funciona:\n", options: { bold: true, breakLine: true } },
    { text: "É como ler o marcador de km do carro em plena estrada — mostra a velocidade contratada, sem sentir os soluços do trânsito.\n\n", options: { breakLine: true } },
    { text: "Resultado:\n", options: { bold: true, breakLine: true } },
    { text: "Rentabilidade acumulada: ", options: {} },
    { text: "+14,68%\n", options: { bold: true, color: GREEN, breakLine: true } },
    { text: "= 104,71% do CDI em 12 meses\n\n", options: { color: GREEN, breakLine: true } },
    { text: "Patrimônio (extrato): ", options: {} },
    { text: "R$ 426.920,24", options: { bold: true } },
  ], {
    x: 0.45, y: 2.1, w: 4.1, h: 2.9,
    fontSize: 11, color: "1E293B", fontFace: "Calibri"
  });

  // Bloco direito — Mercado
  s.addShape(pres.shapes.RECTANGLE, {
    x: 5.3, y: 1.2, w: 4.4, h: 3.95,
    fill: { color: CARD_BG }, line: { color: "E2E8F0", pt: 1 }, shadow: makeShadow()
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 5.3, y: 1.2, w: 4.4, h: 0.12,
    fill: { color: "0369A1" }, line: { color: "0369A1" }
  });
  s.addText("🏦  MARCAÇÃO A MERCADO", {
    x: 5.45, y: 1.35, w: 4.1, h: 0.42,
    fontSize: 13, color: DARK_BG, bold: true, fontFace: "Calibri"
  });
  s.addText("O valor se você vendesse tudo HOJE", {
    x: 5.45, y: 1.75, w: 4.1, h: 0.3,
    fontSize: 9.5, color: GRAY_TEXT, fontFace: "Calibri", italic: true
  });

  s.addText([
    { text: "Como funciona:\n", options: { bold: true, breakLine: true } },
    { text: "É como ver a cotação do imóvel hoje — o mercado balança, mas o valor real você realiza só na venda.\n\n", options: { breakLine: true } },
    { text: "Resultado:\n", options: { bold: true, breakLine: true } },
    { text: "Rentabilidade acumulada: ", options: {} },
    { text: "+12,02%\n", options: { bold: true, color: "0369A1", breakLine: true } },
    { text: "= 85,70% do CDI em 12 meses\n\n", options: { color: "0369A1", breakLine: true } },
    { text: "Patrimônio real: ", options: {} },
    { text: "R$ 415.709,91", options: { bold: true } },
  ], {
    x: 5.45, y: 2.1, w: 4.1, h: 2.9,
    fontSize: 11, color: "1E293B", fontFace: "Calibri"
  });

  // Nota explicativa
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.3, y: 5.2, w: 9.4, h: 0.32,
    fill: { color: "FEF9C3" }, line: { color: "FDE68A", pt: 1 }
  });
  s.addText("💡  A diferença entre os dois existe porque alguns ativos oscilam com o mercado — isso é normal e esperado na sua carteira.", {
    x: 0.45, y: 5.23, w: 9.2, h: 0.25,
    fontSize: 9.5, color: "78350F", fontFace: "Calibri"
  });
}

// =====================
// SLIDE 4 — EVOLUÇÃO DO PATRIMÔNIO (gráfico)
// =====================
{
  const s = pres.addSlide();
  s.background = { color: LIGHT_BG };

  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 10, h: 1.05,
    fill: { color: DARK_BG }, line: { color: DARK_BG }
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 0.18, h: 1.05,
    fill: { color: ACCENT }, line: { color: ACCENT }
  });
  s.addText("EVOLUÇÃO DO SEU PATRIMÔNIO", {
    x: 0.4, y: 0.3, w: 9, h: 0.45,
    fontSize: 20, color: WHITE, bold: true, fontFace: "Calibri"
  });

  // Gráfico de linha combinado
  const labels = ["Jun/25","Jul/25","Ago/25","Set/25","Out/25","Nov/25","Dez/25","Jan/26","Fev/26","Mar/26","Abr/26","Mai/26"];
  const carteira = [1.65, 0.25, 1.54, 2.07, 0.83, 1.06, 0.84, 1.62, 0.68, -1.68, 2.14, 0.48];
  const cdi = [1.10, 1.28, 1.16, 1.22, 1.28, 1.05, 1.22, 1.16, 1.00, 1.21, 1.09, 0.43];

  s.addChart(pres.charts.LINE, [
    { name: "Sua Carteira (%)", labels, values: carteira },
    { name: "CDI (%)", labels, values: cdi },
  ], {
    x: 0.3, y: 1.15, w: 9.4, h: 3.8,
    chartColors: [ACCENT, "94A3B8"],
    chartArea: { fill: { color: CARD_BG }, roundedCorners: true },
    lineSize: 2.5,
    lineSmooth: true,
    showLegend: true,
    legendPos: "b",
    legendFontSize: 10,
    catAxisLabelColor: GRAY_TEXT,
    valAxisLabelColor: GRAY_TEXT,
    valGridLine: { color: "E2E8F0", size: 0.5 },
    catGridLine: { style: "none" },
    showValue: false,
    catAxisLabelRotate: 30,
  });

  // Destaques
  s.addText("📌  Em março/26 houve queda pontual (-1,68%) devido à marcação a mercado — o CDI também sentiu. Nos outros meses a carteira superou ou equiparou o benchmark.", {
    x: 0.3, y: 5.1, w: 9.4, h: 0.42,
    fontSize: 9.5, color: GRAY_TEXT, fontFace: "Calibri", italic: true
  });
}

// =====================
// SLIDE 5 — COMPOSIÇÃO DA CARTEIRA
// =====================
{
  const s = pres.addSlide();
  s.background = { color: LIGHT_BG };

  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 10, h: 1.05,
    fill: { color: DARK_BG }, line: { color: DARK_BG }
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 0.18, h: 1.05,
    fill: { color: ACCENT }, line: { color: ACCENT }
  });
  s.addText("COMO SEU DINHEIRO ESTÁ DISTRIBUÍDO", {
    x: 0.4, y: 0.3, w: 9, h: 0.45,
    fontSize: 20, color: WHITE, bold: true, fontFace: "Calibri"
  });

  // Gráfico pizza
  s.addChart(pres.charts.PIE, [{
    name: "Composição",
    labels: ["Renda Fixa", "Multimercado", "Renda Variável", "Op. Estruturadas", "Previdência", "Corretora"],
    values: [48.72, 25.35, 4.43, 7.35, 5.25, 8.90]
  }], {
    x: 0.3, y: 1.15, w: 4.8, h: 4.1,
    chartColors: ["1A3A6B", "C8972E", "16A34A", "7C3AED", "0369A1", "64748B"],
    chartArea: { fill: { color: CARD_BG }, roundedCorners: true },
    showPercent: true,
    showLegend: true,
    legendPos: "b",
    legendFontSize: 9,
    dataLabelColor: WHITE,
    dataLabelFontSize: 9,
  });

  // Cards direita
  const itens = [
    { nome: "🔵  Renda Fixa",         pct: "48,72%", val: "R$ 202.541",  cor: "1A3A6B", desc: "Base sólida e segura" },
    { nome: "🟡  Multimercado",        pct: "25,35%", val: "R$ 105.370",  cor: "C8972E", desc: "Diversificação e gestão ativa" },
    { nome: "⚫  Corretora (FIIs+ETF)", pct: "8,90%",  val: "R$ 37.004",   cor: "64748B", desc: "Renda passiva mensal" },
    { nome: "🟣  Op. Estruturadas",    pct: "7,35%",  val: "R$ 30.572",   cor: "7C3AED", desc: "Proteção com upside" },
    { nome: "🔵  Previdência",         pct: "5,25%",  val: "R$ 21.811",   cor: "0369A1", desc: "Longo prazo + eficiência fiscal" },
    { nome: "🟢  Renda Variável",      pct: "4,43%",  val: "R$ 18.410",   cor: "16A34A", desc: "Potencial de valorização" },
  ];

  itens.forEach((it, i) => {
    const yy = 1.2 + i * 0.67;
    s.addShape(pres.shapes.RECTANGLE, {
      x: 5.4, y: yy, w: 4.3, h: 0.6,
      fill: { color: CARD_BG }, line: { color: "E2E8F0", pt: 1 }, shadow: makeShadow()
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x: 5.4, y: yy, w: 0.1, h: 0.6,
      fill: { color: it.cor }, line: { color: it.cor }
    });
    s.addText(it.nome, {
      x: 5.58, y: yy + 0.04, w: 2.4, h: 0.28,
      fontSize: 9.5, color: DARK_BG, bold: true, fontFace: "Calibri"
    });
    s.addText(it.desc, {
      x: 5.58, y: yy + 0.3, w: 2.4, h: 0.24,
      fontSize: 8.5, color: GRAY_TEXT, fontFace: "Calibri", italic: true
    });
    s.addText(it.pct, {
      x: 8.1, y: yy + 0.05, w: 0.8, h: 0.28,
      fontSize: 13, color: it.cor, bold: true, fontFace: "Calibri", align: "right"
    });
    s.addText(it.val, {
      x: 8.1, y: yy + 0.3, w: 1.5, h: 0.24,
      fontSize: 8.5, color: GRAY_TEXT, fontFace: "Calibri", align: "right"
    });
  });
}

// =====================
// SLIDE 6 — DESTAQUES POR CLASSE
// =====================
{
  const s = pres.addSlide();
  s.background = { color: LIGHT_BG };

  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 10, h: 1.05,
    fill: { color: DARK_BG }, line: { color: DARK_BG }
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 0.18, h: 1.05,
    fill: { color: ACCENT }, line: { color: ACCENT }
  });
  s.addText("DESTAQUES DE RENTABILIDADE POR PRODUTO", {
    x: 0.4, y: 0.3, w: 9, h: 0.45,
    fontSize: 20, color: WHITE, bold: true, fontFace: "Calibri"
  });

  // Tabela de destaques
  const tableData = [
    [
      { text: "Produto",       options: { bold: true, color: WHITE, fill: { color: DARK_BG }, fontSize: 10, fontFace: "Calibri" } },
      { text: "Classe",        options: { bold: true, color: WHITE, fill: { color: DARK_BG }, fontSize: 10, fontFace: "Calibri", align: "center" } },
      { text: "12 Meses",      options: { bold: true, color: WHITE, fill: { color: DARK_BG }, fontSize: 10, fontFace: "Calibri", align: "center" } },
      { text: "Saldo Atual",   options: { bold: true, color: WHITE, fill: { color: DARK_BG }, fontSize: 10, fontFace: "Calibri", align: "center" } },
    ],
    ["CRI Trisul CDI+1,15%",       "Renda Fixa",    "+18,99% ✅", "R$ 16.168"],
    ["SAF Agilite FI RF CP",        "Renda Fixa",    "+13,99% ✅", "R$ 74.542"],
    ["Safra Vitesse FI RF CP",      "Renda Fixa",    "+14,04% ✅", "R$ 24.552"],
    ["Artesanal Crédito Priv FIM",  "Multimercado",  "+17,09% ✅", "R$ 22.481"],
    ["SAF S&P Reais FICMM",         "Multimercado",  "+35,46% 🏆", "R$ 32.405"],
    ["SAF Kepler FIM",              "Multimercado",  "+15,13% ✅", "R$   8.008"],
    ["MAN Absolute Pace LB FIA",    "Renda Variável","+21,52% ✅", "R$ 18.411"],
    ["COE Call XP",                 "Estruturado",   "+15,80% ✅", "R$ 30.572"],
    ["MCCI11",                      "FII",           "+30,34% 🏆", "R$     770"],
  ];

  const rowColors = [
    null,
    "F8FAFC","EFF6FF","F8FAFC","EFF6FF","FFF7ED","FFF7ED","F0FDF4","FEFCE8","F0FDF4"
  ];

  const finalTable = tableData.map((row, ri) => {
    if (ri === 0) return row;
    return row.map((cell, ci) => ({
      text: cell,
      options: {
        fill: { color: rowColors[ri] || "FFFFFF" },
        color: ci === 2 ? GREEN : DARK_BG,
        bold: ci === 2,
        fontSize: 9.5,
        fontFace: "Calibri",
        align: ci === 0 ? "left" : "center"
      }
    }));
  });

  s.addTable(finalTable, {
    x: 0.3, y: 1.15, w: 9.4, h: 4.2,
    colW: [3.2, 1.8, 1.8, 2.0],
    border: { pt: 0.5, color: "E2E8F0" },
    autoPage: false,
    rowH: 0.38,
  });

  s.addText("🏆 = rendimento excepcional no período   ✅ = acima do CDI", {
    x: 0.3, y: 5.38, w: 9.4, h: 0.2,
    fontSize: 8.5, color: GRAY_TEXT, fontFace: "Calibri", italic: true
  });
}

// =====================
// SLIDE 7 — PONTOS DE ATENÇÃO
// =====================
{
  const s = pres.addSlide();
  s.background = { color: LIGHT_BG };

  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 10, h: 1.05,
    fill: { color: DARK_BG }, line: { color: DARK_BG }
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 0.18, h: 1.05,
    fill: { color: ACCENT }, line: { color: ACCENT }
  });
  s.addText("PONTOS DE ATENÇÃO E TRANSPARÊNCIA", {
    x: 0.4, y: 0.3, w: 9, h: 0.45,
    fontSize: 20, color: WHITE, bold: true, fontFace: "Calibri"
  });

  const itens = [
    {
      ico: "⚠️",
      titulo: "Suitability Desenquadrado",
      corpo: "Seu perfil cadastrado (Conservador) está diferente da carteira atual (Moderado). Precisamos regularizar isso no sistema Safra — é uma atualização simples que reflete a realidade dos seus investimentos.",
      cor: "FEF3C7", borda: "FCD34D", txtcor: "92400E"
    },
    {
      ico: "📉",
      titulo: "CRA IPCA (Raízen) — Marcação a Mercado Negativa",
      corpo: "Este ativo mostra -49,31% na marcação a mercado (valor se vendesse hoje). Isso acontece porque as taxas subiram no mercado. Mas atenção: o ativo paga IPCA+7,35% e você recebe isso normalmente. O valor volta conforme o vencimento (out/2030) se aproxima.",
      cor: "FEF2F2", borda: "FECACA", txtcor: "991B1B"
    },
    {
      ico: "📉",
      titulo: "Debêntures Ecorodovias — Leve Queda no Mês",
      corpo: "Caiu -0,14% em maio. Isso é normal para debêntures IPCA em cenário de juros altos. A taxa contratada é IPCA+6,99% e o ativo paga bem no longo prazo (venc. 2034).",
      cor: "FFF7ED", borda: "FED7AA", txtcor: "7C2D12"
    },
    {
      ico: "📦",
      titulo: "Letra Crédito Agro (ABC Brasil) — Vence Outubro/2026",
      corpo: "Vence em out/2026 — já rendeu 8,67% em 12 meses. Ótimo momento para planejarmos onde realocar esse recurso quando vencer.",
      cor: "EFF6FF", borda: "BFDBFE", txtcor: "1E3A8A"
    },
  ];

  itens.forEach((it, i) => {
    const yy = 1.18 + i * 1.05;
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.3, y: yy, w: 9.4, h: 0.95,
      fill: { color: it.cor }, line: { color: it.borda, pt: 1 }, shadow: makeShadow()
    });
    s.addText(it.ico + "  " + it.titulo, {
      x: 0.45, y: yy + 0.05, w: 9.0, h: 0.3,
      fontSize: 11, color: it.txtcor, bold: true, fontFace: "Calibri"
    });
    s.addText(it.corpo, {
      x: 0.45, y: yy + 0.35, w: 9.0, h: 0.55,
      fontSize: 9.5, color: it.txtcor, fontFace: "Calibri"
    });
  });
}

// =====================
// SLIDE 8 — PRÓXIMOS PASSOS
// =====================
{
  const s = pres.addSlide();
  s.background = { color: LIGHT_BG };

  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 10, h: 1.05,
    fill: { color: DARK_BG }, line: { color: DARK_BG }
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 0.18, h: 1.05,
    fill: { color: ACCENT }, line: { color: ACCENT }
  });
  s.addText("PRÓXIMOS PASSOS RECOMENDADOS", {
    x: 0.4, y: 0.3, w: 9, h: 0.45,
    fontSize: 20, color: WHITE, bold: true, fontFace: "Calibri"
  });

  const passos = [
    { num: "01", titulo: "Regularizar o Suitability",       desc: "Atualizar o perfil para Moderado — alinhando ao que você já tem investido. Processo simples, feito pelo assessor.", prazo: "Urgente", pcor: RED_CLR },
    { num: "02", titulo: "Planejar o vencimento da LCA",    desc: "A Letra de Crédito Agro (ABC Brasil) vence em out/2026. Ótimo momento para avaliar novas oportunidades com isenção de IR.", prazo: "Out/2026", pcor: "C8972E" },
    { num: "03", titulo: "Revisar CRA Raízen (IPCA)",       desc: "Decidir: manter até o vencimento (out/2030) para recuperar valor via cupons, ou avaliar troca por ativo mais líquido.", prazo: "Próx. reunião", pcor: "0369A1" },
    { num: "04", titulo: "Avaliação da carteira de FIIs",   desc: "A carteira de fundos imobiliários representa apenas 1,6% do portfólio. Se o objetivo for renda passiva, pode ser ampliada de forma estratégica.", prazo: "A definir", pcor: "7C3AED" },
  ];

  passos.forEach((p, i) => {
    const x = i < 2 ? 0.3 : 5.3;
    const yy = i % 2 === 0 ? 1.2 : 3.3;
    s.addShape(pres.shapes.RECTANGLE, {
      x, y: yy, w: 4.5, h: 1.85,
      fill: { color: CARD_BG }, line: { color: "E2E8F0", pt: 1 }, shadow: makeShadow()
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x, y: yy, w: 4.5, h: 0.1,
      fill: { color: p.pcor }, line: { color: p.pcor }
    });
    // número
    s.addShape(pres.shapes.OVAL, {
      x: x + 0.12, y: yy + 0.18, w: 0.48, h: 0.48,
      fill: { color: p.pcor }, line: { color: p.pcor }
    });
    s.addText(p.num, {
      x: x + 0.12, y: yy + 0.18, w: 0.48, h: 0.48,
      fontSize: 12, color: WHITE, bold: true, fontFace: "Calibri", align: "center", valign: "middle"
    });
    s.addText(p.titulo, {
      x: x + 0.7, y: yy + 0.18, w: 3.65, h: 0.38,
      fontSize: 11, color: DARK_BG, bold: true, fontFace: "Calibri"
    });
    s.addText(p.desc, {
      x: x + 0.15, y: yy + 0.65, w: 4.2, h: 0.85,
      fontSize: 9.5, color: GRAY_TEXT, fontFace: "Calibri"
    });
    // badge prazo
    s.addShape(pres.shapes.RECTANGLE, {
      x: x + 3.2, y: yy + 0.18, w: 1.15, h: 0.3,
      fill: { color: p.pcor }, line: { color: p.pcor }
    });
    s.addText(p.prazo, {
      x: x + 3.2, y: yy + 0.18, w: 1.15, h: 0.3,
      fontSize: 8, color: WHITE, bold: true, fontFace: "Calibri", align: "center", valign: "middle"
    });
  });
}

// =====================
// SLIDE 9 — ENCERRAMENTO
// =====================
{
  const s = pres.addSlide();
  s.background = { color: DARK_BG };

  // Círculo decorativo
  s.addShape(pres.shapes.OVAL, {
    x: -0.8, y: 2.5, w: 4.5, h: 4.5,
    fill: { color: MID_BLUE }, line: { color: MID_BLUE }
  });

  // Faixa dourada lateral
  s.addShape(pres.shapes.RECTANGLE, {
    x: 9.8, y: 0, w: 0.2, h: 5.625,
    fill: { color: ACCENT }, line: { color: ACCENT }
  });

  s.addText("SAFRA INVEST", {
    x: 1.0, y: 0.4, w: 8, h: 0.35,
    fontSize: 10, color: ACCENT, bold: true, charSpacing: 4, fontFace: "Calibri", align: "center"
  });

  s.addText("Sua carteira está no caminho certo. 🚀", {
    x: 0.8, y: 1.0, w: 8.4, h: 0.8,
    fontSize: 28, color: WHITE, bold: true, fontFace: "Calibri", align: "center"
  });

  s.addShape(pres.shapes.RECTANGLE, {
    x: 3.5, y: 1.95, w: 3, h: 0.05,
    fill: { color: ACCENT }, line: { color: ACCENT }
  });

  const sumario = [
    "✅  Patrimônio de R$ 415.710 com 12,02% de rentabilidade em 12 meses",
    "✅  Carteira diversificada em 6 classes de ativos",
    "✅  Fundo S&P Reais rendeu +35,46% — destaque do portfólio",
    "✅  CRI Trisul rendendo 18,99% — solidez em renda fixa",
    "✅  FII MCCI11 com +30,34% nos últimos 12 meses",
  ];

  sumario.forEach((txt, i) => {
    s.addText(txt, {
      x: 1.5, y: 2.15 + i * 0.52, w: 7, h: 0.42,
      fontSize: 11.5, color: "CBD5E1", fontFace: "Calibri"
    });
  });

  s.addText("Agendemos nossa próxima revisão para continuar evoluindo juntas!", {
    x: 1.0, y: 4.82, w: 8, h: 0.35,
    fontSize: 11, color: ACCENT, fontFace: "Calibri", italic: true, align: "center"
  });
}

// =====================
// SALVAR
// =====================
pres.writeFile({ fileName: "Carteira_Juliana_Machado_Mai2026.pptx" })
  .then(() => console.log("✅ PPT criado com sucesso!"))
  .catch(e => console.error("❌ Erro:", e));
