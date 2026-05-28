const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const cors = require('cors');
const pptxgen = require('pptxgenjs');
const path = require('path');
const fs = require('fs');
const console = require('console');

const app = express();
const upload = multer({ storage: multer.memoryStorage() });
app.use(cors());
app.use(express.static(path.join(__dirname)));

// ─── Column names ─────────────────────────────────────────────────────────────
const lineColumn = 'Jakou linku provozované společností ARRIVA vlaky s.r.o., využívám nejčastěji?';

// ─── Column config ────────────────────────────────────────────────────────────
const COLUMN_CONFIG = [
  // 1. Věkové kategorie dotazovaných
  {
    column: 'Do jaké věkové kategorie patřím',
    title: 'Věkové kategorie dotazovaných cestujících',
    chartType: 'pie',
    categories: ["6 - 18 let", "19 - 25 let", "26 - 40 let", "41 - 64 let"],
    labels: ["6–18", "19–25", "26–40", "41–64"],
  },
  // 2. Četnost cest
  {
    column: 'Vlakem ARRIVA cestuji',
    title: 'Jak často cestujete vlaky ARRIVA?',
    chartType: 'pie',
    categories: [
      "Pravidelně dojíždím (například v neděli a pátek)",
      "Přibližně jednou měsíčně",
      "Příležitostně, několikrát v roce",
      "V rámci denního dojíždění ",
    ],
    labels: ["Pravidelně (např. pá/ne)", "Přibližně jednou měsíčně", "Příležitostně, několikrát ročně", "Denní dojíždění"],
  },
  // 3. Hodnocení zákaznické péče ve vlacích ARRIVA
  {
    chartType: 'multiBar',
    title: 'Hodnocení zákaznické péče ve vlacích ARRIVA (průvodčí, čistota, jízdní řád)',
    columns: [
      'Jak hodnotím vstřícnost a ochotu průvodčích ve vlacích ARRIVA',
      'Jak hodnotím úroveň čistoty ve vlacích ARRIVA',
      'Jak hodnotím dodržování jízdního řádu vlaky ARRIVA',
    ],
    seriesLabels: ["Průvodčí", "Čistota", "Jízdní řád"],
    categories: [1, 2, 3, 4, 5],
    labels: ["1", "2", "3", "4", "5"],
  },
  // 4. Místo nákupu jízdenky
  {
    chartType: 'pie',
    column: 'Jak zpravidla nakupuji jízdenky',
    title: 'Místo nákupu jízdenky',
    categories: ['V aplikaci', 'Ve vlaku', 'Na pokladně', 'V e-shopu', 'V automatu '],
    labels: ['V aplikaci', 'Ve vlaku', 'Na pokladně', 'V e-shopu', 'V automatu'],
  },
  // 4b. Druhy jízdenek podle místa nákupu
  {
    chartType: 'stackedBar',
    title: 'Druhy jízdenek podle místa nákupu',
    xCol: 'Jak zpravidla nakupuji jízdenky',
    seriesCol: 'Jaké využívám druhy jízdenek',
    xValues: ['Na pokladně', 'V aplikaci', 'V automatu ', 'V e-shopu', 'Ve vlaku'],
    xLabels: ['Na pokladně', 'V aplikaci', 'V automatu', 'V e-shopu', 'Ve vlaku'],
    seriesValues: [
      'Zpáteční v tarifu ARRIVA',
      'Jednotlivé v tarifu ARRIVA',
      'Jednotlivé jízdenky ve Státním jednotném tarifu (ONETICKET)',
      'Jednotlivé jízdenky Integrovaných dopravních systémů (PID, IDOL, IREDO, DÚK, ID ZK, IDPK)',
      'Časové jízdenky ve Státním jednotném tarifu (ONETICKET)',
      'Časové jízdenky integrovaných dopravních systémů',
    ],
    seriesLabels: [
      'Zpáteční ARRIVA',
      'Jednotlivé ARRIVA',
      'Jednotlivé ONETICKET',
      'Jednotlivé IDS (PID, IDOL, IREDO…)',
      'Časové ONETICKET',
      'Časové IDS',
    ],
  },
  // 5. Prodejní místa se vzdálenou obsluhou
  {
    column: 'Využívám prodejní místa se vzdálenou obsluhou?',
    title: 'Využití prodejních míst se vzdálenou obsluhou',
    chartType: 'pie',
    categories: [
      "Využívám",
      "Nevyužívám, ve stanicích odkud/kam cestuji není k dispozici",
      "Nevyužívám, sleva mě nemotivuje",
    ],
    labels: ["Využívám vzdálenou obsluhu", "Není k dispozici v mých stanicích", "Nevyužívám – sleva mě nemotivuje"],
  },
  // 6. Hodnocení e-shopu ARRIVA
  {
    column: 'Jak hodnotím e-shop ARRIVA a nakupování v tomto e-shopu',
    title: 'Spokojenost s e-shopem ARRIVA',
    chartType: 'bar',
    categories: [1, 2, 3, 4, 5],
    labels: ["NEJHORŠÍ 1", "2", "3", "4", "NEJLEPŠÍ 5"],
  },
  // 7. Cestujete s jízdním kolem?
  {
    column: 'Cestuji někdy vlakem s jízdním kolem?',
    title: 'Cestování s jízdním kolem – využití a spokojenost',
    chartType: 'pie',
    categories: [
      "Necestuji s jízdním kolem",
      "Ano, ale služby v tomto ohledu mi nevyhovují",
      "Ano a jsem spokojen/a se službami vlaků ARRIVA",
    ],
    labels: ["Necestuji s kolem", "Cestuji s kolem – služby nevyhovují", "Cestuji s kolem – spokojen/a"],
  },
  // 8. Celkové hodnocení pohodlí přepravy ve vlacích ARRIVA
  {
    column: 'Jak hodnotím pohodlí přepravy ve vlacích ARRIVA (kvalita vlaků, sedadel, hluk během jízdy)',
    title: 'Spokojenost s pohodlím přepravy (kvalita vlaků, sedadla, hluk)',
    chartType: 'bar',
    categories: [1, 2, 3, 4, 5],
    labels: ["NEJHORŠÍ 1", "2", "3", "4", "NEJLEPŠÍ 5"],
  },
  // 9. WI-FI ve vlaku
  {
    column: 'Využíváte při svých jízdách vlakové Wi-Fi připojení?',
    title: 'Využití a spokojenost s Wi-Fi ve vlacích ARRIVA',
    chartType: 'pie',
    categories: [
      "Ano, využívám a jsem spokojen/a",
      "Ano, ale nejsem spokojen/a s její kvalitou",
      "Nevyužívám",
    ],
    labels: ["Využívám – spokojen/a s kvalitou", "Využívám – nespokojen/a s kvalitou", "Nevyužívám Wi-Fi"],
  },
  // 10. Ceny jízdenek
  {
    column: 'Jak jsem spokojený s cenami nabízených jízdenek',
    title: 'Spokojenost s cenami jízdenek',
    chartType: 'bar',
    categories: [1, 2, 3, 4, 5],
    labels: ["NEJHORŠÍ 1", "2", "3", "4", "NEJLEPŠÍ 5"],
  },
  // 11. Využití zákaznické linky, spokojenost s odpovědí
  {
    column: 'Nechal/a jsem si někdy poradit na zákaznické lince ARRIVA 725 100 725',
    title: 'Využití zákaznické linky ARRIVA a spokojenost s odpovědí',
    chartType: 'pie',
    categories: [
      "ano a byl/a jsem spokojen/a s odpovědí",
      "ano, ale nebyl/a jsem spokojen/a s odpovědí",
      "ne, informace jsem si zjistil/a vždy jinak",
    ],
    labels: ["Využil/a – spokojen/a s odpovědí", "Využil/a – nespokojen/a s odpovědí", "Nevyužil/a zákaznickou linku"],
  },
  // 12. Hodnocení nabízené kapacity míst k sezení
  {
    column: 'Jak hodnotím nabízenou kapacitu míst k sezení',
    title: 'Spokojenost s kapacitou míst k sezení',
    chartType: 'bar',
    categories: [1, 2, 3, 4, 5],
    labels: ["NEJHORŠÍ 1", "2", "3", "4", "NEJLEPŠÍ 5"],
  },
  // 13a. Spokojenost s časovými polohami vlaků
  {
    chartType: 'pie',
    column: 'Považujete časové polohy vlaků na linkách, kterými cestujete, za vyhovující',
    title: 'Jsou časové polohy vlaků na vašich linkách vyhovující?',
    categories: ['ano'],
    labels: ["Vyhovující", "Nevyhovující"],
    yesNo: true,
  },
  // 13b. Spokojenost s počtem vlaků
  {
    chartType: 'pie',
    column: 'Považujete nabízený počet vlaků na linkách, kterými cestujete za dostatečný?',
    title: 'Je nabízený počet vlaků na vašich linkách dostatečný?',
    categories: ['ano'],
    labels: ["Dostatečný", "Nedostatečný"],
    yesNo: true,
  },
  // 14. Cíle cest cestujících
  {
    chartType: 'multiSelect',
    column: 'Vlak ARRIVA využívám',
    title: 'Za jakým účelem využíváte vlaky ARRIVA?',
    categories: [
      "Pro cesty za volným časem (například turistika, divadlo, nakupování);",
      "Pro cesty do zaměstnání;",
      "Pro cesty do školy;",
      "Na služební cesty;",
      "Pro cesty na úřad či k lékaři;",
      "Pro cesty k péči o jinou osobu;",
      "Jiné;",
    ],
    labels: ["Volný čas", "Zaměstnání", "Škola", "Služební", "Úřad/Lékař", "Péče", "Jiné"],
  },
];

// ─── Layout ───────────────────────────────────────────────────────────────────
let chartX = [0.5, 3.5, 6.5];
let chartY = [1.1, 3.3];
let chartW = 3.0;
let chartH = 2.0;
let headerH = 0.5;
let labelY = [0.9, 3.1];

// ─── Colors ───────────────────────────────────────────────────────────────────
let titleColor   = "FFFFFF";
let headerColor  = "00becd";
let pieColors    = ["0047a5", "ff6e1d", "2d146e", "911d8b"];
let seriesColors = ["0047a5", "ff6e1d", "2d146e"];

// ─── Helper: add header bar + title to a slide ────────────────────────────────
function addHeader(pres, slide, title) {
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: '100%', h: headerH,
    fill: { color: headerColor },
    line: { color: headerColor },
  });
  slide.addText(title, {
    x: 0.2, y: 0, w: 9.6, h: headerH,
    fontSize: 20, bold: true, color: titleColor,
    valign: 'middle',
  });
}

// ─── Helper: add logo to bottom right of slide ────────────────────────────────
function addLogo(slide) {
  const logoPath = path.join(__dirname, 'logo.png');
  if (fs.existsSync(logoPath)) {
    slide.addImage({ path: logoPath, x: 8.3, y: 5.15, w: 1.5, h: 0.45, sizing: { type: 'contain', w: 1.5, h: 0.45 } });
  }
}

// ─── Helper: extract short line label from full line name ─────────────────────
function getLineLabel(fullName) {
  const match = fullName.match(/([A-Z]\d+[A-Z]?(?:\s*(?:jih|sever))?)/);
  return match ? match[1] : fullName.substring(0, 10);
}

// ─── Helper: add pie legend below charts ──────────────────────────────────────
function addPieLegend(pres, slide, labels) {
  const top = legendTopY(labels.length);
  labels.forEach((text, k) => {
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 0.2, y: top + k * legendPitch, w: 0.14, h: 0.14,
      fill: { color: pieColors[k % pieColors.length] },
      line: { color: pieColors[k % pieColors.length] },
    });
    slide.addText(text, {
      x: 0.42, y: top + k * legendPitch, w: 5.0, h: 0.18,
      fontSize: 7, color: '1A1A2E',
    });
  });
}

// ─── Helper: pie chart options ────────────────────────────────────────────────
function pieOpts(x, y, w, h, colors) {
  return {
    x, y, w, h,
    chartColors: colors || pieColors,
    showPercent: true,
    dataLabelColor: 'FFFFFF',
    dataLabelFontSize: 8,
    dataLabelFontBold: true,
    showLegend: false,
    shadow: { type: 'outer', blur: 4, offset: 2, angle: 45, color: '888888', opacity: 0.3 },
  };
}

// ─── Helper: bar chart options ────────────────────────────────────────────────
function barOpts(x, y, w, h, colors) {
  return {
    x, y, w, h,
    chartColors: colors || seriesColors,
    showValue: true,
    dataLabelFontSize: 8,
    dataLabelFormatCode: '0"%"',
    barDir: 'col',
    valAxisHidden: true,
    valGridLine: { style: 'none' },
    catAxisLabelFontSize: 8,
  };
}

// ─── Helper: legend y position ────────────────────────────────────────────────
const slideH = 6.03;
const legendPitch = 0.15;
const legendMargin = 0.5;
function legendTopY(itemCount) {
  return slideH - legendMargin - itemCount * legendPitch;
}

// ─── Helper: suppress zeros ───────────────────────────────────────────────────
function suppressZeros(vals) {
  return vals.map(v => v === 0 ? null : v);
}

// ─── Slide generator: standard pie or bar ────────────────────────────────────
function addCategorySlides(pres, lines, config) {
  const lineNames = Object.keys(lines);
  const slideCount = Math.ceil(lineNames.length / 6);

  for (let i = 0; i < slideCount; i++) {
    if (i === 0) pres.addSection({ title: config.title });
    const slide = pres.addSlide({ sectionTitle: config.title });
    addHeader(pres, slide, config.title);
    addLogo(slide);

    for (let j = 0; j < 6; j++) {
      const lineIndex = i * 6 + j;
      if (lineIndex >= lineNames.length) break;

      const fullName = lineNames[lineIndex];
      const responses = lines[fullName];
      const total = responses.length;
      const label = getLineLabel(fullName);

      let vals;
      if (config.yesNo) {
        const yes = responses.filter(r => String(r[config.column] || '').startsWith('Ano')).length;
        vals = [
          Math.round(yes / total * 100),
          Math.round((total - yes) / total * 100),
        ];
      } else {
        vals = config.categories.map(cat =>
          Math.round(responses.filter(r => r[config.column] === cat).length / total * 100)
        );
      }

      const col = j % 3;
      const row = Math.floor(j / 3);

      const avgText = config.chartType === 'bar'
        ? (() => {
            const avg = config.categories.reduce((sum, cat, i) =>
              sum + cat * (vals[i] || 0), 0) / 100;
            return isNaN(avg) ? '' : `  průměr: ${avg.toFixed(2)}`;
          })()
        : '';

      slide.addText([
        { text: label, options: { bold: true, color: '1A1A2E' } },
        { text: avgText, options: { bold: true, color: '00becd' } },
      ], {
        x: chartX[col], y: labelY[row], w: chartW, h: 0.2,
        fontSize: 9,
      });

      if (config.chartType === 'pie') {
        slide.addChart(pres.charts.PIE, [{
          name: label, labels: config.labels, values: suppressZeros(vals),
        }], pieOpts(chartX[col], chartY[row], chartW, chartH));
      } else if (config.chartType === 'bar') {
        slide.addChart(pres.charts.BAR, [{
          name: label, labels: config.labels, values: suppressZeros(vals),
        }], barOpts(chartX[col], chartY[row], chartW, chartH));
      }
    }

    if (config.chartType === 'pie') {
      addPieLegend(pres, slide, config.labels);
    }
  }
}

// ─── Slide generator: multiBar ───────────────────────────────────────────────
function addMultiBarSlides(pres, lines, config) {
  const lineNames = Object.keys(lines);
  const slideCount = Math.ceil(lineNames.length / 6);

  for (let i = 0; i < slideCount; i++) {
    if (i === 0) pres.addSection({ title: config.title });
    const slide = pres.addSlide({ sectionTitle: config.title });
    addHeader(pres, slide, config.title);
    addLogo(slide);

    for (let j = 0; j < 6; j++) {
      const lineIndex = i * 6 + j;
      if (lineIndex >= lineNames.length) break;

      const fullName = lineNames[lineIndex];
      const responses = lines[fullName];
      const total = responses.length;
      const label = getLineLabel(fullName);

      const col = j % 3;
      const row = Math.floor(j / 3);

      const seriesAvgs = config.columns.map(colName =>
        config.categories.reduce((sum, cat) =>
          sum + cat * responses.filter(r => r[colName] === cat).length, 0) / total
      );

      slide.addText([
        { text: label, options: { bold: true, color: '1A1A2E' } },
        ...config.seriesLabels.map((sl, si) => ({
          text: `  ${sl}: ${seriesAvgs[si].toFixed(2)}`,
          options: { bold: true, color: '00becd' },
        })),
      ], {
        x: chartX[col], y: labelY[row], w: chartW, h: 0.2,
        fontSize: 7,
      });

      const chartData = config.columns.map((colName, s) => ({
        name: config.seriesLabels[s],
        labels: config.labels,
        values: suppressZeros(config.categories.map(cat =>
          Math.round(responses.filter(r => r[colName] === cat).length / total * 100)
        )),
      }));

      slide.addChart(pres.charts.BAR, chartData, {
        ...barOpts(chartX[col], chartY[row], chartW, chartH),
        barGrouping: 'clustered',
        showValue: false,
        valAxisHidden: false,
        valAxisLabelFontSize: 7,
        valAxisLabelFormatCode: '0"%"',
        valGridLine: { style: 'solid', color: 'DDDDDD', pt: 0.5 },
      });
    }

    const top = legendTopY(config.seriesLabels.length);
    config.seriesLabels.forEach((text, k) => {
      slide.addShape(pres.shapes.RECTANGLE, {
        x: 0.2, y: top + k * legendPitch, w: 0.14, h: 0.14,
        fill: { color: seriesColors[k % seriesColors.length] },
        line: { color: seriesColors[k % seriesColors.length] },
      });
      slide.addText(text, {
        x: 0.42, y: top + k * legendPitch, w: 5.0, h: 0.18,
        fontSize: 7, color: '1A1A2E',
      });
    });
  }
}

// ─── Slide generator: multiSelect ────────────────────────────────────────────
function addMultiSelectSlides(pres, lines, config) {
  const lineNames = Object.keys(lines);
  const slideCount = Math.ceil(lineNames.length / 6);

  for (let i = 0; i < slideCount; i++) {
    if (i === 0) pres.addSection({ title: config.title });
    const slide = pres.addSlide({ sectionTitle: config.title });
    addHeader(pres, slide, config.title);
    addLogo(slide);

    for (let j = 0; j < 6; j++) {
      const lineIndex = i * 6 + j;
      if (lineIndex >= lineNames.length) break;

      const fullName = lineNames[lineIndex];
      const responses = lines[fullName];
      const total = responses.length;
      const label = getLineLabel(fullName);

      const col = j % 3;
      const row = Math.floor(j / 3);

      slide.addText(label, {
        x: chartX[col], y: labelY[row], w: chartW, h: 0.2,
        fontSize: 9, bold: true, color: '1A1A2E',
      });

      const vals = config.categories.map(cat =>
        Math.round(responses.filter(r => String(r[config.column] || '').includes(cat)).length / total * 100)
      );

      slide.addChart(pres.charts.BAR, [{
        name: label, labels: config.labels, values: suppressZeros(vals),
      }], barOpts(chartX[col], chartY[row], chartW, chartH));
    }
  }
}

// ─── Slide generator: stackedBar ─────────────────────────────────────────────
// Stacked bar chart: X-axis = xCol values, series = seriesCol values.
// Each bar shows what % of rows with that X value belong to each series value.
function addStackedBarSlides(pres, lines, config) {
  const lineNames = Object.keys(lines);
  const slideCount = Math.ceil(lineNames.length / 6);

  for (let i = 0; i < slideCount; i++) {
    if (i === 0) pres.addSection({ title: config.title });
    const slide = pres.addSlide({ sectionTitle: config.title });
    addHeader(pres, slide, config.title);
    addLogo(slide);

    for (let j = 0; j < 6; j++) {
      const lineIndex = i * 6 + j;
      if (lineIndex >= lineNames.length) break;

      const fullName = lineNames[lineIndex];
      const responses = lines[fullName];
      const label = getLineLabel(fullName);
      const col = j % 3;
      const row = Math.floor(j / 3);

      slide.addText(label, {
        x: chartX[col], y: labelY[row], w: chartW, h: 0.2,
        fontSize: 9, bold: true, color: '1A1A2E',
      });

      // One series per seriesValue; one data point per xValue.
      // Value = % of rows matching that xValue that also match that seriesValue.
      const chartData = config.seriesValues.map((sVal, si) => ({
        name: config.seriesLabels[si],
        labels: config.xLabels,
        values: config.xValues.map(xVal => {
          const atX = responses.filter(r => r[config.xCol] === xVal);
          if (atX.length === 0) return null;
          const count = atX.filter(r => r[config.seriesCol] === sVal).length;
          return Math.round(count / atX.length * 100) || null;
        }),
      }));

      slide.addChart(pres.charts.BAR, chartData, {
        x: chartX[col], y: chartY[row], w: chartW, h: chartH,
        chartColors: pieColors,
        barDir: 'col',
        barGrouping: 'stacked',
        showValue: true,
        dataLabelFontSize: 7,
        dataLabelFormatCode: '0"%"',
        valAxisHidden: true,
        valGridLine: { style: 'none' },
        catAxisLabelFontSize: 7,
        showLegend: false,
      });
    }

    // Legend for series
    const top = legendTopY(config.seriesLabels.length);
    config.seriesLabels.forEach((text, k) => {
      slide.addShape(pres.shapes.RECTANGLE, {
        x: 0.2, y: top + k * legendPitch, w: 0.14, h: 0.14,
        fill: { color: pieColors[k % pieColors.length] },
        line: { color: pieColors[k % pieColors.length] },
      });
      slide.addText(text, {
        x: 0.42, y: top + k * legendPitch, w: 8.5, h: 0.18,
        fontSize: 7, color: '1A1A2E',
      });
    });
  }
}


app.post('/upload', upload.single('file'), async (req, res) => {
  try {
    const wb = XLSX.read(req.file.buffer, { type: 'buffer' });
    const data = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]);

    // Normalize non-breaking spaces in column names
    const normalizedData = data.map(row => {
      const newRow = {};
      Object.keys(row).forEach(k => {
        newRow[k.replace(/\xa0/g, ' ')] = row[k];
      });
      return newRow;
    });

    // Group rows by line
    const normalizedLineColumn = lineColumn.replace(/\xa0/g, ' ');
    const lines = {};
    normalizedData.forEach(row => {
      const line = row[normalizedLineColumn];
      if (!line) return;
      if (!lines[line]) lines[line] = [];
      lines[line].push(row);
    });

    const pres = new pptxgen();

    COLUMN_CONFIG.forEach(config => {
      const c = JSON.parse(JSON.stringify(config));
      if (c.column) c.column = c.column.replace(/\xa0/g, ' ');
      if (c.columns) c.columns = c.columns.map(col => col.replace(/\xa0/g, ' '));
      if (c.xCol) c.xCol = c.xCol.replace(/\xa0/g, ' ');
      if (c.seriesCol) c.seriesCol = c.seriesCol.replace(/\xa0/g, ' ');

      switch (c.chartType) {
        case 'pie':
        case 'bar':
          addCategorySlides(pres, lines, c);
          break;
        case 'multiBar':
          addMultiBarSlides(pres, lines, c);
          break;
        case 'multiSelect':
          addMultiSelectSlides(pres, lines, c);
          break;
        case 'stackedBar':
          addStackedBarSlides(pres, lines, c);
          break;
      }
    });

    const buffer = await pres.write({ outputType: 'nodebuffer' });
    res.setHeader('Content-Disposition', 'attachment; filename=report.pptx');
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.presentationml.presentation');
    res.send(buffer);

  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

const port = process.env.PORT || 3000;
app.listen(port, () => console.log('running on ' + port));