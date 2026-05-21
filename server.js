const express = require('express');
const multer  = require('multer');
const XLSX    = require('xlsx');
const path    = require('path');
const fs      = require('fs');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  WidthType, AlignmentType, BorderStyle, ShadingType, HeadingLevel,
  ImageRun, PageBreak, Header, Footer
} = require('docx');

const app    = express();
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 50 * 1024 * 1024 } });

app.use(express.json({ limit: '50mb' }));
app.use(express.static(path.join(__dirname, 'public')));

// ─────────────────────────────────────────────
// Helper: parse Excel/CSV buffer → array of rows
// ─────────────────────────────────────────────
function parseFile(buffer, originalname) {
  const wb  = XLSX.read(buffer, { type: 'buffer' });
  const ws  = wb.Sheets[wb.SheetNames[0]];
  return XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
}

// ─────────────────────────────────────────────
// Helper: detect header row and question columns
// ─────────────────────────────────────────────
function detectStructure(rows) {
  let headerIdx = 0;
  for (let i = 0; i < Math.min(5, rows.length); i++) {
    const r = rows[i].map(c => String(c).trim());
    if (r.some(c => /سؤال|question|رقم|no\.|#/i.test(c)) ||
        r.filter(c => c.length > 3).length > 3) {
      headerIdx = i;
      break;
    }
  }
  const headers = rows[headerIdx].map(c => String(c).trim());
  const dataRows = rows.slice(headerIdx + 1).filter(r =>
    r.some(c => c !== '' && c !== null && c !== undefined)
  );
  return { headers, dataRows };
}

// ─────────────────────────────────────────────
// Helper: compute Likert stats for one question column
// ─────────────────────────────────────────────
function likertStats(values) {
  const counts = [0, 0, 0, 0, 0]; // 1..5
  let total = 0;
  for (const v of values) {
    const n = parseFloat(v);
    if (!isNaN(n) && n >= 1 && n <= 5) {
      counts[Math.round(n) - 1]++;
      total++;
    }
  }
  if (total === 0) return null;
  const mean = counts.reduce((s, c, i) => s + c * (i + 1), 0) / total;
  const variance = counts.reduce((s, c, i) => s + c * Math.pow(i + 1 - mean, 2), 0) / total;
  const pct = counts.map(c => +((c / total) * 100).toFixed(1));
  return { counts, pct, total, mean: +mean.toFixed(2), std: +Math.sqrt(variance).toFixed(2) };
}

// ─────────────────────────────────────────────
// Helper: classify mean
// ─────────────────────────────────────────────
function classify(mean) {
  if (mean >= 4.5) return { ar: 'مرتفعة جداً', en: 'Very High', color: '1A7340' };
  if (mean >= 3.5) return { ar: 'مرتفعة',      en: 'High',      color: '2E86AB' };
  if (mean >= 2.5) return { ar: 'متوسطة',       en: 'Moderate',  color: 'F0A500' };
  if (mean >= 1.5) return { ar: 'منخفضة',       en: 'Low',       color: 'E05C34' };
  return               { ar: 'منخفضة جداً', en: 'Very Low',  color: 'C0392B' };
}

// ─────────────────────────────────────────────
// API: /api/analyze  → returns JSON result
// ─────────────────────────────────────────────
app.post('/api/analyze', upload.array('files', 20), (req, res) => {
  try {
    const { surveyTitle, sectionName, sampleSize } = req.body;
    const allFiles = req.files || [];
    if (!allFiles.length) return res.status(400).json({ error: 'لم يتم رفع أي ملف' });

    const allStats = []; // { qLabel, stats }[]
    let totalRespondents = 0;

    for (const file of allFiles) {
      const rows = parseFile(file.buffer, file.originalname);
      if (!rows.length) continue;
      const { headers, dataRows } = detectStructure(rows);
      totalRespondents += dataRows.length;

      // find question columns (numeric-ish columns after first 2)
      const qCols = [];
      headers.forEach((h, i) => {
        if (i < 1) return; // skip ID col
        const sample = dataRows.slice(0, 10).map(r => parseFloat(r[i])).filter(n => !isNaN(n));
        if (sample.length > 0) qCols.push({ idx: i, label: h || `سؤال ${i}` });
      });

      qCols.forEach(({ idx, label }) => {
        const values = dataRows.map(r => r[idx]);
        const stats = likertStats(values);
        if (!stats) return;
        const existing = allStats.find(x => x.qLabel === label);
        if (existing) {
          // merge: weighted average
          const total2 = existing.stats.total + stats.total;
          const mean2  = (existing.stats.mean * existing.stats.total + stats.mean * stats.total) / total2;
          const pct2   = existing.stats.pct.map((p, i) => +((p * existing.stats.total + stats.pct[i] * stats.total) / total2).toFixed(1));
          const counts2 = existing.stats.counts.map((c, i) => c + stats.counts[i]);
          existing.stats = { ...existing.stats, mean: +mean2.toFixed(2), pct: pct2, counts: counts2, total: total2 };
        } else {
          allStats.push({ qLabel: label, stats });
        }
      });
    }

    if (!allStats.length) return res.status(400).json({ error: 'لم يتم العثور على بيانات رقمية في الملف' });

    const overallMean = +(allStats.reduce((s, q) => s + q.stats.mean, 0) / allStats.length).toFixed(2);
    const sampleN     = sampleSize || totalRespondents;

    res.json({
      surveyTitle: surveyTitle || 'استبيان',
      sectionName: sectionName || '',
      sampleSize:  sampleN,
      questions:   allStats,
      overallMean,
      classification: classify(overallMean),
      date: new Date().toLocaleDateString('ar-SA', { year: 'numeric', month: 'long', day: 'numeric' })
    });

  } catch (err) {
    console.error('analyze error:', err);
    res.status(500).json({ error: err.message });
  }
});

// ─────────────────────────────────────────────
// Helper: generate inline SVG bar chart as PNG-like base64 via Buffer
// We'll embed a simple bar chart as a table in the Word doc since
// Railway has no canvas. Instead: we draw bars as colored cells.
// ─────────────────────────────────────────────

// ─────────────────────────────────────────────
// Word helpers
// ─────────────────────────────────────────────
const C = {
  primary: '4A1A7A', accent: '7B2FBE', light: 'F4EDFF',
  pale: 'FAF7FF', gray: 'F2F2F2', white: 'FFFFFF',
  darkText: '1A1A2E', midText: '4A4A6A', green: '1A7340',
  yellow: 'B8860B', red: 'C0392B', blue: '2E86AB', border: 'D0B8F0'
};

const noBorder = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
const thinBorder = { style: BorderStyle.SINGLE, size: 4, color: C.border };

function hCell(text, w, bg = C.primary) {
  return new TableCell({
    width: { size: w, type: WidthType.DXA },
    shading: { fill: bg, type: ShadingType.CLEAR },
    borders: { top: noBorder, bottom: noBorder, left: noBorder, right: { style: BorderStyle.SINGLE, size: 4, color: 'FFFFFF' } },
    children: [new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text, bold: true, color: 'FFFFFF', size: 20, font: 'Arial' })]
    })]
  });
}

function dCell(text, w, bg = C.white, bold = false, color = C.darkText, align = AlignmentType.CENTER) {
  return new TableCell({
    width: { size: w, type: WidthType.DXA },
    shading: { fill: bg, type: ShadingType.CLEAR },
    borders: { top: thinBorder, bottom: thinBorder, left: thinBorder, right: thinBorder },
    children: [new Paragraph({
      alignment: align,
      children: [new TextRun({ text: String(text), bold, color, size: 19, font: 'Arial' })]
    })]
  });
}

function spacer(before = 200, after = 0) {
  return new Paragraph({ spacing: { before, after } });
}

function heading(text, level = 1) {
  const sizes = { 1: 32, 2: 26, 3: 22 };
  return new Paragraph({
    alignment: AlignmentType.RIGHT,
    spacing: { before: level === 1 ? 400 : 240, after: 160 },
    children: [new TextRun({
      text, bold: true, size: sizes[level] || 24,
      color: level === 1 ? C.primary : C.accent, font: 'Arial'
    })]
  });
}

function para(text, size = 20, color = C.darkText, before = 80, after = 80) {
  return new Paragraph({
    alignment: AlignmentType.RIGHT,
    spacing: { before, after },
    children: [new TextRun({ text, size, color, font: 'Arial' })]
  });
}

// Bar chart as Word table (colored cells proportional to percentage)
function barChartTable(stats, qLabel, qNum) {
  const labels = ['موافق بشدة (5)', 'موافق (4)', 'محايد (3)', 'غير موافق (2)', 'غير موافق بشدة (1)'];
  const colors = ['1A7340', '2E86AB', 'F0A500', 'E05C34', 'C0392B'];
  const CW = 9200;
  const labelW = 2200;
  const barMaxW = 5000;
  const pctW = 900;
  const countW = 1100;

  const rows = stats.pct.map((pct, i) => {
    const filledW = Math.round((pct / 100) * barMaxW);
    const emptyW  = barMaxW - filledW;

    // bar cell: filled part
    const barCells = [];
    if (filledW > 10) {
      barCells.push(new TableCell({
        width: { size: filledW, type: WidthType.DXA },
        shading: { fill: colors[i], type: ShadingType.CLEAR },
        borders: { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder },
        children: [new Paragraph({ children: [new TextRun({ text: ' ', size: 18 })] })]
      }));
    }
    if (emptyW > 10) {
      barCells.push(new TableCell({
        width: { size: emptyW, type: WidthType.DXA },
        shading: { fill: 'EEEEEE', type: ShadingType.CLEAR },
        borders: { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder },
        children: [new Paragraph({ children: [new TextRun({ text: ' ', size: 18 })] })]
      }));
    }

    return new TableRow({
      children: [
        // label
        new TableCell({
          width: { size: labelW, type: WidthType.DXA },
          borders: { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder },
          children: [new Paragraph({
            alignment: AlignmentType.RIGHT,
            children: [new TextRun({ text: labels[i], size: 17, font: 'Arial', color: C.darkText })]
          })]
        }),
        // bar (nested table trick — use a simple colored cell instead)
        new TableCell({
          width: { size: barMaxW, type: WidthType.DXA },
          borders: { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder },
          children: [
            new Table({
              width: { size: barMaxW, type: WidthType.DXA },
              columnWidths: filledW > 10 ? (emptyW > 10 ? [filledW, emptyW] : [barMaxW]) : [barMaxW],
              borders: { insideH: noBorder, insideV: noBorder },
              rows: [new TableRow({ children: barCells })]
            })
          ]
        }),
        // pct
        new TableCell({
          width: { size: pctW, type: WidthType.DXA },
          borders: { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder },
          children: [new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [new TextRun({ text: `${pct}%`, size: 18, bold: true, color: colors[i], font: 'Arial' })]
          })]
        }),
        // count
        new TableCell({
          width: { size: countW, type: WidthType.DXA },
          borders: { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder },
          children: [new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [new TextRun({ text: `(${stats.counts[i]})`, size: 17, color: '888888', font: 'Arial' })]
          })]
        }),
      ]
    });
  });

  return [
    spacer(200, 0),
    new Paragraph({
      alignment: AlignmentType.RIGHT,
      spacing: { before: 120, after: 60 },
      children: [
        new TextRun({ text: `س${qNum}: `, bold: true, size: 20, color: C.accent, font: 'Arial' }),
        new TextRun({ text: qLabel, size: 20, color: C.darkText, font: 'Arial' })
      ]
    }),
    new Table({
      width: { size: CW, type: WidthType.DXA },
      columnWidths: [labelW, barMaxW, pctW, countW],
      rows,
    }),
    spacer(80, 0),
  ];
}

// ─────────────────────────────────────────────
// API: /api/word  → returns .docx buffer
// ─────────────────────────────────────────────
app.post('/api/word', express.json({ limit: '10mb' }), async (req, res) => {
  try {
    const { surveyTitle, sectionName, sampleSize, questions, overallMean, classification, date } = req.body;
    if (!questions || !questions.length) return res.status(400).json({ error: 'لا توجد بيانات' });

    const CW = 9200;
    const children = [];

    // ── Cover ──────────────────────────────────────────────────
    children.push(
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 600, after: 200 },
        children: [new TextRun({ text: 'كليات الغد للعلوم الطبية التطبيقية', bold: true, size: 44, color: C.primary, font: 'Arial' })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 100 },
        children: [new TextRun({ text: 'AlGhad Colleges for Applied Medical Sciences', size: 24, color: C.midText, font: 'Arial' })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 600 },
        children: [new TextRun({ text: '─────────────────────────────', color: C.accent, size: 20, font: 'Arial' })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 200 },
        children: [new TextRun({ text: 'تقرير تحليل الاستبيان', bold: true, size: 52, color: C.accent, font: 'Arial' })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 600 },
        children: [new TextRun({ text: surveyTitle, bold: true, size: 36, color: C.darkText, font: 'Arial' })]
      }),
    );

    // Cover info table
    const infoRows = [
      ['الاستبيان', surveyTitle],
      ['القسم / التخصص', sectionName || '—'],
      ['حجم العينة', `${sampleSize} مستجيب`],
      ['المتوسط العام', `${overallMean} / 5.00`],
      ['التصنيف العام', classification.ar],
      ['تاريخ التقرير', date],
      ['أعدّه', 'روان علي البارقي'],
    ];
    children.push(
      new Table({
        width: { size: 7000, type: WidthType.DXA },
        alignment: AlignmentType.CENTER,
        columnWidths: [2500, 4500],
        rows: infoRows.map(([k, v], i) => new TableRow({ children: [
          dCell(k, 2500, i % 2 === 0 ? C.light : C.white, true, C.primary, AlignmentType.RIGHT),
          dCell(v, 4500, i % 2 === 0 ? C.pale  : C.white, false, C.darkText, AlignmentType.RIGHT),
        ]}))
      }),
      new Paragraph({ children: [new PageBreak()] })
    );

    // ── Section 1: Overview ────────────────────────────────────
    children.push(heading('١. نظرة عامة على الاستبيان'));
    children.push(para(`يهدف هذا التقرير إلى تحليل نتائج استبيان "${surveyTitle}" المُطبَّق على عينة مؤلفة من ${sampleSize} مستجيباً. تم جمع البيانات وتحليلها إحصائياً باستخدام مقياس ليكرت الخماسي.`));
    children.push(spacer(100));

    // Summary stats table
    children.push(heading('٢. ملخص النتائج الإحصائية', 2));
    children.push(
      new Table({
        width: { size: CW, type: WidthType.DXA },
        columnWidths: [2300, 2300, 2300, 2300],
        rows: [
          new TableRow({ children: [hCell('إجمالي المستجيبين', 2300), hCell('عدد الأسئلة', 2300), hCell('المتوسط العام', 2300), hCell('التصنيف', 2300)] }),
          new TableRow({ children: [
            dCell(String(sampleSize), 2300, C.pale, true, C.primary),
            dCell(String(questions.length), 2300, C.pale, true, C.primary),
            dCell(`${overallMean} / 5.00`, 2300, C.pale, true, C.accent),
            dCell(classification.ar, 2300, C.pale, true, classification.color || C.green),
          ]})
        ]
      }),
      spacer(200)
    );

    // ── Section 2: Detailed questions table ───────────────────
    children.push(heading('٣. جدول النتائج التفصيلية', 2));
    const qColW  = [400, 3400, 1000, 1000, 1000, 800, 1600];
    children.push(
      new Table({
        width: { size: CW, type: WidthType.DXA },
        columnWidths: qColW,
        rows: [
          new TableRow({ children: [
            hCell('#', qColW[0]), hCell('السؤال', qColW[1]), hCell('موافق بشدة %', qColW[2]),
            hCell('موافق %', qColW[3]), hCell('محايد %', qColW[4]),
            hCell('المتوسط', qColW[5]), hCell('التصنيف', qColW[6])
          ]}),
          ...questions.map((q, i) => {
            const cls = classify(q.stats.mean);
            return new TableRow({ children: [
              dCell(String(i + 1), qColW[0], i % 2 === 0 ? C.pale : C.white),
              dCell(q.qLabel, qColW[1], i % 2 === 0 ? C.pale : C.white, false, C.darkText, AlignmentType.RIGHT),
              dCell(`${q.stats.pct[4]}%`, qColW[2], i % 2 === 0 ? C.pale : C.white),
              dCell(`${q.stats.pct[3]}%`, qColW[3], i % 2 === 0 ? C.pale : C.white),
              dCell(`${q.stats.pct[2]}%`, qColW[4], i % 2 === 0 ? C.pale : C.white),
              dCell(String(q.stats.mean), qColW[5], i % 2 === 0 ? C.pale : C.white, true, cls.color),
              dCell(cls.ar, qColW[6], i % 2 === 0 ? C.pale : C.white, false, cls.color),
            ]});
          })
        ]
      }),
      new Paragraph({ children: [new PageBreak()] })
    );

    // ── Section 3: Charts per question ────────────────────────
    children.push(heading('٤. الرسوم البيانية لكل سؤال', 2));
    children.push(para('يُمثّل كل رسم بياني توزيع استجابات المستجيبين وفق مقياس ليكرت الخماسي.'));
    children.push(spacer(100));

    questions.forEach((q, i) => {
      const chartElements = barChartTable(q.stats, q.qLabel, i + 1);
      children.push(...chartElements);
      // stats row under chart
      const cls = classify(q.stats.mean);
      children.push(
        new Table({
          width: { size: CW, type: WidthType.DXA },
          columnWidths: [2300, 2300, 2300, 2300],
          rows: [
            new TableRow({ children: [
              hCell('إجمالي الاستجابات', 2300, C.accent),
              hCell('المتوسط الحسابي', 2300, C.accent),
              hCell('الانحراف المعياري', 2300, C.accent),
              hCell('التصنيف', 2300, C.accent),
            ]}),
            new TableRow({ children: [
              dCell(String(q.stats.total), 2300, C.pale),
              dCell(`${q.stats.mean} / 5.00`, 2300, C.pale, true, C.accent),
              dCell(String(q.stats.std), 2300, C.pale),
              dCell(cls.ar, 2300, C.pale, true, cls.color),
            ]})
          ]
        }),
        spacer(300)
      );

      // page break every 2 questions
      if ((i + 1) % 2 === 0 && i < questions.length - 1) {
        children.push(new Paragraph({ children: [new PageBreak()] }));
      }
    });

    // ── Section 4: Conclusion ──────────────────────────────────
    children.push(
      new Paragraph({ children: [new PageBreak()] }),
      heading('٥. الخلاصة والتوصيات')
    );
    const cls = classify(overallMean);
    children.push(
      para(`بلغ المتوسط العام لنتائج الاستبيان "${surveyTitle}" (${overallMean} / 5.00)، وهو ما يُصنَّف ضمن فئة "${cls.ar}".`),
      para(`استجاب للاستبيان ${sampleSize} مستجيباً على ${questions.length} سؤالاً مُوزَّعاً وفق مقياس ليكرت الخماسي.`),
      spacer(120),
      heading('التوصيات', 3),
      para('• مواصلة تعزيز الجوانب التي حققت تقييماً مرتفعاً والاستمرار في تطويرها.'),
      para('• إيلاء اهتمام خاص للأسئلة التي سجّلت متوسطات دون 3.5 ووضع خطط تحسين مستهدفة.'),
      para('• تكرار قياس رضا المستجيبين بصفة دورية لرصد التحسينات عبر الزمن.'),
      para('• مشاركة النتائج مع الجهات المعنية لاتخاذ قرارات مبنية على البيانات.'),
      spacer(400),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 200 },
        children: [new TextRun({ text: '─── نهاية التقرير ───', color: C.accent, size: 22, font: 'Arial' })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: 'أعدّه: روان علي البارقي  |  كليات الغد جدة', color: C.midText, size: 18, font: 'Arial' })]
      }),
    );

    // ── Build document ─────────────────────────────────────────
    const doc = new Document({
      sections: [{
        properties: { page: { margin: { top: 720, bottom: 720, left: 900, right: 900 } } },
        children
      }]
    });

    const buf = await Packer.toBuffer(doc);
    const filename = encodeURIComponent(`تقرير_${surveyTitle}_${Date.now()}.docx`);
    res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''${filename}`);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.send(buf);

  } catch (err) {
    console.error('word error:', err);
    res.status(500).json({ error: err.message });
  }
});

// ─────────────────────────────────────────────
// Fallback
// ─────────────────────────────────────────────
app.get('*', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

const PORT = process.env.PORT || 8080;
app.listen(PORT, () => console.log(`✅ AlGhad Survey Analyzer v3.0 — port ${PORT}`));
