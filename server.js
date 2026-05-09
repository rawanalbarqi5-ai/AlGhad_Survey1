const express = require('express');
const multer  = require('multer');
const XLSX    = require('xlsx');
const path    = require('path');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType, VerticalAlign, PageOrientation
} = require('docx');

const app    = express();
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 20 * 1024 * 1024 } });

app.use(express.json({ limit: '20mb' }));
// Serve from public/ folder or root (handles both structures)
app.use(express.static(path.join(__dirname, 'public')));
app.use(express.static(__dirname));

// Fallback: serve index.html from public/ or root
app.get('/', (req, res) => {
  const publicIndex = path.join(__dirname, 'public', 'index.html');
  const rootIndex   = path.join(__dirname, 'index.html');
  if (require('fs').existsSync(publicIndex)) return res.sendFile(publicIndex);
  if (require('fs').existsSync(rootIndex))   return res.sendFile(rootIndex);
  res.send('AlGhad Survey Analyzer is running! Upload index.html to public/ folder.');
});

// ── Word generation helpers ────────────────────────────────────────────────
const DARK='1F4E79', MID='2E75B6', PALE='EBF3FB', WHITE='FFFFFF';
const GREEN='375623', GREEN2='E2EFDA', AMBER='7F6000', AMBER2='FFEB9C';
const RED='9C0006', RED2='FFC7CE', FSHADE='FCE4D6', MSHADE='DDEBF7', CSHADE='E2EFDA';
const CW = 14400;

const brd = () => ({ style: BorderStyle.SINGLE, size: 4, color: 'AAAAAA' });
const allB = () => { const b = brd(); return { top:b, bottom:b, left:b, right:b }; };
const mg   = () => ({ top:80, bottom:80, left:120, right:120 });

const clf = m => {
  if (m <= 1.5) return { l:'Excellent', ar:'ممتاز',  hx:GREEN2, hc:GREEN };
  if (m <= 2.0) return { l:'Good',      ar:'جيد',    hx:GREEN2, hc:GREEN };
  if (m <= 2.5) return { l:'Acceptable',ar:'مقبول',  hx:AMBER2, hc:AMBER };
  if (m <= 3.0) return { l:'Weakness',  ar:'ضعف',    hx:RED2,   hc:RED   };
  return               { l:'Critical',  ar:'حرج',    hx:RED2,   hc:RED   };
};

const mC = (text, w, shade, opts={}) => new TableCell({
  width: { size:w, type:WidthType.DXA }, borders: allB(),
  shading: shade ? { fill:shade, type:ShadingType.CLEAR } : undefined,
  margins: mg(), verticalAlign: VerticalAlign.CENTER, rowSpan: opts.rowSpan,
  children: [new Paragraph({
    alignment: opts.align || AlignmentType.CENTER,
    children: [new TextRun({ text: String(text ?? ''), bold: opts.bold||false,
      color: opts.color||'000000', size: opts.size||18, font:'Arial' })]
  })]
});

const mH = (lines, w, shade=DARK) => new TableCell({
  width: { size:w, type:WidthType.DXA }, borders: allB(),
  shading: { fill:shade, type:ShadingType.CLEAR }, margins: mg(), verticalAlign: VerticalAlign.CENTER,
  children: lines.map(l => new Paragraph({
    alignment: AlignmentType.CENTER, spacing: { before:0, after:0 },
    children: [new TextRun({ text:l, bold:true, color:WHITE, size:17, font:'Arial' })]
  }))
});

const mP  = (text, opts={}) => new Paragraph({
  alignment: opts.align || AlignmentType.LEFT,
  spacing: { before: opts.before||60, after: opts.after||80 },
  children: [new TextRun({ text, bold:opts.bold||false, color:opts.color||'000000', size:opts.size||18, font:'Arial' })]
});
const mPR = (text, opts={}) => mP(text, { ...opts, align: AlignmentType.RIGHT });
const sp  = () => new Paragraph({ spacing:{ before:120, after:120 }, children:[] });

// ── Build Word document from RES + CFG ─────────────────────────────────────
async function buildWord(res, cfg) {
  const { nF, nM, n, secs, overall, isMulti, secFileNames, nPerSection } = res;
  const { cname, sname, obj, semester, gmode } = cfg;
  const showGender = gmode === 'col';
  const children = [];

  // Title
  children.push(
    mP(sname || 'Survey Analysis', { align:AlignmentType.CENTER, bold:true, size:52, color:DARK, before:0, after:80 }),
    mPR('تحليل نتائج الاستبيان', { bold:true, size:36, color:MID, before:0, after:60 }),
    mP(cname || 'AlGhad College', { align:AlignmentType.CENTER, size:20, color:'555555', before:0, after:60 }),
    mP(semester || '', { align:AlignmentType.CENTER, size:18, color:'777777', before:0, after:200 }),
  );

  // Multi-section note
  if (isMulti && secFileNames) {
    children.push(
      new Paragraph({ spacing:{before:0,after:60}, children:[
        new TextRun({ text:'🔗 نتائج مدمجة من '+secFileNames.length+' شعب: '+
          secFileNames.map((n,i)=>`${n} (${nPerSection[i]} مستجيب)`).join(' · '),
          bold:true, color:MID, size:18, font:'Arial' })
      ]})
    );
  }

  // Objective
  if (obj) {
    children.push(
      mPR('هدف الاستبيان', { bold:true, size:26, color:DARK, before:200, after:100 }),
      mPR(obj, { size:18, before:0, after:200, color:'222222' }),
    );
  }

  // Likert scale
  children.push(
    sp(),
    mP('المقياس المستخدم | Scale Used', { bold:true, size:24, color:DARK, before:0, after:100 }),
  );
  const SLC = Math.floor(CW / 5);
  const SLW = [SLC, SLC, SLC, SLC, CW - SLC*4];
  children.push(new Table({ width:{ size:CW, type:WidthType.DXA }, columnWidths:SLW,
    rows:[new TableRow({ children:[
      ['1 = Strongly Agree | موافق بشدة', GREEN2, GREEN],
      ['2 = Agree | موافق',               GREEN2, GREEN],
      ['3 = Neutral | محايد',             'F2F2F2','444444'],
      ['4 = Disagree | غير موافق',        AMBER2, AMBER],
      ['5 = Strongly Disagree | غير موافق بشدة', RED2, RED],
    ].map(([t,bg,c],i) => new TableCell({
      width:{size:SLW[i],type:WidthType.DXA}, borders:allB(),
      shading:{fill:bg,type:ShadingType.CLEAR}, margins:mg(),
      children:[new Paragraph({ alignment:AlignmentType.CENTER,
        children:[new TextRun({ text:t, bold:true, color:c, size:16, font:'Arial' })] })]
    })) })]
  }));

  // Classification table
  children.push(sp(), mP('Classification Scale | مقياس التصنيف', { bold:true, size:22, color:DARK, before:150, after:100 }));
  const CC = [1600, 3000, 3000, 1200, 5600];
  children.push(new Table({ width:{size:CW,type:WidthType.DXA}, columnWidths:CC,
    rows:[
      new TableRow({ children:[mH(['Range'],CC[0]),mH(['Classification'],CC[1]),mH(['التصنيف'],CC[2]),mH([''],CC[3]),mH(['Interpretation'],CC[4])] }),
      ...[ ['≤ 1.50','Excellent','ممتاز', GREEN2,GREEN,'Strong positive outcome'],
           ['1.51–2.00','Good','جيد',     GREEN2,GREEN,'Positive; meets expectations'],
           ['2.01–2.50','Acceptable','مقبول',AMBER2,AMBER,'Moderate; requires monitoring'],
           ['2.51–3.00','Weakness','ضعف',  RED2,  RED,  'Below expectations; improvement needed'],
           ['> 3.00','Critical','حرج',     RED2,  RED,  'Significant weakness; immediate action required'],
      ].map(([r,en,ar,bg,c,interp],i) => new TableRow({ children:[
        mC(r,CC[0],bg,{bold:true,color:c}), mC(en,CC[1],bg,{bold:true,color:c}), mC(ar,CC[2],bg,{bold:true,color:c}),
        new TableCell({width:{size:CC[3],type:WidthType.DXA},borders:allB(),shading:{fill:bg,type:ShadingType.CLEAR},margins:mg(),children:[new Paragraph({children:[]})]}),
        mC(interp,CC[4],i%2===0?'FAFAFA':WHITE,{align:AlignmentType.LEFT,size:16}),
      ]}))
    ]
  }));

  // Sample profile
  children.push(sp(), mP('Sample Profile | بيانات العينة', { bold:true, size:24, color:DARK, before:200, after:100 }));
  const SP = [Math.round(CW*.36), Math.round(CW*.14), Math.round(CW*.36), CW-Math.round(CW*.36)*2-Math.round(CW*.14)];
  const spRows = [['Total Respondents', n, 'إجمالي المشاركين', n]];
  if (showGender) { spRows.push(['Female', nF, 'إناث', nF], ['Male', nM, 'ذكور', nM]); }
  else if (gmode === 'all_f') { spRows.push(['Gender','All Female','الجنس','إناث']); }
  else { spRows.push(['Gender','All Male','الجنس','ذكور']); }
  spRows.push(
    ['Sections', secs.length, 'عدد المحاور', secs.length],
    ['Total Questions', secs.reduce((a,s)=>a+s.qs.length,0), 'عدد الأسئلة', secs.reduce((a,s)=>a+s.qs.length,0)],
    ['Overall Mean', overall, 'المتوسط العام', overall],
    ['Survey Period', semester||'—', 'الفصل الدراسي', semester||'—'],
  );
  children.push(new Table({ width:{size:CW,type:WidthType.DXA}, columnWidths:SP,
    rows:[
      new TableRow({ children:[mH(['Detail'],SP[0]),mH(['Value'],SP[1]),mH(['التفاصيل'],SP[2]),mH(['القيمة'],SP[3])] }),
      ...spRows.map(([en,ev,ar,av],i) => new TableRow({ children:[
        mC(en,SP[0],i%2===0?PALE:WHITE,{align:AlignmentType.LEFT}),
        mC(ev,SP[1],i%2===0?PALE:WHITE,{bold:true,color:DARK}),
        mC(ar,SP[2],i%2===0?PALE:WHITE,{align:AlignmentType.RIGHT}),
        mC(av,SP[3],i%2===0?PALE:WHITE,{bold:true,color:DARK}),
      ]}))
    ]
  }));

  // Section summary
  const SSC = showGender
    ? [3200,2900,1050,1200,1200,850,CW-3200-2900-1050-1200-1200-850]
    : [3800,3500,1200,CW-3800-3500-1200];
  children.push(sp(), mP('Section Summary | ملخص المحاور', { bold:true, size:24, color:DARK, before:200, after:100 }));
  children.push(new Table({ width:{size:CW,type:WidthType.DXA}, columnWidths:SSC,
    rows:[
      new TableRow({ children: showGender
        ? [mH(['Section'],SSC[0]),mH(['المحور'],SSC[1]),mH(['Mean'],SSC[2]),mH(['F.Mean'],SSC[3]),mH(['M.Mean'],SSC[4]),mH(['Gap'],SSC[5]),mH(['Classification'],SSC[6])]
        : [mH(['Section'],SSC[0]),mH(['المحور'],SSC[1]),mH(['Mean'],SSC[2]),mH(['Classification'],SSC[3])]
      }),
      ...secs.map((s,i) => {
        const bg = i%2===0 ? PALE : WHITE;
        return new TableRow({ children: showGender
          ? [mC(s.name,SSC[0],bg,{align:AlignmentType.LEFT,size:16}), mC(s.ar,SSC[1],bg,{align:AlignmentType.RIGHT,size:15}),
             mC(s.mean.toFixed(2),SSC[2],bg,{bold:true}), mC(s.fMean.toFixed(2),SSC[3],bg), mC(s.mMean.toFixed(2),SSC[4],bg),
             mC(Math.abs(s.fMean-s.mMean).toFixed(2),SSC[5],bg), mC(s.cl.l+' | '+s.cl.ar,SSC[6],s.cl.hx,{bold:true,color:s.cl.hc,size:15})]
          : [mC(s.name,SSC[0],bg,{align:AlignmentType.LEFT,size:16}), mC(s.ar,SSC[1],bg,{align:AlignmentType.RIGHT,size:15}),
             mC(s.mean.toFixed(2),SSC[2],bg,{bold:true}), mC(s.cl.l+' | '+s.cl.ar,SSC[3],s.cl.hx,{bold:true,color:s.cl.hc,size:15})]
        });
      })
    ]
  }));

  // Executive summary
  children.push(sp(), mPR('اللمحة العامة', { bold:true, size:24, color:DARK, before:200, after:100 }));
  const bestSec  = secs.reduce((a,b) => a.mean < b.mean ? a : b);
  const worstSec = secs.reduce((a,b) => a.mean > b.mean ? a : b);
  [
    `تُظهر النتائج مستوى رضا ${overall<=1.5?'ممتازاً':overall<=2?'جيداً':'مقبولاً'} بمتوسط عام (${overall}) — تصنيف "${clf(overall).ar}".`,
    `أقوى المحاور: "${bestSec.ar}" بمتوسط (${bestSec.mean}) — ${bestSec.cl.ar}.`,
    `المحور الأولى بالتطوير: "${worstSec.ar}" بمتوسط (${worstSec.mean}) — ${worstSec.cl.ar}.`,
    showGender ? `الفجوة بين الجنسين: ${secs.every(s=>Math.abs(s.fMean-s.mMean)<0.3)?'طفيفة تدل على تجانس التجربة':'ملحوظة وتستوجب الدراسة'}.`
               : `عدد المستجيبين: ${n} ${gmode==='all_f'?'(إناث)':'(ذكور)'}.`,
    'نسب الموافقة الإيجابية المرتفعة تعكس جودة تدريبية تستحق التوثيق.',
  ].forEach(ar => children.push(mPR('• '+ar, { size:17, before:80, after:50, color:'1a1a2e' })));

  // Overall analysis
  const OWg = [800,700,1700,1100,1100,1050,1050,1150,1150,1150,3450];
  const OWs = [900,800,2000,1300,1200,1200,CW-900-800-2000-1300-1200-1200];
  children.push(sp(), mP('Overall Analysis | التحليل الإجمالي', { bold:true, size:24, color:DARK, before:200, after:100 }));

  const oaR = [new TableRow({ children: showGender
    ? [mH(['Q#'],OWg[0]),mH(['Sec.Q'],OWg[1]),mH(['Section'],OWg[2]),mH(['F.Mean'],OWg[3]),mH(['M.Mean'],OWg[4]),mH(['Max'],OWg[5]),mH(['Min'],OWg[6]),mH(['Mean'],OWg[7]),mH(['Pos%'],OWg[8]),mH(['Neg%'],OWg[9]),mH(['Classification'],OWg[10])]
    : [mH(['Q#'],OWs[0]),mH(['Sec.Q'],OWs[1]),mH(['Section'],OWs[2]),mH(['Mean'],OWs[3]),mH(['Pos%'],OWs[4]),mH(['Neg%'],OWs[5]),mH(['Classification'],OWs[6])]
  })];

  secs.forEach(s => {
    const span = showGender ? 11 : 7;
    oaR.push(new TableRow({ children:[new TableCell({ columnSpan:span, width:{size:CW,type:WidthType.DXA},
      borders:allB(), shading:{fill:MID,type:ShadingType.CLEAR}, margins:mg(),
      children:[new Paragraph({ alignment:AlignmentType.CENTER,
        children:[new TextRun({ text:s.name+' | '+s.ar, bold:true, color:WHITE, size:18, font:'Arial' })] })]
    })]}));
    s.qs.forEach((q,si) => {
      const cl = q.cl; const bg = si%2===0 ? PALE : WHITE;
      oaR.push(new TableRow({ children: showGender
        ? [mC('Q'+q.qn,OWg[0],bg,{bold:true}),mC(si+1,OWg[1],bg),mC(s.name,OWg[2],bg,{size:14}),
           mC(q.fM.toFixed(2),OWg[3],bg),mC(q.mM.toFixed(2),OWg[4],bg),
           mC(q.maxM.toFixed(2),OWg[5],bg),mC(q.minM.toFixed(2),OWg[6],bg),
           mC(q.cM.toFixed(2),OWg[7],cl.hx,{bold:true,color:cl.hc}),
           mC(q.pos+'%',OWg[8],q.pos>=70?GREEN2:AMBER2,{bold:true,color:q.pos>=70?GREEN:AMBER}),
           mC(q.neg+'%',OWg[9],q.neg>=15?RED2:GREEN2,{bold:true,color:q.neg>=15?RED:GREEN}),
           mC(cl.l+' | '+cl.ar,OWg[10],cl.hx,{bold:true,color:cl.hc,size:14})]
        : [mC('Q'+q.qn,OWs[0],bg,{bold:true}),mC(si+1,OWs[1],bg),mC(s.name,OWs[2],bg,{size:14}),
           mC(q.cM.toFixed(2),OWs[3],cl.hx,{bold:true,color:cl.hc}),
           mC(q.pos+'%',OWs[4],q.pos>=70?GREEN2:AMBER2,{bold:true,color:q.pos>=70?GREEN:AMBER}),
           mC(q.neg+'%',OWs[5],q.neg>=15?RED2:GREEN2,{bold:true,color:q.neg>=15?RED:GREEN}),
           mC(cl.l+' | '+cl.ar,OWs[6],cl.hx,{bold:true,color:cl.hc,size:14})]
      }));
    });
  });
  children.push(new Table({ width:{size:CW,type:WidthType.DXA}, columnWidths:showGender?OWg:OWs, rows:oaR }));

  // Per-section distribution
  const DC  = [900,900,2800,1700,1700,1700,1700,1700,1300];
  const DCs = [1200,1200,1700,1700,1700,1700,1700,CW-1200-1200-1700*5];

  secs.forEach(sec => {
    children.push(
      sp(),
      new Paragraph({ spacing:{before:0,after:0}, children:[new TextRun({ text:'─'.repeat(80), color:'BDD7EE', font:'Arial', size:16 })] }),
      mP(sec.name+' | '+sec.ar, { bold:true, size:28, color:MID, before:100, after:60 }),
      mP('Mean: '+sec.mean+(showGender?'  |  F.Mean: '+sec.fMean+'  |  M.Mean: '+sec.mMean:'')+'  |  '+sec.cl.l+' ('+sec.cl.ar+')',
        { size:17, color:'444444', before:0, after:100 }),
    );

    const distRows = [];
    if (showGender) {
      distRows.push(new TableRow({ children:[
        mH(['Global Q'],DC[0]), mH(['Sec. Q'],DC[1]), mH(['Group'],DC[2]),
        mH(['%5','SD'],DC[3]), mH(['%4','D'],DC[4]), mH(['%3','N'],DC[5]),
        mH(['%2','A'],DC[6]),  mH(['%1','SA'],DC[7]), mH(['Mean'],DC[8]),
      ]}));
      sec.qs.forEach((q,qi) => {
        distRows.push(new TableRow({ children:[new TableCell({ columnSpan:9, width:{size:CW,type:WidthType.DXA},
          borders:allB(), shading:{fill:PALE,type:ShadingType.CLEAR}, margins:mg(),
          children:[new Paragraph({ alignment:AlignmentType.RIGHT, spacing:{before:0,after:0},
            children:[new TextRun({ text:'Q'+q.qn+'. '+q.lbl, bold:true, size:17, color:DARK, font:'Arial' })] })]
        })]}));
        const bg = qi%2===0 ? PALE : 'FFFFFF';
        const mkRS = (t,w,sh) => new TableCell({ width:{size:w,type:WidthType.DXA}, borders:allB(),
          shading:{fill:sh,type:ShadingType.CLEAR}, margins:mg(), verticalAlign:VerticalAlign.CENTER, rowSpan:3,
          children:[new Paragraph({ alignment:AlignmentType.CENTER,
            children:[new TextRun({ text:String(t), bold:true, size:19, color:DARK, font:'Arial' })] })] });
        [['Female',q.fD,q.fM,'FCE4D6','843C0C',true],
         ['Male',  q.mD,q.mM,'DDEBF7','1F4E79',false],
         ['Combined',q.cD,q.cM,'E2EFDA','375623',false],
        ].forEach(([g,d,m,sh,tc,first]) => {
          distRows.push(new TableRow({ children:[
            ...(first ? [mkRS('Q'+q.qn,DC[0],bg), mkRS(qi+1,DC[1],bg)] : []),
            mC(g,DC[2],sh,{bold:true,color:tc}),
            ...d.map((v,j) => mC(v,DC[3+j],sh)),
            mC((+m).toFixed(2),DC[8],sh,{bold:true,color:tc}),
          ]}));
        });
      });
    } else {
      distRows.push(new TableRow({ children:[
        mH(['Global Q'],DCs[0]), mH(['Sec. Q'],DCs[1]),
        mH(['%5 SD'],DCs[2]), mH(['%4 D'],DCs[3]), mH(['%3 N'],DCs[4]),
        mH(['%2 A'],DCs[5]),  mH(['%1 SA'],DCs[6]), mH(['Mean'],DCs[7]),
      ]}));
      sec.qs.forEach((q,qi) => {
        const bg = qi%2===0 ? PALE : 'FFFFFF';
        distRows.push(new TableRow({ children:[
          mC('Q'+q.qn, DCs[0], bg, {bold:true}),
          mC(qi+1,     DCs[1], bg),
          ...q.cD.map((v,j) => mC(v, DCs[2+j], bg)),
          mC((+q.cM).toFixed(2), DCs[7], q.cl.hx, {bold:true,color:q.cl.hc}),
        ]}));
      });
    }
    children.push(new Table({ width:{size:CW,type:WidthType.DXA}, columnWidths:showGender?DC:DCs, rows:distRows }));
  });

  const doc = new Document({
    styles: { default: { document: { run: { font:'Arial', size:18 } } } },
    sections: [{ properties: { page: {
      size: { width:12240, height:15840, orientation:PageOrientation.LANDSCAPE },
      margin: { top:720, right:720, bottom:720, left:720 }
    }}, children }]
  });

  return Packer.toBuffer(doc);
}

// ── Routes ─────────────────────────────────────────────────────────────────

// Upload & parse Excel/CSV
app.post('/api/upload', upload.single('file'), (req, res) => {
  try {
    const { originalname, buffer } = req.file;
    const ext = originalname.split('.').pop().toLowerCase();
    let headers, rows;
    if (ext === 'csv') {
      const text = buffer.toString('utf-8');
      const lines = text.trim().split(/\r?\n/);
      headers = lines[0].split(',').map(h => h.trim().replace(/^"|"$/g,''));
      rows = lines.slice(1).map(l => l.split(',').map(v => v.trim().replace(/^"|"$/g,'')))
                           .filter(r => r.some(v => v !== ''));
    } else {
      const wb = XLSX.read(buffer, { type:'buffer' });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const json = XLSX.utils.sheet_to_json(ws, { header:1, defval:'' });
      headers = json[0].map(h => String(h||'').trim());
      rows = json.slice(1).filter(r => r.some(v => v !== ''));
    }
    res.json({ ok:true, headers, rows, n:rows.length });
  } catch(e) {
    res.status(400).json({ ok:false, error: e.message });
  }
});

// Generate Word
app.post('/api/generate-word', async (req, res) => {
  try {
    const { result, cfg } = req.body;
    // Re-attach clf to each question
    const secs = result.secs.map(s => ({
      ...s,
      cl: clf(s.mean),
      qs: s.qs.map(q => ({ ...q, cl: clf(q.cM) }))
    }));
    const enriched = { ...result, secs };
    const buf = await buildWord(enriched, cfg);
    res.set({
      'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      'Content-Disposition': `attachment; filename="Survey_Analysis.docx"`,
      'Content-Length': buf.length,
    });
    res.send(buf);
  } catch(e) {
    console.error(e);
    res.status(500).json({ error: e.message });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server running on port ${PORT}`));
