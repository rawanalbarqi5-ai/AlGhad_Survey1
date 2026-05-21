const express = require('express');
const multer  = require('multer');
const XLSX    = require('xlsx');
const path    = require('path');
const fs      = require('fs');

const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType, VerticalAlign, PageOrientation,
  ImageRun
} = require('docx');

const app    = express();
const upload = multer({ storage: multer.memoryStorage(), limits:{ fileSize:30*1024*1024 } });
app.use(express.json({ limit:'20mb' }));

// ── Static ────────────────────────────────────────────────────────────────────
app.use(express.static(path.join(__dirname,'public'), {
  etag: false,
  lastModified: false,
  setHeaders: (res) => {
    res.set('Cache-Control', 'no-store, no-cache, must-revalidate, proxy-revalidate');
    res.set('Pragma', 'no-cache');
    res.set('Expires', '0');
  }
}));
app.use(express.static(__dirname));
// ── Batch chart generation (pure JS + sharp, no Python) ─────────────────────
function makeChartSVG(qn, lbl, fD, mD, fLabel, mLabel, secName){
  const W=580, H=210;
  const PAD={top:52, bottom:38, left:42, right:10};
  const cW=W-PAD.left-PAD.right;
  const cH=H-PAD.top-PAD.bottom;
  const labs=['موافق بشدة\nSA(1)','موافق\nA(2)','محايد\nN(3)','لا أوافق\nD(4)','لا أوافق بشدة\nSD(5)'];
  const labsShort=['SA(1)','A(2)','N(3)','D(4)','SD(5)'];

  // Colors: green for positive, grey for neutral, red for negative
  const fColors=['#1B5E20','#388E3C','#757575','#D32F2F','#B71C1C'];
  const mColors=['#0D47A1','#1976D2','#90A4AE','#EF5350','#E53935'];

  const mx = Math.max(...fD, ...mD, 1);
  const yScale = cH / (mx * 1.30 + 5);
  const grpW = cW / 5;
  const bW = grpW * 0.32;
  const gap = 3;

  let bars='', xLabels='', valLabels='', gridLines='';

  // Grid lines & Y axis labels
  const gridSteps = 4;
  const gridStep = Math.ceil(mx / gridSteps / 10) * 10 || 10;
  for(let v = 0; v <= mx * 1.3 + 5; v += gridStep){
    const y = PAD.top + cH - v * yScale;
    if(y < PAD.top - 2) break;
    gridLines += `<line x1="${PAD.left}" y1="${y.toFixed(1)}" x2="${W-PAD.right}" y2="${y.toFixed(1)}" stroke="#E0E0E0" stroke-width="0.8"/>`;
    gridLines += `<text x="${PAD.left-4}" y="${(y+3.5).toFixed(1)}" text-anchor="end" font-size="8" fill="#888" font-family="Arial">${v}</text>`;
  }

  // Bars
  for(let i=0;i<5;i++){
    const xBase = PAD.left + i * grpW + (grpW - bW*2 - gap) / 2;

    // Female bar
    const fh = Math.max(0, fD[i] * yScale);
    const fy = PAD.top + cH - fh;
    bars += `<rect x="${xBase.toFixed(1)}" y="${fy.toFixed(1)}" width="${bW.toFixed(1)}" height="${fh.toFixed(1)}" fill="${fColors[i]}" rx="1.5"/>`;
    if(fD[i] >= 5){
      valLabels += `<text x="${(xBase+bW/2).toFixed(1)}" y="${(fy-3).toFixed(1)}" text-anchor="middle" font-size="7.5" font-weight="bold" fill="${fColors[i]}" font-family="Arial">${fD[i].toFixed(0)}%</text>`;
    }

    // Male bar
    const mh = Math.max(0, mD[i] * yScale);
    const my = PAD.top + cH - mh;
    const mx2 = xBase + bW + gap;
    bars += `<rect x="${mx2.toFixed(1)}" y="${my.toFixed(1)}" width="${bW.toFixed(1)}" height="${mh.toFixed(1)}" fill="${mColors[i]}" rx="1.5"/>`;
    if(mD[i] >= 5){
      valLabels += `<text x="${(mx2+bW/2).toFixed(1)}" y="${(my-3).toFixed(1)}" text-anchor="middle" font-size="7.5" font-weight="bold" fill="${mColors[i]}" font-family="Arial">${mD[i].toFixed(0)}%</text>`;
    }

    // X axis label
    const xCenter = xBase + bW + gap/2;
    xLabels += `<text x="${xCenter.toFixed(1)}" y="${(H-20).toFixed(1)}" text-anchor="middle" font-size="8.5" fill="#444" font-family="Arial">${labsShort[i]}</text>`;
  }

  // Axes
  const axes = `
    <line x1="${PAD.left}" y1="${PAD.top}" x2="${PAD.left}" y2="${PAD.top+cH}" stroke="#999" stroke-width="1"/>
    <line x1="${PAD.left}" y1="${PAD.top+cH}" x2="${W-PAD.right}" y2="${PAD.top+cH}" stroke="#999" stroke-width="1"/>`;

  // Y axis label
  const yAxisLabel = `<text x="9" y="${(PAD.top+cH/2).toFixed(1)}" text-anchor="middle" font-size="8" fill="#666" font-family="Arial" transform="rotate(-90,9,${(PAD.top+cH/2).toFixed(1)})">%</text>`;

  // Title - Q number + question text
  const titleText = `Q${qn}: ${String(lbl||'').slice(0,65)}${lbl&&lbl.length>65?'...':''}`;
  const secText = secName ? `| ${String(secName).slice(0,30)}` : '';
  const title = `<text x="${W/2}" y="14" text-anchor="middle" font-size="9" font-weight="bold" fill="#1F4E79" font-family="Arial">${titleText}</text>
    ${secText ? `<text x="${W/2}" y="26" text-anchor="middle" font-size="7.5" fill="#2E75B6" font-family="Arial">${secText}</text>` : ''}`;

  // Legend - horizontal at bottom
  const fLbl = String(fLabel||'Female');
  const mLbl = String(mLabel||'Male');
  const legendY = H - 6;
  const legend = `
    <rect x="${W/2-80}" y="${legendY-9}" width="11" height="11" fill="${fColors[0]}" rx="2"/>
    <text x="${W/2-65}" y="${legendY}" font-size="8.5" fill="#1B5E20" font-family="Arial" font-weight="bold">${fLbl}</text>
    <rect x="${W/2+10}" y="${legendY-9}" width="11" height="11" fill="${mColors[0]}" rx="2"/>
    <text x="${W/2+25}" y="${legendY}" font-size="8.5" fill="#0D47A1" font-family="Arial" font-weight="bold">${mLbl}</text>`;

  return `<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}">
    <rect width="${W}" height="${H}" fill="#FAFCFF" rx="4"/>
    <rect x="1" y="1" width="${W-2}" height="${H-2}" fill="none" stroke="#DDE8F5" stroke-width="1" rx="4"/>
    ${title}
    ${gridLines}
    ${bars}
    ${valLabels}
    ${xLabels}
    ${axes}
    ${yAxisLabel}
    ${legend}
  </svg>`;
}

app.post('/api/chart', async(req,res)=>{
  try{
    const sharp=require('sharp');
    const body=req.body;
    const questions=body.questions||[{qn:body.qn,lbl:body.lbl,fD:body.fD,mD:body.mD,
      f_lbl:body.fasli_label,m_lbl:body.imtiaz_label}];
    const charts={};
    await Promise.all(questions.map(async d=>{
      try{
        const svg=makeChartSVG(d.qn,d.lbl||'',d.fD||[0,0,0,0,0],d.mD||[0,0,0,0,0],d.f_lbl||'Female',d.m_lbl||'Male',d.secName||'');
        const png=await sharp(Buffer.from(svg)).png().toBuffer();
        charts[String(d.qn)]=png.toString('base64');
      }catch(e){ charts[String(d.qn)]=null; }
    }));
    if(body.questions) res.json({charts});
    else res.json({png:charts[String(body.qn)]});
  }catch(e){
    console.error('Chart err:',e.message?.slice(0,100));
    res.status(500).json({error:e.message?.slice(0,100)});
  }
});

app.get('/', (req,res) => {
  res.set('Cache-Control','no-store, no-cache, must-revalidate');
  res.set('Pragma','no-cache');
  const pub = path.join(__dirname,'public');
  try {
    const files = fs.readdirSync(pub).filter(f=>f.endsWith('.html')||f.endsWith('.htm'));
    if(files.length){
      // Prefer index.html (no space) over index .html (with space)
      const preferred = files.find(f=>f==='index.html') || files[0];
      console.log('Serving HTML:', preferred, '| All files:', files.join(','));
      return res.sendFile(path.join(pub, preferred));
    }
  } catch(e){ console.error(e.message); }
  const fallback = path.join(__dirname,'index.html');
  if(fs.existsSync(fallback)) return res.sendFile(fallback);
  res.send('<h1>Error</h1><p>No HTML in: '+pub+'</p><p>Files: '+(fs.existsSync(pub)?fs.readdirSync(pub).join(', '):'N/A')+'</p>');
});

// ── Colors & helpers ──────────────────────────────────────────────────────────
const DARK='1F4E79',MID='2E75B6',PALE='EBF3FB',WHITE='FFFFFF';

// Bilingual Q translations
const Q_EN = {
  1: 'Course guidelines and descriptions (including knowledge and skills the course was designed to develop) were clear.',
  2: 'Course requirements (including tests and assignments used for assessment) were clear.',
  3: 'Help resources available to students (including office hours) were helpful.',
  4: 'Course delivery and tasks required were consistent with the course outline.',
  5: 'My instructor(s) were fully committed to the delivery of the course (e.g. classes started on time, always present, material well prepared).',
  6: 'My instructor(s) had thorough knowledge of the content of the course.',
  7: 'My instructor(s) were available for help during office hours.',
  8: 'My instructor(s) showed enthusiasm for teaching.',
  9: 'My instructor(s) were interested in my progress and were helpful to me.',
  10: 'All course materials were current and useful (readings, summaries, references, etc.).',
  11: 'Resources I needed in this course were available whenever I needed them.',
  12: 'Technology was used effectively to support my learning in this course.',
  13: 'I was encouraged to ask questions and develop my own ideas in this course.',
  14: 'I was encouraged to do my best work in this course.',
  15: 'Things I had to do in this course (class activities, assignments, laboratories, etc.) were helpful for developing knowledge and skills the course was intended to teach.',
  16: 'The amount of work I had to do in this course was reasonable for the credit hours allocated.',
  17: 'Marks for assignments and tests were returned within reasonable time.',
  18: 'Grading of my tests and assignments was fair and reasonable.',
  19: 'The links between this course and other courses in my program were made clear to me.',
  20: 'What I learned in this course is important and will benefit me in the future.',
  21: 'This course helped me improve my ability to think and solve problems rather than memorize information.',
  22: 'This course helped me improve my teamwork skills.',
  23: 'This course helped me improve my ability to communicate effectively.',
  24: 'I feel generally satisfied with the overall quality of this course.',
};
const GREEN='375623',GREEN2='E2EFDA',AMBER='7F6000',AMBER2='FFEB9C';
const RED='9C0006',RED2='FFC7CE',PINK='FCE4D6',BLUE2='DDEBF7',ORANGE='ED7D31',ORANGE2='FCE4D6';

const brd=(c='AAAAAA')=>({style:BorderStyle.SINGLE,size:4,color:c});
const allB=(c='AAAAAA')=>{const b=brd(c);return{top:b,bottom:b,left:b,right:b};};
const mg=()=>({top:70,bottom:70,left:100,right:100});

const clf=m=>{
  // 1=موافق بشدة (best), 5=لا أوافق بشدة (worst)
  if(m<=1.80)return{l:'ممتاز',    en:'Excellent', bg:GREEN2,c:GREEN};
  if(m<=2.60)return{l:'جيد جداً', en:'Very Good', bg:GREEN2,c:GREEN};
  if(m<=3.40)return{l:'جيد',      en:'Good',      bg:AMBER2,c:AMBER};
  if(m<=4.20)return{l:'مقبول',    en:'Acceptable',bg:RED2,  c:RED};
  return      {l:'ضعيف',          en:'Weak',       bg:RED2,  c:RED};
};

// clf for 5=best scale (instructor evaluation: 5=موافق بشدة=ممتاز)
const clf5=m=>{
  if(m>=4.50)return{l:'ممتاز',    en:'Excellent', bg:GREEN2,c:GREEN};
  if(m>=3.50)return{l:'جيد جداً', en:'Very Good',  bg:GREEN2,c:GREEN};
  if(m>=2.50)return{l:'جيد',      en:'Good',       bg:AMBER2,c:AMBER};
  if(m>=1.50)return{l:'مقبول',    en:'Acceptable', bg:RED2,  c:RED};
  return      {l:'ضعيف',          en:'Weak',        bg:RED2,  c:RED};
};
const clf1=m=>{
  if(m<=1.50)return{l:'ممتاز',    en:'Excellent',        bg:GREEN2,c:GREEN};
  if(m<=2.00)return{l:'جيد',      en:'Good',              bg:GREEN2,c:GREEN};
  if(m<=2.50)return{l:'مقبول',    en:'Acceptable',        bg:AMBER2,c:AMBER};
  if(m<=3.00)return{l:'ضعف',      en:'Weakness',          bg:RED2,  c:RED};
  return      {l:'ضعف حرج',       en:'Critical Weakness', bg:RED2,  c:'9C0006'};
};

const mC=(text,w,shade,opts={})=>new TableCell({
  width:{size:Math.max(1,w||400),type:WidthType.DXA},borders:allB(),
  shading:shade?{fill:shade,type:ShadingType.CLEAR}:undefined,
  margins:mg(),verticalAlign:VerticalAlign.CENTER,
  rowSpan:opts.rowSpan,columnSpan:opts.colSpan,
  children:[new Paragraph({alignment:opts.align||AlignmentType.CENTER,
    children:[new TextRun({text:String(text??''),bold:opts.bold||false,
      color:opts.color||'000000',size:opts.size||17,font:'Arial'})]})]
});

const mH=(lines,w,shade=DARK,size=16)=>new TableCell({
  width:{size:Math.max(1,w||400),type:WidthType.DXA},borders:allB(shade),
  shading:{fill:shade,type:ShadingType.CLEAR},margins:mg(),verticalAlign:VerticalAlign.CENTER,
  children:lines.map(l=>new Paragraph({alignment:AlignmentType.CENTER,spacing:{before:0,after:0},
    children:[new TextRun({text:l,bold:true,color:WHITE,size,font:'Arial'})]}))
});

const mP=(text,opts={})=>new Paragraph({
  alignment:opts.align||AlignmentType.RIGHT,
  spacing:{before:opts.before||80,after:opts.after||80},
  children:[new TextRun({text,bold:opts.bold||false,color:opts.color||'000000',
    size:opts.size||20,font:'Arial',italics:opts.italic||false})]
});
const sp=(b=100,a=100)=>new Paragraph({spacing:{before:b,after:a},children:[]});

const wMean=secs=>{const tn=secs.reduce((a,s)=>a+s.n,0);if(!tn)return null;return +(secs.reduce((a,s)=>a+s.sec_mean*s.n,0)/tn).toFixed(3);};
const wQMean=(secs,qi)=>{const tn=secs.reduce((a,s)=>a+s.n,0);if(!tn)return null;return +(secs.reduce((a,s)=>a+(s.questions[qi]?.mean||0)*s.n,0)/tn).toFixed(2);};

// ── Parse XLS ─────────────────────────────────────────────────────────────────
function parseXLS(buffer) {
  const wb = XLSX.read(buffer,{type:'buffer'});
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws,{header:1,defval:''});

  const sections=[];
  let current=null;

  for(let i=0;i<rows.length;i++){
    const row=rows[i];
    const rowStr=row.join(' ');

    if(rowStr.includes('المحاضر:')){
      if(current) sections.push(current);
      current={lecturer:String(row[2]||'').trim(),sec_num:String(row[15]||'').trim(),
        enrolled:'',dept:'',loc:'',evaluators:'',course:'',semester:'',questions:[]};
      for(let j=0;j<row.length-1;j++){
        const v=String(row[j]||'').trim();
        if(v==='الشعبة :')   current.sec_num   = String(row[j+2]||row[j+1]||'').trim();
        if(v==='المسجلين :') current.enrolled  = String(row[j+2]||row[j+1]||'').trim();
        if(v==='القسم :')    current.dept      = String(row[j+2]||row[j+1]||'').trim();
        if(v==='المقر :')    current.loc       = String(row[j+2]||row[j+1]||'').trim();
      }
    }

    if(current && rowStr.includes('المقيميين')){
      if(!current.evaluators) current.evaluators=String(row[23]||'').trim();
      if(!current.course)     current.course    =String(row[27]||'').trim();
      if(!current.semester)   current.semester  =String(row[32]||'').trim();
      for(let j=0;j<row.length-1;j++){
        const v=String(row[j]||'').trim();
        if(v.includes('المقيميين')) current.evaluators=String(row[j-1]||row[j+2]||current.evaluators||'').trim();
        if(v==='المقرر :')  current.course   =String(row[j+2]||row[j+1]||current.course||'').trim();
        if(v==='الفصل :')   current.semester =String(row[j+2]||row[j+1]||current.semester||'').trim();
      }
    }

    if(current){
      const qn=parseFloat(row[38]),mean=parseFloat(row[2]),text=String(row[22]||'').trim();
      if(!isNaN(qn)&&!isNaN(mean)&&qn>=1&&qn<=50&&mean>0){
        current.questions.push({
          qn:Math.round(qn),text,mean:+mean.toFixed(3),
          pct_sa:+(parseFloat(row[5])||0).toFixed(1),
          pct_a: +(parseFloat(row[8])||0).toFixed(1),
          pct_n: +(parseFloat(row[13])||0).toFixed(1),
          pct_d: +(parseFloat(row[17])||0).toFixed(1),
          pct_sd:+(parseFloat(row[20])||0).toFixed(1),
          cnt_sa:parseInt(row[4])||0,cnt_a:parseInt(row[7])||0,
          cnt_n:parseInt(row[12])||0,cnt_d:parseInt(row[16])||0,cnt_sd:parseInt(row[19])||0,
        });
      }
    }
  }
  if(current) sections.push(current);

  sections.forEach(s=>{
    s.n=parseInt(s.evaluators)||0;
    s.enrolled_num=parseInt(s.enrolled)||0;
    s.not_responded=Math.max(0,s.enrolled_num-s.n);
    s.participation_pct=s.enrolled_num>0?Math.round(s.n/s.enrolled_num*100):0;
    const qm=s.questions.map(q=>q.mean);
    s.sec_mean=qm.length?+(qm.reduce((a,b)=>a+b,0)/qm.length).toFixed(3):0;
  });

  return sections.filter(s=>s.questions.length>0);
}

// ── Upload ────────────────────────────────────────────────────────────────────
app.post('/api/upload',upload.single('file'),(req,res)=>{
  try{
    const sections=parseXLS(req.file.buffer);
    res.json({ok:true,sections,filename:req.file.originalname});
  }catch(e){res.status(400).json({ok:false,error:e.message});}
});

// ── Generate Word ─────────────────────────────────────────────────────────────
app.post('/api/word',async(req,res)=>{
  try{
    const {groups,meta,reportType}=req.body;
    const buf = reportType==='instructor'
      ? await buildInstructorWord(groups,meta)
      : await buildCourseWord(groups,meta);
    res.set({'Content-Type':'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      'Content-Disposition':`attachment; filename="${reportType}_report.docx"`});
    res.send(buf);
  }catch(e){console.error(e);res.status(500).json({error:e.message});}
});

// ── Participation table builder ───────────────────────────────────────────────
function buildParticipationTable(allSecs, CW, meta) {
  const totalEnrolled = allSecs.reduce((a,s)=>a+s.enrolled_num,0);
  const totalN        = allSecs.reduce((a,s)=>a+s.n,0);
  const totalNot      = allSecs.reduce((a,s)=>a+s.not_responded,0);
  const overallPct    = totalEnrolled>0?Math.round(totalN/totalEnrolled*100):0;

  const catLabel = meta.category==='staff'?'الكادر الإداري':
                   meta.category==='employee'?'الموظفون':
                   meta.category==='female'?'الطالبات':
                   meta.category==='male'?'الطلاب':'المشاركون';

  const rows=[];
  const pC=[600,3200,2400,1200,1200,1200,1400,CW-600-3200-2400-1200-1200-1200-1400];

  rows.push(new TableRow({children:[
    mH(['#'],pC[0]),mH([catLabel,'/ المقرر'],pC[1]),mH(['المحاضر / القسم'],pC[2]),
    mH(['إجمالي','المسجلين'],pC[3]),mH(['عدد','المقيّمين'],pC[4]),
    mH(['لم','يستبينوا'],pC[5]),mH(['نسبة','المشاركة'],pC[6]),
    mH(['المتوسط','العام'],pC[7]),
  ]}));

  allSecs.forEach((s,i)=>{
    const cl=clf(s.sec_mean);
    const bg=i%2===0?PALE:'FFFFFF';
    const notPct=s.enrolled_num>0?Math.round(s.not_responded/s.enrolled_num*100):0;
    rows.push(new TableRow({children:[
      mC(i+1,pC[0],bg,{bold:true,color:DARK,size:14}),
      mC(s.course||s.sec_num,pC[1],bg,{bold:true,color:DARK,align:AlignmentType.RIGHT,size:15}),
      mC(s.lecturer||s.dept,pC[2],bg,{align:AlignmentType.RIGHT,size:14}),
      mC(s.enrolled_num||'—',pC[3],bg,{bold:true}),
      mC(s.n,pC[4],bg,{bold:true,color:GREEN}),
      mC(s.not_responded,pC[5],s.not_responded>0?'FFF0F0':'FFFFFF',{bold:s.not_responded>0,color:s.not_responded>0?RED:'444444'}),
      mC(s.participation_pct+'%',pC[6],
        s.participation_pct>=80?GREEN2:s.participation_pct>=60?AMBER2:RED2,
        {bold:true,color:s.participation_pct>=80?GREEN:s.participation_pct>=60?AMBER:RED}),
      mC(s.sec_mean.toFixed(2),pC[7],cl.bg,{bold:true,color:cl.c,size:17}),
    ]}));
  });

  // Totals row
  rows.push(new TableRow({children:[
    mC('الإجمالي',pC[0]+pC[1]+pC[2],DARK,{bold:true,color:WHITE,colSpan:3}),
    mC(totalEnrolled,pC[3],PALE,{bold:true,color:DARK,size:18}),
    mC(totalN,pC[4],GREEN2,{bold:true,color:GREEN,size:18}),
    mC(totalNot,pC[5],totalNot>0?RED2:GREEN2,{bold:true,color:totalNot>0?RED:GREEN,size:18}),
    mC(overallPct+'%',pC[6],overallPct>=80?GREEN2:overallPct>=60?AMBER2:RED2,
      {bold:true,color:overallPct>=80?GREEN:overallPct>=60?AMBER:RED,size:18}),
    mC('—',pC[7],PALE),
  ]}));

  return {table: new Table({width:{size:CW,type:WidthType.DXA},columnWidths:pC,rows}),
    totalEnrolled, totalN, totalNot, overallPct};
}

// ── Build Course Word ─────────────────────────────────────────────────────────
async function buildCourseWord(groups, meta) {
  const CW=15398;
  const allSecs=groups.flatMap(g=>g.sections.map(s=>({...s,gender:g.gender})));
  const hasGender=groups.some(g=>g.gender);
  const nQ=allSecs[0]?.questions.length||0;
  const qTexts=allSecs[0]?.questions.map(q=>q.text)||[];
  const totalN=allSecs.reduce((a,s)=>a+s.n,0);

  // Group by course
  const courseMap={};
  allSecs.forEach(s=>{
    if(!courseMap[s.course]) courseMap[s.course]={F:[],M:[],all:[]};
    if(s.gender==='F') courseMap[s.course].F.push(s);
    else if(s.gender==='M') courseMap[s.course].M.push(s);
    courseMap[s.course].all.push(s);
  });

  const courses=Object.entries(courseMap).map(([code,g])=>({
    code,n:g.all.reduce((a,s)=>a+s.n,0),
    enrolled:g.all.reduce((a,s)=>a+s.enrolled_num,0),
    notResponded:g.all.reduce((a,s)=>a+s.not_responded,0),
    nF:g.F.reduce((a,s)=>a+s.n,0),nM:g.M.reduce((a,s)=>a+s.n,0),
    mean:wMean(g.all),meanF:wMean(g.F),meanM:wMean(g.M),
    qMeans:Array.from({length:nQ},(_,qi)=>wQMean(g.all,qi)),
    qMeansF:Array.from({length:nQ},(_,qi)=>wQMean(g.F,qi)),
    qMeansM:Array.from({length:nQ},(_,qi)=>wQMean(g.M,qi)),
    secs:g.all,
  }));

  const totalEnrolled=allSecs.reduce((a,s)=>a+s.enrolled_num,0);
  const totalNot=allSecs.reduce((a,s)=>a+s.not_responded,0);
  const overallPct=totalEnrolled>0?Math.round(totalN/totalEnrolled*100):0;
  const grandMean=+(courses.reduce((a,c)=>a+(c.mean||0)*c.n,0)/totalN).toFixed(3);
  const gCl=clf(grandMean);

  const children=[];

  // Title
  children.push(
    sp(0,60),
    mP(meta.sname||'تقرير استبانة تقييم المقررات الدراسية',{align:AlignmentType.CENTER,bold:true,size:44,color:DARK,before:0,after:60}),
    mP(meta.cname||'كليات الغد للعلوم الطبية التطبيقية – جدة',{align:AlignmentType.CENTER,size:22,color:'444444',before:0,after:40}),
    mP([meta.dept,meta.semester,meta.loc].filter(Boolean).join('  |  '),{align:AlignmentType.CENTER,size:18,color:'777777',before:0,after:160}),
  );

  // Main stats banner
  const bannerC=[Math.round(CW/6),Math.round(CW/6),Math.round(CW/6),Math.round(CW/6),Math.round(CW/6),CW-Math.round(CW/6)*5];
  children.push(
    mP('إحصائيات المشاركة والتقييم',{bold:true,size:22,color:DARK,before:0,after:80}),
    new Table({width:{size:CW,type:WidthType.DXA},columnWidths:bannerC,rows:[
      new TableRow({children:[
        mH(['إجمالي المسجلين'],bannerC[0]),
        mH(['المقيّمون'],bannerC[1]),
        mH(['لم يستبينوا'],bannerC[2]),
        mH(['نسبة المشاركة'],bannerC[3]),
        mH(['المقررات'],bannerC[4]),
        mH(['المتوسط العام'],bannerC[5]),
      ]}),
      new TableRow({children:[
        mC(totalEnrolled,bannerC[0],PALE,{bold:true,color:DARK,size:26}),
        mC(totalN,bannerC[1],GREEN2,{bold:true,color:GREEN,size:26}),
        mC(totalNot,bannerC[2],totalNot>0?RED2:GREEN2,{bold:true,color:totalNot>0?RED:GREEN,size:26}),
        mC(overallPct+'%',bannerC[3],overallPct>=80?GREEN2:overallPct>=60?AMBER2:RED2,
          {bold:true,color:overallPct>=80?GREEN:overallPct>=60?AMBER:RED,size:26}),
        mC(courses.length,bannerC[4],PALE,{bold:true,color:DARK,size:26}),
        mC(grandMean,bannerC[5],gCl.bg,{bold:true,color:gCl.c,size:32}),
      ]}),
    ]}),
    sp(200,80),
  );

  // Participation table
  const {table:partTable}=buildParticipationTable(allSecs,CW,meta);
  children.push(
    mP('أولاً: جدول المشاركة التفصيلية',{bold:true,size:24,color:DARK,before:0,after:80}),
    mP(`جدول يوضح إجمالي المسجلين والمقيّمين واللي لم يستبينوا لكل شعبة — ${meta.category==='female'?'الطالبات':meta.category==='male'?'الطلاب':meta.category==='staff'?'الكادر الإداري':'المشاركون'}`,
      {size:16,color:'555555',italic:true,before:0,after:80}),
    partTable,
    sp(200,80),
  );

  // Cross table Q × Course
  const nC=courses.length;
  const qTW=Math.round(CW*0.17);
  const cW=hasGender?Math.floor((CW-700-qTW)/(nC*2)):Math.floor((CW-700-qTW)/nC);
  const ovW=CW-700-qTW-(hasGender?cW*2*nC:cW*nC);
  const colWidths=[700,qTW,...(hasGender?courses.flatMap(()=>[cW,cW]):courses.map(()=>cW)),ovW];

  const h1=new TableRow({children:[
    mH(['Q#'],700),mH(['السؤال / Criteria'],qTW),
    ...courses.map(c=>{
      const cl=clf(c.mean||0);
      if(hasGender) return new TableCell({width:{size:cW*2},columnSpan:2,borders:allB(MID),
        shading:{fill:MID,type:ShadingType.CLEAR},margins:mg(),verticalAlign:VerticalAlign.CENTER,
        children:[
          new Paragraph({alignment:AlignmentType.CENTER,spacing:{before:0,after:0},children:[new TextRun({text:c.code,bold:true,color:WHITE,size:16,font:'Arial'})]}),
          new Paragraph({alignment:AlignmentType.CENTER,spacing:{before:0,after:0},children:[new TextRun({text:'('+( c.mean||0).toFixed(2)+')',color:'BDD7EE',size:13,font:'Arial'})]}),
        ]});
      const bg=cl.bg;
      return new TableCell({width:{size:cW},borders:allB(MID),shading:{fill:MID,type:ShadingType.CLEAR},margins:mg(),verticalAlign:VerticalAlign.CENTER,
        children:[
          new Paragraph({alignment:AlignmentType.CENTER,spacing:{before:0,after:0},children:[new TextRun({text:c.code,bold:true,color:WHITE,size:15,font:'Arial'})]}),
          new Paragraph({alignment:AlignmentType.CENTER,spacing:{before:0,after:0},children:[new TextRun({text:'('+(c.mean||0).toFixed(2)+')',color:'BDD7EE',size:13,font:'Arial'})]}),
        ]});
    }),
    mH(['الإجمالي'],ovW),
  ]});

  const h2=hasGender?[new TableRow({children:[
    new TableCell({width:{size:700},rowSpan:0,borders:allB(),shading:{fill:DARK,type:ShadingType.CLEAR},margins:mg(),children:[new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:'#',bold:true,color:WHITE,size:15,font:'Arial'})]}),]}),
    new TableCell({width:{size:qTW},rowSpan:0,borders:allB(),shading:{fill:DARK,type:ShadingType.CLEAR},margins:mg(),children:[new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:'السؤال',bold:true,color:WHITE,size:15,font:'Arial'})]}),]}),
    ...courses.flatMap(()=>[
      new TableCell({width:{size:cW},borders:allB('843C0C'),shading:{fill:PINK,type:ShadingType.CLEAR},margins:mg(),children:[new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:'F',bold:true,color:'843C0C',size:17,font:'Arial'})]}),]}),
      new TableCell({width:{size:cW},borders:allB('1F4E79'),shading:{fill:BLUE2,type:ShadingType.CLEAR},margins:mg(),children:[new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:'M',bold:true,color:DARK,size:17,font:'Arial'})]}),]}),
    ]),
    new TableCell({width:{size:ovW},rowSpan:0,borders:allB(),shading:{fill:DARK,type:ShadingType.CLEAR},margins:mg(),children:[new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:'الإجمالي',bold:true,color:WHITE,size:15,font:'Arial'})]}),]}),
  ]})]:[]; 

  const overallQM=Array.from({length:nQ},(_,qi)=>+(courses.reduce((a,c)=>a+(c.qMeans[qi]||0)*c.n,0)/totalN).toFixed(2));

  const dataRows=Array.from({length:nQ},(_,qi)=>{
    const oM=overallQM[qi];const oCl=clf(oM);const bg=qi%2===0?PALE:'FFFFFF';
    return new TableRow({children:[
      mC(`Q${qi+1}`,700,bg,{bold:true,color:DARK,size:14}),
      mC((qTexts[qi]||'').slice(0,55),qTW,bg,{align:AlignmentType.RIGHT,size:13}),
      ...courses.flatMap(c=>{
        if(hasGender){
          const fM=c.qMeansF[qi],mM=c.qMeansM[qi];
          return [
            new TableCell({width:{size:cW},borders:allB(),shading:{fill:PINK,type:ShadingType.CLEAR},margins:mg(),children:[new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:fM!=null?fM.toFixed(2):'—',bold:!!fM,color:fM?'843C0C':'AAAAAA',size:14,font:'Arial'})]}),]}),
            new TableCell({width:{size:cW},borders:allB(),shading:{fill:BLUE2,type:ShadingType.CLEAR},margins:mg(),children:[new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:mM!=null?mM.toFixed(2):'—',bold:!!mM,color:mM?DARK:'AAAAAA',size:14,font:'Arial'})]}),]}),
          ];
        }
        const qM=c.qMeans[qi];const qCl=qM!=null?clf(qM):{bg,c:'000000'};
        return [mC(qM!=null?qM.toFixed(2):'—',cW,qCl.bg,{color:qCl.c,size:14})];
      }),
      mC(oM.toFixed(2),ovW,oCl.bg,{bold:true,color:oCl.c,size:16}),
    ]});
  });

  const meanRow=new TableRow({children:[
    mC('المتوسط',700+qTW,DARK,{bold:true,color:WHITE,colSpan:2}),
    ...courses.flatMap(c=>{
      const cl=clf(c.mean||0);
      if(hasGender){
        return [
          new TableCell({width:{size:cW},borders:allB(),shading:{fill:PINK,type:ShadingType.CLEAR},margins:mg(),children:[new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:c.meanF!=null?c.meanF.toFixed(2):'—',bold:true,color:'843C0C',size:17,font:'Arial'})]}),]}),
          new TableCell({width:{size:cW},borders:allB(),shading:{fill:BLUE2,type:ShadingType.CLEAR},margins:mg(),children:[new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:c.meanM!=null?c.meanM.toFixed(2):'—',bold:true,color:DARK,size:17,font:'Arial'})]}),]}),
        ];
      }
      return [mC((c.mean||0).toFixed(2),cW,cl.bg,{bold:true,color:cl.c,size:18})];
    }),
    mC(grandMean.toFixed(2),ovW,gCl.bg,{bold:true,color:gCl.c,size:20}),
  ]});

  const nRow=new TableRow({children:[
    mC('المقيّمون',700+qTW,MID,{bold:true,color:WHITE,colSpan:2}),
    ...courses.flatMap(c=>{
      if(hasGender) return [
        new TableCell({width:{size:cW},borders:allB(),shading:{fill:PINK,type:ShadingType.CLEAR},margins:mg(),children:[new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:String(c.nF||'—'),bold:true,color:'843C0C',size:15,font:'Arial'})]}),]}),
        new TableCell({width:{size:cW},borders:allB(),shading:{fill:BLUE2,type:ShadingType.CLEAR},margins:mg(),children:[new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:String(c.nM||'—'),bold:true,color:DARK,size:15,font:'Arial'})]}),]}),
      ];
      return [mC(c.n,cW,PALE,{bold:true,color:DARK})];
    }),
    mC(totalN,ovW,PALE,{bold:true,color:DARK,size:18}),
  ]});

  children.push(
    mP('ثانياً: جدول المتوسطات الحسابية لكل سؤال عبر المقررات',{bold:true,size:24,color:DARK,before:0,after:60}),
    hasGender?mP('الأعمدة الوردية = طالبات (F) | الأعمدة الزرقاء = طلاب (M)',{size:15,color:'777777',italic:true,before:0,after:80}):sp(40,60),
    new Table({width:{size:CW,type:WidthType.DXA},columnWidths:colWidths,rows:[h1,...h2,...dataRows,meanRow,nRow]}),
    sp(200,80),
  );

  // Per-section detail
  children.push(
    new Paragraph({pageBreakBefore:true,children:[]}),
    mP('ثالثاً: التحليل التفصيلي لكل شعبة',{bold:true,size:28,color:DARK,before:0,after:120}),
  );

  allSecs.forEach((sec,si)=>{
    const cl=clf(sec.sec_mean);
    const gTag=sec.gender==='F'?'👩 إناث':sec.gender==='M'?'👨 ذكور':'';
    children.push(
      mP(`${sec.course} — الشعبة ${sec.sec_num} ${gTag}`,{bold:true,size:22,color:MID,before:si===0?0:200,after:40}),
      mP([`المحاضر: ${sec.lecturer}`,`المقيّمون: ${sec.n}`,`المسجلون: ${sec.enrolled_num||'—'}`,
        sec.not_responded>0?`لم يستبينوا: ${sec.not_responded} (${100-sec.participation_pct}%)`:'',
        `نسبة المشاركة: ${sec.participation_pct}%`,`المتوسط: ${sec.sec_mean.toFixed(2)}`,cl.l
      ].filter(Boolean).join('  |  '),{size:15,color:'444444',before:0,after:80}),
    );
    const dC=[500,3600,1300,1300,1300,1300,1300,1400,CW-500-3600-1300*5-1400];
    children.push(new Table({width:{size:CW,type:WidthType.DXA},columnWidths:dC,rows:[
      new TableRow({children:[mH(['Q#'],dC[0]),mH(['السؤال'],dC[1]),
        mH(['موافق','بشدة%'],dC[2]),mH(['موافق%'],dC[3]),mH(['حد ما%'],dC[4]),
        mH(['لا أوافق%'],dC[5]),mH(['لا أوافق','بشدة%'],dC[6]),
        mH(['المتوسط'],dC[7]),mH(['التصنيف'],dC[8])]}),
      ...sec.questions.map((q,qi)=>{
        const qcl=clf(q.mean);const bg=qi%2===0?PALE:'FFFFFF';
        return new TableRow({children:[
          mC(`Q${q.qn}`,dC[0],bg,{bold:true,color:DARK,size:13}),
          mC(q.text.slice(0,60),dC[1],bg,{align:AlignmentType.RIGHT,size:12}),
          mC(q.pct_sa+'%',dC[2],q.pct_sa>=80?GREEN2:bg,{color:q.pct_sa>=80?GREEN:'000000',size:13}),
          mC(q.pct_a+'%',dC[3],bg,{size:13}),mC(q.pct_n+'%',dC[4],bg,{size:13}),
          mC(q.pct_d+'%',dC[5],q.pct_d>15?RED2:bg,{color:q.pct_d>15?RED:'000000',size:13}),
          mC(q.pct_sd+'%',dC[6],q.pct_sd>10?RED2:bg,{color:q.pct_sd>10?RED:'000000',size:13}),
          mC(q.mean.toFixed(2),dC[7],qcl.bg,{bold:true,color:qcl.c,size:15}),
          mC(qcl.l,dC[8],qcl.bg,{bold:true,color:qcl.c,size:13}),
        ]});
      }),
      new TableRow({children:[
        mC('المتوسط',dC[0]+dC[1],MID,{bold:true,color:WHITE,colSpan:2}),
        mC('—',dC[2],PALE),mC('—',dC[3],PALE),mC('—',dC[4],PALE),mC('—',dC[5],PALE),mC('—',dC[6],PALE),
        mC(sec.sec_mean.toFixed(2),dC[7],cl.bg,{bold:true,color:cl.c,size:17}),
        mC(cl.l,dC[8],cl.bg,{bold:true,color:cl.c,size:13}),
      ]}),
    ]}),sp(60,40));
  });

  return buildDoc(children);
}

// ── Build Instructor Word ─────────────────────────────────────────────────────
async function buildInstructorWord(groups,meta){
  const CW=15398;
  const allSecs=groups.flatMap(g=>g.sections.map(s=>({...s,gender:g.gender})));
  const hasGender=groups.some(g=>g.gender);
  const nQ=allSecs[0]?.questions.length||0;
  const qTexts=allSecs[0]?.questions.map(q=>q.text)||[];
  const totalN=allSecs.reduce((a,s)=>a+s.n,0);
  const totalEnrolled=allSecs.reduce((a,s)=>a+s.enrolled_num,0);
  const totalNot=allSecs.reduce((a,s)=>a+s.not_responded,0);
  const overallPct=totalEnrolled>0?Math.round(totalN/totalEnrolled*100):0;
  const grandMean=+(allSecs.reduce((a,s)=>a+s.sec_mean,0)/(allSecs.length||1)).toFixed(3);
  const gCl=clf(grandMean);

  const lecMap={};
  allSecs.forEach(s=>{
    if(!lecMap[s.lecturer]) lecMap[s.lecturer]={F:[],M:[],all:[]};
    if(s.gender==='F') lecMap[s.lecturer].F.push(s);
    else if(s.gender==='M') lecMap[s.lecturer].M.push(s);
    lecMap[s.lecturer].all.push(s);
  });

  const lecturers=Object.entries(lecMap).map(([name,g])=>({
    name,n:g.all.reduce((a,s)=>a+s.n,0),
    enrolled:g.all.reduce((a,s)=>a+s.enrolled_num,0),
    notResponded:g.all.reduce((a,s)=>a+s.not_responded,0),
    nF:g.F.reduce((a,s)=>a+s.n,0),nM:g.M.reduce((a,s)=>a+s.n,0),
    mean:wMean(g.all),meanF:wMean(g.F),meanM:wMean(g.M),
    qMeans:Array.from({length:nQ},(_,qi)=>wQMean(g.all,qi)),
    secs:g.all,courses:[...new Set(g.all.map(s=>s.course))],
  })).sort((a,b)=>b.mean-a.mean);

  const children=[];

  // Title
  children.push(
    sp(0,60),
    mP(meta.sname||'تقرير استبانة تقييم المحاضرين',{align:AlignmentType.CENTER,bold:true,size:44,color:DARK,before:0,after:60}),
    mP(meta.cname||'كليات الغد للعلوم الطبية التطبيقية – جدة',{align:AlignmentType.CENTER,size:22,color:'444444',before:0,after:40}),
    mP([meta.dept,meta.semester].filter(Boolean).join('  |  '),{align:AlignmentType.CENTER,size:18,color:'777777',before:0,after:160}),
  );

  // Banner
  const bannerC=[Math.round(CW/6),Math.round(CW/6),Math.round(CW/6),Math.round(CW/6),Math.round(CW/6),CW-Math.round(CW/6)*5];
  children.push(
    mP('إحصائيات المشاركة والتقييم',{bold:true,size:22,color:DARK,before:0,after:80}),
    new Table({width:{size:CW,type:WidthType.DXA},columnWidths:bannerC,rows:[
      new TableRow({children:[mH(['إجمالي المسجلين'],bannerC[0]),mH(['المقيّمون'],bannerC[1]),
        mH(['لم يستبينوا'],bannerC[2]),mH(['نسبة المشاركة'],bannerC[3]),
        mH(['المحاضرون'],bannerC[4]),mH(['المتوسط العام'],bannerC[5])]}),
      new TableRow({children:[
        mC(totalEnrolled,bannerC[0],PALE,{bold:true,color:DARK,size:26}),
        mC(totalN,bannerC[1],GREEN2,{bold:true,color:GREEN,size:26}),
        mC(totalNot,bannerC[2],totalNot>0?RED2:GREEN2,{bold:true,color:totalNot>0?RED:GREEN,size:26}),
        mC(overallPct+'%',bannerC[3],overallPct>=80?GREEN2:overallPct>=60?AMBER2:RED2,{bold:true,color:overallPct>=80?GREEN:overallPct>=60?AMBER:RED,size:26}),
        mC(lecturers.length,bannerC[4],PALE,{bold:true,color:DARK,size:26}),
        mC(grandMean,bannerC[5],gCl.bg,{bold:true,color:gCl.c,size:32}),
      ]}),
    ]}),sp(180,80),
  );

  // Participation table
  const {table:partTable}=buildParticipationTable(allSecs,CW,meta);
  children.push(
    mP('أولاً: جدول المشاركة التفصيلية',{bold:true,size:24,color:DARK,before:0,after:80}),
    partTable,sp(180,80),
  );

  // Lecturer summary
  const lqW=Math.floor((CW-500-2800-700-700-1400)/Math.max(nQ,1));
  const t2C=[500,2800,700,700,...Array(nQ).fill(lqW),1400,CW-500-2800-700-700-lqW*nQ-1400];
  children.push(
    mP('ثانياً: ملخص المحاضرين (متوسط موزون)',{bold:true,size:24,color:DARK,before:0,after:80}),
    new Table({width:{size:CW,type:WidthType.DXA},columnWidths:t2C,rows:[
      new TableRow({children:[
        mH(['#'],t2C[0]),mH(['المحاضر'],t2C[1]),mH(['الشعب'],t2C[2]),mH(['المقيّمون'],t2C[3]),
        ...Array.from({length:nQ},(_,i)=>mH([`Q${i+1}`],t2C[4+i],MID,13)),
        mH(['المتوسط','الموزون'],t2C[4+nQ]),mH(['التصنيف'],t2C[5+nQ]),
      ]}),
      ...lecturers.map((lec,i)=>{
        const cl=clf(lec.mean||0);const bg=i%2===0?PALE:'FFFFFF';
        return new TableRow({children:[
          mC(i+1,t2C[0],bg,{bold:true,color:DARK,size:13}),
          mC(lec.name,t2C[1],bg,{align:AlignmentType.RIGHT,size:13}),
          mC((lec.secs||[]).length,t2C[2],bg),mC(lec.n,t2C[3],bg,{bold:true}),
          ...lec.qMeans.map((qm,qi)=>{const qcl=qm!=null?clf(qm):{bg,c:'000000'};
            return mC(qm!=null?qm.toFixed(2):'—',t2C[4+qi],qcl.bg,{color:qcl.c,size:12});}),
          mC((lec.mean||0).toFixed(2),t2C[4+nQ],cl.bg,{bold:true,color:cl.c,size:16}),
          mC(cl.l,t2C[5+nQ],cl.bg,{bold:true,color:cl.c,size:12}),
        ]});
      }),
    ]}),sp(200,80),
  );

  // Q × Course cross table
  const uniqueCourses=[...new Set(allSecs.map(s=>s.course))];
  const nCC=uniqueCourses.length;
  const ccW=Math.max(500,Math.floor((CW-1800-1400)/Math.max(hasGender?nCC*2:nCC,1)));
  const t3LastW=Math.max(400,CW-1800-(hasGender?ccW*2*nCC:ccW*nCC)-1400);
  const t3C=[1800,...(hasGender?uniqueCourses.flatMap(()=>[ccW,ccW]):uniqueCourses.map(()=>ccW)),t3LastW];

  const cQM={},cQMF={},cQMM={},cMeans={};
  uniqueCourses.forEach(code=>{
    const cs=allSecs.filter(s=>s.course===code);
    const csF=cs.filter(s=>s.gender==='F'),csM=cs.filter(s=>s.gender==='M');
    const tn=cs.reduce((a,s)=>a+s.n,0),tnF=csF.reduce((a,s)=>a+s.n,0),tnM=csM.reduce((a,s)=>a+s.n,0);
    cQM[code]=Array.from({length:nQ},(_,qi)=>tn?+(cs.reduce((a,s)=>a+(s.questions[qi]?.mean||0)*s.n,0)/tn).toFixed(2):null);
    cQMF[code]=Array.from({length:nQ},(_,qi)=>tnF?+(csF.reduce((a,s)=>a+(s.questions[qi]?.mean||0)*s.n,0)/tnF).toFixed(2):null);
    cQMM[code]=Array.from({length:nQ},(_,qi)=>tnM?+(csM.reduce((a,s)=>a+(s.questions[qi]?.mean||0)*s.n,0)/tnM).toFixed(2):null);
    cMeans[code]=tn?+(cs.reduce((a,s)=>a+s.sec_mean*s.n,0)/tn).toFixed(2):0;
  });

  const ch_h1=new TableRow({children:[
    mH(['السؤال'],t3C[0]),
    ...uniqueCourses.map((code,ci)=>{
      const sz=hasGender?ccW*2:ccW;
      const opt=hasGender?{columnSpan:2}:{};
      return new TableCell({width:{size:sz},columnSpan:hasGender?2:1,borders:allB(MID),
        shading:{fill:MID,type:ShadingType.CLEAR},margins:mg(),verticalAlign:VerticalAlign.CENTER,
        children:[
          new Paragraph({alignment:AlignmentType.CENTER,spacing:{before:0,after:0},children:[new TextRun({text:code,bold:true,color:WHITE,size:15,font:'Arial'})]}),
          new Paragraph({alignment:AlignmentType.CENTER,spacing:{before:0,after:0},children:[new TextRun({text:'('+cMeans[code].toFixed(2)+')',color:'BDD7EE',size:12,font:'Arial'})]}),
        ]});
    }),
    mH(['الإجمالي'],t3C[t3C.length-1]),
  ]});

  const ch_h2=hasGender?[new TableRow({children:[
    new TableCell({width:{size:1800},rowSpan:0,borders:allB(),shading:{fill:DARK,type:ShadingType.CLEAR},margins:mg(),children:[new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:'السؤال',bold:true,color:WHITE,size:14,font:'Arial'})]}),]}),
    ...uniqueCourses.flatMap(()=>[
      new TableCell({width:{size:ccW},borders:allB('843C0C'),shading:{fill:PINK,type:ShadingType.CLEAR},margins:mg(),children:[new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:'F',bold:true,color:'843C0C',size:16,font:'Arial'})]}),]}),
      new TableCell({width:{size:ccW},borders:allB('1F4E79'),shading:{fill:BLUE2,type:ShadingType.CLEAR},margins:mg(),children:[new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:'M',bold:true,color:DARK,size:16,font:'Arial'})]}),]}),
    ]),
    new TableCell({width:{size:t3C[t3C.length-1]},rowSpan:0,borders:allB(),shading:{fill:DARK,type:ShadingType.CLEAR},margins:mg(),children:[new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:'الإجمالي',bold:true,color:WHITE,size:14,font:'Arial'})]}),]}),
  ]})]:[]; 

  const ovQM=Array.from({length:nQ},(_,qi)=>+(allSecs.reduce((a,s)=>a+(s.questions[qi]?.mean||0)*s.n,0)/totalN).toFixed(2));

  const ch_data=Array.from({length:nQ},(_,qi)=>{
    const oM=ovQM[qi];const oCl=clf(oM);const bg=qi%2===0?PALE:'FFFFFF';
    return new TableRow({children:[
      mC((`Q${qi+1} — `+(qTexts[qi]||'').slice(0,32)),t3C[0],bg,{align:AlignmentType.RIGHT,size:12}),
      ...uniqueCourses.flatMap(code=>{
        if(hasGender) return [
          new TableCell({width:{size:ccW},borders:allB(),shading:{fill:PINK,type:ShadingType.CLEAR},margins:mg(),children:[new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:cQMF[code][qi]!=null?cQMF[code][qi].toFixed(2):'—',bold:!!cQMF[code][qi],color:cQMF[code][qi]?'843C0C':'AAAAAA',size:13,font:'Arial'})]}),]}),
          new TableCell({width:{size:ccW},borders:allB(),shading:{fill:BLUE2,type:ShadingType.CLEAR},margins:mg(),children:[new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:cQMM[code][qi]!=null?cQMM[code][qi].toFixed(2):'—',bold:!!cQMM[code][qi],color:cQMM[code][qi]?DARK:'AAAAAA',size:13,font:'Arial'})]}),]}),
        ];
        const qM=cQM[code][qi];const qCl=qM!=null?clf(qM):{bg,c:'000000'};
        return [mC(qM!=null?qM.toFixed(2):'—',ccW,qCl.bg,{color:qCl.c,size:13})];
      }),
      mC(oM.toFixed(2),t3C[t3C.length-1],oCl.bg,{bold:true,color:oCl.c,size:15}),
    ]});
  });

  const ch_mean=new TableRow({children:[
    mC('المتوسط',t3C[0],DARK,{bold:true,color:WHITE}),
    ...uniqueCourses.flatMap(code=>{
      const cl=clf(cMeans[code]);
      if(hasGender){
        const csF=allSecs.filter(s=>s.course===code&&s.gender==='F');
        const csM=allSecs.filter(s=>s.course===code&&s.gender==='M');
        const tnF=csF.reduce((a,s)=>a+s.n,0),tnM=csM.reduce((a,s)=>a+s.n,0);
        const mF=tnF?+(csF.reduce((a,s)=>a+s.sec_mean*s.n,0)/tnF).toFixed(2):null;
        const mM=tnM?+(csM.reduce((a,s)=>a+s.sec_mean*s.n,0)/tnM).toFixed(2):null;
        return [
          new TableCell({width:{size:ccW},borders:allB(),shading:{fill:PINK,type:ShadingType.CLEAR},margins:mg(),children:[new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:mF!=null?mF.toFixed(2):'—',bold:true,color:'843C0C',size:15,font:'Arial'})]}),]}),
          new TableCell({width:{size:ccW},borders:allB(),shading:{fill:BLUE2,type:ShadingType.CLEAR},margins:mg(),children:[new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:mM!=null?mM.toFixed(2):'—',bold:true,color:DARK,size:15,font:'Arial'})]}),]}),
        ];
      }
      return [mC(cMeans[code].toFixed(2),ccW,cl.bg,{bold:true,color:cl.c,size:16})];
    }),
    mC(grandMean.toFixed(2),t3C[t3C.length-1],gCl.bg,{bold:true,color:gCl.c,size:18}),
  ]});

  children.push(
    mP('ثالثاً: جدول مقارنة الأسئلة عبر المقررات',{bold:true,size:24,color:DARK,before:0,after:60}),
    hasGender?mP('الوردي = طالبات (F) | الأزرق = طلاب (M)',{size:15,color:'777777',italic:true,before:0,after:80}):sp(40,60),
    new Table({width:{size:CW,type:WidthType.DXA},columnWidths:t3C,rows:[ch_h1,...ch_h2,...ch_data,ch_mean]}),
  );

  return buildDoc(children);
}

function buildDoc(children){
  const doc=new Document({
    numbering:{config:[{reference:'bullets',levels:[{level:0,format:'bullet',text:'•',
      alignment:AlignmentType.RIGHT,
      style:{paragraph:{indent:{left:500,hanging:300}},run:{font:'Arial',size:20}}}]}]},
    styles:{default:{document:{run:{font:'Arial',size:18}}}},
    sections:[{properties:{page:{
      size:{width:12240,height:15840,orientation:PageOrientation.LANDSCAPE},
      margin:{top:720,right:720,bottom:720,left:720}
    }},children}]
  });
  return Packer.toBuffer(doc);
}


// ── Batch chart fetch helper ─────────────────────────────────────────────────
async function fetchChartsAll(questions, f_lbl, m_lbl){
  try{
    const http=require('http');
    const qs=questions.map(q=>({
      qn:q.qn, lbl:q.lbl||'',
      fD:q.fD||q.cD||[0,0,0,0,0],
      mD:q.mD||q.cD||[0,0,0,0,0],
      f_lbl:f_lbl||'Female', m_lbl:m_lbl||'Male',
      secName:q.secName||''
    }));
    return await new Promise((resolve)=>{
      const body=JSON.stringify({questions:qs});
      const req=http.request({
        hostname:'localhost', port:process.env.PORT||3000,
        path:'/api/chart', method:'POST',
        headers:{'Content-Type':'application/json','Content-Length':Buffer.byteLength(body)}
      }, res=>{
        let d=''; res.on('data',c=>d+=c);
        res.on('end',()=>{ try{resolve(JSON.parse(d).charts||{});}catch{resolve({});} });
      });
      req.on('error',()=>resolve({}));
      req.setTimeout(60000,()=>{req.destroy();resolve({});});
      req.write(body); req.end();
    });
  }catch{return {};}
}

// New word endpoint matching new frontend
app.post('/api/generate-word',async(req,res)=>{
  try{
    const {result,cfg}=req.body;
    if(!result){return res.status(400).json({error:'No result data'});}
    if(!result.instructorMode && (!result.secs || !Array.isArray(result.secs))){
      return res.status(400).json({error:'بيانات الاستبيان غير مكتملة — حاولي مرة ثانية'});
    }

    // If instructor mode
    if(result.instructorMode){
      const buf=await buildInstructorWordFromResult(result,cfg);
      res.set({'Content-Type':'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        'Content-Disposition':'attachment; filename="Instructor_Report.docx"'});
      return res.send(buf);
    }

    // Regular course report
    const buf=await buildWordFromResult(result,cfg);
    res.set({'Content-Type':'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      'Content-Disposition':'attachment; filename="Survey_Report.docx"'});
    res.send(buf);
  }catch(e){console.error(e);res.status(500).json({error:e.message});}
});

// ── Build Word from result (course) ──────────────────────────────────────────
async function fetchChartsAll(questions, f_lbl, m_lbl){
  try{
    const http=require('http');
    const qs=questions.map(q=>({
      qn:q.qn, lbl:q.lbl||'',
      fD:q.fD||q.cD||[0,0,0,0,0],
      mD:q.mD||q.cD||[0,0,0,0,0],
      f_lbl:f_lbl||'Female', m_lbl:m_lbl||'Male',
      secName:q.secName||''
    }));
    return await new Promise((resolve)=>{
      const body=JSON.stringify({questions:qs});
      const req=http.request({
        hostname:'localhost', port:process.env.PORT||3000,
        path:'/api/chart', method:'POST',
        headers:{'Content-Type':'application/json','Content-Length':Buffer.byteLength(body)}
      }, res=>{
        let d=''; res.on('data',c=>d+=c);
        res.on('end',()=>{try{resolve(JSON.parse(d).charts||{});}catch{resolve({});}});
      });
      req.on('error',()=>resolve({}));
      req.setTimeout(90000,()=>{req.destroy();resolve({});});
      req.write(body); req.end();
    });
  }catch{return {};}
}

async function buildWordFromResult(result, cfg){
  // Use clf1 (1=best) for single survey, clf5 (5=best) for XLS multi-course
  // Use scale from cfg: '5best' = XLS scale, '1best' = single survey scale
  const scale=cfg.scale||(cfg.isMulti?'5best':'1best');
  const clfR=scale==='5best'?clf5:clf1;
  console.log('Report scale:', scale, '| isMulti:', cfg.isMulti);
  const CW=15398;
  const {nF,nM,n,secs,overall}=result;
  const showG=(cfg.gmode==='col');
  const children=[];
  const gCl=clfR(overall);
  const totalQ=secs.reduce((a,s)=>a+(s.qs||[]).length,0);
  const courseResults=cfg.courseResults||null;
  const isMulti=cfg.isMulti&&courseResults&&Object.keys(courseResults).length>1;
  const courses=isMulti?Object.keys(courseResults):[];

  // ── TITLE ──────────────────────────────────────────────────────────────
  children.push(
    sp(0,200),
    mP(cfg.sname||'تقرير استبانة',{align:AlignmentType.CENTER,bold:true,size:52,color:DARK,before:0,after:80}),
    mP('تحليل نتائج الاستبيان',{align:AlignmentType.CENTER,bold:true,size:30,color:MID,before:0,after:80}),
    mP(cfg.cname||'كليات الغد للعلوم الطبية التطبيقية – جدة',{align:AlignmentType.CENTER,size:22,color:'555555',before:0,after:60}),
    mP(cfg.semester||'',{align:AlignmentType.CENTER,size:20,color:'777777',before:0,after:300}),
  );

  // ── SCALE ──────────────────────────────────────────────────────────────
  const sCols=[Math.round(CW*0.22),Math.round(CW*0.15),Math.round(CW*0.15),Math.round(CW*0.15),Math.round(CW*0.15)];
  children.push(
    mP('المقياس المستخدم | Scale',{bold:true,size:22,color:DARK,before:0,after:80}),
    new Table({width:{size:CW,type:WidthType.DXA},columnWidths:sCols,rows:[
      new TableRow({children:[
        ...(scale==='5best'?[
          mH(['5=موافق بشدة\nStrongly Agree'],sCols[0],GREEN),
          mH(['4=موافق\nAgree'],sCols[1],GREEN),
          mH(['3=محايد\nNeutral'],sCols[2],'7F7F7F'),
          mH(['2=لا أوافق\nDisagree'],sCols[3],RED),
          mH(['1=لا أوافق بشدة\nStr. Disagree'],sCols[4],RED),
        ]:[
          mH(['1=موافق بشدة\nStrongly Agree'],sCols[0],GREEN),
          mH(['2=موافق\nAgree'],sCols[1],GREEN),
          mH(['3=محايد\nNeutral'],sCols[2],'7F7F7F'),
          mH(['4=لا أوافق\nDisagree'],sCols[3],RED),
          mH(['5=لا أوافق بشدة\nStr. Disagree'],sCols[4],RED),
        ]),
      ]}),
    ]}),sp(100,60),
  );

  // ── CLASSIFICATION ──────────────────────────────────────────────────────
  const clsC=[Math.round(CW*0.10),Math.round(CW*0.16),Math.round(CW*0.14),CW-Math.round(CW*0.40)];
  children.push(
    mP('Classification Scale | مقياس التصنيف',{bold:true,size:22,color:DARK,before:0,after:80}),
    new Table({width:{size:CW,type:WidthType.DXA},columnWidths:clsC,rows:[
      new TableRow({children:[mH(['Range'],clsC[0]),mH(['Classification'],clsC[1]),mH(['التصنيف'],clsC[2]),mH(['Interpretation'],clsC[3])]}),
      ...(scale==='1best'?[
        ['≤1.50','Excellent / Clear Strength','ممتاز / نقطة قوة','Excellent — sustain',GREEN2,GREEN],
        ['1.51–2.00','Good / Strength','جيد / نقطة قوة','Good — maintain',GREEN2,GREEN],
        ['2.01–2.50','Acceptable','مقبول / يحتاج متابعة','Acceptable — monitor',AMBER2,AMBER],
        ['2.51–3.00','Weakness','ضعف / يحتاج تحسين','Weakness — action plan',RED2,RED],
        ['>3.00','Critical Weakness','ضعف حرج','Critical — immediate action',RED2,'9C0006'],
      ]:[
        ['≥4.50','Excellent','ممتاز','Strong positive outcome',GREEN2,GREEN],
        ['3.50–4.49','Very Good','جيد جداً','Good performance',GREEN2,GREEN],
        ['2.50–3.49','Good','جيد','Acceptable — monitor',AMBER2,AMBER],
        ['1.50–2.49','Acceptable','مقبول','Needs improvement',RED2,RED],
        ['<1.50','Weak','ضعيف','Immediate action required',RED2,RED],
      ]).map(([r,cl,ar,interp,bg,c])=>new TableRow({children:[
        mC(r,clsC[0],bg,{bold:true,color:c,align:AlignmentType.CENTER}),
        mC(cl,clsC[1],bg,{bold:true,color:c}),
        mC(ar,clsC[2],bg,{bold:true,color:c}),
        mC(interp,clsC[3],WHITE,{size:16,align:AlignmentType.LEFT}),
      ]}))
    ]}),sp(140,80),
  );

  // ── SAMPLE PROFILE ──────────────────────────────────────────────────────
  const spC=[Math.round(CW*0.22),Math.round(CW*0.12),Math.round(CW*0.22),Math.round(CW*0.12)];
  const profileRows=[
    ['Total Respondents',n,'إجمالي المشاركين',n],
    ['Female (إناث)',nF||'—','إناث',nF||'—'],
    ['Male (ذكور)',nM||'—','ذكور',nM||'—'],
    ...(isMulti?[['No. of Courses',courses.length,'عدد المقررات',courses.length]]:[] ),
    ['No. of Questions',totalQ,'عدد الأسئلة',totalQ],
    ['Overall Mean',overall,'المتوسط العام',overall],
    ['Period',cfg.semester||'—','الفصل الدراسي',cfg.semester||'—'],
  ];
  children.push(
    mP('Sample Profile | بيانات العينة',{bold:true,size:22,color:DARK,before:0,after:80}),
    new Table({width:{size:CW,type:WidthType.DXA},columnWidths:spC,rows:[
      new TableRow({children:[mH(['Detail'],spC[0]),mH(['Value'],spC[1]),mH(['التفاصيل'],spC[2]),mH(['القيمة'],spC[3])]}),
      ...profileRows.map(([e,ev,a,av],i)=>new TableRow({children:[
        mC(e,spC[0],i%2===0?PALE:WHITE,{align:AlignmentType.LEFT}),
        mC(ev,spC[1],i%2===0?PALE:WHITE,{bold:true,color:DARK}),
        mC(a,spC[2],i%2===0?PALE:WHITE,{align:AlignmentType.RIGHT}),
        mC(av,spC[3],i%2===0?PALE:WHITE,{bold:true,color:DARK}),
      ]}))
    ]}),sp(200,80),
  );

  // ══════════════════════════════════════════════════════════════════════
  // MULTI-COURSE MODE: same structure as instructor report
  // ══════════════════════════════════════════════════════════════════════
  if(isMulti){
    const allQs=secs.reduce((acc,sec)=>acc.concat(sec.qs||[]),[]);
    const nQ=allQs.length;
    const nC=courses.length;

    // ── SECTION 1: ALL SECTIONS SUMMARY ──────────────────────────────
    // Build allSecs from courseResults + secs data
    // s1C: # | Course | Total n | F sections | F n | M sections | M n | Mean | Class
    const s1CW=[Math.round(CW*0.06),Math.round(CW*0.16),Math.round(CW*0.08),Math.round(CW*0.07),Math.round(CW*0.08),Math.round(CW*0.07),Math.round(CW*0.08),Math.round(CW*0.08),Math.max(300,CW-Math.round(CW*0.06)-Math.round(CW*0.16)-Math.round(CW*0.08)*3-Math.round(CW*0.07)*2-Math.round(CW*0.08)-Math.round(CW*0.08))];
    const s1C=s1CW;
    children.push(
      mP('أولاً: ملخص المقررات | Course Summary',{bold:true,size:22,color:DARK,before:0,after:80}),
      new Table({width:{size:CW,type:WidthType.DXA},columnWidths:s1C,rows:[
        new TableRow({children:[
          mH(['#'],s1C[0]),
          mH(['المقرر / Course'],s1C[1]),
          mH(['إجمالي مستجيبين'],s1C[2]),
          mH(['شعب إناث F'],s1C[3]),
          mH(['مستجيبين إناث F'],s1C[4]),
          mH(['شعب ذكور M'],s1C[5]),
          mH(['مستجيبين ذكور M'],s1C[6]),
          mH(['المتوسط / Mean'],s1C[7]),
          mH(['التصنيف / Class.'],s1C[8]),
        ]}),
        ...courses.map((cn,ci)=>{
          const cd=courseResults[cn]; const cl=clfR(cd.mean||0); const bg=ci%2===0?PALE:WHITE;
          return new TableRow({children:[
            mC(ci+1,s1C[0],bg,{bold:true,color:DARK,size:13}),
            mC(cn,s1C[1],bg,{bold:true,color:DARK,align:AlignmentType.LEFT,size:13}),
            mC(cd.n||0,s1C[2],bg,{bold:true,size:14}),
            mC(cd.sectionsG||cd.sectionCount||'—',s1C[3],'FCE4D6',{color:'843C0C',size:13}),
            mC(cd.nG||'—',s1C[4],'FCE4D6',{bold:true,color:'843C0C',size:14}),
            mC(cd.sectionsM||'—',s1C[5],'DDEBF7',{color:'1F4E79',size:13}),
            mC(cd.nB||'—',s1C[6],'DDEBF7',{bold:true,color:'1F4E79',size:14}),
            mC((cd.mean||0).toFixed(2),s1C[7],cl.bg,{bold:true,color:cl.c,size:14}),
            mC(cl.l,s1C[8],cl.bg,{bold:true,color:cl.c,size:12}),
          ]});
        }),
        new TableRow({children:[
          mC('الإجمالي',s1C[0]+s1C[1],DARK,{bold:true,color:WHITE,colSpan:2}),
          mC(n,s1C[2],GREEN2,{bold:true,color:GREEN,size:16}),
          mC(nF||'—',s1C[3],'FCE4D6',{bold:true,color:'843C0C'}),
          mC(nM||'—',s1C[4],'DDEBF7',{bold:true,color:DARK}),
          mC(overall,s1C[5],gCl.bg,{bold:true,color:gCl.c,size:18}),
          mC(gCl.l,s1C[6],gCl.bg,{bold:true,color:gCl.c}),
        ]}),
      ]}),sp(200,80),
    );

    // ── SECTION 2: COURSE Q SUMMARY (like instructor Q summary) ──────
    // Show all Q columns - no limit
    const lqN=nQ;
    const lqW=Math.max(380,Math.floor((CW-400-2800-600-600-1500)/Math.max(lqN,1)));
    const lqUsed=lqW*lqN;
    const s2LW=Math.max(400,CW-400-2800-600-600-lqUsed-1500);
    const s2C=[400,2800,600,600,...Array(lqN).fill(lqW),1500,s2LW];
    children.push(
      new Paragraph({pageBreakBefore:true,children:[]}),
      mP('ثانياً: ملخص تقييم المقررات | Course Evaluation Summary',{bold:true,size:22,color:DARK,before:0,after:80}),
      lqN<nQ?mP(`يعرض أول ${lqN} سؤال من ${nQ} — المتوسط يشمل جميع الأسئلة`,{size:15,color:'777777',italic:true,before:0,after:60}):sp(0,0),
      new Table({width:{size:CW,type:WidthType.DXA},columnWidths:s2C,rows:[
        new TableRow({children:[
          mH(['#'],s2C[0]),mH(['المقرر / Course'],s2C[1]),
          mH(['الشعب'],s2C[2]),mH(['المستجيبون'],s2C[3]),
          ...Array.from({length:lqN},(_,i)=>mH(['Q'+(i+1)],s2C[4+i],MID,13)),
          mH(['المتوسط\nالموزون'],s2C[4+lqN]),mH(['التصنيف'],s2C[5+lqN]),
        ]}),
        ...courses.flatMap((cn,ci)=>{
          const cd=courseResults[cn]; const cl=clfR(cd.mean||0); const bg=ci%2===0?PALE:WHITE;
          const rows=[];
          // Combined row
          rows.push(new TableRow({children:[
            mC(ci+1,s2C[0],bg,{bold:true,color:DARK,size:13}),
            mC(cn,s2C[1],bg,{bold:true,align:AlignmentType.LEFT,size:13}),
            mC(cd.sectionCount||1,s2C[2],bg),
            mC(cd.n||0,s2C[3],bg,{bold:true}),
            ...(cd.qMeans||[]).slice(0,lqN).map((qm,qi)=>{const qcl=clfR(qm||0);return mC((qm||0).toFixed(2),s2C[4+qi],qcl.bg,{color:qcl.c,size:11});}),
            mC((cd.mean||0).toFixed(2),s2C[4+lqN],cl.bg,{bold:true,color:cl.c,size:14}),
            mC(cl.l,s2C[5+lqN],cl.bg,{bold:true,color:cl.c,size:11}),
          ]}));
          // Female row
          if(cd.nG>0&&cd.qMeansF){
            const clf=clfR(cd.meanF||0);
            rows.push(new TableRow({children:[
              mC('',s2C[0],'FCE4D6'),mC('👧 Female',s2C[1],'FCE4D6',{color:'843C0C',size:11}),
              mC('—',s2C[2],'FCE4D6'),mC(cd.nG,s2C[3],'FCE4D6',{color:'843C0C',size:12}),
              ...(cd.qMeansF||[]).slice(0,lqN).map((qm,qi)=>{const qcl=clfR(qm||0);return mC((qm||0).toFixed(2),s2C[4+qi],'FCE4D6',{color:'843C0C',size:11});}),
              mC((cd.meanF||0).toFixed(2),s2C[4+lqN],'FCE4D6',{bold:true,color:'843C0C',size:13}),
              mC(clf.l,s2C[5+lqN],'FCE4D6',{bold:true,color:'843C0C',size:11}),
            ]}));
          }
          // Male row
          if(cd.nB>0&&cd.qMeansM){
            const clm=clfR(cd.meanM||0);
            rows.push(new TableRow({children:[
              mC('',s2C[0],'DDEBF7'),mC('👦 Male',s2C[1],'DDEBF7',{color:'1F4E79',size:11}),
              mC('—',s2C[2],'DDEBF7'),mC(cd.nB,s2C[3],'DDEBF7',{color:'1F4E79',size:12}),
              ...(cd.qMeansM||[]).slice(0,lqN).map((qm,qi)=>{const qcl=clfR(qm||0);return mC((qm||0).toFixed(2),s2C[4+qi],'DDEBF7',{color:'1F4E79',size:11});}),
              mC((cd.meanM||0).toFixed(2),s2C[4+lqN],'DDEBF7',{bold:true,color:'1F4E79',size:13}),
              mC(clm.l,s2C[5+lqN],'DDEBF7',{bold:true,color:'1F4E79',size:11}),
            ]}));
          }
          return rows;
        }),
      ]}),sp(200,80),
    );

    // ── SECTION 3: Q × COURSE (course name on top, M/F columns below) ────
    const hasMF=Object.values(courseResults).some(cd=>cd.nB>0&&cd.nG>0);

    if(hasMF){

      const nC=courses.length;
      const qLblW=Math.round(CW*0.18);
      const cW=Math.max(340,Math.floor((CW-qLblW-600)/(nC*2+1)));
      const ovW=Math.max(500,CW-qLblW-cW*nC*2);
      // cols: [Q_label, C1_M, C1_F, C2_M, C2_F, ..., Overall]
      const mfCols=[qLblW,...Array(nC*2).fill(cW),ovW];

      children.push(
        new Paragraph({pageBreakBefore:true,children:[]}),
        mP('ثالثاً: مقارنة الأسئلة — ذكور وإناث | Q × Course (M/F)',{bold:true,size:22,color:DARK,before:0,after:80}),
        new Table({width:{size:CW,type:WidthType.DXA},columnWidths:mfCols,rows:[

          // ── Header Row 1: السؤال + Course names (each spans 2 cols) + Overall
          new TableRow({children:[
            // Q label cell - spans 2 rows
            new TableCell({
              width:{size:Math.max(1,qLblW),type:WidthType.DXA},
              borders:allB(DARK),shading:{fill:DARK,type:ShadingType.CLEAR},
              margins:mg(),rowSpan:2,verticalAlign:VerticalAlign.CENTER,
              children:[new Paragraph({alignment:AlignmentType.CENTER,
                children:[new TextRun({text:'Criteria / السؤال',bold:true,color:WHITE,size:15,font:'Arial'})]})]
            }),
            // Course names - each spans M+F (2 cols)
            ...courses.map(cn=>new TableCell({
              width:{size:Math.max(1,cW*2),type:WidthType.DXA},
              columnSpan:2,
              borders:allB(MID),shading:{fill:MID,type:ShadingType.CLEAR},
              margins:mg(),verticalAlign:VerticalAlign.CENTER,
              children:[new Paragraph({alignment:AlignmentType.CENTER,
                children:[new TextRun({text:cn.slice(0,12),bold:true,color:WHITE,size:13,font:'Arial'})]})]
            })),
            // Overall - spans 2 rows
            new TableCell({
              width:{size:Math.max(1,ovW),type:WidthType.DXA},
              borders:allB(DARK),shading:{fill:DARK,type:ShadingType.CLEAR},
              margins:mg(),rowSpan:2,verticalAlign:VerticalAlign.CENTER,
              children:[new Paragraph({alignment:AlignmentType.CENTER,
                children:[new TextRun({text:'Overall',bold:true,color:WHITE,size:13,font:'Arial'})]})]
            }),
          ]}),

          // ── Header Row 2: M / F under each course
          new TableRow({children:[
            ...courses.flatMap(cn=>[
              new TableCell({width:{size:Math.max(1,cW),type:WidthType.DXA},borders:allB('1F4E79'),
                shading:{fill:'DDEBF7',type:ShadingType.CLEAR},margins:mg(),
                children:[new Paragraph({alignment:AlignmentType.CENTER,
                  children:[new TextRun({text:'M',bold:true,color:'1F4E79',size:14,font:'Arial'})]})]}),
              new TableCell({width:{size:Math.max(1,cW),type:WidthType.DXA},borders:allB('843C0C'),
                shading:{fill:'FCE4D6',type:ShadingType.CLEAR},margins:mg(),
                children:[new Paragraph({alignment:AlignmentType.CENTER,
                  children:[new TextRun({text:'F',bold:true,color:'843C0C',size:14,font:'Arial'})]})]}),
            ]),
          ]}),

          // ── Data Rows: one per question
          ...allQs.map((q,qi)=>{
            const bg=qi%2===0?PALE:WHITE;
            const oCl=clfR(q.cM||0);
            return new TableRow({children:[
              // Question label - bilingual
              new TableCell({width:{size:Math.max(1,mfCols[0]),type:WidthType.DXA},borders:allB(),
                shading:{fill:bg,type:ShadingType.CLEAR},margins:mg(),verticalAlign:VerticalAlign.CENTER,
                children:[
                  new Paragraph({alignment:AlignmentType.RIGHT,spacing:{before:0,after:20},children:[
                    new TextRun({text:'Q'+q.qn+': '+(q.lbl||''),bold:true,size:12,color:'1F4E79',font:'Arial',rtl:true})
                  ]}),
                  new Paragraph({alignment:AlignmentType.LEFT,spacing:{before:0,after:0},children:[
                    new TextRun({text:Q_EN[q.qn]||'',size:10,color:'555555',font:'Arial',italics:true})
                  ]}),
                ]
              }),
              // M+F per course
              ...courses.flatMap((cn,ci)=>{
                const cd=courseResults[cn];
                const qmM=cd.qMeansM?+(cd.qMeansM[qi]||0).toFixed(2):'—';
                const qmF=cd.qMeansF?+(cd.qMeansF[qi]||0).toFixed(2):'—';
                const mCl=clfR(parseFloat(qmM)||0);
                const fCl=clfR(parseFloat(qmF)||0);
                return [
                  mC(qmM,cW,'DDEBF7',{color:'1F4E79',size:13,bold:true}),
                  mC(qmF,cW,'FCE4D6',{color:'843C0C',size:13,bold:true}),
                ];
              }),
              // Overall
              mC((q.cM||0).toFixed?+(q.cM).toFixed(2):q.cM,mfCols[mfCols.length-1],oCl.bg,{bold:true,color:oCl.c,size:14}),
            ]});
          }),

          // ── Total Row
          new TableRow({children:[
            mC('المتوسط العام',mfCols[0],DARK,{bold:true,color:WHITE,size:14}),
            ...courses.flatMap((cn,ci)=>{
              const cd=courseResults[cn];
              const mV=+(cd.meanM||cd.mean||0).toFixed(2);
              const fV=+(cd.meanF||cd.mean||0).toFixed(2);
              const mCl=clfR(parseFloat(mV));
              const fCl=clfR(parseFloat(fV));
              return [
                mC(mV,cW,mCl.bg,{bold:true,color:mCl.c,size:14}),
                mC(fV,cW,fCl.bg,{bold:true,color:fCl.c,size:14}),
              ];
            }),
            mC(overall,mfCols[mfCols.length-1],gCl.bg,{bold:true,color:gCl.c,size:16}),
          ]}),

        ]}),sp(200,80),
      );

  
    } // end hasMF


    // ── COURSE DETAIL ─────────────────────────────────────────────────────
    children.push(
      new Paragraph({pageBreakBefore:true,children:[]}),
      mP('خامساً: التحليل التفصيلي لكل مقرر | Course Detail',{bold:true,size:22,color:DARK,before:0,after:80}),
    );
    // Batch fetch charts for all Qs
    const multiAllQs=secs.reduce((acc,s)=>acc.concat(s.qs||[]),[]);
    const multiSeenC=new Set();
    const multiUniqueQs=multiAllQs.filter(q=>{if(!q||multiSeenC.has(q.qn))return false;multiSeenC.add(q.qn);return true;});
    const multiSecLookup={};
    secs.forEach(s=>(s.qs||[]).forEach(q=>{multiSecLookup[q.qn]=s.ar||s.name||'';}));
    multiUniqueQs.forEach(q=>{q.secName=multiSecLookup[q.qn]||'';});
    const multiChartMap=await fetchChartsAll(multiUniqueQs,'Female','Male');

    for(const [ci,cn] of courses.entries()){
      const cd=courseResults[cn]; const cl=clfR(cd.mean||0);
      const allQsList=secs.reduce((acc,sec)=>acc.concat(sec.qs||[]),[]);
      const nQ2=allQsList.length;
      children.push(
        mP(`${ci+1}. ${cn}`,{bold:true,size:20,color:MID,before:ci===0?0:160,after:30}),
        mP(`المستجيبون: ${cd.n||0}  |  إناث: ${cd.nG||'—'}  |  ذكور: ${cd.nB||'—'}  |  المتوسط: ${(cd.mean||0).toFixed(2)}  |  ${cl.l}`,
          {size:15,color:'444444',before:0,after:50}),
      );
      const dqW=Math.max(500,Math.floor((CW-600-2600-600)/Math.max(nQ2,1)));
      const dC=[600,2600,600,...Array(nQ2).fill(dqW),Math.max(400,CW-600-2600-600-dqW*nQ2)];
      children.push(new Table({width:{size:CW,type:WidthType.DXA},columnWidths:dC,rows:[
        new TableRow({children:[
          mH(['Group'],dC[0]),mH(['المقرر / Course'],dC[1]),mH(['n'],dC[2]),
          ...Array.from({length:nQ2},(_,i)=>mH(['Q'+(i+1)],dC[3+i],MID,11)),
          mH(['Mean'],dC[3+nQ2]),
        ]}),
        // Combined
        new TableRow({children:[
          mC('All',dC[0],PALE,{bold:true,color:DARK,size:11}),
          mC(cn,dC[1],PALE,{bold:true,color:DARK,size:11}),
          mC(cd.n||0,dC[2],PALE,{size:11}),
          ...(cd.qMeans||[]).map((qm,qi)=>{const qcl=clfR(qm||0);return mC((qm||0).toFixed(2),dC[3+qi],qcl.bg,{color:qcl.c,size:11});}),
          mC((cd.mean||0).toFixed(2),dC[3+nQ2],cl.bg,{bold:true,color:cl.c,size:14}),
        ]}),
        // Female
        ...(cd.nG>0&&cd.qMeansF?[new TableRow({children:[
          mC('Female',dC[0],'FCE4D6',{bold:true,color:'843C0C',size:11}),
          mC(cn,dC[1],'FCE4D6',{color:'843C0C',size:11}),
          mC(cd.nG,dC[2],'FCE4D6',{color:'843C0C',size:11}),
          ...(cd.qMeansF||[]).map((qm,qi)=>{const qcl=clfR(qm||0);return mC((qm||0).toFixed(2),dC[3+qi],'FCE4D6',{color:'843C0C',size:11});}),
          mC((cd.meanF||cd.mean||0).toFixed(2),dC[3+nQ2],'FCE4D6',{bold:true,color:'843C0C',size:13}),
        ]})]:[]),
        // Male
        ...(cd.nB>0&&cd.qMeansM?[new TableRow({children:[
          mC('Male',dC[0],'DDEBF7',{bold:true,color:'1F4E79',size:11}),
          mC(cn,dC[1],'DDEBF7',{color:'1F4E79',size:11}),
          mC(cd.nB,dC[2],'DDEBF7',{color:'1F4E79',size:11}),
          ...(cd.qMeansM||[]).map((qm,qi)=>{const qcl=clfR(qm||0);return mC((qm||0).toFixed(2),dC[3+qi],'DDEBF7',{color:'1F4E79',size:11});}),
          mC((cd.meanM||cd.mean||0).toFixed(2),dC[3+nQ2],'DDEBF7',{bold:true,color:'1F4E79',size:13}),
        ]})]:[]),
      ]}),sp(60,40));

      // charts added at end
    }

    // ── ENHANCEMENT PLANS (isMulti) ────────────────────────────────────────
    const allQsEP=secs.reduce((acc,sec)=>acc.concat(sec.qs||[]),[]);
    const seenEP=new Set();
    const epItemsMulti=allQsEP.filter(q=>{
      if(!q||q.cM===undefined||seenEP.has(q.qn))return false;
      seenEP.add(q.qn);return true;
    }).sort((a,b)=>(b.cM||0)-(a.cM||0));
    const epCM=[700,500,2800,700,1000,600,600,Math.max(300,CW-700-500-2800-700-1000-600-600)];
    children.push(
      new Paragraph({pageBreakBefore:true,children:[]}),
      mP('Enhancement Plans | خطة التحسين والتطوير',{bold:true,size:28,color:DARK,before:0,after:60}),
      mP('مرتبة من الأضعف للأقوى — Sorted by Mean (Lowest = Needs Most Attention)',{size:16,color:'555555',italic:true,before:0,after:80}),
      new Table({width:{size:CW,type:WidthType.DXA},columnWidths:epCM,rows:[
        new TableRow({children:[
          mH(['Priority'],epCM[0]),mH(['Q#'],epCM[1]),mH(['Survey Item | السؤال'],epCM[2]),
          mH(['Mean'],epCM[3]),mH(['Classification'],epCM[4]),
          mH(['Pos%'],epCM[5],'375623'),mH(['Neg%'],epCM[6],'9C0006'),
          mH(['Recommended Action | KPI'],epCM[7]),
        ]}),
        ...epItemsMulti.map((q,i)=>{
          const cD=q.cD||[0,0,0,0,0];
          const posP=Math.round((cD[0]||0)+(cD[1]||0));
          const negP=Math.round((cD[3]||0)+(cD[4]||0));
          const cl=clfR(q.cM||0); const bg=i%2===0?PALE:WHITE;
          const cMv=parseFloat(q.cM)||0;
          const pr=cMv>=4.5?'🔴 Critical':cMv>=3.5?'🔴 High':cMv>=2.5?'🟡 Medium':cMv>=1.5?'🟢 Good':'🟢 Excellent';
          const prBg=cMv>=3.5?RED2:cMv>=2.5?AMBER2:GREEN2;
          const prC=cMv>=3.5?RED:cMv>=2.5?AMBER:GREEN;
          const action=cMv<2.5?'Immediate review & improvement plan required.':cMv<3.5?'Monitor and provide additional support.':'Maintain current practices.';
          const kpi=cMv>=3.5?`Raise to ≥${(cMv+0.5).toFixed(1)}; Pos% ≥${Math.min(95,posP+15)}%`:`Maintain ≥${cMv}; Pos% ≥${posP}%`;
          return new TableRow({children:[
            mC(pr,epCM[0],prBg,{bold:true,color:prC,size:11}),
            mC('Q'+q.qn,epCM[1],bg,{bold:true,color:DARK,size:13}),
            new TableCell({width:{size:Math.max(1,epCM[2]),type:WidthType.DXA},borders:allB(),
              shading:{fill:bg,type:ShadingType.CLEAR},margins:mg(),verticalAlign:VerticalAlign.CENTER,
              children:[
                new Paragraph({alignment:AlignmentType.RIGHT,spacing:{before:0,after:8},
                  children:[new TextRun({text:String(q.lbl||''),size:11,color:'1F4E79',font:'Arial',rtl:true})]}),
                new Paragraph({alignment:AlignmentType.LEFT,spacing:{before:0,after:0},
                  children:[new TextRun({text:Q_EN[q.qn]||'',size:9,color:'666666',font:'Arial',italics:true})]}),
              ]}),
            mC(cMv.toFixed(2),epCM[3],cl.bg,{bold:true,color:cl.c,size:13}),
            mC(cl.l,epCM[4],cl.bg,{bold:true,color:cl.c,size:11}),
            mC(posP+'%',epCM[5],posP>=80?GREEN2:posP>=60?AMBER2:WHITE,{bold:true,color:posP>=80?GREEN:posP>=60?AMBER:'333333',size:12}),
            mC(negP+'%',epCM[6],negP>20?RED2:WHITE,{bold:true,color:negP>20?RED:'333333',size:12}),
            new TableCell({width:{size:Math.max(1,epCM[7]),type:WidthType.DXA},borders:allB(),
              shading:{fill:bg,type:ShadingType.CLEAR},margins:mg(),verticalAlign:VerticalAlign.CENTER,
              children:[
                new Paragraph({alignment:AlignmentType.LEFT,spacing:{before:0,after:4},
                  children:[new TextRun({text:action,size:10,color:'333333',font:'Arial'})]}),
                new Paragraph({alignment:AlignmentType.LEFT,spacing:{before:0,after:0},
                  children:[new TextRun({text:'KPI: '+kpi,size:9,color:'666666',font:'Arial',italics:true})]}),
              ]}),
          ]});
        }),
      ]}),sp(200,80),
    );

  } else {
    // ══════════════════════════════════════════════════════════════════
    // SINGLE COURSE — matches reference report structure exactly
    // Scale: 1=Strongly Agree (best), 5=Strongly Disagree (worst)
    // Pos% = %SA(1)+%A(2), Neg% = %D(4)+%SD(5), Neutral(3) excluded
    // ══════════════════════════════════════════════════════════════════

    // ── 1. SECTION SUMMARY ─────────────────────────────────────────
    const ssC=[Math.round(CW*0.10),Math.round(CW*0.35),Math.round(CW*0.10),Math.round(CW*0.10),Math.round(CW*0.10),Math.round(CW*0.10),Math.max(300,CW-Math.round(CW*0.85))];
    children.push(
      mP('Section Summary Results | ملخص نتائج المحاور',{bold:true,size:22,color:DARK,before:0,after:80}),
      new Table({width:{size:CW,type:WidthType.DXA},columnWidths:ssC,rows:[
        new TableRow({children:[
          mH(['#'],ssC[0]),mH(['Section / المحور'],ssC[1]),
          mH(['Mean'],ssC[2]),mH(['F.Mean'],ssC[3]),mH(['M.Mean'],ssC[4]),
          mH(['Gap'],ssC[5]),mH(['Classification'],ssC[6]),
        ]}),
        ...secs.map((s,i)=>{
          const cl=clfR(s.mean); const bg=i%2===0?PALE:WHITE;
          return new TableRow({children:[
            mC(i+1,ssC[0],bg,{bold:true,color:DARK,size:13}),
            new TableCell({width:{size:Math.max(1,ssC[1]),type:WidthType.DXA},borders:allB(),
              shading:{fill:bg,type:ShadingType.CLEAR},margins:mg(),verticalAlign:VerticalAlign.CENTER,
              children:[
                new Paragraph({alignment:AlignmentType.LEFT,spacing:{before:0,after:4},
                  children:[new TextRun({text:s.name,bold:true,size:13,color:DARK,font:'Arial'})]}),
                new Paragraph({alignment:AlignmentType.RIGHT,spacing:{before:0,after:0},
                  children:[new TextRun({text:s.ar,size:12,color:DARK,font:'Arial',rtl:true})]}),
              ]}),
            mC(s.mean,ssC[2],cl.bg,{bold:true,color:cl.c,size:15}),
            mC(s.fMean||s.mean,ssC[3],'FCE4D6',{bold:true,color:'843C0C',size:13}),
            mC(s.mMean||s.mean,ssC[4],'DDEBF7',{bold:true,color:'1F4E79',size:13}),
            mC(Math.abs((s.fMean||s.mean)-(s.mMean||s.mean)).toFixed(2),ssC[5],bg,{size:12}),
            mC(cl.l,ssC[6],cl.bg,{bold:true,color:cl.c,size:11}),
          ]});
        }),
      ]}),sp(200,80),
    );

    // ── 2. OVERALL ANALYSIS ─────────────────────────────────────────
    // Matches Table 5 in reference: Q | Sec.Q | Section | Female | Male | Max | Min | SD | Mean | Pos% | Neg% | Classification
    const oaC=[600,400,1000,800,800,700,700,600,900,700,700,Math.max(300,CW-600-400-1000-800*2-700*2-600-900-700*2)];
    const oaRows=[];
    const seenOA=new Set();
    secs.forEach(s=>{
      const sortedSQs=(s.qs||[]).filter(q=>{
        if(!q||seenOA.has(q.qn)) return false;
        seenOA.add(q.qn); return true;
      }).sort((a,b)=>a.qn-b.qn);
      if(!sortedSQs.length) return;
      // Section header row
      oaRows.push(new TableRow({children:[
        new TableCell({columnSpan:12,width:{size:CW,type:WidthType.DXA},
          borders:allB(MID),shading:{fill:MID,type:ShadingType.CLEAR},margins:mg(),
          children:[new Paragraph({alignment:AlignmentType.LEFT,
            children:[new TextRun({text:s.name+' | '+s.ar,bold:true,color:WHITE,size:16,font:'Arial'})]})]}),
      ]}));
      sortedSQs.forEach((q,qi)=>{
        const cl=clfR(q.cM||0); const bg=qi%2===0?PALE:WHITE;
        const cD=q.cD||[0,0,0,0,0];
        const posP=Math.round((cD[0]||0)+(cD[1]||0)); // %SA + %A only
        const negP=Math.round((cD[3]||0)+(cD[4]||0)); // %D + %SD only
        oaRows.push(new TableRow({children:[
          mC('Q'+q.qn,oaC[0],bg,{bold:true,color:DARK,size:13}),
          mC(qi+1,oaC[1],bg,{size:11}),
          new TableCell({width:{size:Math.max(1,oaC[2]),type:WidthType.DXA},borders:allB(),
            shading:{fill:bg,type:ShadingType.CLEAR},margins:mg(),verticalAlign:VerticalAlign.CENTER,
            children:[
              new Paragraph({alignment:AlignmentType.RIGHT,spacing:{before:0,after:4},
                children:[new TextRun({text:String(q.lbl||''),size:11,color:'1F4E79',font:'Arial',rtl:true})]}),
            ]}),
          mC(q.fM??q.cM,oaC[3],'FCE4D6',{color:'843C0C',bold:true,size:12}),
          mC(q.mM??q.cM,oaC[4],'DDEBF7',{color:'1F4E79',bold:true,size:12}),
          mC(q.maxM??q.cM,oaC[5],bg,{size:11}),
          mC(q.minM??q.cM,oaC[6],bg,{size:11}),
          mC('—',oaC[7],bg,{size:10}),
          mC(q.cM,oaC[8],cl.bg,{bold:true,color:cl.c,size:14}),
          mC(posP+'%',oaC[9],posP>=80?GREEN2:posP>=60?AMBER2:WHITE,{bold:true,color:posP>=80?GREEN:posP>=60?AMBER:'333333',size:12}),
          mC(negP+'%',oaC[10],negP>20?RED2:WHITE,{bold:true,color:negP>20?RED:'333333',size:12}),
          mC(cl.l,oaC[11],cl.bg,{bold:true,color:cl.c,size:10}),
        ]}));
      });
    });
    children.push(
      new Paragraph({pageBreakBefore:true,children:[]}),
      mP('Overall Analysis | التحليل الإجمالي',{bold:true,size:22,color:DARK,before:0,after:20}),
      mP('Pos% = %(1)SA + %(2)A  |  Neg% = %(4)D + %(5)SD  |  %(3)Neutral excluded',
        {size:13,color:'2E75B6',italic:true,before:0,after:60}),
      new Table({width:{size:CW,type:WidthType.DXA},columnWidths:oaC,rows:[
        new TableRow({children:[
          mH(['Q'],oaC[0]),mH(['Sec.\nQ'],oaC[1]),mH(['Section / السؤال'],oaC[2]),
          mH(['Female Mean'],oaC[3],'843C0C'),mH(['Male Mean'],oaC[4],'1F4E79'),
          mH(['Max\nMean'],oaC[5]),mH(['Min\nMean'],oaC[6]),mH(['SD'],oaC[7]),
          mH(['Mean\nAll'],oaC[8]),
          mH(['Positive\n%'],oaC[9],'375623'),mH(['Negative\n%'],oaC[10],'9C0006'),
          mH(['Classification'],oaC[11]),
        ]}),
        ...oaRows,
      ]}),sp(200,80),
    );

    // ── 3. PER-SECTION DETAIL ──────────────────────────────────────
    // Matches Table 6+ in reference: Section title | Q text row | Female/Male/Combined rows
    // ── Batch fetch ALL charts at once ──────────────────────────
    const allQsForCharts=secs.reduce((acc,s)=>acc.concat(s.qs||[]),[]);
    const seenChart=new Set();
    const uniqueQsForCharts=allQsForCharts.filter(q=>{
      if(!q||seenChart.has(q.qn))return false;
      seenChart.add(q.qn);return true;
    });
    // Add secName to each Q
    const secLookup={};
    secs.forEach(s=>(s.qs||[]).forEach(q=>{secLookup[q.qn]=s.ar||s.name||'';}));
    uniqueQsForCharts.forEach(q=>{q.secName=secLookup[q.qn]||'';});
    const chartMap=await fetchChartsAll(uniqueQsForCharts,'Female','Male');

    for(const [si,sec] of secs.entries()){
      const cl=clfR(sec.mean);
      const seenQns=new Set();
      const sortedQs=(sec.qs||[]).filter(q=>{
        if(!q||seenQns.has(q.qn)) return false;
        seenQns.add(q.qn); return true;
      }).sort((a,b)=>a.qn-b.qn);
      if(!sortedQs.length) return;

      if(si>0) children.push(new Paragraph({pageBreakBefore:true,children:[]}));

      // Section title
      children.push(
        mP(sec.name+' | '+sec.ar,{bold:true,size:20,color:MID,before:0,after:6}),
        mP('Mean: '+sec.mean+' | F: '+(sec.fMean||sec.mean)+' | M: '+(sec.mMean||sec.mean)+' | '+cl.l,
          {size:15,color:'444444',before:0,after:8}),
      );

      // Mini summary row
      const sSumC=[Math.round(CW*0.42),Math.round(CW*0.12),Math.round(CW*0.12),Math.round(CW*0.12),Math.max(300,CW-Math.round(CW*0.78))];
      children.push(
        new Table({width:{size:CW,type:WidthType.DXA},columnWidths:sSumC,rows:[
          new TableRow({children:[
            mH(['Section / المحور'],sSumC[0]),
            mH(['Mean'],sSumC[1]),mH(['F.Mean'],sSumC[2],'843C0C'),
            mH(['M.Mean'],sSumC[3],'1F4E79'),mH(['Classification'],sSumC[4]),
          ]}),
          new TableRow({children:[
            new TableCell({width:{size:Math.max(1,sSumC[0]),type:WidthType.DXA},borders:allB(MID),
              shading:{fill:PALE,type:ShadingType.CLEAR},margins:mg(),verticalAlign:VerticalAlign.CENTER,
              children:[
                new Paragraph({alignment:AlignmentType.LEFT,spacing:{before:0,after:3},
                  children:[new TextRun({text:sec.name,bold:true,size:13,color:DARK,font:'Arial'})]}),
                new Paragraph({alignment:AlignmentType.RIGHT,spacing:{before:0,after:0},
                  children:[new TextRun({text:sec.ar,size:12,color:DARK,font:'Arial',rtl:true})]}),
              ]}),
            mC(sec.mean,sSumC[1],cl.bg,{bold:true,color:cl.c,size:15}),
            mC(sec.fMean||sec.mean,sSumC[2],'FCE4D6',{bold:true,color:'843C0C',size:13}),
            mC(sec.mMean||sec.mean,sSumC[3],'DDEBF7',{bold:true,color:'1F4E79',size:13}),
            mC(cl.l,sSumC[4],cl.bg,{bold:true,color:cl.c,size:12}),
          ]}),
        ]}),sp(40,30),
      );

      // Distribution table - matches reference Table 6+ exactly
      // Cols: Global Q | Sec.Q | Group | %SD(5) | %D(4) | %N(3) | %A(2) | %SA(1) | Mean
      const dC=[1600,400,700,700,700,700,700,700,Math.max(400,CW-1600-400-700*6)];
      const secRows=[];

      // Section header row (merged)
      secRows.push(new TableRow({children:[
        new TableCell({columnSpan:9,width:{size:CW,type:WidthType.DXA},
          borders:allB(DARK),shading:{fill:DARK,type:ShadingType.CLEAR},margins:mg(),
          children:[new Paragraph({alignment:AlignmentType.LEFT,
            children:[new TextRun({text:sec.name+' | '+sec.ar,bold:true,color:WHITE,size:15,font:'Arial'})]})]}),
      ]}));

      // Column headers row
      secRows.push(new TableRow({children:[
        mH(['Global\nQ'],dC[0]),mH(['Section\nQ'],dC[1]),mH(['Group'],dC[2]),
        mH(['%\nStrongly\nDisagree\n(5)'],dC[3],'9C0006',8),
        mH(['%\nDisagree\n(4)'],dC[4],'C55A11',8),
        mH(['%\nNeutral\n(3)'],dC[5],'7F7F7F',8),
        mH(['%\nAgree\n(2)'],dC[6],'375623',8),
        mH(['%\nStrongly\nAgree\n(1)'],dC[7],'1a5216',8),
        mH(['Mean'],dC[8]),
      ]}));

      sortedQs.forEach((q,qi)=>{
        const cD=q.cD||[0,0,0,0,0];
        const fD=q.fD||cD; const mD=q.mD||cD;
        const qcl=clfR(q.cM||0);

        // Q text row (merged across all 9 cols)
        secRows.push(new TableRow({children:[
          new TableCell({columnSpan:9,width:{size:CW,type:WidthType.DXA},
            borders:allB(MID),shading:{fill:PALE,type:ShadingType.CLEAR},margins:mg(),
            children:[
              new Paragraph({alignment:AlignmentType.RIGHT,spacing:{before:0,after:4},
                children:[new TextRun({text:'Q'+q.qn+'. '+String(q.lbl||''),bold:true,size:11,color:'1F4E79',font:'Arial',rtl:true})]}),
              new Paragraph({alignment:AlignmentType.LEFT,spacing:{before:0,after:0},
                children:[new TextRun({text:Q_EN[q.qn]||'',size:9,color:'666666',font:'Arial',italics:true})]}),
            ]}),
        ]}));

        // Female row
        secRows.push(new TableRow({children:[
          mC('Q'+q.qn,dC[0],'FCE4D6',{bold:true,color:'843C0C',size:11}),
          mC(qi+1,dC[1],'FCE4D6',{size:11}),
          mC('Female',dC[2],'FCE4D6',{bold:true,color:'843C0C',size:10}),
          mC((fD[4]||0),dC[3],'FFEBEE',{color:'9C0006',size:11}),
          mC((fD[3]||0),dC[4],'FFF3E0',{color:'C55A11',size:11}),
          mC((fD[2]||0),dC[5],'FCE4D6',{size:11}),
          mC((fD[1]||0),dC[6],'E8F5E9',{color:'375623',size:11}),
          mC((fD[0]||0),dC[7],'E2EFDA',{color:'1a5216',size:11}),
          mC(q.fM??q.cM,dC[8],'FCE4D6',{bold:true,color:'843C0C',size:13}),
        ]}));

        // Male row
        secRows.push(new TableRow({children:[
          mC('Q'+q.qn,dC[0],'DDEBF7',{bold:true,color:'1F4E79',size:11}),
          mC(qi+1,dC[1],'DDEBF7',{size:11}),
          mC('Male',dC[2],'DDEBF7',{bold:true,color:'1F4E79',size:10}),
          mC((mD[4]||0),dC[3],'FFEBEE',{color:'9C0006',size:11}),
          mC((mD[3]||0),dC[4],'FFF3E0',{color:'C55A11',size:11}),
          mC((mD[2]||0),dC[5],'DDEBF7',{size:11}),
          mC((mD[1]||0),dC[6],'E8F5E9',{color:'375623',size:11}),
          mC((mD[0]||0),dC[7],'E2EFDA',{color:'1a5216',size:11}),
          mC(q.mM??q.cM,dC[8],'DDEBF7',{bold:true,color:'1F4E79',size:13}),
        ]}));

        // Combined row
        secRows.push(new TableRow({children:[
          mC('Q'+q.qn,dC[0],'EBF3FB',{bold:true,color:DARK,size:11}),
          mC(qi+1,dC[1],'EBF3FB',{size:11}),
          mC('Combined',dC[2],'EBF3FB',{bold:true,color:DARK,size:10}),
          mC((cD[4]||0),dC[3],'FFEBEE',{color:'9C0006',size:11}),
          mC((cD[3]||0),dC[4],'FFF3E0',{color:'C55A11',size:11}),
          mC((cD[2]||0),dC[5],'EBF3FB',{size:11}),
          mC((cD[1]||0),dC[6],'E8F5E9',{color:'375623',size:11}),
          mC((cD[0]||0),dC[7],'E2EFDA',{color:'1a5216',size:11}),
          mC(q.cM,dC[8],qcl.bg,{bold:true,color:qcl.c,size:13}),
        ]}));
      });

      children.push(
        new Table({width:{size:CW,type:WidthType.DXA},columnWidths:dC,rows:secRows}),
        sp(80,40),
      );

      // charts added at end of report
    }

  // ── ENHANCEMENT PLANS ──────────────────────────────────────────────
  children.push(
    new Paragraph({pageBreakBefore:true,children:[]}),
    mP('Enhancement Plans | خطة التحسين والتطوير',{bold:true,size:28,color:DARK,before:0,after:60}),
    mP('مرتبة من الأضعف للأقوى — Sorted by Mean (Highest = Needs Most Attention)',
      {size:16,color:'555555',italic:true,before:0,after:20}),
    mP('Positive % = %(1) Strongly Agree + %(2) Agree  |  Negative % = %(4) Disagree + %(5) Strongly Disagree  |  Neutral %(3) is excluded from both',
      {size:14,color:'2E75B6',italic:true,before:0,after:60}),
  );

  // Build EP items from all Qs (deduplicated, sorted by mean desc = weakest first)
  const allEPQs=secs.reduce((acc,sec)=>acc.concat(sec.qs||[]),[]);
  const seenEP=new Set();
  const epItems=allEPQs.filter(q=>{
    if(!q||q.cM===undefined||seenEP.has(q.qn)) return false;
    seenEP.add(q.qn); return true;
  }).sort((a,b)=>(b.q&&b.q.cM?b.q.cM:b.cM||0)-(a.q&&a.q.cM?a.q.cM:a.cM||0));

  const epC=[700,500,2800,700,1000,600,600,Math.max(300,CW-700-500-2800-700-1000-600-600)];
  children.push(new Table({width:{size:CW,type:WidthType.DXA},columnWidths:epC,rows:[
    new TableRow({children:[
      mH(['Priority\nالأولوية'],epC[0]),
      mH(['Q#'],epC[1]),
      mH(['Survey Item | السؤال'],epC[2]),
      mH(['Mean\nالمتوسط'],epC[3]),
      mH(['Classification\nالتصنيف'],epC[4]),
      mH(['Pos%\nموافق'],epC[5],'375623'),
      mH(['Neg%\nلا أوافق'],epC[6],'9C0006'),
      mH(['Recommended Action & KPI\nالإجراء المقترح'],epC[7]),
    ]}),
    ...epItems.map((q,i)=>{
      const cD=q.cD||[0,0,0,0,0];
      const posP=Math.round((cD[0]||0)+(cD[1]||0));
      const negP=Math.round((cD[3]||0)+(cD[4]||0));
      const cl=clfR(q.cM||0);
      const bg=i%2===0?PALE:WHITE;
      const cMv=parseFloat(q.cM)||0;
      // Priority using 1=best scale thresholds
      const pr=cMv>3.00?'Critical / حرج':cMv>2.50?'High / مرتفع':cMv>2.00?'Medium / متوسط':cMv>1.50?'Low / منخفض':'Strength / نقطة قوة';
      const prBg=cMv>3.00?'FF0000':cMv>2.50?RED2:cMv>2.00?AMBER2:GREEN2;
      const prC=cMv>3.00?'FFFFFF':cMv>2.50?RED:cMv>2.00?AMBER:GREEN;
      const action=cMv>3.00?'Critical weakness — immediate intervention required. Assign responsible team and set urgent timeline.':
        cMv>2.50?'Weakness — develop targeted improvement plan with clear milestones and accountability.':
        cMv>2.00?'Needs monitoring — provide additional support, resources, and follow-up assessment.':
        cMv>1.50?'Good — maintain current practices and document for benchmarking.':
        'Excellent — document as best practice and share with other departments.';
      const kpi=cMv>2.50?`Reduce mean to ≤${(cMv-0.3).toFixed(2)}; Positive% ≥${Math.min(95,posP+15)}%`:
        `Maintain mean ≤${(cMv+0.2).toFixed(2)}; Positive% ≥${posP}%`;
      return new TableRow({children:[
        mC(pr,epC[0],prBg,{bold:true,color:prC,size:10}),
        mC('Q'+q.qn,epC[1],bg,{bold:true,color:DARK,size:13}),
        // Bilingual Q text
        new TableCell({width:{size:Math.max(1,epC[2]),type:WidthType.DXA},borders:allB(),
          shading:{fill:bg,type:ShadingType.CLEAR},margins:mg(),verticalAlign:VerticalAlign.CENTER,
          children:[
            new Paragraph({alignment:AlignmentType.RIGHT,spacing:{before:0,after:8},
              children:[new TextRun({text:String(q.lbl||''),size:11,color:'1F4E79',font:'Arial',rtl:true})]}),
            new Paragraph({alignment:AlignmentType.LEFT,spacing:{before:0,after:0},
              children:[new TextRun({text:Q_EN[q.qn]||'',size:9,color:'666666',font:'Arial',italics:true})]}),
          ]}),
        mC(cMv.toFixed(2),epC[3],cl.bg,{bold:true,color:cl.c,size:14}),
        mC(cl.l,epC[4],cl.bg,{bold:true,color:cl.c,size:11}),
        mC(posP+'%',epC[5],posP>=80?GREEN2:posP>=60?AMBER2:WHITE,
          {bold:true,color:posP>=80?GREEN:posP>=60?AMBER:'333333',size:12}),
        mC(negP+'%',epC[6],negP>20?RED2:negP>10?AMBER2:WHITE,
          {bold:true,color:negP>20?RED:negP>10?AMBER:'333333',size:12}),
        new TableCell({width:{size:Math.max(1,epC[7]),type:WidthType.DXA},borders:allB(),
          shading:{fill:bg,type:ShadingType.CLEAR},margins:mg(),verticalAlign:VerticalAlign.CENTER,
          children:[
            new Paragraph({alignment:AlignmentType.LEFT,spacing:{before:0,after:4},
              children:[new TextRun({text:action,size:10,color:'333333',font:'Arial'})]}),
            new Paragraph({alignment:AlignmentType.LEFT,spacing:{before:0,after:0},
              children:[new TextRun({text:'KPI: '+kpi,size:9,color:'666666',font:'Arial',italics:true})]}),
          ]}),
      ]});
    }),
  ]}),sp(200,80));

  } // end else single course

  return buildDoc(children);
}

async function buildInstructorWordFromResult(result, cfg){
  const CW=15398;
  const {allSecs,lecturers,qTexts,totalN,totalEnrolled,totalNot,grandMean}=result;
  // Instructor uses 5=best scale always
  const clfR=clf5;
  const nQ=qTexts.length;
  const gCl=clfR(grandMean);
  const pct=totalEnrolled>0?Math.round(totalN/totalEnrolled*100):0;
  const children=[];
  const courses=[...new Set(allSecs.map(s=>s.course))];

  // ── TITLE ──────────────────────────────────────────────────────────────
  children.push(
    sp(0,200),
    mP('تقرير استبانة تقييم المحاضرين',{align:AlignmentType.CENTER,bold:true,size:52,color:DARK,before:0,after:80}),
    mP('Instructor Evaluation Report',{align:AlignmentType.CENTER,bold:true,size:30,color:MID,before:0,after:80}),
    mP(cfg.cname||'كليات الغد للعلوم الطبية التطبيقية – جدة',{align:AlignmentType.CENTER,size:22,color:'555555',before:0,after:60}),
    mP(cfg.semester||'',{align:AlignmentType.CENTER,size:20,color:'777777',before:0,after:300}),
  );

  // ── GOAL / INFO ────────────────────────────────────────────────────────
  if(cfg.obj){
    children.push(
      mP('هدف الاستبيان',{bold:true,size:24,color:DARK,before:200,after:60}),
      mP(cfg.obj,{size:20,before:0,after:200}),
    );
  }

  // ── SCALE ──────────────────────────────────────────────────────────────
  const sCols=[Math.round(CW*0.22),Math.round(CW*0.15),Math.round(CW*0.15),Math.round(CW*0.15),Math.round(CW*0.15)];
  children.push(
    mP('المقياس المستخدم | Likert Scale (5=Strongly Agree)',{bold:true,size:22,color:DARK,before:0,after:80}),
    new Table({width:{size:CW,type:WidthType.DXA},columnWidths:sCols,rows:[
      new TableRow({children:[
        mH(['5 = موافق بشدة\nStrongly Agree'],sCols[0],GREEN),
        mH(['4 = موافق\nAgree'],sCols[1],GREEN),
        mH(['3 = محايد\nNeutral'],sCols[2],'7F7F7F'),
        mH(['2 = لا أوافق\nDisagree'],sCols[3],RED),
        mH(['1 = لا أوافق بشدة\nStrongly Disagree'],sCols[4],RED),
      ]}),
    ]}),sp(120,80),
  );

  // ── CLASSIFICATION SCALE ───────────────────────────────────────────────
  const clsCols=[Math.round(CW*0.10),Math.round(CW*0.16),Math.round(CW*0.14),CW-Math.round(CW*0.40)];
  children.push(
    mP('Classification Scale | مقياس التصنيف',{bold:true,size:22,color:DARK,before:0,after:80}),
    new Table({width:{size:CW,type:WidthType.DXA},columnWidths:clsCols,rows:[
      new TableRow({children:[mH(['Range'],clsCols[0]),mH(['Classification'],clsCols[1]),mH(['التصنيف'],clsCols[2]),mH(['Interpretation'],clsCols[3])]}),
      ...[
        ['≥4.50','Excellent','ممتاز','Strong positive outcome — maintain and share best practices',GREEN2,GREEN],
        ['3.50–4.49','Very Good','جيد جداً','Good performance — monitor and continue improvement',GREEN2,GREEN],
        ['2.50–3.49','Good','جيد','Acceptable — review and support further development',AMBER2,AMBER],
        ['1.50–2.49','Acceptable','مقبول','Below expectations — improvement plan required',RED2,RED],
        ['<1.50','Weak','ضعيف','Significant weakness — immediate intervention required',RED2,RED],
      ].map(([r,cl,ar,interp,bg,c])=>new TableRow({children:[
        mC(r,clsCols[0],bg,{bold:true,color:c,align:AlignmentType.CENTER}),
        mC(cl,clsCols[1],bg,{bold:true,color:c}),
        mC(ar,clsCols[2],bg,{bold:true,color:c}),
        mC(interp,clsCols[3],WHITE,{size:17,align:AlignmentType.LEFT}),
      ]}))
    ]}),sp(160,80),
  );

  // ── SAMPLE PROFILE ─────────────────────────────────────────────────────
  const spCols=[Math.round(CW*0.22),Math.round(CW*0.12),Math.round(CW*0.22),Math.round(CW*0.12)];
  children.push(
    mP('Sample Profile | بيانات العينة',{bold:true,size:22,color:DARK,before:0,after:80}),
    new Table({width:{size:CW,type:WidthType.DXA},columnWidths:spCols,rows:[
      new TableRow({children:[mH(['Detail'],spCols[0]),mH(['Value'],spCols[1]),mH(['التفاصيل'],spCols[2]),mH(['القيمة'],spCols[3])]}),
      ...[
        ['Total Enrolled',totalEnrolled,'إجمالي المسجلين',totalEnrolled],
        ['Total Evaluators',totalN,'إجمالي المقيّمين',totalN],
        ['Did Not Respond',totalNot,'لم يستبينوا',totalNot],
        ['Participation Rate',pct+'%','نسبة المشاركة',pct+'%'],
        ['No. of Sections',allSecs.length,'عدد الشعب',allSecs.length],
        ['No. of Instructors',lecturers.length,'عدد المحاضرين',lecturers.length],
        ['No. of Courses',courses.length,'عدد المقررات',courses.length],
        ['No. of Questions',nQ,'عدد الأسئلة',nQ],
        ['Overall Mean',grandMean,'المتوسط العام',grandMean],
        ['Survey Period',cfg.semester||'—','الفصل الدراسي',cfg.semester||'—'],
      ].map(([e,ev,a,av],i)=>new TableRow({children:[
        mC(e,spCols[0],i%2===0?PALE:WHITE,{align:AlignmentType.LEFT}),
        mC(ev,spCols[1],i%2===0?PALE:WHITE,{bold:true,color:DARK}),
        mC(a,spCols[2],i%2===0?PALE:WHITE,{align:AlignmentType.RIGHT}),
        mC(av,spCols[3],i%2===0?PALE:WHITE,{bold:true,color:DARK}),
      ]}))
    ]}),sp(200,80),
  );

  // ── SECTION 1: ALL SECTIONS ────────────────────────────────────────────
  const s1C=[600,Math.round(CW*0.17),Math.round(CW*0.16),900,900,1100,1300,Math.max(400,CW-600-Math.round(CW*0.17)-Math.round(CW*0.16)-900-900-1100-1300)];
  children.push(
    mP('أولاً: ملخص جميع الشعب | All Sections Summary',{bold:true,size:22,color:DARK,before:0,after:80}),
    new Table({width:{size:CW,type:WidthType.DXA},columnWidths:s1C,rows:[
      new TableRow({children:[
        mH(['الشعبة'],s1C[0]),mH(['المحاضر'],s1C[1]),mH(['المقرر'],s1C[2]),
        mH(['المسجلون'],s1C[3]),mH(['المقيّمون'],s1C[4]),
        mH(['المشاركة%'],s1C[5]),mH(['المتوسط'],s1C[6]),mH(['التصنيف'],s1C[7]),
      ]}),
      ...allSecs.map((s,i)=>{
        const cl=clfR(s.sec_mean); const bg=i%2===0?PALE:WHITE;
        const pp=s.participation_pct||0;
        return new TableRow({children:[
          mC(s.sec_num,s1C[0],bg,{size:14}),
          mC(s.lecturer,s1C[1],bg,{align:AlignmentType.RIGHT,size:13}),
          mC(s.course,s1C[2],bg,{bold:true,color:DARK,size:14}),
          mC(s.enrolled||s.enrolled_num||'—',s1C[3],bg),
          mC(s.n,s1C[4],bg,{bold:true}),
          mC(pp+'%',s1C[5],pp>=80?GREEN2:pp>=60?AMBER2:RED2,{color:pp>=80?GREEN:pp>=60?AMBER:RED,bold:true}),
          mC(s.sec_mean.toFixed(2),s1C[6],cl.bg,{bold:true,color:cl.c}),
          mC(cl.l,s1C[7],cl.bg,{bold:true,color:cl.c,size:14}),
        ]});
      }),
      new TableRow({children:[
        mC('الإجمالي',s1C[0]+s1C[1]+s1C[2],DARK,{bold:true,color:WHITE,colSpan:3}),
        mC(totalEnrolled,s1C[3],PALE,{bold:true}),
        mC(totalN,s1C[4],GREEN2,{bold:true,color:GREEN,size:16}),
        mC(pct+'%',s1C[5],pct>=80?GREEN2:pct>=60?AMBER2:RED2,{bold:true,color:pct>=80?GREEN:pct>=60?AMBER:RED,size:15}),
        mC(grandMean,s1C[6],gCl.bg,{bold:true,color:gCl.c,size:18}),
        mC(gCl.l,s1C[7],gCl.bg,{bold:true,color:gCl.c}),
      ]}),
    ]}),sp(200,80),
  );

  // ── SECTION 2: Q × COURSE OVERALL ANALYSIS ────────────────────────────
  const nC=courses.length;
  const qLblW=Math.round(CW*0.22);
  const qCW=Math.max(600,Math.floor((CW-qLblW-1200)/(nC+1)));
  const qLastW=Math.max(400,CW-qLblW-qCW*nC-1200);
  const oaC=[qLblW,...Array(nC).fill(qCW),1200,qLastW];

  // Course means per Q
  const courseQData={};
  courses.forEach(code=>{
    const cs=allSecs.filter(s=>s.course===code);
    const ns=cs.length||1;
    courseQData[code]={
      mean:+(cs.reduce((a,s)=>a+s.sec_mean,0)/ns).toFixed(2),
      qMeans:Array.from({length:nQ},(_,qi)=>+(cs.reduce((a,s)=>a+(s.questions[qi]?.mean||0),0)/ns).toFixed(2)),
    };
  });
  const overallQM=Array.from({length:nQ},(_,qi)=>+(allSecs.reduce((a,s)=>a+(s.questions[qi]?.mean||0),0)/(allSecs.length||1)).toFixed(2));

  children.push(
    new Paragraph({pageBreakBefore:true,children:[]}),
    mP('ثانياً: التحليل الإجمالي — كل الأسئلة عبر المقررات | Overall Analysis',{bold:true,size:22,color:DARK,before:0,after:80}),
    new Table({width:{size:CW,type:WidthType.DXA},columnWidths:oaC,rows:[
      new TableRow({children:[
        mH(['السؤال / Question'],oaC[0]),
        ...courses.map((code,ci)=>{
          return new TableCell({width:{size:Math.max(1,qCW),type:WidthType.DXA},borders:allB(MID),
            shading:{fill:MID,type:ShadingType.CLEAR},margins:mg(),verticalAlign:VerticalAlign.CENTER,
            children:[
              new Paragraph({alignment:AlignmentType.CENTER,spacing:{before:0,after:0},children:[new TextRun({text:code,bold:true,color:WHITE,size:15,font:'Arial'})]}),
              new Paragraph({alignment:AlignmentType.CENTER,spacing:{before:0,after:0},children:[new TextRun({text:'('+courseQData[code].mean.toFixed(2)+')',color:'AAAAAA',size:12,font:'Arial'})]}),
            ]
          });
        }),
        mH(['الإجمالي'],oaC[nC+1]),
        mH(['التصنيف'],oaC[nC+2]),
      ]}),
      ...Array.from({length:nQ},(_,qi)=>{
        const oM=overallQM[qi]; const oCl=clfR(oM); const bg=qi%2===0?PALE:WHITE;
        return new TableRow({children:[
          mC('Q'+(qi+1)+' — '+(qTexts[qi]||'').slice(0,45),oaC[0],bg,{align:AlignmentType.RIGHT,size:13}),
          ...courses.map((code,ci)=>{
            const qm=courseQData[code].qMeans[qi]; const qcl=clfR(qm);
            return mC(qm.toFixed(2),oaC[ci+1],qcl.bg,{color:qcl.c,size:14});
          }),
          mC(oM.toFixed(2),oaC[nC+1],oCl.bg,{bold:true,color:oCl.c,size:15}),
          mC(oCl.l,oaC[nC+2],oCl.bg,{bold:true,color:oCl.c,size:13}),
        ]});
      }),
      new TableRow({children:[
        mC('المتوسط العام',oaC[0],DARK,{bold:true,color:WHITE}),
        ...courses.map((code,ci)=>{const cl=clfR(courseQData[code].mean);return mC(courseQData[code].mean.toFixed(2),oaC[ci+1],cl.bg,{bold:true,color:cl.c,size:16});}),
        mC(grandMean,oaC[nC+1],gCl.bg,{bold:true,color:gCl.c,size:18}),
        mC(gCl.l,oaC[nC+2],gCl.bg,{bold:true,color:gCl.c}),
      ]}),
    ]}),sp(200,80),
  );

  // ── SECTION 3: INSTRUCTOR SUMMARY ─────────────────────────────────────
  const lqShowN=Math.min(nQ,12);
  const lqW=Math.max(500,Math.floor((CW-400-2800-600-600-1500)/Math.max(lqShowN,1)));
  const lqUsedW=lqW*lqShowN;
  const s2LastW=Math.max(400,CW-400-2800-600-600-lqUsedW-1500);
  const s2C=[400,2800,600,600,...Array(lqShowN).fill(lqW),1500,s2LastW];

  children.push(
    new Paragraph({pageBreakBefore:true,children:[]}),
    mP('ثالثاً: ملخص تقييم المحاضرين | Instructor Summary',{bold:true,size:22,color:DARK,before:0,after:80}),
    lqShowN<nQ?mP(`* يعرض أول ${lqShowN} سؤال من ${nQ} — الإجمالي يشمل جميع الأسئلة`,{size:15,color:'777777',italic:true,before:0,after:60}):sp(0,0),
    new Table({width:{size:CW,type:WidthType.DXA},columnWidths:s2C,rows:[
      new TableRow({children:[
        mH(['#'],s2C[0]),mH(['المحاضر'],s2C[1]),mH(['الشعب'],s2C[2]),mH(['المقيّمون'],s2C[3]),
        ...Array.from({length:lqShowN},(_,i)=>mH(['Q'+(i+1)],s2C[4+i],MID,13)),
        mH(['المتوسط'],s2C[4+lqShowN]),mH(['التصنيف'],s2C[5+lqShowN]),
      ]}),
      ...lecturers.map((lec,i)=>{
        const cl=clfR(lec.mean||0); const bg=i%2===0?PALE:WHITE;
        return new TableRow({children:[
          mC(i+1,s2C[0],bg,{bold:true,size:13}),
          mC(lec.name,s2C[1],bg,{align:AlignmentType.RIGHT,size:13}),
          mC((lec.secs||[]).length,s2C[2],bg),
          mC(lec.n,s2C[3],bg,{bold:true}),
          ...(lec.qMeans||[]).slice(0,lqShowN).map((qm,qi)=>{const qcl=clfR(qm);return mC(qm.toFixed(2),s2C[4+qi],qcl.bg,{color:qcl.c,size:12});}),
          mC((lec.mean||0).toFixed(2),s2C[4+lqShowN],cl.bg,{bold:true,color:cl.c,size:16}),
          mC(cl.l,s2C[5+lqShowN],cl.bg,{bold:true,color:cl.c,size:12}),
        ]});
      }),
    ]}),sp(200,80),
  );

  // ── SECTION 4: DETAILED PER LECTURER ──────────────────────────────────
  children.push(
    new Paragraph({pageBreakBefore:true,children:[]}),
    mP('رابعاً: التحليل التفصيلي لكل محاضر | Detailed Analysis',{bold:true,size:22,color:DARK,before:0,after:100}),
  );

  for(const [li,lec] of lecturers.entries()){
    const cl=clfR(lec.mean||0);
    children.push(
      mP(`${li+1}. ${lec.name}`,{bold:true,size:22,color:MID,before:li===0?0:180,after:40}),
      mP(`المقررات: ${(lec.courses||[]).join(' | ')}  |  الشعب: ${(lec.secs||[]).length}  |  المقيّمون: ${lec.n}  |  المتوسط: ${(lec.mean||0).toFixed(2)}  |  ${cl.l}`,
        {size:16,color:'444444',before:0,after:70}),
    );
    const dqW=Math.max(400,Math.floor((CW-600-2400-700-1400)/Math.max(nQ,1)));
    const dLastW=Math.max(400,CW-600-2400-700-dqW*nQ-1400);
    const dC=[600,2400,700,...Array(nQ).fill(dqW),dLastW];
    children.push(new Table({width:{size:CW,type:WidthType.DXA},columnWidths:dC,rows:[
      new TableRow({children:[
        mH(['الشعبة'],dC[0]),mH(['المقرر'],dC[1]),mH(['المقيّمون'],dC[2]),
        ...Array.from({length:nQ},(_,i)=>mH(['Q'+(i+1)],dC[3+i],MID,12)),
        mH(['المتوسط'],dC[3+nQ]),
      ]}),
      ...(lec.secs||[]).map((s,si)=>{
        const scl=clfR(s.sec_mean||0); const bg=si%2===0?PALE:WHITE;
        return new TableRow({children:[
          mC(s.sec_num,dC[0],bg,{size:13}),
          mC(s.course,dC[1],bg,{bold:true,color:DARK,size:13}),
          mC(s.n,dC[2],bg,{bold:true}),
          ...(s.questions||[]).map((q,qi)=>{const qcl=clfR(q.mean||0);return mC((q.mean||0).toFixed(2),dC[3+qi],qcl.bg,{color:qcl.c,size:12});}),
          mC((s.sec_mean||0).toFixed(2),dC[3+nQ],scl.bg,{bold:true,color:scl.c,size:15}),
        ]});
      }),
      new TableRow({children:[
        mC('المتوسط',dC[0]+dC[1],DARK,{bold:true,color:WHITE,colSpan:2}),
        mC(lec.n,dC[2],PALE,{bold:true}),
        ...(lec.qMeans||[]).map((qm,qi)=>{const qcl=clfR(qm);return mC(qm.toFixed(2),dC[3+qi],qcl.bg,{bold:true,color:qcl.c,size:13});}),
        mC((lec.mean||0).toFixed(2),dC[3+nQ],cl.bg,{bold:true,color:cl.c,size:17}),
      ]}),
    ]}),sp(60,40));
  }

  // ── ENHANCEMENT PLANS ─────────────────────────────────────────────────
  children.push(
    new Paragraph({pageBreakBefore:true,children:[]}),
    mP('Enhancement Plans | خطة التحسين',{bold:true,size:30,color:DARK,before:0,after:80}),
    mP('بناءً على نتائج التقييم — العناصر التي تحتاج اهتماماً مرتبة حسب الأولوية',{size:17,color:'555555',italic:true,before:0,after:120}),
  );

  // Build EP items per Q overall
  const epItems=[];
  Array.from({length:nQ},(_,qi)=>{
    const oM=overallQM[qi]; if(oM===undefined||oM===0) return;
    const cl=clfR(oM);
    const posP=Math.round(allSecs.reduce((a,s)=>{const q=s.questions[qi];return a+(q?(q.pct_sa||0)+(q.pct_a||0):0);},0)/(allSecs.length||1));
    epItems.push({qi,text:qTexts[qi]||('Q'+(qi+1)),mean:oM,cl,posP,
      priority:oM<2.5?'🔴 High':oM<3.5?'🟡 Medium':'🟢 Good',
      action:oM<2.5?`Immediate improvement required for Q${qi+1}. Develop action plan and assign responsible parties.`
            :oM<3.5?`Review and enhance Q${qi+1} through training and peer feedback sessions.`
            :`Maintain performance on Q${qi+1}. Document as best practice.`,
      kpi:oM<3.5?`Raise mean to ≥${Math.min(5,+(oM+0.5).toFixed(1))}; Positive% ≥${Math.min(95,posP+15)}%`
               :`Maintain mean ≥${oM}; Positive% ≥${posP}%`,
    });
  });
  epItems.sort((a,b)=>a.mean-b.mean); // worst (lowest) first for 5=best scale

  const epC=[900,1200,3500,800,1400,1400,Math.max(400,CW-900-1200-3500-800-1400-1400)];
  children.push(new Table({width:{size:CW,type:WidthType.DXA},columnWidths:epC,rows:[
    new TableRow({children:[
      mH(['Priority'],epC[0]),mH(['Q#'],epC[1]),
      mH(['Survey Item / Question'],epC[2]),
      mH(['Mean'],epC[3]),mH(['Classification'],epC[4]),
      mH(['Positive%\nSA+A'],epC[5]),
      mH(['Recommended Action | KPI'],epC[6]),
    ]}),
    ...epItems.map((item,i)=>{
      const bg=i%2===0?PALE:WHITE;
      return new TableRow({children:[
        mC(item.priority,epC[0],item.cl.bg,{bold:true,color:item.cl.c,size:13}),
        mC('Q'+(item.qi+1),epC[1],bg,{bold:true,color:DARK}),
        mC(item.text.slice(0,60),epC[2],bg,{align:AlignmentType.RIGHT,size:12}),
        mC(item.mean.toFixed(2),epC[3],item.cl.bg,{bold:true,color:item.cl.c}),
        mC(item.cl.l,epC[4],item.cl.bg,{bold:true,color:item.cl.c,size:13}),
        mC(item.posP+'%',epC[5],item.posP>=80?GREEN2:item.posP>=60?AMBER2:RED2,
          {bold:true,color:item.posP>=80?GREEN:item.posP>=60?AMBER:RED,size:14}),
        mC(item.action+'\n'+item.kpi,epC[6],bg,{align:AlignmentType.LEFT,size:12}),
      ]});
    }),
  ]}),sp(200,80));

  return buildDoc(children);
}


const PORT=process.env.PORT||3000;
app.listen(PORT,()=>console.log('Server on port',PORT,'| v2026-05-21-CHARTS'));
