const express = require('express');
const multer  = require('multer');
const XLSX    = require('xlsx');
const path    = require('path');
const fs      = require('fs');

const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType, VerticalAlign, PageOrientation
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
app.get('/', (req,res) => {
  const p1 = path.join(__dirname,'public','index.html');
  if(fs.existsSync(p1)) return res.sendFile(p1);
  try {
    const files = fs.readdirSync(path.join(__dirname,'public')).filter(f=>f.endsWith('.html'));
    if(files.length) return res.sendFile(path.join(__dirname,'public',files[0]));
  } catch(e){}
  const p2 = path.join(__dirname,'index.html');
  if(fs.existsSync(p2)) return res.sendFile(p2);
  res.send('Server running');
});

// ── Colors & helpers ──────────────────────────────────────────────────────────
const DARK='1F4E79',MID='2E75B6',PALE='EBF3FB',WHITE='FFFFFF';
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

const mC=(text,w,shade,opts={})=>new TableCell({
  width:{size:w,type:WidthType.DXA},borders:allB(),
  shading:shade?{fill:shade,type:ShadingType.CLEAR}:undefined,
  margins:mg(),verticalAlign:VerticalAlign.CENTER,
  rowSpan:opts.rowSpan,columnSpan:opts.colSpan,
  children:[new Paragraph({alignment:opts.align||AlignmentType.CENTER,
    children:[new TextRun({text:String(text??''),bold:opts.bold||false,
      color:opts.color||'000000',size:opts.size||17,font:'Arial'})]})]
});

const mH=(lines,w,shade=DARK,size=16)=>new TableCell({
  width:{size:w,type:WidthType.DXA},borders:allB(shade),
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
  const grandMean=+(allSecs.reduce((a,s)=>a+s.sec_mean*s.n,0)/totalN).toFixed(3);
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
          mC(lec.secs.length,t2C[2],bg),mC(lec.n,t2C[3],bg,{bold:true}),
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
  const ccW=Math.floor((CW-1800-1400)/(hasGender?nCC*2:nCC));
  const t3C=[1800,...(hasGender?uniqueCourses.flatMap(()=>[ccW,ccW]):uniqueCourses.map(()=>ccW)),1400];

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
async function buildWordFromResult(result, cfg){
  const CW=15398;
  const {nF,nM,n,secs,overall}=result;
  const showG=(cfg.gmode==='col');
  const children=[];
  const gCl=clf(overall);
  const totalQ=secs.reduce((a,s)=>a+(s.qs||[]).length,0);

  // ── TITLE ──────────────────────────────────────────────────────────────
  children.push(
    sp(0,200),
    mP(cfg.sname||'تقرير استبانة',{align:AlignmentType.CENTER,bold:true,size:52,color:DARK,before:0,after:80}),
    mP('تحليل نتائج الاستبيان',{align:AlignmentType.CENTER,bold:true,size:30,color:MID,before:0,after:80}),
    mP(cfg.cname||'كليات الغد للعلوم الطبية التطبيقية – جدة',{align:AlignmentType.CENTER,size:22,color:'555555',before:0,after:60}),
    mP(cfg.semester||'',{align:AlignmentType.CENTER,size:20,color:'777777',before:0,after:300}),
  );

  // ── GOAL ──────────────────────────────────────────────────────────────
  if(cfg.obj){
    children.push(
      mP('هدف الاستبيان',{bold:true,size:24,color:DARK,before:200,after:80}),
      mP(cfg.obj,{size:20,before:0,after:200}),
    );
  }

  // ── SCALE TABLE ────────────────────────────────────────────────────────
  const sCols=[Math.round(CW*0.28),Math.round(CW*0.18),Math.round(CW*0.18),Math.round(CW*0.18),Math.round(CW*0.18)];
  children.push(
    mP('المقياس المستخدم | Scale Used',{bold:true,size:22,color:DARK,before:0,after:80}),
    new Table({width:{size:CW,type:WidthType.DXA},columnWidths:sCols,rows:[
      new TableRow({children:[
        mH(['1 = Strongly Agree │ موافق بشدة'],sCols[0],GREEN),
        mH(['2 = Agree │ موافق'],sCols[1],GREEN),
        mH(['3 = Neutral │ محايد'],sCols[2],'7F7F7F'),
        mH(['4 = Disagree │ غير موافق'],sCols[3],RED),
        mH(['5 = Strongly Disagree │ غير موافق بشدة'],sCols[4],RED),
      ]}),
    ]}),
    sp(160,80),
  );

  // ── CLASSIFICATION SCALE TABLE ──────────────────────────────────────────
  const clsCols=[Math.round(CW*0.12),Math.round(CW*0.16),Math.round(CW*0.14),Math.round(CW*0.06),CW-Math.round(CW*0.48)];
  children.push(
    mP('Classification Scale | مقياس التصنيف',{bold:true,size:22,color:DARK,before:0,after:80}),
    new Table({width:{size:CW,type:WidthType.DXA},columnWidths:clsCols,rows:[
      new TableRow({children:[mH(['Range'],clsCols[0]),mH(['Classification'],clsCols[1]),mH(['التصنيف'],clsCols[2]),mH([''],clsCols[3]),mH(['Interpretation'],clsCols[4])]}),
      ...[
        ['≤ 1.50','Excellent','ممتاز','','Strong positive outcome',GREEN2,GREEN],
        ['1.51–2.00','Good','جيد','','Positive; meets expectations',GREEN2,GREEN],
        ['2.01–2.50','Acceptable','مقبول','','Moderate; requires monitoring',AMBER2,AMBER],
        ['2.51–3.00','Weakness','ضعف','','Below expectations; improvement needed',RED2,RED],
        ['> 3.00','Critical','حرج','','Significant weakness; immediate action required',RED2,RED],
      ].map(([r,cl,ar,_,interp,bg,c])=>new TableRow({children:[
        mC(r,clsCols[0],bg,{bold:true,color:c,align:AlignmentType.CENTER}),
        mC(cl,clsCols[1],bg,{bold:true,color:c}),
        mC(ar,clsCols[2],bg,{bold:true,color:c}),
        mC('',clsCols[3],bg),
        mC(interp,clsCols[4],WHITE,{size:17,align:AlignmentType.LEFT}),
      ]}))
    ]}),
    sp(160,80),
  );

  // ── SAMPLE PROFILE ─────────────────────────────────────────────────────
  const spCols=[Math.round(CW*0.22),Math.round(CW*0.15),Math.round(CW*0.22),Math.round(CW*0.15)];
  children.push(
    mP('Sample Profile | بيانات العينة',{bold:true,size:22,color:DARK,before:0,after:80}),
    new Table({width:{size:CW,type:WidthType.DXA},columnWidths:spCols,rows:[
      new TableRow({children:[mH(['Detail'],spCols[0]),mH(['Value'],spCols[1]),mH(['التفاصيل'],spCols[2]),mH(['القيمة'],spCols[3])]}),
      ...[
        ['Total Respondents',n,'إجمالي المشاركين',n],
        ['Female',nF||'—','إناث',nF||'—'],
        ['Male',nM||'—','ذكور',nM||'—'],
        ['Sections',secs.length,'عدد المحاور',secs.length],
        ['Total Questions',totalQ,'عدد الأسئلة',totalQ],
        ['Overall Mean',overall,'المتوسط العام',overall],
        ['Survey Period',cfg.semester||'—','الفصل الدراسي',cfg.semester||'—'],
      ].map(([e,ev,a,av],i)=>new TableRow({children:[
        mC(e,spCols[0],i%2===0?PALE:WHITE,{align:AlignmentType.LEFT}),
        mC(ev,spCols[1],i%2===0?PALE:WHITE,{bold:true,color:DARK}),
        mC(a,spCols[2],i%2===0?PALE:WHITE,{align:AlignmentType.RIGHT}),
        mC(av,spCols[3],i%2===0?PALE:WHITE,{bold:true,color:DARK}),
      ]}))
    ]}),
    sp(200,80),
  );

  // ── SECTION SUMMARY ────────────────────────────────────────────────────
  const ssCols=showG
    ?[1400,3000,1200,1200,1200,1200,CW-1400-3000-1200*4]
    :[1400,4500,1400,CW-1400-4500-1400];
  const ssHdrs=[mH(['Section'],ssCols[0]),mH(['المحور'],ssCols[1]),mH(['Mean'],ssCols[2])];
  if(showG) ssHdrs.push(mH(['F.Mean'],ssCols[3]),mH(['M.Mean'],ssCols[4]),mH(['Gap'],ssCols[5]));
  ssHdrs.push(mH(['Classification'],ssCols[ssCols.length-1]));

  children.push(
    mP('Section Summary | ملخص المحاور',{bold:true,size:22,color:DARK,before:0,after:80}),
    new Table({width:{size:CW,type:WidthType.DXA},columnWidths:ssCols,rows:[
      new TableRow({children:ssHdrs}),
      ...secs.map((s,i)=>{
        const cl=clf(s.mean); const bg=i%2===0?PALE:WHITE;
        const cells=[
          mC(s.name,ssCols[0],bg,{align:AlignmentType.LEFT}),
          mC(s.ar,ssCols[1],bg,{align:AlignmentType.RIGHT}),
          mC(s.mean,ssCols[2],cl.bg,{bold:true,color:cl.c}),
        ];
        if(showG) cells.push(
          mC(s.fMean,ssCols[3],'FCE4D6',{color:'843C0C'}),
          mC(s.mMean,ssCols[4],'DDEBF7',{color:DARK}),
          mC(Math.abs(s.fMean-s.mMean).toFixed(2),ssCols[5],bg),
        );
        cells.push(mC(cl.l+' │ '+cl.l,ssCols[ssCols.length-1],cl.bg,{bold:true,color:cl.c}));
        return new TableRow({children:cells});
      }),
    ]}),
    sp(200,80),
  );

  // ── EXECUTIVE OVERVIEW ─────────────────────────────────────────────────
  const bestSec=secs.reduce((a,b)=>a.mean<b.mean?a:b);
  const worstSec=secs.reduce((a,b)=>a.mean>b.mean?a:b);
  const gapNote=showG&&nF>0&&nM>0?`الفجوة بين الجنسين: ${Math.abs(secs.reduce((a,s)=>a+s.fMean,0)/secs.length - secs.reduce((a,s)=>a+s.mMean,0)/secs.length).toFixed(2)} — ${Math.abs(secs.reduce((a,s)=>a+s.fMean,0)/secs.length - secs.reduce((a,s)=>a+s.mMean,0)/secs.length)<0.3?'طفيفة تدل على تجانس التجربة.':'تستحق المراجعة.'}`:'';

  const bullets=[
    `تُظهر النتائج مستوى رضا ${gCl.l} بمتوسط عام (${overall}) — تصنيف "${gCl.l}".`,
    `أقوى المحاور: "${bestSec.ar}" بمتوسط (${bestSec.mean}) — ${clf(bestSec.mean).l}.`,
    `المحور الأولى بالتطوير: "${worstSec.ar}" بمتوسط (${worstSec.mean}) — ${clf(worstSec.mean).l}.`,
    ...(gapNote?[gapNote]:[]),
    overall<=1.5?'نسب الموافقة الإيجابية المرتفعة تعكس جودة تدريبية تستحق التوثيق.':'تُوصى بمراجعة المحاور ذات المتوسط المرتفع وتطوير خطة تحسين.',
  ];
  children.push(
    mP('اللمحة العامة',{bold:true,size:22,color:DARK,before:0,after:80}),
    ...bullets.map(b=>new Paragraph({
      alignment:AlignmentType.RIGHT,spacing:{before:60,after:60},
      numbering:{reference:'bullets',level:0},
      children:[new TextRun({text:b,size:20,font:'Arial',color:'222222'})]
    })),
    sp(200,80),
  );

  // ── OVERALL ANALYSIS TABLE ──────────────────────────────────────────────
  const oaCols=showG
    ?[700,200,1200,900,900,900,900,900,700,700,CW-700-200-1200-900*5-700-700]
    :[700,200,1200,900,900,900,900,900,700,700,CW-700-200-1200-900*5-700-700];

  const oaHdrs=[
    mH(['Q#'],oaCols[0]),mH(['Sec.Q'],oaCols[1]),mH(['Section'],oaCols[2]),
    mH(['F.Mean'],oaCols[3]),mH(['M.Mean'],oaCols[4]),
    mH(['Max'],oaCols[5]),mH(['Min'],oaCols[6]),
    mH(['Mean'],oaCols[7]),mH(['Pos%'],oaCols[8]),mH(['Neg%'],oaCols[9]),
    mH(['Classification'],oaCols[10]),
  ];

  const oaRows=[];
  secs.forEach(s=>{
    oaRows.push(new TableRow({children:[
      new TableCell({width:{size:CW,type:WidthType.DXA},columnSpan:11,
        borders:allB(MID),shading:{fill:MID,type:ShadingType.CLEAR},margins:mg(),
        children:[new Paragraph({alignment:AlignmentType.CENTER,
          children:[new TextRun({text:s.name+' │ '+s.ar,bold:true,color:WHITE,size:18,font:'Arial'})]})],
      })
    ]}));
    (s.qs||[]).forEach((q,qi)=>{
      const cl=clf(q.cM||0); const bg=qi%2===0?PALE:WHITE;
      oaRows.push(new TableRow({children:[
        mC('Q'+q.qn,oaCols[0],bg,{bold:true,color:DARK,size:15}),
        mC(qi+1,oaCols[1],bg,{size:14}),
        mC(s.name,oaCols[2],bg,{size:13,align:AlignmentType.LEFT}),
        mC(q.fM??q.cM,oaCols[3],'FCE4D6',{color:'843C0C',size:14}),
        mC(q.mM??q.cM,oaCols[4],'DDEBF7',{color:DARK,size:14}),
        mC(q.maxM??q.cM,oaCols[5],bg,{size:13}),
        mC(q.minM??q.cM,oaCols[6],bg,{size:13}),
        mC(q.cM,oaCols[7],cl.bg,{bold:true,color:cl.c}),
        mC((q.pos||0)+'%',oaCols[8],(q.pos||0)>=80?GREEN2:bg,{color:(q.pos||0)>=80?GREEN:'000000',size:13}),
        mC((q.neg||0)+'%',oaCols[9],(q.neg||0)>20?RED2:bg,{color:(q.neg||0)>20?RED:'000000',size:13}),
        mC(cl.l+' │ '+cl.l,oaCols[10],cl.bg,{bold:true,color:cl.c,size:13}),
      ]}));
    });
  });

  children.push(
    mP('Overall Analysis | التحليل الإجمالي',{bold:true,size:22,color:DARK,before:0,after:80}),
    new Table({width:{size:CW,type:WidthType.DXA},columnWidths:oaCols,rows:[new TableRow({children:oaHdrs}),...oaRows]}),
    sp(200,80),
    new Paragraph({children:[new TextRun({break:1})]}),
  );

  // ── DETAILED DISTRIBUTION PER SECTION ──────────────────────────────────
  secs.forEach((sec,si)=>{
    const cl=clf(sec.mean);
    if(si>0) children.push(new Paragraph({pageBreakBefore:true,children:[]}));

    children.push(
      mP(sec.name+' | '+sec.ar,{bold:true,size:26,color:MID,before:si===0?0:0,after:40}),
      mP(`Mean: ${sec.mean}  |  F.Mean: ${sec.fMean}  |  M.Mean: ${sec.mMean}  |  ${cl.l} (${cl.l})`,{size:17,color:'444444',before:0,after:80}),
    );

    const dCols=[700,200,1100,1000,1000,1000,1000,1000,900,CW-700-200-1100-1000*4-900];
    children.push(new Table({width:{size:CW,type:WidthType.DXA},columnWidths:dCols,rows:[
      new TableRow({children:[
        mH(['Global Q'],dCols[0]),mH(['Sec. Q'],dCols[1]),mH(['Group'],dCols[2]),
        mH(['%1\nموافق بشدة\nSA'],dCols[3]),mH(['%2\nموافق\nA'],dCols[4]),
        mH(['%3\nمحايد\nN'],dCols[5]),mH(['%4\nلا أوافق\nD'],dCols[6]),
        mH(['%5\nلا أوافق بشدة\nSD'],dCols[7]),mH(['Mean'],dCols[8]),
        mH([' '],dCols[9]),
      ]}),
      ...(sec.qs||[]).flatMap((q,qi)=>{
        const fD=q.fD||q.cD||[0,0,0,0,0];
        const mD=q.mD||q.cD||[0,0,0,0,0];
        const cD=q.cD||q.fD||[0,0,0,0,0];
        const bg=qi%2===0?PALE:WHITE;
        const rows=showG?[
          ['Female',fD,q.fM??q.cM,'FCE4D6','843C0C',true],
          ['Male',  mD,q.mM??q.cM,'DDEBF7',DARK,    false],
          ['Combined',cD,q.cM,    'E2EFDA',GREEN,    false],
        ]:[
          ['Overall',cD,q.cM,'E2EFDA',GREEN,true],
        ];
        return rows.map((row,ri)=>{
          const [grp,d,m,cb,tc,first]=row;
          const cells=[];
          if(first){
            cells.push(new TableCell({width:{size:dCols[0]},rowSpan:rows.length,
              borders:allB(),shading:{fill:bg,type:ShadingType.CLEAR},margins:mg(),
              verticalAlign:VerticalAlign.CENTER,
              children:[new Paragraph({alignment:AlignmentType.CENTER,children:[
                new TextRun({text:'Q'+q.qn,bold:true,color:DARK,size:16,font:'Arial'})
              ]})]}));
            cells.push(new TableCell({width:{size:dCols[1]},rowSpan:rows.length,
              borders:allB(),shading:{fill:bg,type:ShadingType.CLEAR},margins:mg(),
              verticalAlign:VerticalAlign.CENTER,
              children:[new Paragraph({alignment:AlignmentType.CENTER,children:[
                new TextRun({text:String(qi+1),size:14,font:'Arial'})
              ]})]}));
            // Q text as header row before data
          }
          cells.push(mC(grp,dCols[2],cb,{bold:true,color:tc,size:14,align:AlignmentType.LEFT}));
          d.forEach((_,j)=>{
            cells.push(mC(d[j],dCols[3+j],cb,{color:tc,size:14}));
          });
          cells.push(mC(parseFloat(m).toFixed(2),dCols[8],clf(parseFloat(m)).bg,{bold:true,color:clf(parseFloat(m)).c}));
          cells.push(mC('',dCols[9],WHITE));
          return new TableRow({children:cells});
        });
      }),
    ]}),sp(80,40));
  });


  // ── ENHANCEMENT PLANS ─────────────────────────────────────────────────────
  children.push(
    new Paragraph({pageBreakBefore:true,children:[]}),
    mP('خطة التحسين والتطوير | Enhancement Plans',{bold:true,size:30,color:DARK,before:0,after:80}),
    mP('بناءً على نتائج الاستبيان، يوصى بتبني الخطة التالية لتعزيز جودة التعليم:',{size:18,color:'555555',before:0,after:120}),
  );

  // Build enhancement items based on results
  const enhancements=[];
  secs.forEach(s=>{
    const cl=clf(s.mean);
    if(s.mean>2.0){  // Needs improvement
      enhancements.push({
        area:s.ar,
        status:cl.l,
        mean:s.mean,
        priority:s.mean>3.0?'عاجل':'متوسط',
        action:s.mean>3.0
          ?`مراجعة فورية لـ "${s.ar}" وإعداد خطة تدخل خلال الفصل الدراسي القادم`
          :`تعزيز وتطوير "${s.ar}" من خلال ورش عمل وجلسات تدريبية`,
        kpi:`رفع المتوسط من ${s.mean} إلى أقل من 2.0 خلال فصلين دراسيين`,
      });
    }
  });

  // Always add strengths to build on
  const strengths=secs.filter(s=>s.mean<=2.0);
  const weaknesses=secs.filter(s=>s.mean>2.0);

  // Enhancement table
  const epC=[2200,1800,1200,1000,3600,CW-2200-1800-1200-1000-3600];
  children.push(
    new Table({width:{size:CW,type:WidthType.DXA},columnWidths:epC,rows:[
      new TableRow({children:[
        mH(['المجال / Area'],epC[0]),
        mH(['المتوسط / Mean'],epC[1]),
        mH(['الأولوية / Priority'],epC[2]),
        mH(['الحالة / Status'],epC[3]),
        mH(['الإجراء المقترح / Recommended Action'],epC[4]),
        mH(['مؤشر النجاح / KPI'],epC[5]),
      ]}),
      ...secs.map((s,i)=>{
        const cl=clf(s.mean);
        const priority=s.mean>3.0?'🔴 عاجل':s.mean>2.0?'🟡 متوسط':'🟢 قوة';
        const action=s.mean>3.0
          ?`مراجعة فورية وخطة تدخل عاجلة لتحسين "${s.ar}"`
          :s.mean>2.0
          ?`تطوير وتعزيز "${s.ar}" عبر برامج تدريبية`
          :`توثيق وتكرار ممارسات النجاح في "${s.ar}"`;
        const kpi=s.mean>2.0
          ?`رفع المتوسط من ${s.mean} إلى أقل من 2.0`
          :`الحفاظ على المتوسط ${s.mean} أو تحسينه`;
        const bg=i%2===0?PALE:WHITE;
        return new TableRow({children:[
          mC(s.ar,epC[0],bg,{align:AlignmentType.RIGHT,bold:s.mean>2.0}),
          mC(s.mean,epC[1],cl.bg,{bold:true,color:cl.c}),
          mC(priority,epC[2],s.mean>3.0?RED2:s.mean>2.0?AMBER2:GREEN2,{bold:true,color:s.mean>3.0?RED:s.mean>2.0?AMBER:GREEN}),
          mC(cl.l,epC[3],cl.bg,{bold:true,color:cl.c,size:14}),
          mC(action,epC[4],bg,{align:AlignmentType.RIGHT,size:15}),
          mC(kpi,epC[5],bg,{align:AlignmentType.RIGHT,size:14,color:'555555'}),
        ]});
      }),
    ]}),
    sp(200,100),
  );

  // Action plan summary
  children.push(
    mP('ملخص خطة العمل | Action Plan Summary',{bold:true,size:22,color:DARK,before:0,after:80}),
  );

  const planItems=[
    weaknesses.length>0
      ?`📌 يوجد ${weaknesses.length} محور يحتاج تطوير فوري: ${weaknesses.map(s=>s.ar).join(' | ')}`
      :'✅ جميع المحاور ضمن المستوى المقبول — استمر في التميز',
    strengths.length>0
      ?`💪 نقاط القوة (${strengths.length} محور): ${strengths.map(s=>s.ar).join(' | ')} — يُوصى بتوثيقها كممارسات جيدة`
      :'',
    `📊 المتوسط العام الحالي: ${overall} — الهدف المقترح: ${Math.max(1.0,parseFloat((overall-0.3).toFixed(2)))} خلال فصل دراسي`,
    `🗓️ يُقترح مراجعة النتائج في منتصف الفصل القادم وإعداد تقرير متابعة`,
    `👥 مشاركة نتائج الاستبيان مع أعضاء هيئة التدريس في اجتماع القسم القادم`,
  ].filter(Boolean);

  children.push(
    ...planItems.map(item=>new Paragraph({
      alignment:AlignmentType.RIGHT,
      spacing:{before:60,after:60},
      numbering:{reference:'bullets',level:0},
      children:[new TextRun({text:item,size:19,font:'Arial',color:'222222'})]
    })),
    sp(200,80),
  );

  return buildDoc(children);
}

async function buildInstructorWordFromResult(result, cfg){
  const CW=15398;
  const {allSecs,lecturers,qTexts,totalN,totalEnrolled,totalNot,grandMean}=result;
  const nQ=qTexts.length;
  const gCl=clf(grandMean);
  const pct=totalEnrolled>0?Math.round(totalN/totalEnrolled*100):0;
  const children=[];

  // unique courses list
  const courses=[...new Set(allSecs.map(s=>s.course))];

  // ── TITLE ──────────────────────────────────────────────────────────────
  children.push(
    sp(0,200),
    mP('تقرير استبانة تقييم المحاضرين',{align:AlignmentType.CENTER,bold:true,size:52,color:DARK,before:0,after:80}),
    mP('Instructor Evaluation Report',{align:AlignmentType.CENTER,bold:true,size:32,color:MID,before:0,after:80}),
    mP(cfg.cname||'كليات الغد للعلوم الطبية التطبيقية – جدة',{align:AlignmentType.CENTER,size:22,color:'555555',before:0,after:60}),
    mP((cfg.dept||'')+(cfg.semester?' | '+cfg.semester:''),{align:AlignmentType.CENTER,size:20,color:'777777',before:0,after:300}),
  );

  // ── OVERALL STATS ──────────────────────────────────────────────────────
  const stC=[Math.round(CW/5),Math.round(CW/5),Math.round(CW/5),Math.round(CW/5),CW-Math.round(CW/5)*4];
  children.push(
    mP('إحصائيات عامة | Overall Statistics',{bold:true,size:24,color:DARK,before:0,after:80}),
    new Table({width:{size:CW,type:WidthType.DXA},columnWidths:stC,rows:[
      new TableRow({children:[mH(['إجمالي الشعب'],stC[0]),mH(['إجمالي المقيّمين'],stC[1]),mH(['المحاضرون'],stC[2]),mH(['المقررات'],stC[3]),mH(['المتوسط العام'],stC[4])]}),
      new TableRow({children:[
        mC(allSecs.length,stC[0],PALE,{bold:true,color:DARK,size:28}),
        mC(totalN,stC[1],GREEN2,{bold:true,color:GREEN,size:28}),
        mC(lecturers.length,stC[2],PALE,{bold:true,color:DARK,size:28}),
        mC(courses.length,stC[3],PALE,{bold:true,color:DARK,size:28}),
        mC(grandMean,stC[4],gCl.bg,{bold:true,color:gCl.c,size:36}),
      ]}),
    ]}),
    sp(200,100),
  );

  // ── SECTION 1: ALL SECTIONS SUMMARY ────────────────────────────────────
  const s1C=[600,2800,2600,1000,1000,1200,1400,CW-600-2800-2600-1000-1000-1200-1400];
  children.push(
    mP('أولاً: ملخص جميع الشعب | All Sections Summary',{bold:true,size:22,color:DARK,before:0,after:80}),
    new Table({width:{size:CW,type:WidthType.DXA},columnWidths:s1C,rows:[
      new TableRow({children:[
        mH(['الشعبة'],s1C[0]),mH(['المقرر'],s1C[1]),mH(['المحاضر'],s1C[2]),
        mH(['المسجلون'],s1C[3]),mH(['المقيّمون'],s1C[4]),
        mH(['نسبة\nالمشاركة'],s1C[5]),mH(['المتوسط'],s1C[6]),mH(['التصنيف'],s1C[7]),
      ]}),
      ...allSecs.map((s,i)=>{
        const cl=clf(s.sec_mean); const bg=i%2===0?PALE:WHITE;
        const pp=s.participation_pct||0;
        return new TableRow({children:[
          mC(s.sec_num,s1C[0],bg,{size:14}),
          mC(s.course,s1C[1],bg,{bold:true,color:DARK,size:15}),
          mC(s.lecturer,s1C[2],bg,{align:AlignmentType.RIGHT,size:14}),
          mC(s.enrolled,s1C[3],bg),mC(s.n,s1C[4],bg,{bold:true}),
          mC(pp+'%',s1C[5],pp>=80?GREEN2:pp>=60?AMBER2:RED2,{color:pp>=80?GREEN:pp>=60?AMBER:RED,bold:true}),
          mC(s.sec_mean.toFixed(2),s1C[6],cl.bg,{bold:true,color:cl.c}),
          mC(cl.l,s1C[7],cl.bg,{bold:true,color:cl.c,size:14}),
        ]});
      }),
      new TableRow({children:[
        mC('الإجمالي',s1C[0]+s1C[1]+s1C[2],DARK,{bold:true,color:WHITE,colSpan:3}),
        mC('—',s1C[3],PALE),mC(totalN,s1C[4],GREEN2,{bold:true,color:GREEN,size:18}),
        mC(pct+'%',s1C[5],pct>=80?GREEN2:pct>=60?AMBER2:RED2,{bold:true,color:pct>=80?GREEN:pct>=60?AMBER:RED,size:16}),
        mC(grandMean,s1C[6],gCl.bg,{bold:true,color:gCl.c,size:20}),
        mC(gCl.l,s1C[7],gCl.bg,{bold:true,color:gCl.c}),
      ]}),
    ]}),
    sp(200,100),
  );

  // ── SECTION 2: INSTRUCTOR SUMMARY ──────────────────────────────────────
  const lqW=Math.max(700,Math.floor((CW-400-2800-700-700-1500)/Math.max(nQ,1)));
  const s2C=[400,2800,700,700,...Array(nQ).fill(lqW),1500,CW-400-2800-700-700-lqW*nQ-1500];
  children.push(
    new Paragraph({pageBreakBefore:true,children:[]}),
    mP('ثانياً: ملخص تقييم المحاضرين (متوسط موزون) | Instructor Summary',{bold:true,size:22,color:DARK,before:0,after:80}),
    new Table({width:{size:CW,type:WidthType.DXA},columnWidths:s2C,rows:[
      new TableRow({children:[
        mH(['#'],s2C[0]),mH(['المحاضر'],s2C[1]),mH(['الشعب'],s2C[2]),mH(['عدد\nالمقيّمين'],s2C[3]),
        ...qTexts.map((_,i)=>mH(['Q'+(i+1)],s2C[4+i],MID,13)),
        mH(['المتوسط\nالموزون'],s2C[4+nQ]),mH(['التصنيف'],s2C[5+nQ]),
      ]}),
      ...lecturers.map((lec,i)=>{
        const cl=clf(lec.mean||0); const bg=i%2===0?PALE:WHITE;
        return new TableRow({children:[
          mC(i+1,s2C[0],bg,{bold:true,color:DARK,size:13}),
          mC(lec.name,s2C[1],bg,{align:AlignmentType.RIGHT,size:13}),
          mC(lec.secs.length,s2C[2],bg),mC(lec.n,s2C[3],bg,{bold:true}),
          ...lec.qMeans.map((qm,qi)=>{const qcl=clf(qm);return mC(qm.toFixed(2),s2C[4+qi],qcl.bg,{color:qcl.c,size:12});}),
          mC((lec.mean||0).toFixed(2),s2C[4+nQ],cl.bg,{bold:true,color:cl.c,size:16}),
          mC(cl.l,s2C[5+nQ],cl.bg,{bold:true,color:cl.c,size:12}),
        ]});
      }),
    ]}),
    sp(200,100),
  );

  // ── SECTION 3: COURSE COMPARISON TABLE ─────────────────────────────────
  const nC=courses.length;
  const qLabelW=Math.round(CW*0.16);
  const cqW=Math.floor((CW-qLabelW-1200)/(nC+1));
  const s3C=[qLabelW,...Array(nC).fill(cqW),1200,CW-qLabelW-cqW*nC-1200];

  // Course means
  const courseData={};
  courses.forEach(code=>{
    const cs=allSecs.filter(s=>s.course===code);
    const tn=cs.reduce((a,s)=>a+s.n,0);
    courseData[code]={
      n:tn,
      mean:tn?+(cs.reduce((a,s)=>a+s.sec_mean*s.n,0)/tn).toFixed(2):0,
      qMeans:Array.from({length:nQ},(_,qi)=>tn?+(cs.reduce((a,s)=>a+(s.questions[qi]?.mean||0)*s.n,0)/tn).toFixed(2):0),
    };
  });
  const overallQMeans=Array.from({length:nQ},(_,qi)=>+(allSecs.reduce((a,s)=>a+(s.questions[qi]?.mean||0)*s.n,0)/totalN).toFixed(2));

  children.push(
    new Paragraph({pageBreakBefore:true,children:[]}),
    mP('ثالثاً: جدول المقارنة بين المقررات | Course Comparison Table',{bold:true,size:22,color:DARK,before:0,after:60}),
    mP('المتوسط الحسابي لكل سؤال عبر جميع المقررات — الخلايا الملونة تشير للتصنيف',{size:16,color:'777777',italic:true,before:0,after:80}),
    new Table({width:{size:CW,type:WidthType.DXA},columnWidths:s3C,rows:[
      // Header row: course codes
      new TableRow({children:[
        mH(['السؤال / Criteria'],s3C[0]),
        ...courses.map((code,ci)=>new TableCell({
          width:{size:cqW,type:WidthType.DXA},borders:allB(MID),
          shading:{fill:MID,type:ShadingType.CLEAR},margins:mg(),verticalAlign:VerticalAlign.CENTER,
          children:[
            new Paragraph({alignment:AlignmentType.CENTER,spacing:{before:0,after:0},children:[new TextRun({text:code,bold:true,color:WHITE,size:15,font:'Arial'})]}),
            new Paragraph({alignment:AlignmentType.CENTER,spacing:{before:0,after:0},children:[new TextRun({text:'('+courseData[code].mean.toFixed(2)+')',color:'AAAAAA',size:12,font:'Arial'})]}),
          ]
        })),
        mH(['الإجمالي'],s3C[nC+1]),
        mH(['التصنيف'],s3C[nC+2]),
      ]}),
      // Q rows
      ...Array.from({length:nQ},(_,qi)=>{
        const oM=overallQMeans[qi]; const oCl=clf(oM); const bg=qi%2===0?PALE:WHITE;
        return new TableRow({children:[
          mC('Q'+(qi+1)+' — '+(qTexts[qi]||'').slice(0,35),s3C[0],bg,{align:AlignmentType.RIGHT,size:13}),
          ...courses.map((code,ci)=>{const qm=courseData[code].qMeans[qi];const qcl=clf(qm);return mC(qm.toFixed(2),cqW,qcl.bg,{color:qcl.c,size:14});}),
          mC(oM.toFixed(2),s3C[nC+1],oCl.bg,{bold:true,color:oCl.c,size:15}),
          mC(oCl.l,s3C[nC+2],oCl.bg,{bold:true,color:oCl.c,size:13}),
        ]});
      }),
      // Mean row
      new TableRow({children:[
        mC('المتوسط العام',s3C[0],DARK,{bold:true,color:WHITE}),
        ...courses.map((code,ci)=>{const cl=clf(courseData[code].mean);return mC(courseData[code].mean.toFixed(2),cqW,cl.bg,{bold:true,color:cl.c,size:16});}),
        mC(grandMean,s3C[nC+1],gCl.bg,{bold:true,color:gCl.c,size:18}),
        mC(gCl.l,s3C[nC+2],gCl.bg,{bold:true,color:gCl.c}),
      ]}),
      // N row
      new TableRow({children:[
        mC('عدد المقيّمين',s3C[0],BLUE2||'DDEBF7',{bold:true,color:DARK}),
        ...courses.map((code,ci)=>mC(courseData[code].n,cqW,PALE,{color:DARK,size:13})),
        mC(totalN,s3C[nC+1],PALE,{bold:true,color:DARK,size:16}),
        mC('',s3C[nC+2],PALE),
      ]}),
    ]}),
    sp(200,100),
  );

  // ── SECTION 4: DETAILED PER LECTURER ───────────────────────────────────
  children.push(
    new Paragraph({pageBreakBefore:true,children:[]}),
    mP('رابعاً: التحليل التفصيلي لكل محاضر | Detailed Instructor Analysis',{bold:true,size:22,color:DARK,before:0,after:100}),
  );

  lecturers.forEach((lec,li)=>{
    const cl=clf(lec.mean||0);
    children.push(
      mP(`${li+1}. ${lec.name}`,{bold:true,size:20,color:MID,before:li===0?0:200,after:40}),
      mP(`المقررات: ${lec.courses.join(' | ')}  |  الشعب: ${lec.secs.length}  |  المقيّمون: ${lec.n}  |  المتوسط: ${(lec.mean||0).toFixed(2)}  |  ${cl.l}`,
        {size:15,color:'444444',before:0,after:70}),
    );

    const dC=[600,2600,800,...Array(nQ).fill(Math.floor((CW-600-2600-800-1400)/nQ)),1400];
    children.push(new Table({width:{size:CW,type:WidthType.DXA},columnWidths:dC,rows:[
      new TableRow({children:[
        mH(['الشعبة'],dC[0]),mH(['المقرر'],dC[1]),mH(['المقيّمون'],dC[2]),
        ...Array.from({length:nQ},(_,i)=>mH(['Q'+(i+1)],dC[3+i],MID,13)),
        mH(['المتوسط'],dC[3+nQ]),
      ]}),
      ...lec.secs.map((s,si)=>{
        const scl=clf(s.sec_mean); const bg=si%2===0?PALE:WHITE;
        return new TableRow({children:[
          mC(s.sec_num,dC[0],bg,{size:13}),
          mC(s.course,dC[1],bg,{bold:true,color:DARK,size:13}),
          mC(s.n,dC[2],bg,{bold:true}),
          ...(s.questions||[]).map((q,qi)=>{const qcl=clf(q.mean);return mC(q.mean.toFixed(2),dC[3+qi],qcl.bg,{color:qcl.c,size:12});}),
          mC(s.sec_mean.toFixed(2),dC[3+nQ],scl.bg,{bold:true,color:scl.c,size:15}),
        ]});
      }),
      // Weighted mean row
      new TableRow({children:[
        mC('المتوسط الموزون',dC[0]+dC[1],DARK,{bold:true,color:WHITE,colSpan:2}),
        mC(lec.n,dC[2],PALE,{bold:true}),
        ...lec.qMeans.map((qm,qi)=>{const qcl=clf(qm);return mC(qm.toFixed(2),dC[3+qi],qcl.bg,{bold:true,color:qcl.c,size:14});}),
        mC((lec.mean||0).toFixed(2),dC[3+nQ],cl.bg,{bold:true,color:cl.c,size:17}),
      ]}),
    ]}),sp(60,40));
  });

  return buildDoc(children);
}


const PORT=process.env.PORT||3000;
app.listen(PORT,()=>console.log('Server on port',PORT));
