const pptxgen = require('pptxgenjs');
const { warnIfSlideHasOverlaps, warnIfSlideElementsOutOfBounds } = require('/home/oai/skills/slides/pptxgenjs_helpers');

const pptx = new pptxgen();
pptx.layout = 'LAYOUT_WIDE'; // 13.333 x 7.5
pptx.author = 'OpenAI / ChatGPT';
pptx.subject = 'Assessment of whether MOC2 mangrove project can achieve 9.4 tCO2e/rai/year';
pptx.title = '9.4 tCO₂e/ไร่/ปี — ป่าของเราจะทำได้จริงไหม?';
pptx.company = 'MOC2 Carbon Analysis';
pptx.lang = 'th-TH';
pptx.theme = {
  headFontFace: 'Noto Sans Thai',
  bodyFontFace: 'Noto Sans Thai',
  lang: 'th-TH'
};
pptx.defineSlideMaster({
  title: 'MASTER',
  background: { color: '070A3D' },
  objects: [
    { line: { x: 0.55, y: 0.46, w: 0.36, h: 0, line: { color: '00FFAC', width: 2.5 } } },
    { text: { text: 'MOC2 • CARBON SCIENCE NOTE', options: { x: 0.98, y: 0.30, w: 3.3, h: 0.28, fontFace: 'Noto Sans Thai', fontSize: 9.5, color: 'B7C2FF', bold: true, charSpacing: 1.2, margin: 0 } } },
  ],
  slideNumber: { x: 12.22, y: 7.05, w: 0.42, h: 0.2, color: '7F8EDB', fontFace: 'Noto Sans Thai', fontSize: 9, align: 'right', margin: 0 }
});

const C = {
  bg: '070A3D', panel: '11165D', panel2: '171C8F', text: 'F0F6F7', muted: 'C6D2FF', dim: '7F8EDB',
  green: '00FFAC', green2: '097EF6', mint: 'BFEFFF', lime: '4DFFCF', yellow: '55B6FF', red: 'FF5C8A',
  white: 'FFFFFF', grid: '3036A3', gray: 'D8E0FF', black: '070A3D'
};
const FONT = 'Noto Sans Thai';

function addTitle(slide, kicker, title, sub='') {
  if (kicker) slide.addText(kicker.toUpperCase(), { x:0.72, y:0.85, w:4.8, h:0.28, fontFace:FONT, fontSize:10, bold:true, color:C.green, charSpacing:1.4, margin:0 });
  slide.addText(title, { x:0.72, y:1.18, w:11.95, h:0.78, fontFace:FONT, fontSize:31, bold:true, color:C.text, margin:0, breakLine:false, valign:'mid', fit:'shrink' });
  if (sub) slide.addText(sub, { x:0.72, y:2.02, w:11.65, h:0.55, fontFace:FONT, fontSize:15.2, color:C.muted, margin:0, fit:'shrink', breakLine:false });
}
function addFooter(slide, text) {
  slide.addText(text, { x:0.72, y:6.93, w:10.8, h:0.22, fontFace:FONT, fontSize:8.5, color:C.dim, margin:0, fit:'shrink' });
}
function roundedRect(slide, x,y,w,h, fill, line='none', radius=0.12) {
  slide.addShape(pptx.ShapeType.roundRect, { x,y,w,h, rectRadius: radius, fill:{color:fill}, line: line==='none'?{color:fill, transparency:100}:{color:line, width:1} });
}
function pill(slide, text, x,y,w, fill=C.panel2, color=C.green, fs=11) {
  roundedRect(slide,x,y,w,0.42,fill);
  slide.addText(text,{x:x+0.12,y:y+0.055,w:w-0.24,h:0.26,fontFace:FONT,fontSize:fs,bold:true,color,align:'center',margin:0,fit:'shrink'});
}
function metric(slide, x,y,w, label, value, suffix='', accent=C.green, note='') {
  roundedRect(slide,x,y,w,1.35,C.panel);
  slide.addText(label,{x:x+0.18,y:y+0.16,w:w-0.36,h:0.24,fontFace:FONT,fontSize:10.5,bold:true,color:C.muted,margin:0,fit:'shrink'});
  slide.addText(value,{x:x+0.18,y:y+0.45,w:w-0.36,h:0.5,fontFace:'Noto Sans',fontSize:27,bold:true,color:accent,margin:0,fit:'shrink'});
  if (suffix) slide.addText(suffix,{x:x+0.2,y:y+0.98,w:w-0.4,h:0.2,fontFace:FONT,fontSize:10,color:C.gray,margin:0,fit:'shrink'});
  if (note) slide.addText(note,{x:x+0.18,y:y+1.05,w:w-0.36,h:0.18,fontFace:FONT,fontSize:8.5,color:C.dim,margin:0,fit:'shrink'});
}
function lineText(slide, icon, title, body, x,y,w, accent=C.green) {
  slide.addShape(pptx.ShapeType.ellipse,{x,y:y+0.03,w:0.34,h:0.34,fill:{color:accent},line:{color:accent,transparency:100}});
  slide.addText(icon,{x:x,y:y+0.03,w:0.34,h:0.32,fontFace:FONT,fontSize:12,bold:true,color:C.bg,align:'center',valign:'mid',margin:0});
  slide.addText(title,{x:x+0.48,y:y,w:w-0.48,h:0.27,fontFace:FONT,fontSize:15,bold:true,color:C.text,margin:0,fit:'shrink'});
  slide.addText(body,{x:x+0.48,y:y+0.34,w:w-0.48,h:0.58,fontFace:FONT,fontSize:11.3,color:C.muted,margin:0,fit:'shrink',breakLine:false});
}
function addReferenceTag(slide, text, x, y, w) {
  slide.addText(text,{x,y,w,h:0.2,fontFace:FONT,fontSize:7.6,color:C.dim,margin:0,italic:true,fit:'shrink'});
}

// ---------- Slide 1 ----------
{
  const s = pptx.addSlide('MASTER');
  s.background = {color:C.bg};
  // abstract canopy + waterways
  for (let i=0;i<20;i++) {
    const cx=8.7+(i%5)*0.72 + ((i%2)*0.18), cy=0.9+Math.floor(i/5)*0.82;
    const r=0.42+((i*7)%5)*0.05;
    s.addShape(pptx.ShapeType.ellipse,{x:cx,y:cy,w:r*2,h:r*2,fill:{color: i%3===0?'2029B8':'0F145D',transparency:12},line:{color:C.green,transparency:86,width:0.6}});
  }
  s.addShape(pptx.ShapeType.arc,{x:8.05,y:3.1,w:4.5,h:3.2,adjustPoint:0.25,rotate:15,line:{color:'097EF6',width:10,transparency:20},fill:{color:C.bg,transparency:100}});
  s.addShape(pptx.ShapeType.arc,{x:8.55,y:3.6,w:4.1,h:2.6,adjustPoint:0.28,rotate:-8,line:{color:'0B1050',width:22,transparency:5},fill:{color:C.bg,transparency:100}});
  pill(s,'MOC2 • SCIENCE / MRV',0.72,1.28,2.2,C.panel2,C.green,10.2);
  s.addText('9.4',{x:0.72,y:1.98,w:3.0,h:1.42,fontFace:FONT,fontSize:76,bold:true,color:C.green,margin:0});
  s.addText('tCO₂e / ไร่ / ปี',{x:0.82,y:3.25,w:3.5,h:0.42,fontFace:FONT,fontSize:19,bold:true,color:C.text,margin:0});
  s.addText('ป่าของเราจะทำได้จริงไหม?',{x:0.72,y:4.02,w:6.8,h:0.68,fontFace:FONT,fontSize:33,bold:true,color:C.text,margin:0,fit:'shrink'});
  s.addText('แยกให้ออกระหว่าง “ค่าคาดการณ์จากแปลงปลูกโกงกาง” กับ “การเพิ่มขึ้นจริงของคาร์บอนในป่าธรรมชาติของโครงการ”',{x:0.72,y:4.86,w:6.7,h:1.0,fontFace:FONT,fontSize:16.5,color:C.muted,margin:0,fit:'shrink'});
  s.addText('BluFinance indigo theme • MOC2 baseline workbook + TGO / DMCR reference study',{x:0.72,y:6.35,w:6.7,h:0.28,fontFace:FONT,fontSize:9.5,color:C.dim,margin:0});
}

// ---------- Slide 2 ----------
{
  const s = pptx.addSlide('MASTER');
  addTitle(s,'Executive answer','คำตอบสั้นที่สุด','9.4 เป็นไปได้ในบางแปลงหรือบางช่วงอายุ แต่ยังไม่มีหลักฐานว่าทั้งโครงการจะรักษาค่าเฉลี่ยนี้ได้');
  roundedRect(s,0.72,2.86,11.9,2.05,C.panel);
  s.addText('ไม่ควรสัญญา 9.4 เป็นผลจริงของทั้งโครงการ',{x:1.08,y:3.22,w:7.4,h:0.55,fontFace:FONT,fontSize:28,bold:true,color:C.text,margin:0,fit:'shrink'});
  s.addText('จนกว่าจะมีการวัดซ้ำและคำนวณ ΔC/Δt ของแปลงถาวรจริง',{x:1.08,y:3.92,w:8.1,h:0.42,fontFace:FONT,fontSize:15,color:C.muted,margin:0});
  s.addShape(pptx.ShapeType.chevron,{x:9.55,y:3.18,w:1.72,h:1.22,fill:{color:C.yellow},line:{color:C.yellow,transparency:100}});
  s.addText('OPTIMISTIC\nBENCHMARK',{x:9.72,y:3.47,w:1.3,h:0.58,fontFace:FONT,fontSize:10.8,bold:true,color:C.bg,align:'center',margin:0});
  lineText(s,'1','9.4 ไม่ใช่ “โตปีต่อปี”','ต้นทางคือ Mean Annual Increment: biomass สะสม ÷ อายุแปลงปลูก',0.88,5.25,3.8,C.green);
  lineText(s,'2','ชุดอ้างอิงต่างจากเรา','Rhizophora plantation อายุทราบ 7–30 ปี; ของเราเป็นป่าธรรมชาติผสม',4.66,5.25,3.95,C.green2);
  lineText(s,'3','ต้องพิสูจน์ด้วย monitoring','เครดิตจริงขึ้นกับ stock ที่เพิ่มสุทธิ รวม recruitment และหัก mortality',8.72,5.25,3.6,C.yellow);
  addFooter(s,'Verdict: possible at some plots / periods ≠ defensible whole-project average.');
}

// ---------- Slide 3 ----------
{
  const s=pptx.addSlide('MASTER');
  addTitle(s,'Where 9.4 comes from','9.4 มาจาก “แปลงปลูกโกงกาง”','งานศึกษาของ ทช./อบก. ใช้ ๑๑๖ แปลง • ๙ จังหวัด • ๑๖ ชั้นอายุ • อายุ ๗–๓๐ ปี');
  metric(s,0.72,2.72,2.55,'แปลงที่ใช้วิเคราะห์','116','แปลง',C.green);
  metric(s,3.48,2.72,2.55,'พื้นที่ศึกษา','9','จังหวัด',C.green2);
  metric(s,6.24,2.72,2.55,'ช่วงอายุ','7–30','ปี',C.lime);
  metric(s,9.00,2.72,3.1,'ผลเฉลี่ย','9.4 ± 4.6','tCO₂e/ไร่/ปี',C.yellow);
  roundedRect(s,0.72,4.45,11.38,1.5,C.panel2);
  s.addText('เกณฑ์คัดข้อมูลสำคัญ',{x:1,y:4.75,w:2.1,h:0.3,fontFace:FONT,fontSize:11,bold:true,color:C.green,margin:0});
  s.addText('ตัดออก ๘ แปลง — เพราะสัดส่วนโกงกางต่ำกว่าร้อยละ ๘๐',{x:1,y:5.10,w:5.0,h:0.36,fontFace:FONT,fontSize:15,bold:true,color:C.text,margin:0,fit:'shrink'});
  s.addText('ดังนั้น reference population ไม่ใช่แปลงป่าธรรมชาติผสม แต่คือ Rhizophora plantation ที่ผ่านการคัดกรองแล้ว',{x:6.25,y:4.86,w:5.45,h:0.72,fontFace:FONT,fontSize:13,color:C.muted,margin:0,fit:'shrink'});
  addFooter(s,'Source: TGO — การศึกษาการกักเก็บก๊าซคาร์บอนไดออกไซด์ของไม้สกุลโกงกาง (Rhizophora spp.)');
}

// ---------- Slide 4 ----------
{
  const s=pptx.addSlide('MASTER');
  addTitle(s,'Concept','เลขเดียวกัน แต่ความหมายคนละแบบ','นี่คือเหตุผลหลักที่ 9.4 ไม่ควรถูกอ่านว่า “ป่าของเราต้องโตเพิ่ม 9.4 ทุกปี”');
  roundedRect(s,0.72,2.72,5.55,3.6,C.panel);
  roundedRect(s,6.55,2.72,5.55,3.6,C.panel);
  pill(s,'TGO REFERENCE • MAI',1.02,3.02,2.15,C.panel2,C.green,9.5);
  pill(s,'OUR CREDIT • PERIODIC CHANGE',6.86,3.02,2.65,C.panel2,C.yellow,9.5);
  s.addText('MAI = Biomass ที่อายุ a ÷ a',{x:1.02,y:3.7,w:4.7,h:0.55,fontFace:FONT,fontSize:24,bold:true,color:C.text,margin:0,fit:'shrink'});
  s.addText('ค่าเฉลี่ยตั้งแต่เริ่มปลูกจนถึงอายุปัจจุบัน',{x:1.02,y:4.44,w:4.65,h:0.52,fontFace:FONT,fontSize:14.5,color:C.muted,margin:0});
  s.addShape(pptx.ShapeType.line,{x:1.15,y:5.55,w:4.15,h:0,line:{color:C.grid,width:4,endArrowType:'triangle'}});
  s.addShape(pptx.ShapeType.ellipse,{x:1.10,y:5.43,w:0.23,h:0.23,fill:{color:C.green},line:{color:C.green,transparency:100}});
  s.addShape(pptx.ShapeType.ellipse,{x:4.98,y:5.43,w:0.23,h:0.23,fill:{color:C.green},line:{color:C.green,transparency:100}});
  s.addText('ปลูก',{x:0.95,y:5.82,w:0.6,h:0.22,fontFace:FONT,fontSize:9,color:C.dim,align:'center',margin:0});
  s.addText('อายุ a',{x:4.82,y:5.82,w:0.8,h:0.22,fontFace:FONT,fontSize:9,color:C.dim,align:'center',margin:0});
  s.addText('Increment = (Cₜ − C₀) ÷ Δt',{x:6.86,y:3.7,w:4.65,h:0.55,fontFace:FONT,fontSize:24,bold:true,color:C.text,margin:0,fit:'shrink'});
  s.addText('การเปลี่ยนแปลงจริงในช่วง monitoring',{x:6.86,y:4.44,w:4.7,h:0.52,fontFace:FONT,fontSize:14.5,color:C.muted,margin:0});
  s.addShape(pptx.ShapeType.line,{x:7.02,y:5.55,w:4.15,h:0,line:{color:C.grid,width:4,endArrowType:'triangle'}});
  s.addShape(pptx.ShapeType.ellipse,{x:7.00,y:5.43,w:0.23,h:0.23,fill:{color:C.yellow},line:{color:C.yellow,transparency:100}});
  s.addShape(pptx.ShapeType.ellipse,{x:10.87,y:5.43,w:0.23,h:0.23,fill:{color:C.yellow},line:{color:C.yellow,transparency:100}});
  s.addText('Baseline',{x:6.78,y:5.82,w:0.9,h:0.22,fontFace:FONT,fontSize:9,color:C.dim,align:'center',margin:0});
  s.addText('Monitoring',{x:10.62,y:5.82,w:1.05,h:0.22,fontFace:FONT,fontSize:9,color:C.dim,align:'center',margin:0});
  addFooter(s,'TGO formula: MAIWT = WTop+R / a; carbon = MAI × 0.47; CO₂ = carbon × 3.67');
}

// ---------- Slide 5 ----------
{
  const s=pptx.addSlide('MASTER');
  addTitle(s,'Reference variability','แม้ในแปลงปลูกอ้างอิง 9.4 ก็ไม่ใช่ค่าคงที่','ค่ารายชั้นอายุ/จังหวัดกระจายประมาณ 3.5–19.1 tCO₂e/ไร่/ปี');
  roundedRect(s,0.72,2.72,11.38,3.78,C.panel);
  const x0=1.0,y0=6.0,w=10.65,h=3.25;
  const px=a=>x0+((a-7)/(30-7))*w;
  const py=v=>y0-(v/20)*h;
  for(let i=0;i<=4;i++){
    const v=i*5, y=py(v);
    s.addShape(pptx.ShapeType.line,{x:x0,y,w,h:0,line:{color:C.grid,width:1}});
    s.addText(String(v),{x:0.78,y:y-0.11,w:0.42,h:0.22,fontFace:'Noto Sans',fontSize:8,color:C.dim,align:'right',margin:0});
  }
  [7,10,15,20,25,30].forEach(a=>{
    const x=px(a);
    s.addShape(pptx.ShapeType.line,{x,y:y0,w:0,h:-h,line:{color:C.grid,width:0.6,transparency:50}});
    s.addText(String(a),{x:x-0.22,y:6.08,w:0.44,h:0.22,fontFace:'Noto Sans',fontSize:8,color:C.dim,align:'center',margin:0});
  });
  const my=py(9.4);
  s.addShape(pptx.ShapeType.line,{x:x0,y:my,w,h:0,line:{color:C.yellow,width:2.3,dash:'dash'}});
  s.addText('ค่าเฉลี่ย 9.4',{x:9.55,y:my-0.34,w:1.6,h:0.25,fontFace:FONT,fontSize:9.5,bold:true,color:C.yellow,align:'right',margin:0});
  const pts=[[7,10.5],[7,7.4],[7,8.3],[7,11.2],[7,15.2],[8,7.5],[8,14.2],[8,7.2],[8,7.0],[9,8.1],[10,8.7],[10,11.9],[10,8.1],[10,8.0],[10,10.9],[12,19.1],[13,12.3],[13,18.3],[14,8.0],[14,7.6],[14,17.4],[15,16.0],[16,6.3],[18,6.1],[19,4.1],[21,3.5],[27,5.6],[28,6.4],[29,5.2],[29,4.2],[30,9.0]];
  pts.forEach((p,i)=>{
    const jitter=((i%5)-2)*0.028;
    s.addShape(pptx.ShapeType.ellipse,{x:px(p[0])+jitter-0.055,y:py(p[1])-0.055,w:0.11,h:0.11,fill:{color:i<15?C.green:C.green2,transparency:8},line:{color:C.text,transparency:70,width:0.4}});
  });
  s.addText('อายุแปลงปลูก (ปี)',{x:5.0,y:6.28,w:2.0,h:0.22,fontFace:FONT,fontSize:9.5,color:C.muted,align:'center',margin:0});
  s.addText('tCO₂e/ไร่/ปี',{x:0.78,y:2.72,w:0.42,h:0.22,fontFace:FONT,fontSize:8,color:C.dim,rotate:270,margin:0});
  pill(s,'3.5',10.35,2.95,0.72,C.panel2,C.red,9);
  pill(s,'19.1',11.18,2.95,0.78,C.panel2,C.lime,9);
  addFooter(s,'Source: TGO Rhizophora study, Table 3. Variation by age/site explains why a single average should not be treated as site-specific growth.');
}

// ---------- Slide 6 ----------
{
  const s=pptx.addSlide('MASTER');
  addTitle(s,'Apples vs oranges','โครงการเราไม่ใช่ประชากรเดียวกับชุด 9.4','ข้อมูลตั้งต้นของโครงการเป็นป่าธรรมชาติ/ป่าที่มีอยู่เดิม หลายชนิด และไม่ทราบอายุ');
  roundedRect(s,0.72,2.72,5.52,2.92,C.panel);
  roundedRect(s,6.58,2.72,5.52,2.92,C.panel);
  pill(s,'REFERENCE 9.4',1.0,2.98,1.75,C.panel2,C.green,9.5);
  pill(s,'MOC2 BASELINE',6.86,2.98,1.85,C.panel2,C.yellow,9.5);
  const left=[['โครงสร้าง','แปลงปลูก'],['ชนิดเด่น','Rhizophora spp. ≥80% ในชุดวิเคราะห์'],['อายุ','ทราบ 7–30 ปี'],['ตัวชี้วัด','MAI จาก stock ÷ age']];
  const right=[['โครงสร้าง','ป่าธรรมชาติ/ป่าที่มีอยู่เดิม'],['ชนิดเด่น','หลายชนิดผสม'],['อายุ','ไม่ทราบ และไม่ใช่อายุเดียวทั้ง stand'],['ตัวชี้วัด','ต้องวัด ΔC จาก monitoring']];
  function comparisonTable(items,x,accent){
    items.forEach((it,i)=>{
      const y=3.58+i*0.53;
      s.addText(it[0],{x,y,w:1.12,h:0.25,fontFace:FONT,fontSize:9.5,bold:true,color:accent,margin:0});
      s.addText(it[1],{x:x+1.22,y,w:3.65,h:0.38,fontFace:FONT,fontSize:12,color:C.text,margin:0,fit:'shrink'});
      if(i<items.length-1) s.addShape(pptx.ShapeType.line,{x,y:y+0.4,w:4.88,h:0,line:{color:C.grid,width:0.6}});
    });
  }
  comparisonTable(left,1.0,C.green);
  comparisonTable(right,6.86,C.yellow);
  s.addText('ดังนั้น 9.4 ใช้เป็น “ค่าคาดการณ์อ้างอิง” ได้ แต่ไม่ใช่หลักฐานว่าป่าของเราจะเพิ่มคาร์บอนเท่ากัน',{x:1.0,y:5.95,w:10.78,h:0.38,fontFace:FONT,fontSize:15.5,bold:true,color:C.text,align:'center',margin:0,fit:'shrink'});
  addFooter(s,'PDD describes the 6,775.53 rai project area as heterogeneous and uses stratified sampling.');
}

// ---------- Slide 7 ----------
{
  const s=pptx.addSlide('MASTER');
  addTitle(s,'Our baseline','Baseline ของเราบอกว่า “ป่าไม่เหมือนกันหมด”','32 แปลงตัวอย่าง • 6 ชั้นภูมิ (2 ฝั่งทะเล × 3 ระดับความหนาแน่นเรือนยอด) • คาร์บอนสะสมต่างกันมาก');
  roundedRect(s,0.72,2.72,8.0,3.85,C.panel);
  const data=[
    ['อ่าวไทย / มาก',54.57,23.58,74.44,C.green],
    ['อ่าวไทย / กลาง',29.39,25.25,33.67,C.green2],
    ['อ่าวไทย / น้อย',5.05,2.07,7.52,C.mint],
    ['อันดามัน / มาก',71.66,52.65,87.30,C.green],
    ['อันดามัน / กลาง',41.24,34.24,50.31,C.green2],
    ['อันดามัน / น้อย',18.91,0.51,42.22,C.red],
  ];
  const bx=3.25,bw=4.95,scale=90;
  data.forEach((d,i)=>{
    const y=3.05+i*0.52;
    s.addText(d[0],{x:1.02,y:y-0.04,w:2.0,h:0.25,fontFace:FONT,fontSize:10.3,color:C.text,margin:0});
    s.addShape(pptx.ShapeType.line,{x:bx,y:y+0.1,w:bw,h:0,line:{color:C.grid,width:1.1}});
    const x1=bx+(d[2]/scale)*bw,x2=bx+(d[3]/scale)*bw,xm=bx+(d[1]/scale)*bw;
    s.addShape(pptx.ShapeType.line,{x:x1,y:y+0.1,w:x2-x1,h:0,line:{color:d[4],width:5.4}});
    s.addShape(pptx.ShapeType.ellipse,{x:xm-0.07,y:y+0.03,w:0.14,h:0.14,fill:{color:C.text},line:{color:C.text,transparency:100}});
    s.addText(d[1].toFixed(1),{x:x2+0.08,y:y-0.04,w:0.58,h:0.25,fontFace:'Noto Sans',fontSize:9.4,bold:true,color:d[4],margin:0});
  });
  s.addText('0',{x:3.15,y:6.15,w:0.3,h:0.2,fontFace:'Noto Sans',fontSize:8,color:C.dim,align:'center',margin:0});
  s.addText('90',{x:8.02,y:6.15,w:0.35,h:0.2,fontFace:'Noto Sans',fontSize:8,color:C.dim,align:'right',margin:0});
  metric(s,9.02,2.72,3.1,'พื้นที่โครงการ','6,775.53','ไร่',C.green);
  metric(s,9.02,4.08,3.1,'Baseline tree carbon','362,962.81','tCO₂e',C.green2);
  metric(s,9.02,5.44,3.1,'เฉลี่ยถ่วงพื้นที่','53.57','tCO₂e/ไร่',C.yellow);
  addFooter(s,'Range bars show min–max across baseline plots; dot = mean. This is stock variability, not future growth variability.');
}

// ---------- Slide 8 ----------
{
  const s=pptx.addSlide('MASTER');
  addTitle(s,'Scale check','ถ้าตีความ 9.4 เป็น “เพิ่มจริงใน 1 ปี” มันใหญ่แค่ไหน?','เทียบกับคาร์บอนสะสมของต้นไม้ในปีฐาน เฉลี่ย 53.57 tCO₂e/ไร่');
  roundedRect(s,0.72,2.72,11.38,3.86,C.panel);
  s.addText('Baseline stock',{x:1.02,y:3.08,w:2.0,h:0.28,fontFace:FONT,fontSize:11,bold:true,color:C.muted,margin:0});
  s.addText('53.57',{x:1.02,y:3.48,w:2.2,h:0.65,fontFace:'Noto Sans',fontSize:40,bold:true,color:C.text,margin:0});
  s.addText('tCO₂e/ไร่',{x:1.08,y:4.15,w:1.7,h:0.26,fontFace:FONT,fontSize:11,color:C.dim,margin:0});
  s.addText('+',{x:3.16,y:3.52,w:0.55,h:0.55,fontFace:'Noto Sans',fontSize:36,bold:true,color:C.dim,align:'center',margin:0});
  s.addText('9.4',{x:3.75,y:3.48,w:1.8,h:0.65,fontFace:'Noto Sans',fontSize:40,bold:true,color:C.yellow,margin:0});
  s.addText('เพิ่มต่อปี',{x:3.82,y:4.15,w:1.2,h:0.26,fontFace:FONT,fontSize:11,color:C.dim,margin:0});
  s.addText('=',{x:5.35,y:3.52,w:0.55,h:0.55,fontFace:'Noto Sans',fontSize:36,bold:true,color:C.dim,align:'center',margin:0});
  s.addText('+17.5%',{x:5.95,y:3.42,w:2.5,h:0.72,fontFace:'Noto Sans',fontSize:42,bold:true,color:C.red,margin:0});
  s.addText('ของ stock ปัจจุบันในหนึ่งปี',{x:6.02,y:4.15,w:2.6,h:0.26,fontFace:FONT,fontSize:11,color:C.muted,margin:0});
  s.addText('นี่ไม่ใช่การบอกว่า DBH ต้องเพิ่ม 17.5%',{x:1.02,y:5.07,w:4.2,h:0.32,fontFace:FONT,fontSize:12.5,bold:true,color:C.text,margin:0});
  s.addText('เพราะ biomass–DBH เป็นสมการไม่เชิงเส้น และ stock ยังเปลี่ยนจาก recruitment / mortality ด้วย',{x:1.02,y:5.48,w:5.0,h:0.46,fontFace:FONT,fontSize:11.3,color:C.muted,margin:0});
  roundedRect(s,8.62,3.03,2.95,2.6,C.panel2);
  s.addText('INTERPRETATION',{x:8.92,y:3.3,w:2.35,h:0.24,fontFace:FONT,fontSize:9.5,bold:true,color:C.green,charSpacing:1.1,align:'center',margin:0});
  s.addText('9.4 เป็น “hurdle ที่สูง”\nสำหรับป่าที่มี stock อยู่แล้ว',{x:8.93,y:3.76,w:2.28,h:1.02,fontFace:FONT,fontSize:20,bold:true,color:C.text,align:'center',margin:0,fit:'shrink'});
  s.addText('โดยเฉพาะเมื่อ stand มีต้นแก่และ mortality',{x:8.95,y:5.05,w:2.25,h:0.38,fontFace:FONT,fontSize:9.5,color:C.muted,align:'center',margin:0});
  addFooter(s,'53.57 = weighted baseline tree-carbon stock from the 32-plot MOC2 workbook.');
}

// ---------- Slide 9 ----------
{
  const s=pptx.addSlide('MASTER');
  addTitle(s,'Biological logic','อะไรทำให้ “ป่าธรรมชาติ” ไม่โตถึง 9.4 เร็วกว่าบางแปลงปลูก','ผลจริงจะเป็น net effect ของการโต + ต้นเข้า − ต้นตาย ไม่ใช่การโตของต้นที่รอดอย่างเดียว');
  const items=[
    ['1','Stand age / saturation','ต้นที่ stock สูงแล้วมีโอกาสโตช้าลง เพราะใกล้เพดานเชิงนิเวศของพื้นที่',C.green],
    ['2','Species mix','ป่าของเราเป็นหลายชนิด ไม่ใช่แปลงปลูกโกงกางล้วน Rhizophora',C.green2],
    ['3','Mortality','ต้นตายหรือเสียหายสามารถหักล้างการเติบโตของต้นอื่นได้ โดยเฉพาะในป่าธรรมชาติ',C.red],
    ['4','Site quality','น้ำท่วม ความเค็ม ดินเลน และความหนาแน่น ทำให้ growth ต่างกัน',C.yellow]
  ];
  items.forEach((it,i)=>{
    const x=0.72+(i%2)*5.8, y=2.78+Math.floor(i/2)*1.38;
    roundedRect(s,x,y,5.42,1.06,C.panel);
    s.addShape(pptx.ShapeType.ellipse,{x:x+0.22,y:y+0.28,w:0.38,h:0.38,fill:{color:it[3]},line:{color:it[3],transparency:100}});
    s.addText(it[0],{x:x+0.22,y:y+0.35,w:0.38,h:0.2,fontFace:'Noto Sans',fontSize:8.8,bold:true,color:C.bg,align:'center',margin:0});
    s.addText(it[1],{x:x+0.78,y:y+0.22,w:4.2,h:0.28,fontFace:FONT,fontSize:12.2,bold:true,color:C.text,margin:0});
    s.addText(it[2],{x:x+0.78,y:y+0.56,w:4.25,h:0.34,fontFace:FONT,fontSize:9.8,color:C.muted,margin:0,fit:'shrink'});
  });
  roundedRect(s,0.72,6.02,11.38,0.55,C.panel2);
  s.addText('จึงไม่ควรถามแค่ว่า “DBH โตไหม?” — ต้องถามว่า “tree carbon stock ทั้ง stand เพิ่มสุทธิเท่าไรต่อปี?”',{x:1.0,y:6.17,w:10.8,h:0.25,fontFace:FONT,fontSize:13.2,bold:true,color:C.text,align:'center',margin:0,fit:'shrink'});
  addFooter(s,'The crediting number must be net stand-level stock change, not selected-tree growth.');
}

// ---------- Slide 10 ----------
{
  const s=pptx.addSlide('MASTER');
  addTitle(s,'Proof','วิธีพิสูจน์ว่าเราได้ 9.4 หรือไม่','ใช้แปลงถาวรเดิมและวัดซ้ำ—not เดาจากอายุป่า');
  const steps=[
    ['01','กลับไปแปลงเดิม','ใช้ permanent plots / tags เดิม และขอบเขตแปลงเดิม'],
    ['02','วัด DBH ซ้ำ','ให้ protocol เดิม + คนวัดเทียบกันก่อนลงสนาม'],
    ['03','เก็บ recruitment + mortality','ต้นเข้าใหม่และต้นตายต้องอยู่ใน net change'],
    ['04','คำนวณ Cₜ และ ΔC','allometry เดิม → stock รอบใหม่ → ลบ baseline'],
    ['05','ถ่วงตาม strata + uncertainty','รวม 6 ชั้นภูมิ พร้อมรายงาน precision / uncertainty'],
    ['06','เทียบกับ 9.4','ถ้า weighted ΔC/Δt ≥ 9.4 จึงถือว่าพิสูจน์ได้']
  ];
  steps.forEach((st,i)=>{
    const y=2.72+i*0.6;
    s.addText(st[0],{x:0.82,y:y+0.03,w:0.45,h:0.25,fontFace:'Noto Sans',fontSize:9.5,bold:true,color:C.green,margin:0});
    s.addShape(pptx.ShapeType.line,{x:1.28,y:y+0.17,w:0.42,h:0,line:{color:C.green,width:1.8}});
    s.addText(st[1],{x:1.84,y,w:2.65,h:0.28,fontFace:FONT,fontSize:13,bold:true,color:C.text,margin:0});
    s.addText(st[2],{x:4.62,y,w:6.72,h:0.34,fontFace:FONT,fontSize:11.1,color:C.muted,margin:0,fit:'shrink'});
  });
  roundedRect(s,8.72,6.16,3.38,0.62,C.panel2);
  s.addText('PASS CONDITION',{x:8.95,y:6.29,w:1.2,h:0.22,fontFace:FONT,fontSize:9,bold:true,color:C.green,margin:0});
  s.addText('Weighted ΔC/Δt ≥ 9.4',{x:10.04,y:6.24,w:1.78,h:0.28,fontFace:'Noto Sans',fontSize:12.3,bold:true,color:C.text,align:'right',margin:0,fit:'shrink'});
  addFooter(s,'Use the monitoring protocol and methodology version approved for the registered project.');
}

// ---------- Slide 11 ----------
{
  const s=pptx.addSlide('MASTER');
  addTitle(s,'Decision','ดังนั้น “จะได้ 9.4 ไหม?”','คำตอบที่ defend ได้ต่อ auditor / verifier วันนี้');
  roundedRect(s,0.72,2.72,11.38,1.18,C.panel);
  s.addText('ยังตอบว่า “ได้แน่” ไม่ได้',{x:1.04,y:3.03,w:4.2,h:0.5,fontFace:FONT,fontSize:30,bold:true,color:C.red,margin:0});
  s.addText('และยังตอบว่า “เป็นไปไม่ได้” ก็เร็วเกินไป',{x:5.48,y:3.09,w:5.98,h:0.4,fontFace:FONT,fontSize:18,bold:true,color:C.text,align:'right',margin:0});
  const cards=[
    ['ใช้ 9.4 เพื่ออะไร?','Ex-ante benchmark','ใช้เป็นค่าคาดการณ์อ้างอิงได้ แต่ต้อง label ว่าไม่ใช่ measured result',C.green],
    ['ความน่าจะเป็นเชิงวิทยาศาสตร์','Optimistic for whole project','baseline เป็นป่าธรรมชาติผสม มี stock อยู่แล้ว และ heterogeneous',C.yellow],
    ['ตัวเลขที่จะใช้เครดิตจริง','Measured ΔC wins','ผลวัดซ้ำ + recruitment − mortality + uncertainty ต้องเป็นตัวตัดสิน',C.green2]
  ];
  cards.forEach((c,i)=>{
    const x=0.72+i*3.94;
    roundedRect(s,x,4.22,3.5,2.05,C.panel);
    s.addText(c[0],{x:x+0.2,y:4.44,w:3.1,h:0.25,fontFace:FONT,fontSize:9.4,bold:true,color:c[3],margin:0});
    s.addText(c[1],{x:x+0.2,y:4.82,w:3.08,h:0.46,fontFace:FONT,fontSize:17,bold:true,color:C.text,margin:0,fit:'shrink'});
    s.addText(c[2],{x:x+0.2,y:5.38,w:3.08,h:0.62,fontFace:FONT,fontSize:10.3,color:C.muted,margin:0,fit:'shrink'});
  });
  addFooter(s,'Recommendation wording: “9.4 is a reference benchmark; project performance will be determined by monitoring-based stock change.”');
}

// ---------- Slide 12 ----------
{
  const s=pptx.addSlide('MASTER');
  addTitle(s,'Sources & caveats','แหล่งข้อมูลและข้อจำกัด','สไลด์นี้ตอบ “ต้องอ้างอิงจากเอกสาร” ออกจาก “การตีความเชิงวิทยาศาสตร์ของโครงการ”');
  const sources=[
    ['TGO / DMCR study','การศึกษาการกักเก็บคาร์บอนของไม้สกุลโกงกาง (Rhizophora spp.)','116 plots; 9 provinces; MAI formula = biomass / age'],
    ['DMCR announcement','แจ้งรับทราบศักยภาพ 9.40 และอ้างเป็นค่า “คาดการณ์” สำหรับกิจกรรมปลูกป่าชายเลน','issued 1 Mar 2023'],
    ['MOC2 baseline workbook','MOC2_VSD-25-06-24 3.xlsx','32 plots; 6 strata; baseline stock'],
    ['MOC2 PDD','T-VER-S-F001-PDD V2.1-MOC2-25-06-24 1.pdf','Project area described as heterogeneous; stratified sampling; CTTt monitored under approved tool']
  ];
  sources.forEach((d,i)=>{
    const y=2.72+i*0.85;
    s.addText(String(i+1).padStart(2,'0'),{x:0.78,y:y+0.08,w:0.4,h:0.25,fontFace:FONT,fontSize:9.5,bold:true,color:C.green,margin:0});
    s.addText(d[0],{x:1.35,y,w:2.3,h:0.28,fontFace:FONT,fontSize:13,bold:true,color:C.text,margin:0,fit:'shrink'});
    s.addText(d[1],{x:3.72,y,w:5.6,h:0.42,fontFace:FONT,fontSize:11.2,color:C.gray,margin:0,fit:'shrink'});
    s.addText(d[2],{x:9.52,y,w:2.25,h:0.42,fontFace:FONT,fontSize:9.5,color:C.muted,margin:0,fit:'shrink'});
    if(i<3) s.addShape(pptx.ShapeType.line,{x:1.35,y:y+0.64,w:10.45,h:0,line:{color:C.grid,width:0.6}});
  });
  roundedRect(s,0.72,6.30,11.38,0.52,C.panel2);
  s.addText('ข้อจำกัด: baseline เป็น snapshot ก่อนเริ่มโครงการ จึงยังไม่สามารถสรุป actual annual increment ได้จนกว่าจะมีข้อมูลวัดซ้ำ',{x:0.98,y:6.43,w:10.86,h:0.24,fontFace:FONT,fontSize:11.7,bold:true,color:C.yellow,margin:0,align:'center',fit:'shrink'});
  addFooter(s,'Prepared 2026-10-05 • For technical discussion; not a verification opinion or credit issuance decision.');
}

for (const slide of pptx._slides) {
  warnIfSlideHasOverlaps(slide, pptx, { ignoreLines: true, ignoreDecorativeShapes: true });
  warnIfSlideElementsOutOfBounds(slide, pptx);
}
pptx.writeFile({ fileName: '/mnt/data/ppt-carbon/output/MOC2_9_4_carbon_assessment.pptx' });
