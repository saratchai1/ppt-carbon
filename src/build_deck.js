const pptxgen=require('pptxgenjs');
const pptx=new pptxgen();
pptx.layout='LAYOUT_WIDE';
pptx.author='OpenAI / ChatGPT';
pptx.title='9.4 tCO₂e/ไร่/ปี — ป่าของเราจะทำได้จริงไหม?';
pptx.subject='MOC2 mangrove carbon assessment';
pptx.lang='th-TH';
pptx.theme={headFontFace:'Noto Sans Thai',bodyFontFace:'Noto Sans Thai',lang:'th-TH'};
const C={bg:'07110E',panel:'10201A',panel2:'132A22',text:'F3FAF6',muted:'A9BBB3',dim:'6F837A',green:'4FD1A1',teal:'2DD4BF',lime:'B7F34B',yellow:'F2C14E',red:'FF6B6B',grid:'294239'};
const F='Noto Sans Thai';
pptx.defineSlideMaster({title:'M',background:{color:C.bg},objects:[
 {line:{x:.55,y:.46,w:.36,h:0,line:{color:C.green,width:2.5}}},
 {text:{text:'MOC2 • CARBON SCIENCE NOTE',options:{x:.98,y:.30,w:3.4,h:.28,fontFace:F,fontSize:9.5,color:'8FA59B',bold:true,charSpacing:1.2,margin:0}}}
],slideNumber:{x:12.25,y:7.05,w:.35,h:.2,color:C.dim,fontFace:F,fontSize:9,align:'right',margin:0}});
function rr(s,x,y,w,h,fill=C.panel){s.addShape(pptx.ShapeType.roundRect,{x,y,w,h,fill:{color:fill},line:{color:fill,transparency:100}})}
function title(s,k,t,sub=''){s.addText(k.toUpperCase(),{x:.72,y:.85,w:4.8,h:.25,fontFace:F,fontSize:10,bold:true,color:C.green,charSpacing:1.2,margin:0});s.addText(t,{x:.72,y:1.18,w:11.8,h:.78,fontFace:F,fontSize:31,bold:true,color:C.text,margin:0,fit:'shrink'});if(sub)s.addText(sub,{x:.72,y:2.02,w:11.7,h:.55,fontFace:F,fontSize:15,color:C.muted,margin:0,fit:'shrink'})}
function foot(s,t){s.addText(t,{x:.72,y:6.93,w:10.8,h:.22,fontFace:F,fontSize:8.3,color:C.dim,margin:0,fit:'shrink'})}
function pill(s,t,x,y,w,c=C.green){rr(s,x,y,w,.42,C.panel2);s.addText(t,{x:x+.1,y:y+.06,w:w-.2,h:.24,fontFace:F,fontSize:9.5,bold:true,color:c,align:'center',margin:0,fit:'shrink'})}
function metric(s,x,y,w,l,v,u,c=C.green){rr(s,x,y,w,1.35);s.addText(l,{x:x+.18,y:y+.16,w:w-.36,h:.24,fontFace:F,fontSize:10.3,bold:true,color:C.muted,margin:0});s.addText(v,{x:x+.18,y:y+.44,w:w-.36,h:.5,fontFace:'Noto Sans',fontSize:27,bold:true,color:c,margin:0,fit:'shrink'});s.addText(u,{x:x+.2,y:y+.98,w:w-.4,h:.2,fontFace:F,fontSize:10,color:'D5E0DB',margin:0,fit:'shrink'})}
function bullet(s,n,h,b,x,y,w,c=C.green){s.addShape(pptx.ShapeType.ellipse,{x,y:y+.03,w:.34,h:.34,fill:{color:c},line:{color:c,transparency:100}});s.addText(n,{x,y:y+.03,w:.34,h:.32,fontFace:F,fontSize:11,bold:true,color:C.bg,align:'center',margin:0});s.addText(h,{x:x+.48,y,w:w-.48,h:.27,fontFace:F,fontSize:14.2,bold:true,color:C.text,margin:0,fit:'shrink'});s.addText(b,{x:x+.48,y:y+.34,w:w-.48,h:.58,fontFace:F,fontSize:10.9,color:C.muted,margin:0,fit:'shrink'})}

// 1 cover
{
 const s=pptx.addSlide('M');
 for(let i=0;i<20;i++){const cx=8.7+(i%5)*.72+((i%2)*.18),cy=.9+Math.floor(i/5)*.82,r=.42+((i*7)%5)*.05;s.addShape(pptx.ShapeType.ellipse,{x:cx,y:cy,w:r*2,h:r*2,fill:{color:i%3===0?'1E5B45':'174532',transparency:12},line:{color:C.green,transparency:86,width:.6}})}
 s.addShape(pptx.ShapeType.arc,{x:8.05,y:3.1,w:4.5,h:3.2,rotate:15,line:{color:'2B8E6B',width:10,transparency:20},fill:{color:C.bg,transparency:100}});
 pill(s,'MOC2 • SCIENCE / MRV',.72,1.28,2.2,C.green);
 s.addText('9.4',{x:.72,y:1.98,w:3,h:1.42,fontFace:F,fontSize:76,bold:true,color:C.green,margin:0});
 s.addText('tCO₂e / ไร่ / ปี',{x:.82,y:3.25,w:3.5,h:.42,fontFace:F,fontSize:19,bold:true,color:C.text,margin:0});
 s.addText('ป่าของเราจะทำได้จริงไหม?',{x:.72,y:4.02,w:6.8,h:.68,fontFace:F,fontSize:33,bold:true,color:C.text,margin:0,fit:'shrink'});
 s.addText('แยก “ค่าคาดการณ์จากแปลงปลูกโกงกาง” ออกจาก “การเพิ่มขึ้นจริงของคาร์บอนในป่าธรรมชาติของโครงการ”',{x:.72,y:4.86,w:6.8,h:1,fontFace:F,fontSize:16.5,color:C.muted,margin:0,fit:'shrink'});
 s.addText('MOC2 baseline workbook + TGO / DMCR reference study',{x:.72,y:6.35,w:6.7,h:.28,fontFace:F,fontSize:9.5,color:C.dim,margin:0});
}
// 2 answer
{
 const s=pptx.addSlide('M');title(s,'Executive answer','คำตอบสั้นที่สุด','9.4 เป็นไปได้ในบางแปลงหรือบางช่วงอายุ แต่ยังไม่มีหลักฐานว่าทั้งโครงการจะรักษาค่าเฉลี่ยนี้ได้');
 rr(s,.72,2.86,11.9,2.05);
 s.addText('ไม่ควรสัญญา 9.4 เป็นผลจริงของทั้งโครงการ',{x:1.08,y:3.22,w:7.4,h:.55,fontFace:F,fontSize:28,bold:true,color:C.text,margin:0,fit:'shrink'});
 s.addText('จนกว่าจะมีการวัดซ้ำและคำนวณ ΔC/Δt ของแปลงถาวรจริง',{x:1.08,y:3.92,w:8.1,h:.42,fontFace:F,fontSize:15,color:C.muted,margin:0});
 s.addShape(pptx.ShapeType.chevron,{x:9.52,y:3.2,w:1.65,h:1.18,fill:{color:C.yellow},line:{color:C.yellow,transparency:100}});
 s.addText('OPTIMISTIC\nBENCHMARK',{x:9.68,y:3.47,w:1.25,h:.58,fontFace:F,fontSize:10.5,bold:true,color:C.bg,align:'center',margin:0});
 bullet(s,'1','9.4 ไม่ใช่ “โตปีต่อปี”','ต้นทางคือ Mean Annual Increment: biomass สะสม ÷ อายุแปลงปลูก',.88,5.25,3.8,C.green);
 bullet(s,'2','ชุดอ้างอิงต่างจากเรา','Rhizophora plantation อายุทราบ 7–30 ปี; ของเราเป็นป่าธรรมชาติผสม',4.66,5.25,3.95,C.teal);
 bullet(s,'3','ต้องพิสูจน์ด้วย monitoring','เครดิตจริงขึ้นกับ stock ที่เพิ่มสุทธิ รวม recruitment และหัก mortality',8.72,5.25,3.6,C.yellow);
 foot(s,'Verdict: possible at some plots / periods ≠ defensible whole-project average.');
}
// 3 origin
{
 const s=pptx.addSlide('M');title(s,'Where 9.4 comes from','9.4 มาจาก “แปลงปลูกโกงกาง”','งานศึกษาของ ทช./อบก. ใช้ ๑๑๖ แปลง • ๙ จังหวัด • ๑๖ ชั้นอายุ • อายุ ๗–๓๐ ปี');
 metric(s,.72,2.72,2.55,'แปลงที่ใช้วิเคราะห์','116','แปลง',C.green);metric(s,3.48,2.72,2.55,'พื้นที่ศึกษา','9','จังหวัด',C.teal);metric(s,6.24,2.72,2.55,'ช่วงอายุ','7–30','ปี',C.lime);metric(s,9,2.72,3.1,'ผลเฉลี่ย','9.4 ± 4.6','tCO₂e/ไร่/ปี',C.yellow);
 rr(s,.72,4.45,11.38,1.5,C.panel2);
 s.addText('เกณฑ์คัดข้อมูลสำคัญ',{x:1,y:4.75,w:2.1,h:.3,fontFace:F,fontSize:11,bold:true,color:C.green,margin:0});
 s.addText('ตัดออก ๘ แปลง — เพราะสัดส่วนโกงกางต่ำกว่าร้อยละ ๘๐',{x:1,y:5.10,w:5,h:.36,fontFace:F,fontSize:15,bold:true,color:C.text,margin:0,fit:'shrink'});
 s.addText('reference population จึงเข้มข้นไปทาง Rhizophora plantation มากกว่าป่าธรรมชาติผสมของเรา',{x:6.25,y:4.86,w:5.45,h:.72,fontFace:F,fontSize:13,color:C.muted,margin:0,fit:'shrink'});
 foot(s,'Source: TGO — การศึกษาการกักเก็บก๊าซคาร์บอนไดออกไซด์ของไม้สกุลโกงกาง (Rhizophora spp.)');
}
// 4 concept
{
 const s=pptx.addSlide('M');title(s,'Concept','เลขเดียวกัน แต่ความหมายคนละแบบ','นี่คือเหตุผลหลักที่ 9.4 ไม่ควรถูกอ่านว่า “ป่าของเราต้องโตเพิ่ม 9.4 ทุกปี”');
 rr(s,.72,2.72,5.55,3.6);rr(s,6.55,2.72,5.55,3.6);pill(s,'TGO REFERENCE • MAI',1.02,3.02,2.15,C.green);pill(s,'OUR CREDIT • PERIODIC CHANGE',6.86,3.02,2.65,C.yellow);
 s.addText('MAI = Biomass ที่อายุ a ÷ a',{x:1.02,y:3.7,w:4.7,h:.55,fontFace:F,fontSize:24,bold:true,color:C.text,margin:0,fit:'shrink'});
 s.addText('ค่าเฉลี่ยตั้งแต่เริ่มปลูกจนถึงอายุปัจจุบัน',{x:1.02,y:4.44,w:4.65,h:.52,fontFace:F,fontSize:14.5,color:C.muted,margin:0});
 s.addShape(pptx.ShapeType.line,{x:1.15,y:5.55,w:4.15,h:0,line:{color:C.grid,width:4,endArrowType:'triangle'}});s.addShape(pptx.ShapeType.ellipse,{x:1.1,y:5.43,w:.23,h:.23,fill:{color:C.green},line:{color:C.green,transparency:100}});s.addShape(pptx.ShapeType.ellipse,{x:4.98,y:5.43,w:.23,h:.23,fill:{color:C.green},line:{color:C.green,transparency:100}});
 s.addText('ปลูก',{x:.95,y:5.82,w:.6,h:.22,fontFace:F,fontSize:9,color:C.dim,align:'center',margin:0});s.addText('อายุ a',{x:4.82,y:5.82,w:.8,h:.22,fontFace:F,fontSize:9,color:C.dim,align:'center',margin:0});
 s.addText('Increment = (Cₜ − C₀) ÷ Δt',{x:6.86,y:3.7,w:4.65,h:.55,fontFace:F,fontSize:24,bold:true,color:C.text,margin:0,fit:'shrink'});
 s.addText('การเปลี่ยนแปลงจริงในช่วง monitoring',{x:6.86,y:4.44,w:4.7,h:.52,fontFace:F,fontSize:14.5,color:C.muted,margin:0});
 s.addShape(pptx.ShapeType.line,{x:7.02,y:5.55,w:4.15,h:0,line:{color:C.grid,width:4,endArrowType:'triangle'}});s.addShape(pptx.ShapeType.ellipse,{x:7,y:5.43,w:.23,h:.23,fill:{color:C.yellow},line:{color:C.yellow,transparency:100}});s.addShape(pptx.ShapeType.ellipse,{x:10.87,y:5.43,w:.23,h:.23,fill:{color:C.yellow},line:{color:C.yellow,transparency:100}});
 s.addText('Baseline',{x:6.78,y:5.82,w:.9,h:.22,fontFace:F,fontSize:9,color:C.dim,align:'center',margin:0});s.addText('Monitoring',{x:10.62,y:5.82,w:1.05,h:.22,fontFace:F,fontSize:9,color:C.dim,align:'center',margin:0});
 foot(s,'TGO formula: MAIWT = WTop+R / a; carbon = MAI × 0.47; CO₂ = carbon × 3.67');
}
// 5 reference variability
{
 const s=pptx.addSlide('M');title(s,'Reference variability','แม้ในแปลงปลูกอ้างอิง 9.4 ก็ไม่ใช่ค่าคงที่','ค่ารายชั้นอายุ/จังหวัดกระจายประมาณ 3.5–19.1 tCO₂e/ไร่/ปี');
 const pts=[[7,10.5],[7,7.4],[7,8.3],[7,11.2],[7,15.2],[8,7.5],[8,14.2],[8,7.2],[8,7],[9,8.1],[10,8.7],[10,11.9],[10,8.1],[10,8],[10,10.9],[12,19.1],[13,12.3],[13,18.3],[14,8],[14,7.6],[14,17.4],[15,16],[16,6.3],[18,6.1],[19,4.1],[21,3.5],[27,5.6],[28,6.4],[29,5.2],[29,4.2],[30,9]];
 rr(s,.72,2.72,11.38,3.78);const x0=1,y0=6,w=10.65,h=3.25,px=a=>x0+((a-7)/23)*w,py=v=>y0-(v/20)*h;
 [0,5,10,15,20].forEach(v=>{const y=py(v);s.addShape(pptx.ShapeType.line,{x:x0,y,w,h:0,line:{color:C.grid,width:1}});s.addText(String(v),{x:.78,y:y-.11,w:.42,h:.22,fontFace:'Noto Sans',fontSize:8,color:C.dim,align:'right',margin:0})});
 [7,10,15,20,25,30].forEach(a=>{const x=px(a);s.addShape(pptx.ShapeType.line,{x,y:y0,w:0,h:-h,line:{color:C.grid,width:.6,transparency:50}});s.addText(String(a),{x:x-.22,y:6.08,w:.44,h:.22,fontFace:'Noto Sans',fontSize:8,color:C.dim,align:'center',margin:0})});
 const my=py(9.4);s.addShape(pptx.ShapeType.line,{x:x0,y:my,w,h:0,line:{color:C.yellow,width:2.3,dash:'dash'}});s.addText('ค่าเฉลี่ย 9.4',{x:9.55,y:my-.34,w:1.6,h:.25,fontFace:F,fontSize:9.5,bold:true,color:C.yellow,align:'right',margin:0});
 pts.forEach((p,i)=>{const j=((i%5)-2)*.028;s.addShape(pptx.ShapeType.ellipse,{x:px(p[0])+j-.055,y:py(p[1])-.055,w:.11,h:.11,fill:{color:i<15?C.green:C.teal,transparency:8},line:{color:C.text,transparency:70,width:.4}})});
 s.addText('อายุแปลงปลูก (ปี)',{x:5,y:6.28,w:2,h:.22,fontFace:F,fontSize:9.5,color:C.muted,align:'center',margin:0});
 pill(s,'3.5',10.35,2.95,.72,C.red);pill(s,'19.1',11.18,2.95,.78,C.lime);
 foot(s,'Source: TGO Rhizophora study, Table 3. Age and site quality produce a wide range around the mean.');
}
// 6 compare
{
 const s=pptx.addSlide('M');title(s,'Apples vs oranges','โครงการเราไม่ใช่ประชากรเดียวกับชุด 9.4','ข้อมูลตั้งต้นของโครงการเป็นป่าธรรมชาติ/ป่าที่มีอยู่เดิม หลายชนิด และไม่ทราบอายุ');
 rr(s,.72,2.72,5.52,2.92);rr(s,6.58,2.72,5.52,2.92);pill(s,'REFERENCE 9.4',1,2.98,1.75,C.green);pill(s,'MOC2 BASELINE',6.86,2.98,1.85,C.yellow);
 const L=[['โครงสร้าง','แปลงปลูก'],['ชนิดเด่น','Rhizophora spp. ≥80%'],['อายุ','ทราบ 7–30 ปี'],['ตัวชี้วัด','MAI = stock ÷ age']],R=[['โครงสร้าง','ป่าธรรมชาติ/ป่าที่มีอยู่เดิม'],['ชนิดเด่น','หลายชนิดผสม'],['อายุ','ไม่ทราบ / ไม่ใช่อายุเดียวทั้ง stand'],['ตัวชี้วัด','ต้องใช้ ΔC จาก monitoring']];
 function list(a,x,c){a.forEach((it,i)=>{const y=3.58+i*.53;s.addText(it[0],{x,y,w:1.12,h:.25,fontFace:F,fontSize:9.5,bold:true,color:c,margin:0});s.addText(it[1],{x:x+1.22,y,w:3.65,h:.38,fontFace:F,fontSize:12,color:C.text,margin:0,fit:'shrink'});if(i<3)s.addShape(pptx.ShapeType.line,{x,y:y+.4,w:4.88,h:0,line:{color:C.grid,width:.6}})})}
 list(L,1,C.green);list(R,6.86,C.yellow);
 s.addText('ดังนั้น 9.4 ใช้เป็น “ค่าคาดการณ์อ้างอิง” ได้ แต่ไม่ใช่หลักฐานว่าป่าของเราจะเพิ่มคาร์บอนเท่ากัน',{x:1,y:5.95,w:10.78,h:.38,fontFace:F,fontSize:15.5,bold:true,color:C.text,align:'center',margin:0,fit:'shrink'});
 foot(s,'PDD describes the 6,775.53 rai project area as heterogeneous and uses stratified sampling.');
}
// 7 baseline
{
 const s=pptx.addSlide('M');title(s,'Our baseline','Baseline ของเราบอกว่า “ป่าไม่เหมือนกันหมด”','32 แปลงตัวอย่าง • 6 ชั้นภูมิ • คาร์บอนสะสมต่างกันมากระหว่างกลุ่มและภายในกลุ่ม');
 const d=[['อ่าวไทย / มาก',54.57,23.58,74.44,C.green],['อ่าวไทย / กลาง',29.39,25.25,33.67,C.teal],['อ่าวไทย / น้อย',5.05,2.07,7.52,C.lime],['อันดามัน / มาก',71.66,52.65,87.3,C.green],['อันดามัน / กลาง',41.24,34.24,50.31,C.teal],['อันดามัน / น้อย',18.91,.51,42.22,C.red]];
 rr(s,.72,2.72,8,3.85);const x0=3.25,w=4.95,sc=90;
 d.forEach((q,i)=>{const y=3.05+i*.52;s.addText(q[0],{x:1.02,y:y-.04,w:2,h:.25,fontFace:F,fontSize:10.3,color:C.text,margin:0});s.addShape(pptx.ShapeType.line,{x:x0,y:y+.1,w,h:0,line:{color:C.grid,width:1.1}});const a=x0+q[2]/sc*w,b=x0+q[3]/sc*w,m=x0+q[1]/sc*w;s.addShape(pptx.ShapeType.line,{x:a,y:y+.1,w:b-a,h:0,line:{color:q[4],width:5.4}});s.addShape(pptx.ShapeType.ellipse,{x:m-.07,y:y+.03,w:.14,h:.14,fill:{color:C.text},line:{color:C.text,transparency:100}});s.addText(q[1].toFixed(1),{x:b+.08,y:y-.04,w:.58,h:.25,fontFace:'Noto Sans',fontSize:9.4,bold:true,color:q[4],margin:0})});
 s.addText('0',{x:3.15,y:6.15,w:.3,h:.2,fontFace:'Noto Sans',fontSize:8,color:C.dim,align:'center',margin:0});s.addText('90',{x:8.02,y:6.15,w:.35,h:.2,fontFace:'Noto Sans',fontSize:8,color:C.dim,align:'right',margin:0});
 metric(s,9.02,2.72,3.1,'พื้นที่โครงการ','6,775.53','ไร่',C.green);metric(s,9.02,4.08,3.1,'Baseline tree carbon','362,962.81','tCO₂e',C.teal);metric(s,9.02,5.44,3.1,'เฉลี่ยถ่วงพื้นที่','53.57','tCO₂e/ไร่',C.yellow);
 foot(s,'Range = min–max across baseline plots; dot = mean. This is stock variability, not future growth variability.');
}
// 8 scale
{
 const s=pptx.addSlide('M');title(s,'Scale check','ถ้าตีความ 9.4 เป็น “เพิ่มจริงใน 1 ปี” มันใหญ่แค่ไหน?','เทียบกับคาร์บอนสะสมของต้นไม้ในปีฐาน เฉลี่ย 53.57 tCO₂e/ไร่');
 rr(s,.72,2.72,11.38,3.86);
 s.addText('Baseline stock',{x:1.02,y:3.08,w:2,h:.28,fontFace:F,fontSize:11,bold:true,color:C.muted,margin:0});s.addText('53.57',{x:1.02,y:3.48,w:2.2,h:.65,fontFace:'Noto Sans',fontSize:40,bold:true,color:C.text,margin:0});s.addText('tCO₂e/ไร่',{x:1.08,y:4.15,w:1.7,h:.26,fontFace:F,fontSize:11,color:C.dim,margin:0});
 s.addText('+',{x:3.16,y:3.52,w:.55,h:.55,fontFace:'Noto Sans',fontSize:36,bold:true,color:C.dim,align:'center',margin:0});s.addText('9.4',{x:3.75,y:3.48,w:1.8,h:.65,fontFace:'Noto Sans',fontSize:40,bold:true,color:C.yellow,margin:0});s.addText('เพิ่มต่อปี',{x:3.82,y:4.15,w:1.2,h:.26,fontFace:F,fontSize:11,color:C.dim,margin:0});
 s.addText('=',{x:5.35,y:3.52,w:.55,h:.55,fontFace:'Noto Sans',fontSize:36,bold:true,color:C.dim,align:'center',margin:0});s.addText('+17.5%',{x:5.95,y:3.42,w:2.5,h:.72,fontFace:'Noto Sans',fontSize:42,bold:true,color:C.red,margin:0});s.addText('ของ stock ปัจจุบันในหนึ่งปี',{x:6.02,y:4.15,w:2.6,h:.26,fontFace:F,fontSize:11,color:C.muted,margin:0});
 s.addText('นี่ไม่ใช่การบอกว่า DBH ต้องเพิ่ม 17.5%',{x:1.02,y:5.07,w:4.2,h:.32,fontFace:F,fontSize:12.5,bold:true,color:C.text,margin:0});s.addText('biomass–DBH ไม่เชิงเส้น และ stock ยังเปลี่ยนจาก recruitment / mortality',{x:1.02,y:5.48,w:5,h:.46,fontFace:F,fontSize:11.3,color:C.muted,margin:0});
 rr(s,8.62,3.03,2.95,2.6,C.panel2);s.addText('INTERPRETATION',{x:8.92,y:3.3,w:2.35,h:.24,fontFace:F,fontSize:9.5,bold:true,color:C.green,charSpacing:1.1,align:'center',margin:0});s.addText('9.4 เป็น “hurdle ที่สูง”\nสำหรับป่าที่มี stock อยู่แล้ว',{x:8.93,y:3.76,w:2.28,h:1.02,fontFace:F,fontSize:20,bold:true,color:C.text,align:'center',margin:0,fit:'shrink'});s.addText('โดยเฉพาะเมื่อมีต้นแก่และ mortality',{x:8.95,y:5.05,w:2.25,h:.38,fontFace:F,fontSize:9.5,color:C.muted,align:'center',margin:0});
 foot(s,'53.57 = weighted baseline tree-carbon stock from the 32-plot MOC2 workbook.');
}
// 9 proof
{
 const s=pptx.addSlide('M');title(s,'Proof','วิธีพิสูจน์ว่าเราได้ 9.4 หรือไม่','ใช้แปลงถาวรเดิมและวัดซ้ำ—not เดาจากอายุป่า');
 const a=[['01','กลับไปแปลงเดิม','ใช้ permanent plots / tags เดิม'],['02','วัด DBH ซ้ำ','protocol และเกณฑ์เดียวกับ baseline'],['03','เก็บ recruitment + mortality','ต้นเข้าใหม่และต้นตายต้องอยู่ใน net change'],['04','คำนวณ Cₜ และ ΔC','allometry เดิม → stock รอบใหม่ → ลบ baseline'],['05','ถ่วงตาม strata + uncertainty','รวม 6 ชั้นภูมิและรายงาน precision'],['06','เทียบกับ 9.4','weighted ΔC/Δt ≥ 9.4 จึงถือว่าพิสูจน์ได้']];
 a.forEach((q,i)=>{const y=2.72+i*.6;s.addText(q[0],{x:.82,y:y+.03,w:.45,h:.25,fontFace:'Noto Sans',fontSize:9.5,bold:true,color:C.green,margin:0});s.addShape(pptx.ShapeType.line,{x:1.28,y:y+.17,w:.42,h:0,line:{color:C.green,width:1.8}});s.addText(q[1],{x:1.84,y,w:2.65,h:.28,fontFace:F,fontSize:13,bold:true,color:C.text,margin:0});s.addText(q[2],{x:4.62,y,w:6.72,h:.34,fontFace:F,fontSize:11.1,color:C.muted,margin:0,fit:'shrink'})});
 rr(s,8.72,6.16,3.38,.62,C.panel2);s.addText('PASS CONDITION',{x:8.95,y:6.29,w:1.2,h:.22,fontFace:F,fontSize:9,bold:true,color:C.green,margin:0});s.addText('Weighted ΔC/Δt ≥ 9.4',{x:10.04,y:6.24,w:1.78,h:.28,fontFace:'Noto Sans',fontSize:12.3,bold:true,color:C.text,align:'right',margin:0,fit:'shrink'});
 foot(s,'Use the monitoring protocol and methodology version approved for the registered project.');
}
// 10 decision + sources
{
 const s=pptx.addSlide('M');title(s,'Decision','ดังนั้น “จะได้ 9.4 ไหม?”','คำตอบที่ defend ได้ต่อ auditor / verifier วันนี้');
 rr(s,.72,2.72,11.38,1.18);s.addText('ยังตอบว่า “ได้แน่” ไม่ได้',{x:1.04,y:3.03,w:4.2,h:.5,fontFace:F,fontSize:30,bold:true,color:C.red,margin:0});s.addText('และยังตอบว่า “เป็นไปไม่ได้” ก็เร็วเกินไป',{x:5.48,y:3.09,w:5.98,h:.4,fontFace:F,fontSize:18,bold:true,color:C.text,align:'right',margin:0});
 const boxes=[[.72,'ใช้ 9.4 เพื่ออะไร?','Ex-ante benchmark','ใช้เป็นค่าคาดการณ์อ้างอิงได้ แต่ต้อง label ว่าไม่ใช่ measured result',C.green],[4.66,'ความน่าจะเป็นเชิงวิทยาศาสตร์','Optimistic for whole project','baseline เป็นป่าธรรมชาติผสม มี stock อยู่แล้ว และ heterogeneous',C.yellow],[8.6,'ตัวเลขที่จะใช้เครดิตจริง','Measured ΔC wins','ผลวัดซ้ำ + recruitment − mortality + uncertainty ต้องเป็นตัวตัดสิน',C.teal]];
 boxes.forEach(q=>{rr(s,q[0],4.22,3.5,2.05);s.addText(q[1],{x:q[0]+.2,y:4.44,w:3.1,h:.25,fontFace:F,fontSize:9.4,bold:true,color:q[4],margin:0});s.addText(q[2],{x:q[0]+.2,y:4.82,w:3.08,h:.46,fontFace:F,fontSize:17,bold:true,color:C.text,margin:0,fit:'shrink'});s.addText(q[3],{x:q[0]+.2,y:5.38,w:3.08,h:.62,fontFace:F,fontSize:10.3,color:C.muted,margin:0,fit:'shrink'})});
 foot(s,'Sources: TGO Rhizophora study; DMCR 9.40 announcement; MOC2 baseline workbook; MOC2 PDD. Baseline is a snapshot—not actual annual increment.');
}
require('fs').mkdirSync('output',{recursive:true});
pptx.writeFile({fileName:'output/MOC2_9_4_carbon_assessment.pptx'});