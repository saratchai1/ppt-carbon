# MOC2 Carbon — 9.4 tCO₂e/ไร่/ปี

ชุดสไลด์วิเคราะห์คำถามว่า **ป่าของโครงการ MOC2 จะสามารถเพิ่มคาร์บอนเฉลี่ย 9.4 tCO₂e/ไร่/ปี ได้จริงหรือไม่**

## ข้อสรุป

- ค่า **9.4 ± 4.6 tCO₂e/ไร่/ปี** มาจากงานศึกษาของ ทช./อบก. ใน **แปลงปลูกไม้สกุลโกงกาง (Rhizophora spp.)** 116 แปลง 9 จังหวัด อายุ 7–30 ปี
- งานอ้างอิงใช้ **Mean Annual Increment (MAI)** ซึ่งคำนวณจาก biomass สะสม ณ อายุ `a` หารด้วยอายุ `a` แล้วแปลงเป็นคาร์บอนและ CO₂ ดังนั้นไม่ใช่การวัดว่าป่าเพิ่ม stock จริง 9.4 ในทุกปี
- Baseline ของ MOC2 เป็นป่าธรรมชาติ/ป่าที่มีอยู่เดิม หลายชนิด อายุไม่ทราบ และมีความแปรปรวนสูง จึงไม่ใช่ประชากรเดียวกับชุดอ้างอิง 9.4
- จาก baseline workbook: 32 sample plots, พื้นที่โครงการ 6,775.53 ไร่, baseline tree carbon 362,962.81 tCO₂e หรือเฉลี่ยถ่วงพื้นที่ประมาณ 53.57 tCO₂e/ไร่
- หากตีความ 9.4 เป็น net stock increase ในหนึ่งปี จะเท่ากับประมาณ **17.5% ของ baseline tree-carbon stock ต่อปี** ซึ่งเป็น hurdle ที่สูงสำหรับป่าที่มี stock อยู่แล้ว
- ข้อสรุปที่ defend ได้วันนี้คือ **9.4 เป็น ex-ante benchmark ที่ optimistic สำหรับทั้งโครงการ แต่ยังไม่สามารถสรุปว่าได้หรือไม่ได้จนกว่าจะมี repeated monitoring**

## Files

- `MOC2_9_4_carbon_assessment.pptx` — PowerPoint 16:9
- `MOC2_9_4_carbon_assessment.pdf` — preview/export
- `src/build_deck.js` — source สำหรับสร้าง PowerPoint ด้วย PptxGenJS

## Main sources

1. TGO — การศึกษาการกักเก็บก๊าซคาร์บอนไดออกไซด์ของไม้สกุลโกงกาง (Rhizophora spp.)  
   https://ghgreduction.tgo.or.th/en/component/flexicontent/download/8997/74/17.html
2. DMCR — ประกาศค่าศักยภาพ 9.40 tCO₂eq/ไร่-ปี สำหรับใช้เป็นค่าคาดการณ์กิจกรรมปลูกป่าชายเลน  
   https://www.dmcr.go.th/detailAll/64566/nws/0/
3. MOC2 baseline workbook — `MOC2_VSD-25-06-24 3.xlsx`
4. MOC2 PDD — `T-VER-S-F001-PDD V2.1-MOC2-25-06-24 1.pdf`

## Important caveat

Baseline เป็น snapshot ก่อนเริ่มโครงการ จึงยังไม่สามารถใช้สรุป **actual annual increment** ได้ ผลจริงต้องมาจาก repeated permanent plots และคำนวณ `(C_t - C_0) / Δt` โดยรวม recruitment, mortality และ uncertainty ตาม methodology/PDD ที่ได้รับอนุมัติ
