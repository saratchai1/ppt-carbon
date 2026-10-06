# Revised PDD growth scenario — 2026-10-06

This note replaces the earlier flat 9.4 tCO2e/rai/year assumption for existing forest with a working scenario requested by the project owner.

## Scenario inputs

| Class | Interpretation | Annual assumption |
|---|---|---:|
| Green / FCD high | Existing forest stock | 5% stock growth per year, compounded |
| Yellow / FCD medium | Existing forest stock | 6% stock growth per year, compounded |
| Red / FCD low | Existing forest stock | 10% stock growth per year, compounded |
| Zero-stock / bare planting | Starts from 0 carbon stock | Ramp: Y1 0.5, Y2 1, Y3 2, Y4 4, Y5 6, Y6 8, Y7+ 9 tCO2e/rai/year |

The zero-stock ramp is a working assumption, not an approved methodology: it implements the instruction that early years should be small and that after year 6 the plot can reach about 9 tCO2e/rai/year.

## MOC2-VSD calculation

MOC2-VSD uses the baseline workbook values:

- Project area: 6,775.53 rai
- Baseline tree carbon stock: 362,962.81 tCO2e
- Old PDD calculation: 6,775.53 × 9.4 × 10 = 636,899.82 tCO2e over 10 years

Under the revised stock-growth scenario:

| Metric | Value |
|---|---:|
| New 10-year increment | 244,042.03 tCO2e |
| Scenario / old PDD | 38.3% |
| Difference vs old PDD | -392,857.79 tCO2e |
| Year-1 annual increment | 18,999.30 tCO2e/year |
| Year-10 annual increment | 30,553.67 tCO2e/year |

Interpretation: this is much lower than the flat 9.4 assumption, because the existing forest already has standing stock and is modeled as percentage stock growth rather than a constant 9.4 tCO2e/rai/year added to every rai.

## PDD / legacy comparison

| Project/PDD | Basis | Old / legacy total | New scenario total | Scenario / old | Notes |
|---|---|---:|---:|---:|---|
| MOC2-VSD | Existing forest + small zero-stock area | 636,899.82 | 244,042.03 | 38.3% | Directly modeled from MOC2 baseline workbook |
| MOC1-STC | Zero-stock planting + SOC note | 86,085.90 | 64,744.58 | 75.2% | Keeps legacy SOC note unchanged; tree stock uses ramp |
| MOC4 / group 1 PDD | Zero-stock planting | 27,293.98 | 18,014.38 | 66.0% | PDD table implied about 10.35 tCO2e/rai/year, not 9.4 |
| MOC1-VSD | Legacy summary only | 169,725.08 | not recalculated | n/a | Area and SOC/tree split not confirmed |

## Working conclusion

The project should stop treating 9.4 tCO2e/rai/year as a direct annual growth rate for existing forest. The revised model is more defensible as an internal sensitivity case, but it is still not a submission-ready credit calculation. The next defensible step is to validate these assumptions with repeated permanent plots and to agree the model boundary with the verifier/TGO.

## Local generated artifacts

- `PDD_carbon_growth_scenario_model.xlsx` — formula model with inputs and comparison tables
- `MOC2_9_4_carbon_assessment_scenario.pptx` — updated deck with revised scenario slides
- `MOC2_9_4_carbon_assessment_scenario.pdf` — preview export
