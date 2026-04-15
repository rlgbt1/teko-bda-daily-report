# Reference Alignment Notes

Reference deck:
`/Users/ReinaldoLuigi/Desktop/Teko AI/RESUMO DIÁRIO 20260330.pptx`

Current generator test deck:
`/Users/ReinaldoLuigi/Desktop/Teko AI/teko-bda-daily-report/output/test1_sample_01042026.pptx`

## Global mapping

- Slide size is effectively 13.33" x 7.50" in both decks.
- Cover keeps the `Uma Missão de Futuro` branding band.
- Inner slides use the slim header system only:
  - orange accent block at `L=0.000 T=0.401 W=0.277 H=0.397`
  - title at `L=0.396 T=0.224 W=11.321 H=0.351`
  - orange rule at `T=0.607`
- Inner slides in the reference do not show the report date in the header.
- Footer is always present as the BDA Talatona footer strip.

## Slide-by-slide notes

### Slide 1

- Keep existing cover composition.
- Banner and cover art match the reference pattern closely enough.

### Slide 2

- Agenda structure matches reference.
- Remaining refinement is mostly typography spacing, not structural layout.

### Slide 3

- Reference uses a dashboard-style summary centered on one brown KPI circle.
- Current generator already follows the dashboard direction better than the old prose version.
- Important mapping points:
  - central KPI circle at about `L=5.741 T=3.009 W=1.779 H=1.491`
  - metric cards orbit around the center rather than stack as a table
  - inner slide date removed to match the reference header system

### Slide 4

- Left stack should remain the main data zone:
  - Liquidez MN block at `L=0.407 T=0.700 W=8.557`
  - Transações block around `T=2.204`
  - Operações Vivas block around `T=3.012`
- KPI bubbles stay in the lower-right region, away from the tables.
- Inner slide date removed.

### Slide 5

- This slide stays as the MN dashboard with cash-flow plus compact supporting visuals.
- Reference composition uses visual elements in the lower half rather than one full-page table.
- Current generator keeps the pie-chart direction.
- Inner slide date removed.

### Slide 6

- Reference is a four-part dashboard:
  - Liquidez ME table
  - Transações / Operações Vivas stack
  - composition pie chart in upper-right
  - Fluxos de Caixa table in lower-right
- Added a dedicated Liquidez ME pie chart to match the missing top-right component.
- Key mapped areas:
  - pie chart at about `L=8.571 T=0.377 W=4.630 H=2.981`
  - Fluxos table at about `L=8.230 T=3.078 W=5.021 H=2.191`
  - bottom KPI bubbles near `L=2.817` and `L=4.350`
- Inner slide date removed.

### Slide 7

- Reference layout is strongly chart-led.
- Key mapped regions:
  - Cambiais table at `L=0.396 T=0.705 W=6.112`
  - Transações BDA at `T=1.865`
  - Transações do Mercado at `T=3.047`
  - Posição Cambial bar chart at `L=6.212 T=3.260 W=6.521 H=2.809`
  - Taxa de Cambio line chart at `L=0.163 T=4.812 W=6.385 H=1.858`
  - KPI bubbles at `L=8.502 T=1.422` and `L=10.308 T=1.454`
- Added explicit chart title bars to match the reference labeling.
- Inner slide date removed.

### Slide 8

- Reference uses:
  - segment table on top-left
  - single KPI bubble near center
  - equities table in the lower half
- Current generator already follows that structure.
- Inner slide date removed.

### Slide 9

- Portfolio remains split into `Custo Amortizado` and `Justo Valor`.
- KPI bubbles belong at the bottom band, not mixed into the main table area.
- Inner slide date removed.

### Slide 10

- Reference has:
  - capital markets table on left
  - commentary box on right
  - note tag near the middle-right
  - crypto table in lower-left
  - crypto commentary box lower-right
- Generator coordinates were nudged closer to the reference left alignment.
- Inner slide date removed.

### Slide 11

- Reference uses:
  - commodities table left
  - commentary box upper-right
  - minerals table lower-left
  - commentary box mid-right
  - note area at the bottom
- Generator already removed the globe and uses `PETRÓLEO`.
- Bottom `Nota` box is kept as the standardized commentary band.
- Inner slide date removed.

## Remaining tuning targets

- Reduce visible “shape fragmentation” on table slides where possible by favoring larger grouped visual blocks.
- Fine-tune Slide 3 bubble spacing against the exact reference visual balance.
- Fine-tune Slide 9 vertical compression if real data makes the portfolio sections too tall.
- Adjust typography and paragraph spacing on Slides 10 and 11 after visual review of the regenerated deck.
