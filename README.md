# Budget vs Actual P&L Report with Scenario Model (Excel)

A P&L variance analysis built in Excel, comparing monthly actuals with budget and forecasting the next year under three scenarios.

## Business question

Revenue came in above budget, so why did profit end up below plan?

## Key findings

- **Revenue +3.6% vs budget** ($9.86M vs $9.52M). Cloud Subscriptions (+8.1%) more than offset Hardware Sales (-6.9%).
- **Gross margin 74.8% vs 76.6% budgeted** (about -1.8 pp), driven by a **15.5% overrun in Server Hosting & Infrastructure**.
- **EBIT $3.21M, 0.5% below budget** (-$15k), as higher infrastructure and R&D spend outweighed the revenue gain.
- Hosting cost per $1 of subscription revenue rose from 26.7% in January to 29.9% in December.

## What is in the workbook

| Sheet | Content |
|---|---|
| `P&L_Report` | Actual vs Budget P&L, variance in $ and %, gross and EBIT margins, next-year forecast, scenario selector, executive summary |
| `Actuals` | Monthly actuals by account (Jan-Dec), stored as the table `tblActuals` |
| `Budget` | Monthly budget by account, stored as the table `tblBudget` |

## Scenario model

A selector cell switches the forecast between three scenarios; lookup formulas apply the growth rates to every P&L line.

| Scenario | Revenue growth | Cost growth |
|---|---|---|
| Worst | -3% | 2% |
| Base | 5% | 3% |
| Best | 12% | 5% |

## Excel techniques used

- Structured tables and `SUMIFS` to aggregate monthly data into the report
- `VLOOKUP` driven by a scenario selector
- `IFERROR` for safe variance percentages
- Variance analysis ($ and %), margin calculations and an executive summary

## Files

- `Excel_projekt.xlsx` - the workbook
- Presentation (`.pptx`) - 8-slide summary with findings, recommendations, user stories and acceptance criteria

## Notes

The dataset is illustrative sample data, not figures from a real company.

## Author

Michał Mucha - Economic Analytics student, Wrocław University of Economics and Business
[LinkedIn](https://www.linkedin.com/in/micha%C5%82-mucha-4692872b9/)
