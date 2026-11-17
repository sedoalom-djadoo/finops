# Changelog: public version

I reviewed the production model before publishing it. The table lists the calculation problems I found and how this version handles them.

## Calculation fixes

| # | Area | Before | Consequence | Now |
|---|---|---|---|---|
| 1 | Gross margin by site | Company direct costs split by revenue share, presented as a site margin | Every clinic had the same margin %, so the chart could not compare them | Renamed `Allocated Gross Margin by Site`. The chart title shows the method and `Class Tagging Coverage %`. `Measured Gross Margin by Site` is ready for when tagging improves |
| 2 | Revenue YTD | Current year cut at `TODAY()`, prior year computed with `TOTALYTD` | Two windows of different lengths in the same variance | Both use `TOTALYTD` with the September fiscal year. The prior year is limited to the dates that have data this year |
| 3 | Gross and net margin | QuickBooks costs and net income divided by billing-team revenue | Two systems on two calendars in the same ratio | Margins use QuickBooks revenue (account 40001). Billing revenue only drives the split by clinic, sponsor or study |
| 4 | Gross margin under a site filter | Site billing revenue minus the company's total direct costs | Meaningless number when a clinic was selected | Shows the site's allocated share |
| 5 | Average invoice | USD and CAD amounts added together | USD invoices undervalued by about 27 % | USD converted at the monthly Bank of Canada rate |
| 6 | Prior year | `SAMEPERIODLASTYEAR` in some measures, `FiscalYear - 1` in others | A partial year compared with a full one on the net income and EBITDA cards | Flows use a like-for-like shift (`DimDate[IsBeforeDataCutoff]`). Balance sheet measures keep the previous year-end snapshot on purpose |
| 7 | Long-term debt | Eight account numbers written into the DAX | A new loan account would be missed | Uses the `Lender` attribute of `DimAccount` |
| 8 | EBITDA bridge | Step labeled "- Tax" while the tax amount is added back | Label contradicted the bar | Renamed "+ Tax" |
| 9 | Direct cost composition | `ALL(DimAccount)` this year, `ALLSELECTED(DimAccount)` last year | The two percentages had different denominators under a filter | `ALLSELECTED` for both |

## Model cleanup

- Removed the bidirectional, one-to-one relationship between `FactPLAdjustments` and `DimDate`.
- Turned off Auto date/time and removed the hidden `LocalDateTable` it had created.
- Replaced the hard-coded "Profit for the year" amounts in `FactBalanceSheet` with a parsing rule on the balance sheet export.
- Moved every source path into a `DataFolder` parameter.
- Removed the auto-generated Q&A linguistic schema and internal names from columns and comments.
- Fixed two charts (net income, assets vs liabilities) whose info tooltip had been copied from the site margin chart.

## Not changed yet

About fifteen "variance label" measures repeat the same text-formatting logic (arrow, percentage, amount vs prior year). They work. Merging them would need DAX user-defined functions, which I have left for a later version.
