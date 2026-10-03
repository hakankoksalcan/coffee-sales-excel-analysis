# Coffee Sales Analysis in Excel

I put together this project to practise a small freelance reporting task: reviewing coffee sales and explaining what the numbers could mean for the business.

Using Excel, I analysed 669 transactions from 1 March to 31 May 2024. I focused on a few practical questions: which drinks brought in the most revenue, how sales changed each month, and when customers were most likely to buy.

## Files

- [Excel workbook](Coffee_Sales_Analysis.xlsx): the data, quality checks, PivotTables, charts and recommendations.
- [Executive summary](Executive_Summary.pdf): a one-page overview of the main findings and suggested actions.

The workbook uses Turkish worksheet labels. The PDF summary is in English.

## Working with the data

Before building the report, I checked missing values, date and number formats, beverage names, exact duplicates and unusual transaction amounts.

There were 85 missing card identifiers, all belonging to cash payments. I kept those records because a cash transaction does not need a card identifier. No exact duplicates were found, and all 669 transactions in the selected period were retained.

I then used Excel formulas, PivotTables and charts to compare beverages, months, hours and payment methods.

## What I found

| Measure | Result |
| --- | ---: |
| Total revenue | 22,834.18 |
| Transactions | 669 |
| Average transaction value | 34.13 |

The currency was not independently verified, so amounts are shown in the dataset's currency units.

Latte generated the most revenue, accounting for 23.06% of the total. May had the highest monthly revenue, with a 34.86% increase compared with April.

The busiest hour was 10:00–10:59, with 87 transactions across the period. Card payments made up 87.29% of transactions.

## Suggested next steps

Based on these results, I suggested prioritising ingredient checks for the four most frequently purchased drinks. Actual ingredient use would need to be recorded before deciding how much to order.

I also suggested testing short maintenance or refill tasks during quieter morning hours, provided the business is open then and the change does not interrupt sales.

These are starting points for the business to test. The data does not include costs, stock shortages, opening hours or downtime, so it cannot explain profitability or the reasons behind every sales pattern.

## Data source

- Original dataset: [Coffee Sales on Kaggle](https://www.kaggle.com/datasets/ihelon/coffee-sales)
- Working CSV: [CoffeeSales on Hugging Face](https://huggingface.co/datasets/tablegpt/CoffeeSales)

This project uses public data and was completed as a simulated client assignment for my portfolio.
