# Business Analytics Portfolio

Ruth Selorme Acolatse | MSc AI for Sustainable Societies (EMJM), Tampere, Finland

Work samples in forecasting, Excel modelling and spreadsheet automation. All data is fictional or dummy data. Nothing here comes from a real employer, client or personal account.

Contact: [LinkedIn](https://www.linkedin.com/in/ruth-acolatse-0450181a1?utm_source=share&utm_campaign=share_via&utm_content=profile&utm_medium=android_app) | [Website](https://www.ruthselormeacolatse.info)

## Contents

| Project | What it shows | Tools |
|---|---|---|
| [Pharmacy demand forecasting summary](Ruth_Acolatse_Pharmacy_Demand_Forecasting_Business_Summary.pdf) | Business version of a machine learning class project. Nine model families were compared on 2,106 days of pharmacy demand. The simplest model (linear regression) won, but every model under-predicted demand by 26% to 43%. | Python, forecasting metrics (RMSE, WAPE, MASE, bias) |
| [Licensee screening tracker](Excel%20tracker) | A scoring model that ranks 15 fictional automakers as licensing targets, with a company profile card and a summary sheet. | Excel: XLOOKUP, named ranges, data validation, SUMIF(S), COUNTIF, Power Query |
| [Finance tracker template](Finance%20tracker) | Monthly budget tracker with a dashboard, month picker, budget vs used flags and account balances that roll forward month to month. | Excel: SUMIFS, INDIRECT, named ranges, conditional formatting, charts |
| [Calendar sync script](Calendar%20sync) | Google Apps Script that keeps a calendar in sync with a spreadsheet. Includes a demo workbook. | Google Apps Script, Google Sheets |

## How to use the files

- **Licensee screening tracker:** open the .xlsx in Excel. Change the company on the Profile sheet, or edit the weights on the Lookups sheet, and the scores and ranks update. The `.pq` file is the Power Query (M) code for cleaning the messy `Raw_Shipments` table.
- **Finance tracker template:** pick a month on the Dashboard. Add transactions at the bottom of a month sheet. Blue text marks inputs.
- **Calendar sync script:** see the README in that folder for setup.

## Notes and limits

- The demand forecasting results come from a class project. The business deck states that no model met the industry accuracy benchmark.
- The tracker's scores and weights are illustrative, not real licensing data.
- Exchange rates in the finance template are typed in by hand.

## About

I am completing an MSc in AI for Sustainable Societies (expected June 2027). Before that I managed a retail spare parts store, worked in customer service and sales, and did freelance web development. I am interested in market intelligence, strategy and business analysis roles.
