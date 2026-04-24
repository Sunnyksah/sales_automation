## What It Does

- Reads 3 source files — products, regions, and raw sales transactions
- Filters out Pending and Refunded orders, keeps only Completed
- Merges all 3 tables into one combined dataset
- Aggregates into monthly KPIs, region performance, and product breakdown
- Writes a 5-sheet styled Excel report with charts and conditional formatting

## Report Output

The generated `monthly_report_YYYY_MM.xlsx` contains 5 sheets:

| Sheet | Contents |
|---|---|
| Dashboard | KPI cards + monthly revenue trend table |
| Monthly Summary | Revenue, growth %, units, orders per month + line chart |
| Region Performance | Actual vs target revenue per region + bar chart |
| Product Breakdown | Revenue by category and individual product |
| Raw Data | Full cleaned transaction-level data |

## Tech Stack

- Python 3.x
- pandas — data cleaning, merging, aggregation
- openpyxl — Excel file creation and styling

## Setup

**1. Clone the repository**
```bash
git clone https://github.com/your-username/sales_automation.git
cd sales_automation
```

**2. Create and activate virtual environment**

Mac/Linux:
```bash
python3 -m venv venv
source venv/bin/activate
```

Windows:
```bash
python -m venv venv
venv\Scripts\activate
```

**3. Install dependencies**
```bash
pip install -r requirements.txt
```

**4. Add your data files**

Place your Excel files in the `data/` folder:
- `products.xlsx`
- `regions.xlsx`
- `sales_raw.xlsx`

**5. Run the pipeline**
```bash
python main.py
```

The report will appear in the `output/` folder.

## Key Concepts Demonstrated

- ETL pipeline design (Extract → Transform → Load)
- Multi-source Excel data merging with pandas
- Data cleaning and filtering
- Aggregation and KPI calculation
- Professional Excel report generation with openpyxl
- Modular, reusable project structure

## Author

 
[LinkedIn](https://linkedin.com/in/sunny-sah-666598239) | [GitHub](https://github.com/SunnyKsah)
