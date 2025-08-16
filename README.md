🧾 Automated Excel Report Generator

A Python script to automate CSV-to-Excel report generation with formatting, totals, and professional styling. Perfect for retail, field audits, or any dataset that needs structured, ready-to-use Excel reports.

✨ Features

✅ Bulk Processing: Reads all CSV files from an input folder and outputs formatted Excel files.

✅ Standardized Reports: Includes columns like:

Sr No, Shop Name, Count, Elements, Product Name

Measurements in inches and feet

Quantity, Total Sqft

Contact and location details

✅ Fill-Down Logic: Automatically fills missing Contact Number and Contact Person (like Excel Ctrl+D).

✅ Auto Numbering: Generates Sr No and Count per shop.

✅ Professional Formatting:

Bold headers

Center-middle alignment for all cells

Thin borders for all cells

Numeric columns (W in Ft, H in Ft, Total Sqft) formatted to two decimal places

✅ Total Row: Adds a row with shop count and sum of Total Sqft.

✅ Extendable: Add more columns, formatting rules, or calculations easily.

🛠 Installation
pip install pandas openpyxl

🚀 Usage

Place all CSV files in the InputFiles folder.

Update input_folder and output_folder paths in the script if needed.

Run the script:

python excel_report_automation.py


Formatted Excel reports will appear in the Output folder with _output appended to the original filename.

⏱ Results
Metric	Manual	Automated
Processing Time (100 files)	16–25 hours	~1 second
Time Saved per Batch	—	~99%
Accuracy	—	100%

Ongoing Impact: Since CSV files are received daily, this automation eliminates a permanent manual workload, saving thousands of hours annually.

💡 Business Impact

Efficiency: Eliminates repetitive tasks from daily operations.

Productivity: Staff can focus on high-value work.

Consistency: Standardized dashboards improve data quality.

Scalability: Handles growing data volumes effortlessly.

Cost Savings: Reduces operational costs permanently using free tools.
