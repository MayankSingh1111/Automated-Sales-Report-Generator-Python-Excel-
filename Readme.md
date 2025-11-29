Automated Sales Report Generator (Python & Excel)

This project automates the entire sales reporting process using Python, Pandas, and Excel.
Instead of manually combining multiple monthly Excel files and preparing summaries, this tool automatically:

- Reads all Excel files from the data/ folder
- Merges them into a single dataset
- Cleans and processes the data
- Creates product, category, and region-wise summaries
- Generates a final Excel report with multiple formatted sheets

This project reduces manual effort by 90% and produces clean, consistent reports in seconds.

FEATURES:

1. Automatic File Reading  
   Reads all Excel files from the data/ folder.

2. Data Cleaning  
   - Removes duplicates  
   - Handles invalid dates  
   - Drops missing rows  

3. Automated Summary Reports  
   Product-wise, category-wise, and region-wise sales sheets.

4. Final Excel Output  
   Final_Sales_Report.xlsx with multiple sheets:
   - Merged Data  
   - Product Sales  
   - Category Sales  
   - Region Sales  

TECH STACK:
- Python
- Pandas
- OpenPyXL
- Jupyter Notebook
- Excel

PROJECT STRUCTURE:

Automated-Sales-Report-Generator/
│
├── data/
│     ├── sales_2013-06.xlsx
│     ├── sales_2013-09.xlsx
│     ├── sales_2012-11.xlsx
│
├── main.ipynb
├── Final_Sales_Report.xlsx
└── README.txt

HOW TO RUN:

1. Install Required Libraries:
   pip install pandas openpyxl

2. Open the Jupyter Notebook:
   jupyter notebook main.ipynb

3. Run All Cells  
   The script will:
   - Read all Excel files  
   - Merge and clean the data  
   - Create summary reports  
   - Export the final Excel file  

OUTPUT FILE:
Final_Sales_Report.xlsx  
Includes:
- Merged data  
- Product-wise totals  
- Category totals  
- Region totals  

AUTHOR:
Mayank Singh  
Python Developer | Data Analyst | Automation Enthusiast
