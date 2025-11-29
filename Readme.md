# 📊 Automated Sales Report Generator (Python & Excel)

This project automates the entire sales reporting process using **Python**, **Pandas**, and **Excel**.  
Instead of manually combining multiple monthly Excel files and preparing summaries, this tool automatically:

- Reads all Excel files from the `data/` folder  
- Merges them into a single dataset  
- Cleans and processes the data  
- Creates product, category, and region-wise summaries  
- Generates a final Excel report with multiple formatted sheets  

This automation reduces manual effort by **90%** and produces clean, consistent reports in seconds.

---

## 🚀 Features

### **1️⃣ Automatic File Reading**
Reads all Excel files from the `data/` folder.

### **2️⃣ Data Cleaning**
- Removes duplicate records  
- Handles invalid or inconsistent date formats  
- Drops missing or corrupted rows  

### **3️⃣ Automated Summary Reports**
Generates clean summary sheets:
- Product-wise Sales  
- Category-wise Sales  
- Region-wise Sales  

### **4️⃣ Excel Output (Final Report)**
Creates `Final_Sales_Report.xlsx` containing:
- **Merged Data**  
- **Product Sales**  
- **Category Sales**  
- **Region Sales**  

---

## 🛠️ Tech Stack

- Python  
- Pandas  
- OpenPyXL  
- Jupyter Notebook  
- Excel  

---

## 📁 Project Structure

```
Automated-Sales-Report-Generator/
│
├── data/
│     ├── sales_2013-06.xlsx
│     ├── sales_2013-09.xlsx
│     ├── sales_2012-11.xlsx
│
├── main.ipynb
├── Final_Sales_Report.xlsx
└── README.md
```

---

## ▶️ How to Run

### **1️⃣ Install Required Libraries**
```bash
pip install pandas openpyxl
```

### **2️⃣ Open the Notebook**
```bash
jupyter notebook main.ipynb
```

### **3️⃣ Run All Cells**
The automation will:
- Read all Excel files  
- Merge and clean the data  
- Generate product/category/region summaries  
- Export the final Excel report  

---

## 📦 Output File

### **Final_Sales_Report.xlsx** includes:
- Merged full dataset  
- Product-wise total sales  
- Category-wise totals  
- Region-wise totals  

---

## 👨‍💻 Author

**Mayank Singh**  
Python Developer | Data Analyst | Automation Enthusiast
