# multi_file_sales_data_cleaning_automation
# Excel Sales Data Cleaning & Reporting Automation

## 📌 Project Overview
This project demonstrates how messy sales data from multiple CSV files can be cleaned, standardized, and transformed into a structured Excel report using Python.

The goal is to simulate a real-world client scenario where data comes from different sources with inconsistencies.

---

## 🧹 Data Cleaning Tasks Performed

- Merged multiple CSV files into one dataset
- Standardized inconsistent column names
- Cleaned text data (e.g., product names)
- Handled missing values:
  - Filled missing quantity and price with 0
  - Kept missing dates unchanged
- Removed duplicate rows
- Fixed formatting issues

---

## 📊 Data Processing & Analysis

- Grouped data to generate summary insights
- Calculated total quantity per product
- Sorted data for better readability

---

## 📁 Output

The script generates a clean Excel file with:

- **Cleaned Data Sheet**
- **Summary Sheet**
- Formatted headers (bold)
- Adjusted column widths

---

## 🛠️ Tools & Technologies

- Python
- pandas
- openpyxl

---

## 📂 Project Structure
project/ │
├── input/ │  
    ├── sales_january.csv │   
    ├── sales_february.csv │  
    ├── sales_march.csv │
├── output/ │  
    └── cleaned_report.xlsx │ 
  ├── script.py 
  ├── requirements.txt
  └── README.md













