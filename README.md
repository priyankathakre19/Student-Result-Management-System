# 🎓 Student Result Management System

A simple and user-friendly **Student Result Management System** built using **Python Tkinter** and **Excel**. The application allows users to add student details and marks, automatically calculate total marks, percentage, and result status, search individual results using Roll No., and view all student results in a Treeview table.

## 📌 Features

- 📝 Add student details and subject marks
- 🧮 Automatically calculate total marks
- 📊 Automatically calculate percentage
- ✅ Automatically determine Pass/Fail result
- 🔍 Search student result using Roll No.
- 📋 Display all student results in a Tkinter Treeview
- 💾 Store student records in an Excel file
- ⚠️ Input validation for marks and duplicate Roll Numbers
- 📁 Automatically create the Excel file if it does not exist

## 🛠️ Technologies Used

- **Python**
- **Tkinter** – GUI development
- **OpenPyXL** – Excel file handling
- **Microsoft Excel** – Data storage

## 📂 Project Structure

Student-Result-Management-System/
│
├── result_management.py
├── student_results.xlsx
└── README.md

## 🔄 Project Flow

Add Student
     ↓
Enter Student Details & Marks
     ↓
Calculate Total & Percentage
     ↓
Determine Pass / Fail
     ↓
Save Record to Excel

Get Result
     ↓
Enter Roll No.
     ↓
Search Student Record
     ↓
Display Result

Show All Results
     ↓
Read Excel Data
     ↓
Display Records in Treeview

## ▶️ How to Run

### 1. Install Python

Make sure Python is installed on your system.

### 2. Install OpenPyXL

Open the terminal and run:

python -m pip install openpyxl

### 3. Run the application

python result_management.py

## 📊 Excel Data

Student records are stored in:

student_results.xlsx

The Excel file contains:

| Field | Description |
|---|---|
| Name | Student name |
| Roll No. | Unique student roll number |
| Class | Student class |
| Subject 1–5 | Marks obtained |
| Total Marks | Total of all subjects |
| Percentage | Calculated percentage |
| Result | Pass / Fail |

## 🎯 Purpose

This project demonstrates the use of **Python GUI programming, Excel file handling, data validation, calculations, and Treeview-based data presentation** in a practical mini project.

## 👩‍💻 Author

**Priyanka Thakre**

---

⭐ If you find this project useful, consider giving the repository a star!