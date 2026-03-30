[README.md](https://github.com/user-attachments/files/26339475/README.md)
# 🚀 Auto Employee Card Generator

A Python-based automation tool that generates professional employee ID
cards in bulk using Excel data.

## 📌 Project Overview

This tool reads employee information from an Excel (`.xlsx`) file and
automatically generates ID cards including:

-   Employee Name
-   Employee ID
-   Designation
-   Office / Branch
-   Contact Information
-   QR Code (optional)
-   Front & Back side design

## ✨ Features

-   Bulk card generation from Excel
-   Professional card design
-   Logo integration
-   QR Code support
-   Export as PDF / Image (JPEG/PNG)

## 🛠 Technologies Used

-   Python 3
-   Pillow (PIL)
-   Pandas
-   OpenPyXL
-   qrcode
-   ReportLab

## 📂 Project Structure

Auto-Employee-Card-Generator/ │ ├── card_generator.py ├── assets/ ├──
data/ ├── output/ └── README.md

## 📥 Installation

``` bash
git clone https://github.com/your-username/auto-employee-card-generator.git
cd auto-employee-card-generator
pip install pandas pillow openpyxl qrcode reportlab
```

## 📊 Excel Format

| EmployeeName \| EmployeeID \| Designation \| OfficeName \| Mobile \|
  Email \|

## ▶️ Usage

``` bash
python card_generator.py
```

## 📤 Output

Generated cards will be saved in the output folder.

## 🎯 Use Cases

-   Banks
-   Corporate Offices
-   Educational Institutions

## 👨‍💻 Author

Susam Das
