# Karza GST Master Consolidator

An industrial-strength automation tool for consolidating Karza GST Transaction Reports across multiple states and entities.

## 🚀 Features
- **Multi-Entity Support:** Automatically groups files by PAN/Trade Name and creates separate workbooks.
- **Revenue Netting:** Reconciles GSTR-3B data against internal branch transfers (Inter-company sales/purchases).
- **Interactive Matrices:** Generates customer and supplier drill-downs with state-level grouping.
- **Visual Audit Key:** Highlighted rows for Related Parties (Yellow) and Third Parties (Grey).
- **Dashboard UI:** Real-time terminal progress bar and status updates.

## 🛠️ Requirements
- Windows 10/11
- Microsoft Excel (Desktop version)
- Windows PowerShell 5.1 (Standard on most PCs)

## 📖 How to Use
1.  Download `Consolidate_Karza` zip file and extract the zip to retrieve the .bat file.
2.  Place the `.bat` file in a folder containing your `.xlsx` Karza reports.
3.  Double-click the file.
4.  The script will scan the folder, process the data, and generate a `CONSOLIDATED_[PAN]_[Name].xlsx` file for each entity found.

## ⚖️ License
This project is licensed under the MIT License.
