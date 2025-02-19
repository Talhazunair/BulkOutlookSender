# **OutlookMailer**

**OutlookMailer** is a Python automation tool built with **Selenium** and **PyQt5** to send bulk emails via Outlook. It reads recipient emails and attachments from an Excel file and automates the process through a user-friendly GUI. 🚀  

## **Features**  
✅ Automates sending emails via Outlook  
✅ Reads recipient emails from an Excel file  
✅ Simple and user-friendly **PyQt5 GUI**  
✅ Handles multiple emails efficiently  

---

## **Installation**  

### **1. Clone the Repository**  
```bash
git clone https://github.com/Talhazunair/BulkOutlookSender.git
cd BulkOutlookSender
```

### **2. Install Dependencies**  
Ensure Python (3.x) is installed, then run:  
```bash
pip install -r requirements.txt
```

---

## **Setup: Prepare Your Excel File**  

1. Open **Excel** and create a new file (`emails.xlsx`).  
2. Add the following column headers in the first row:  

   | Email | Password | To (Recipient) |
   |-------|---------|---------------|
   | your_email@outlook.com | your_password | recipient1@example.com | 
   | your_email@outlook.com | your_password | recipient2@example.com | 

3. Save the file as **`emails.xlsx`** in the same directory as `main.py`.

---

## **Usage**  

Run the software using:  
```bash
python main.py
```
This will launch the **GUI**, allowing you to send bulk emails with attachments via Outlook.

---

## **License**  
📜 This project is licensed under the **MIT License**.  

---

