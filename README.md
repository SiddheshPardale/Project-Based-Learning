# 📋 Job Application Tracker

A **Python-based Job Application Tracker** that allows users to manage their job applications using **CRUD operations**. The application provides a Rich-based terminal interface and allows application data to be permanently stored in an **Excel file** using `openpyxl`.

---

## 🚀 Project Overview

The Job Application Tracker is a **console-based Python application** developed to maintain and manage job application records.

The application allows the user to:

* ➕ Add a new job application
* 👀 View all applications in a formatted table
* ✏️ Update an existing application
* 🔍 Search for an application using Application ID
* 🗑️ Delete an application
* 📊 Export/save application data into Excel
* 💾 Preserve data in an Excel file for future use

---

# 🛠️ Technologies & Libraries Used

| Technology / Library | Purpose                                          |
| -------------------- | ------------------------------------------------ |
| **Python**           | Main programming language                        |
| **datetime**         | Handling and validating application dates        |
| **Rich**             | Creating colorful and formatted terminal UI      |
| **Rich Table**       | Displaying application records in tabular format |
| **Rich Panel**       | Creating the welcome interface                   |
| **openpyxl**         | Creating, reading and updating Excel files       |
| **os**               | Checking whether the Excel file already exists   |

### Installation

Install the required external libraries using:

```bash
pip install rich openpyxl
```

---

# 📚 Python Concepts Learned
 CRUD Operations

This project implements the basic **CRUD concept**:

| Operation  | Function              |
| ---------- | --------------------- |
| **Create** | `Add()`               |
| **Read**   | `View()` / `Search()` |
| **Update** | `Update()`            |
| **Delete** | `Delete()`            |

The Excel export functionality provides persistent storage.

---

# 🔄 Application Execution Flow

The complete execution flow of the application is:

```text
                ┌──────────────────────┐
                │ Start Program        │
                └──────────┬───────────┘
                           ↓
                ┌──────────────────────┐
                │ Display Welcome Panel│
                └──────────┬───────────┘
                           ↓
                ┌──────────────────────┐
                │ Display Main Menu    │
                └──────────┬───────────┘
                           ↓
              ┌────────────┴────────────┐
              │     User Selects Option │
              └────────────┬────────────┘
                           ↓
        ┌─────────┬────────┼────────┬─────────┬─────────┬─────────┐
        ↓         ↓        ↓        ↓         ↓         ↓
      Add()    View()   Update() Search()  Delete() ExportExcel()
        │         │        │        │         │         │
        └─────────┴────────┴────────┴─────────┴─────────┘
                           ↓
                ┌──────────────────────┐
                │ Return to Main Menu  │
                └──────────┬───────────┘
                           ↓
                    User selects 7
                           ↓
                ┌──────────────────────┐
                │ Exit Program        │
                └──────────────────────┘
```

---

# 🔍 Detailed Execution Flow

## Step 1 — Program Starts

The program imports the required modules and creates an empty list:

```python
Applications = []
```

This list temporarily holds the application records while the program is running.

---

## Step 2 — Welcome Screen

The `Rich` library is used to display a formatted welcome panel.

```python
Panel(
    "\n[bold cyan]Welcome to Job Application Tracker ![/bold cyan]\n"
)
```

This makes the console application more user-friendly.

---

## Step 3 — Main Menu

The program displays seven options:

```text
1. Add Job Application
2. View Application
3. Update Application
4. Search Application
5. Delete Application
6. Save the Data to Excel
7. Exit
```

The user enters a choice.

---

# ➕ Step 4 — Add Application

When the user selects:

```text
1
```

the `Add()` function is called.

The program asks for:

* Application ID
* First Name
* Last Name
* Company Name
* Position
* Location
* Application Date
* CTC
* Application Status

The information is stored in the `Applications` list.

Example:

```text
[101, "Rahul", "Patil", "Amazon", "Software Engineer",
 "Pune", 2026-09-04, 8.5, "Applied"]
```

---

# 👀 Step 5 — View Applications

When the user selects:

```text
2
```

the `View()` function displays all records.

The Rich `Table` class is used to create a formatted table.

Application status is also displayed using different colors:

* 🟡 Applied
* 🔵 Interview
* 🔴 Rejected
* 🟢 Selected

---

# ✏️ Step 6 — Update Application

When the user selects:

```text
3
```

the program asks for the Application ID.

It searches the list:

```python
for i in Applications:
    if i[0] == roll_a:
```

If the ID is found, the existing information is replaced with the new information.

---

# 🔎 Step 7 — Search Application

When the user selects:

```text
4
```

the program asks for an Application ID.

It searches through the `Applications` list.

If the ID exists:

```text
Found.....
```

and the application details are displayed.

If it does not exist:

```text
Not Found the Application ID
```

---

# 🗑️ Step 8 — Delete Application

When the user selects:

```text
5
```

the program searches for the Application ID.

If found, the record is removed:

```python
Applications.remove(i)
```

The user receives:

```text
Deleted Application Successfully!
```

---

# 📊 Step 9 — Save Data to Excel

When the user selects:

```text
6
```

the `ExportExcel()` function is executed.

The program checks:

```python
os.path.exists("JobApplications.xlsx")
```

### If the file does not exist

A new Excel workbook is created:

```python
workbook = Workbook()
```

Headers are added:

```text
Application ID | First Name | Last Name | Company |
Position | Location | Date | CTC | Status
```

### If the file already exists

The existing workbook is opened:

```python
workbook = load_workbook(file_name)
```

New records are then added to the existing worksheet.

Finally:

```python
workbook.save(file_name)
```

saves the data permanently.

---

# 💾 Data Persistence

One important concept learned from this project is **persistent storage**.

Initially:

```python
Applications = []
```

stores data only in memory.

Therefore, if the Python program is closed, the list is destroyed.

To solve this problem, the project uses an Excel file:

```text
JobApplications.xlsx
```

The flow becomes:

```text
Python List
     ↓
ExportExcel()
     ↓
JobApplications.xlsx
     ↓
Data remains after program closes
```

When the program runs again, the existing Excel file can be loaded using:

```python
load_workbook()
```

This demonstrates the difference between:

**Temporary/In-memory data**

and

**Permanent/File-based data storage.**

---

# 🧠 Important Programming Concepts Demonstrated

This project combines several concepts into one practical application:

```text
Python Basics
     │
     ├── Variables
     ├── Input / Output
     ├── Lists
     ├── Loops
     ├── Conditions
     ├── Functions
     ├── Match-Case
     ├── Type Conversion
     ├── Date Handling
     │
     ↓
File Handling
     │
     ├── os.path.exists()
     ├── Workbook()
     ├── load_workbook()
     └── workbook.save()
     │
     ↓
External Libraries
     │
     ├── Rich
     └── OpenPyXL
     │
     ↓
Application Development
     │
     └── CRUD Operations
```

---

# ▶️ How to Run

### 1. Clone the repository

```bash
git clone <your-github-repository-url>
```

### 2. Open the project directory

```bash
cd Job-Application-Tracker
```

### 3. Install dependencies

```bash
pip install rich openpyxl
```

### 4. Run the Python file

```bash
python JobApplicationTracker.py
```

---

# 🔮 Future Improvements

Possible improvements for future versions:

* Add **SQLite/MySQL database** instead of Excel
* Prevent duplicate Application IDs
* Add login/authentication
* Add search by company or position
* Add sorting by application date
* Add automatic date/time
* Build a web version using **Flask/Django**

---

# 👨‍💻 Conclusion

The Job Application Tracker is a practical Python project that demonstrates how fundamental Python concepts can be combined to create a useful real-world application.

The project helped me move from writing individual Python programs to building a **structured, menu-driven application with CRUD functionality, external libraries, formatted CLI output, and persistent Excel-based storage**.
