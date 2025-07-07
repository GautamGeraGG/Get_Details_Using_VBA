# Excel VBA Name Filter Macro

## 📌 Project Overview
This VBA macro is designed to filter and retrieve details about a person from a database (Excel Sheet1) based on their name input.  
It searches for a name entered by the user, retrieves the related score data from predefined columns, and displays it in a designated area.  

---

## 🚀 Features
- Takes a name from **cell L2 (Sheet1)** as input.
- Performs **case-insensitive search** in **column A** of Sheet1.
- Retrieves corresponding data from **columns D to F (rows relative to the matched name)**.
- Displays the retrieved data in the range **K8:M11**.
- Displays the entered name in **uppercase in K6**.
- Shows a message if:
  - The name is not found.
  - The name input is empty.
- Automatically disables Cut/Copy mode after execution.

---

## 📂 How It Works
1️⃣ The user enters a name in **L2**.  
2️⃣ The macro builds a dictionary of names from **column A**.  
3️⃣ It looks for a matching name (case-insensitive).  
4️⃣ If found:
- Copies data from **D to F (relative rows)** to **K8:M11**.
- Displays the entered name in uppercase in **K6**.

5️⃣ If not found:
- Shows a message box: `"The name you entered is not found"`.
- Clears **K8:M11**.

6️⃣ If L2 is empty:
- Shows a message box: `"Please Enter The Name First"`.

---

## 💻 How to Use
1. Open the Excel file.
2. Go to **Developer → Macros**.
3. Run `match`.
4. Enter a name in **L2** before running.

---

## 📌 Notes
- The macro uses `Scripting.Dictionary` for fast lookups.
- It disables Cut/Copy mode at the end (`Application.CutCopyMode = False`).
- Designed for **Sheet1** — modify the code if using a different sheet.

---

## 📝 Code Summary
```vba
' Builds dictionary of names (lowercase)
' Matches user input (lowercase) to keys
' Copies corresponding data if found
' Shows messages if not found or empty
' Cleans up Cut/Copy mode
