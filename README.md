# V-Type Truss to Column AutoCAD Generator

This project is a Python tool that automatically generates a **technical drawing of a V-type truss connection to a column in AutoCAD**.  
The application reads parameters from an Excel configuration file and uses them to create a complete drawing with sections, plates, and a summary table.

The program was created during my university studies when I needed to prepare this type of drawing manually. To speed up the process and reduce repetitive work, I decided to **automate the entire workflow**.

⚠️ **Important:** The program currently supports **only V-type trusses**.

---

## Features

- Automatically generates a **V-type truss to column connection drawing**
- Creates:
  - the main drawing
  - cross-sections
  - a table listing structural elements and their lengths
  - plates drawn in top view for additional dimensioning and description
- Uses **Excel (`data.xlsm`) as a parameter input panel**
- Supports **multiple predefined cross-sections** in Excel
- Section parameters are **automatically updated when a section type is selected**
- Customizable design parameters

---

## Excel Configuration

All parameters are controlled from the **`data.xlsm`** file.

Cell colors indicate how they should be used:

- 🔴 **Red cells**: parameters that **must be adjusted** for your truss design  
- 🟡 **Yellow cells**: parameters that **can be modified**, but it is **not recommended unless necessary**

The Excel file also contains **multiple predefined structural cross-sections**.  
When you select a section, its parameters are **automatically updated in the spreadsheet**.

⚠️ **Important:** After making any changes, **you must save the Excel file**, otherwise the new parameters will not be applied to the drawing.

---

## How It Works

The workflow is simple:

1. Define parameters in **Excel**
2. Run the **Python script**
3. The program communicates with **AutoCAD**
4. The drawing is generated automatically

The output includes:

- main connection drawing
- cross sections
- element list with lengths
- plates drawn in **top view** for additional detailing

---

## How to Run the Program

1. Open **AutoCAD**
2. Create a **new drawing**
3. Change the drawing **scale to 1:10**
4. Open **`data.xlsm`**
5. Modify the parameters according to your project
6. **Save the Excel file**
7. Run the script: `start.py`

---

## Troubleshooting

If an error occurs during drawing:

1. Delete the partially generated drawing
2. Run `start.py` again

If the drawing scale appears incorrect after generation, use the AutoCAD command: `regenw`


This should resolve the issue.

---

## Demo

VIDEO_PLACEHOLDER

---

## Notes

- The program was developed as part of a university assignment to automate repetitive engineering drawing tasks.
- It significantly reduces the time required to create a detailed truss-to-column connection drawing.

