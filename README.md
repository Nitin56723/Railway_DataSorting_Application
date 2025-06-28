# 🚆 Crew Duty Filter – Excel Automation for Railway Data

This project is a **Tkinter-based desktop application** for filtering crew duty Excel sheets based on zone-specific station mappings. It helps automate the generation of structured reports (From/To) for Indian Railways or similar use cases.

---

## 📌 Project Objectives

- Filter large Excel datasets based on zone-wise crew movements.
- Generate two separate Excel files (`FromZone.xlsx` and `ToZone.xlsx`) for each region.
- Count and summarize `SP` and `WR` duty types.
- Allow easy addition and editing of zones and station mappings through a GUI.

---

## 📁 Input Data Requirements

- **Input Format**: Excel (.xlsx or .xls)
- **Header Row**: Should contain `S.No.` or `S.No`
- **Required Columns**:
  - `S.No.`
  - `CREW ID`
  - `SIGNON STTN`
  - `SIGNOFF STTN`
  - `DUTY TYPE` (values like `SP` and `WR`)

---

## 🔧 Tools & Technologies

- **Python**
- **Tkinter** for the GUI
- **Pandas** for data processing
- **OpenPyXL** for Excel file manipulation
- **JSON** for saving station mappings
- **itertools.product** for pair-wise station analysis

---

## 🧠 Core Features

- 📂 Select and load Excel file from GUI  
- 🧪 Automatically detect header row  
- 📤 Generate `FromZone` and `ToZone` Excel files for each region  
- 📊 Append duty counts and SP duty summaries to each sheet  
- 📝 Add/Edit/Delete zones and mapped stations via interface  
- 💾 Station mappings persist in `stations.json`  

---

## 📊 Example Output

Each output file (`FromErode.xlsx`, `ToJolarpettai.xlsx`, etc.) contains:

- Filtered crew records relevant to the selected zone
- Summary rows:
  - `SP COUNT`
  - `WR COUNT`
- Table of SP duties:


---

## ▶️ How to Run

1. **Clone the repository**:

```bash
git clone https://github.com/yourusername/crew-duty-filter.git
cd crew-duty-filter


