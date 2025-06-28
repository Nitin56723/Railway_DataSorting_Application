Crew Duty Filter
A Python GUI application built using Tkinter for filtering Indian Railways crew duty Excel data based on specific station zones. The app reads Excel files, filters crew duty information by sign-on/sign-off stations, and generates new structured Excel files with additional summaries and counts.

🧩 Features:-


  📁 Load crew data from Excel files.

  🔍 Filter based on predefined station zones.

  📝 Edit or add new zones dynamically (GUI-based).

  📊 Auto-generates filtered Excel reports:

  From-Zone and To-Zone files.

  SP/WR duty counts and summaries.

  💾 Persist zone data in a stations.json file.

  📐 Automatically adjusts Excel column widths.

📸 GUI Preview:-

  Main Window: File selection and action buttons.

  Add Station: Add new zone name, zone code, and mapped stations.

  View/Edit Station: Modify or review existing zone mappings.

🛠️ Installation:-


Requirements
Python 3.7+

Required libraries:

pip install pandas openpyxl
Optional (For .exe build via PyInstaller)

pip install pyinstaller


🚀 How to Run
Clone the repository:

git clone https://github.com/yourusername/crew-duty-filter.git
cd crew-duty-filter


Run the application:


python crew_duty_filter.py

📂 Input Excel Format


The Excel file must contain a row with "S.No." or "S.No" in any column to identify headers.

Columns used:

S.No.

CREW ID

SIGNON STTN

SIGNOFF STTN

DUTY TYPE (Should contain values like "SP" and "WR")

📤 Output
For each zone, two files are generated:

From<ZoneName>.xlsx

To<ZoneName>.xlsx

Each file contains:

Filtered rows based on relevant station mappings.

Counts of SP and WR duty types.

A table summary of SP duties at the bottom.

📘 Zone Definitions (stations.json)
Stores:

{
  "fg": ["Erode", "Jolarpettai"],
  "dt": {
    "ED": ["ED", "TPMR", "PYR", "..."],
    "JTJ": ["JTJ", "TPT", "KEY", "..."]
  }
}
Can be updated via "Add Station" or "View Station" GUI buttons.

💡 Use Cases
Railway crew duty filtering and reporting

Automating Excel report generation

Educational GUI projects in Python

📦 Packaging (Optional)
To convert into a .exe for Windows:


pyinstaller --onefile --add-data "stations.json;." crew_duty_filter.py


📄 License
This project is licensed under the MIT License.
