# 🖥️ Asset Management Tool

This project is an **Asset Management Tool** built with **Python**, designed to parse asset report `.html` files and convert them into structured **Excel reports**.

It extracts system, hardware, and employee-related details from the provided HTML file and organizes them into a clean Excel sheet for easy reporting and analysis.

---

## 🚀 Features
- Extracts metadata from the HTML filename (System Name, Department, Employee, etc.).
- Parses hardware and system details from structured HTML sections.
- Captures key information such as:
  - System Model & Service Tag
  - Processor & Motherboard
  - HDD & RAM details
  - Operating System & Display details
  - Monitor information
- Outputs a neatly formatted **Excel file** with auto-adjusted column widths.
---

## 📂 Project Structure

```bash
Asset-Management-Tool/
│── app.py               # Main script
│── README.md            # Project documentation
│── requirements.txt     # Python dependencies (to be created)
│── sample.html          # Example input (add your asset report HTML)
│── output.xlsx          # Example generated Excel output
```

## ⚙️ Installation

1. Clone the Repository
```bash
git clone https://github.com/yourusername/asset-management-tool.git
cd asset-management-tool
```
2. Create a virtual environment (recommended):
```bash
python -m venv venv
source venv/bin/activate   # On Mac/Linux
venv\Scripts\activate      # On Windows
```
3. Install dependencies:
```bash
pip install -r requirements.txt
```
## 📦 Dependencies

This project requires the following Python libraries:
- pandas
- beautifulsoup4
- openpyxl
Create a requirements.txt file with:
```nginx
pandas
beautifulsoup4
openpyxl
```
## ▶️ Usage

Run the script with:
```bash
python app.py
```
Example function call inside app.py:
```python
hlo('SYMXC039_ESD_K.SURESH_VSPITP_1B_D69.html', 'output.xlsx')
```
- Input: SYMXC039_ESD_K.SURESH_VSPITP_1B_D69.html (asset report in HTML format)
- Output: output.xlsx (formatted Excel file)

## 📊 Output Example:

| SYSTEM NAME | DEPARTMENT | EMPLOYEE NAME | LOCATION | WORK PLACE | PORT NUMBER | COMPUTER NAME | OPERATING SYSTEM | SYSTEM MODEL  | SERVICE TAG | PROCESSOR | BOARD | HDD (GB) | MEMORY (MB) | RAM SLOTS | DISPLAY  | MONITOR     |
| ----------- | ---------- | ------------- | -------- | ---------- | ----------- | ------------- | ---------------- | ------------- | ----------- | --------- | ----- | -------- | ----------- | --------- | -------- | ----------- |
| SYMXC039    | ESD        | K.SURESH      | VSPITP   | 1B         | D69         | PC-12345      | Windows 10       | Dell OptiPlex | XYZ12345    | Intel i5  | Intel | 512      | 8192        | 2         | Dell 22" | Dell P2219H |

## 🛠️ Future Enhancements

- Support for multiple HTML files in batch mode.
- Add a GUI interface for easier usage.
- Export reports in CSV and PDF formats.
