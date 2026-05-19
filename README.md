<div align="center">

<img src="https://upload.wikimedia.org/wikipedia/commons/f/fa/Microsoft_Azure.svg" width="80" alt="Azure Logo"/>

# Azure Route Table — JSON to Excel

**Convert Azure Route Table JSON exports into clean, formatted Excel reports — available as a web app, a desktop script, or a Jupyter notebook.**

[![Live App](https://img.shields.io/badge/🚀%20Live%20App-Streamlit-FF4B4B?style=for-the-badge)](https://azure-rt-json2excel.streamlit.app/)
[![Python](https://img.shields.io/badge/Python-3.9%2B-3776AB?style=for-the-badge&logo=python&logoColor=white)](https://www.python.org/)
[![License](https://img.shields.io/badge/License-Apache%202.0-blue?style=for-the-badge)](LICENSE)
[![Streamlit](https://img.shields.io/badge/Built%20with-Streamlit-FF4B4B?style=for-the-badge&logo=streamlit&logoColor=white)](https://streamlit.io/)

</div>

---

## Overview

Azure Route Tables control how traffic flows through your Virtual Networks — but reviewing UDR (User Defined Route) configurations and associated subnet bindings is tedious when you're stuck reading raw JSON from the Portal or CLI. This tool parses an Azure Route Table JSON export and produces a structured, formatted Excel workbook covering route entries, subnet associations, and resource metadata.

Available in three forms to suit any workflow: a hosted web app, a local Python script with a GUI file picker, and a Jupyter notebook for interactive exploration.

**Try it live →** [azure-rt-json2excel.streamlit.app](https://azure-rt-json2excel.streamlit.app/)

---

## Features

- **Three usage modes** — hosted web app, standalone Python desktop script, and Jupyter notebook
- **Full route extraction** — name, address prefix, next hop type, and next hop IP address for every UDR
- **Subnet association table** — lists each bound subnet with its address range, parent VNet, and attached NSG
- **Resource metadata header** — captures Route Table name, Resource Group, Location (human-readable region name), and Subscription ID
- **Formatted Excel output** — titled workbook with colour-coded headers, section dividers, borders, and auto-sized columns
- **Comprehensive region mapping** — converts Azure location codes (e.g. `australiaeast`) to friendly names across all regions globally
- **Live data preview** — view all three tables directly in the browser before downloading (Streamlit app)

---

## Excel Output Structure

The generated workbook (`ROUTE_TABLE` sheet) is organised into three sections:

| Section | Contents |
|---|---|
| **Metadata** | Route Table name, Resource Group, Location, Subscription ID |
| **ROUTES** | Name · Address Prefix · Next Hop Type · Next Hop IP Address |
| **SUBNETS** | Subnet Name · Address Range · Virtual Network · Security Group |

---

## Getting Started

### Option 1 — Hosted Web App (no setup)

Visit **[azure-rt-json2excel.streamlit.app](https://azure-rt-json2excel.streamlit.app/)**, upload your JSON file, preview the tables, and download the Excel report.

---

### Option 2 — Run Locally (Streamlit)

**Prerequisites:** Python 3.9+

```bash
# 1. Clone the repository
git clone https://github.com/HashimsGitHub/Azure--Route-Table_JSON_to_Excel.git
cd Azure--Route-Table_JSON_to_Excel

# 2. Install dependencies
pip install -r requirements.txt

# 3. Launch the app
streamlit run streamlit_app.py
```

Opens at `http://localhost:8501`.

---

### Option 3 — Desktop Script (GUI file picker)

For users who prefer working locally without a browser. Uses a Tkinter file dialog to select the input JSON and choose where to save the output.

```bash
pip install -r requirements.txt
python Azure_Route_Table_JSON_to_Excel.py
```

A file picker will prompt for the input `.json` file, then a save dialog will prompt for the output `.xlsx` path.

---

### Option 4 — Jupyter Notebook

Open `Azure_Route_Table_JSON_to_Excel.ipynb` for an interactive, cell-by-cell walkthrough of the parsing and export logic — useful for customisation or learning.

```bash
pip install -r requirements.txt
jupyter notebook Azure_Route_Table_JSON_to_Excel.ipynb
```

---

## How to Export an Azure Route Table (JSON)

Use the Azure CLI to export the JSON file this tool expects:

```bash
# Export a specific route table by name and resource group
az network route-table show \
  --name <RouteTableName> \
  --resource-group <ResourceGroupName> \
  --output json > route_table.json

# Export all route tables in a resource group
az network route-table list \
  --resource-group <ResourceGroupName> \
  --output json > all_route_tables.json

# Export all route tables in a subscription
az network route-table list --output json > all_route_tables.json
```

Then upload or point the script at the resulting `.json` file.

---

## Project Structure

```
Azure--Route-Table_JSON_to_Excel/
├── streamlit_app.py                       # Streamlit web application
├── Azure_Route_Table_JSON_to_Excel.py     # Standalone desktop script (GUI file picker)
├── Azure_Route_Table_JSON_to_Excel.ipynb  # Jupyter notebook
├── requirements.txt                       # Python dependencies
├── .devcontainer/                         # Dev container configuration
├── .github/                               # GitHub Actions workflows
├── .gitignore
└── LICENSE
```

---

## Dependencies

| Package | Purpose |
|---|---|
| `streamlit` | Web application framework |
| `pandas` | DataFrame creation and manipulation |
| `openpyxl` | Excel workbook generation and styling |

Install all with:

```bash
pip install -r requirements.txt
```

> The desktop script additionally uses `tkinter` for the GUI file picker, which is bundled with standard Python installations.

---

## Contributing

Contributions are welcome! To contribute:

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/your-feature`)
3. Commit your changes (`git commit -m 'Add your feature'`)
4. Push to the branch (`git push origin feature/your-feature`)
5. Open a Pull Request

Ideas for future improvements: multi-route-table batch processing, NSG rule extraction alongside subnets, BGP route summary support, and VNet peering visualisation.

---

## License

Distributed under the [Apache 2.0 License](LICENSE).

---

<div align="center">

Built with ❤️ using Streamlit & Microsoft Azure

**Hashim Hilal** — Cloud Architect

</div>
