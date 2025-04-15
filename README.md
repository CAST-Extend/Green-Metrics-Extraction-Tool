
# 🌱 CAST Highlight Green Metrics Extractor

This Python script extracts **Green Impact Metrics** for a specific application from **CAST Highlight** using its REST API. The data is then exported to an Excel file with calculated efforts and placeholders for cost-based tech debt estimation.

---

## 📦 Features

- Fetches **Green Impact** metric data for a specific domain/application.
- Extracts rule occurrences, associated technologies, and estimated effort.
- Exports results to a styled Excel file with formulas and total rows.
- Adds placeholders for manual cost input to compute **Tech Debt ($)**.
- Adds Excel formulas to compute **Tech Debt** automatically when costs are filled.

---

## 🛠️ Prerequisites

- Python 3.8+
- `pip install` the following packages:
  ```bash
  pip install requests pandas openpyxl
  ```

---

## 🔧 Configuration

Create a `config.json` file in the same directory with the following format:

```json
{
  "HLInstance": "your_instance_name",
  "domain_id": 123456,
  "application_id": 789012,
  "api_key": "your_api_key_here"
}
```

> 🔐 Keep your API key secure and do not share it publicly.

---

## ▶️ Usage

Run the script via command line:

```bash
python green_metrics_extractor.py
```

If the configuration is valid and data is available, you'll get an Excel file in the `output/` directory.

---

## 📊 Excel Output

The generated Excel file includes:

| Rule/Pattern | Technology | Number of Occurrences | Effort by Occurrence (Person-day) | Cost (FTE/Day) | Tech Debt ($) Effort x Cost |
|--------------|------------|------------------------|-----------------------------------|----------------|-----------------------------|
| ...          | ...        | ...                    | ...                               | (Enter Manually)| (Auto-calculated in Excel)  |

- **Effort by Occurrence**: Rounded person-day estimation (1 day = 480 mins).
- **Cost (FTE/Day)**: Placeholder column for your input.
- **Tech Debt ($)**: Excel formula auto-calculates this using `Effort * Cost`.
- **Total Row**: Summed values and a total formula for tech debt.

---

## 📁 Output File Naming

Files are saved in the format:

```
output/green_metrics_d{domain_id}_a{application_id}_{timestamp}.xlsx
```

---

## 💡 Notes

- Only rules with **non-zero occurrences** are included.
- The script automatically handles empty or malformed API responses.
- Make sure your API key has access to the selected domain and application.

---

## ❓ Troubleshooting

- **Config load failure**: Check the format and existence of `config.json`.
- **API request failed**: Verify your instance name, domain/app IDs, and API key.
- **No metrics or data**: Ensure the selected application has Green Impact analysis.

---

## 📄 License

This script is intended for internal use. Modify and adapt it as needed for your workflows.
