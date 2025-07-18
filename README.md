# 📊 Automated Excel Report Generator

A Python script that automatically generates professional Excel reports with formatting, charts, and multiple sheets using pandas and openpyxl.

## ✨ Features

- **Automated Report Generation**: Creates daily/weekly reports automatically
- **Professional Formatting**: Beautiful styling with headers, borders, and colors
- **Multiple Sheets**: Summary sheet with key metrics + detailed data sheets
- **Configurable**: JSON configuration file for easy customization
- **Sample Data**: Built-in sample data generator for testing
- **Charts Support**: Automatic chart generation (Bar and Line charts)
- **Auto-sizing**: Columns automatically adjust to content width

## 🚀 Quick Start

### Prerequisites

- Python 3.7+
- pip package manager

### Installation

1. Clone this repository:
```bash
git clone https://github.com/yourusername/automated-excel-reports.git
cd automated-excel-reports
```

2. Install dependencies:
```bash
pip install -r requirements.txt
```

### Usage

#### Basic Usage
```python
from excel_report_generator import ExcelReportGenerator

# Initialize generator
generator = ExcelReportGenerator()

# Generate daily report
daily_report = generator.generate_report("daily")
print(f"Report saved: {daily_report}")

# Generate weekly report
weekly_report = generator.generate_report("weekly")
print(f"Report saved: {weekly_report}")
```

#### Using Custom Data
```python
import pandas as pd

# Prepare your data
custom_data = {
    'sales_data': pd.DataFrame({
        'Date': ['2024-01-01', '2024-01-02'],
        'Revenue': [1000, 1500],
        'Orders': [25, 30]
    }),
    'customer_data': pd.DataFrame({
        'Customer': ['ABC Corp', 'XYZ Ltd'],
        'Total_Spent': [5000, 3000]
    })
}

# Generate report with custom data
generator = ExcelReportGenerator()
report_path = generator.generate_report("daily", data=custom_data)
```

#### Run the Demo
```bash
python excel_report_generator.py
```

## ⚙️ Configuration

The script uses a `report_config.json` file for customization:

```json
{
    "company_name": "Your Company Name",
    "report_title": "Automated Report",
    "output_directory": "reports",
    "date_format": "%Y-%m-%d",
    "colors": {
        "header": "366092",
        "subheader": "5B9BD5",
        "accent": "70AD47"
    }
}
```

## 📁 Project Structure

```
automated-excel-reports/
├── excel_report_generator.py    # Main script
├── requirements.txt             # Dependencies
├── report_config.json          # Configuration file (auto-generated)
├── reports/                    # Generated reports folder
└── README.md                   # This file
```

## 🔧 Customization

### Adding New Report Types

```python
def generate_custom_data(self):
    """Add your custom data generation logic"""
    return {
        'your_data': pd.DataFrame({
            'column1': ['value1', 'value2'],
            'column2': [100, 200]
        })
    }

# Use in main script
generator = ExcelReportGenerator()
custom_data = generator.generate_custom_data()
report = generator.generate_report("custom", data=custom_data)
```

### Modifying Styles

Edit the `colors` section in `report_config.json`:
- `header`: Main header background color
- `subheader`: Secondary header color
- `accent`: Accent color for highlights

### Adding Charts

The script supports automatic chart generation. Modify the `_add_chart` method to customize chart types and positioning.

## 📊 Sample Output

The generated Excel files include:

1. **Summary Sheet**: Key metrics and overview
2. **Sales Data Sheet**: Detailed sales information with formatting
3. **Customer Data Sheet**: Customer analysis and statistics

Features:
- Professional styling with alternating row colors
- Auto-sized columns
- Bordered tables
- Header formatting
- Summary statistics

## 🛠️ Advanced Usage

### Scheduling Reports

Create a batch script or use cron (Linux/Mac) or Task Scheduler (Windows) to run reports automatically:

```bash
# Example cron job for daily reports at 9 AM
0 9 * * * /usr/bin/python3 /path/to/excel_report_generator.py
```

### Integration with Data Sources

```python
# Example: Reading from database
import sqlite3

def load_data_from_db():
    conn = sqlite3.connect('your_database.db')
    sales_data = pd.read_sql_query('SELECT * FROM sales', conn)
    customer_data = pd.read_sql_query('SELECT * FROM customers', conn)
    conn.close()
    
    return {
        'sales_data': sales_data,
        'customer_data': customer_data
    }

# Use with report generator
data = load_data_from_db()
generator = ExcelReportGenerator()
report = generator.generate_report("daily", data=data)
```

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add some amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## 📝 License

This project is licensed under the MIT License - see the LICENSE file for details.

## 🆘 Support

If you encounter any issues or have questions:

1. Check the [Issues](https://github.com/yourusername/automated-excel-reports/issues) page
2. Create a new issue with detailed description
3. Include sample data and error messages if applicable

## 🎯 Future Enhancements

- [ ] Email automation for report distribution
- [ ] Dashboard integration
- [ ] More chart types (Pie, Scatter, etc.)
- [ ] Template system for different report layouts
- [ ] Data validation and error handling
- [ ] Export to other formats (PDF, CSV)

---

Made with ❤️ by [Your Name]

⭐ If you found this useful, please give it a star!