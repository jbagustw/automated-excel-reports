#!/usr/bin/env python3
"""
Setup script for Automated Excel Report Generator
"""

import os
import sys
import subprocess
import json
from pathlib import Path

def install_requirements():
    """Install required packages"""
    print("📦 Installing requirements...")
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "-r", "requirements.txt"])
        print("✅ Requirements installed successfully!")
    except subprocess.CalledProcessError as e:
        print(f"❌ Error installing requirements: {e}")
        sys.exit(1)

def create_directories():
    """Create necessary directories"""
    print("📁 Creating directories...")
    directories = ["reports", "logs", "data"]
    
    for directory in directories:
        Path(directory).mkdir(exist_ok=True)
        print(f"✅ Created directory: {directory}")

def create_config_file():
    """Create default configuration file"""
    print("⚙️ Creating configuration file...")
    
    config = {
        "company_name": "Your Company Name",
        "report_title": "Automated Report",
        "output_directory": "reports",
        "date_format": "%Y-%m-%d",
        "colors": {
            "header": "366092",
            "subheader": "5B9BD5",
            "accent": "70AD47"
        },
        "email_settings": {
            "smtp_server": "smtp.gmail.com",
            "smtp_port": 587,
            "sender_email": "your_email@gmail.com",
            "recipients": [
                "manager@company.com"
            ]
        },
        "report_schedule": {
            "daily_time": "09:00",
            "weekly_day": "Monday",
            "weekly_time": "08:00"
        }
    }
    
    if not os.path.exists("report_config.json"):
        with open("report_config.json", "w") as f:
            json.dump(config, f, indent=4)
        print("✅ Configuration file created: report_config.json")
    else:
        print("⚠️ Configuration file already exists")

def create_sample_data():
    """Create sample data file"""
    print("📊 Creating sample data...")
    
    sample_data = """import pandas as pd
import numpy as np
from datetime import datetime, timedelta

def create_sample_sales_data():
    \"\"\"Create sample sales data for testing\"\"\"
    dates = pd.date_range(start=datetime.now() - timedelta(days=30), 
                         end=datetime.now(), freq='D')
    
    return pd.DataFrame({
        'Date': dates,
        'Product_A': np.random.randint(50, 200, len(dates)),
        'Product_B': np.random.randint(30, 150, len(dates)),
        'Product_C': np.random.randint(20, 100, len(dates)),
        'Revenue': np.random.randint(1000, 5000, len(dates)),
        'Orders': np.random.randint(10, 50, len(dates))
    })

def create_sample_customer_data():
    \"\"\"Create sample customer data for testing\"\"\"
    return pd.DataFrame({
        'Customer_ID': [f'CUST_{i:04d}' for i in range(1, 101)],
        'Customer_Name': [f'Customer {i}' for i in range(1, 101)],
        'Total_Orders': np.random.randint(1, 20, 100),
        'Total_Spent': np.random.randint(100, 10000, 100),
        'Category': np.random.choice(['Premium', 'Standard', 'Basic'], 100),
        'Last_Order_Date': pd.date_range(start=datetime.now() - timedelta(days=90), 
                                        end=datetime.now(), periods=100)
    })

if __name__ == "__main__":
    # Test the sample data creation
    sales_df = create_sample_sales_data()
    customer_df = create_sample_customer_data()
    
    print("Sales Data Sample:")
    print(sales_df.head())
    print("\\nCustomer Data Sample:")
    print(customer_df.head())
"""
    
    with open("sample_data.py", "w") as f:
        f.write(sample_data)
    print("✅ Sample data file created: sample_data.py")

def create_run_script():
    """Create easy run script"""
    print("🚀 Creating run script...")
    
    run_script = """#!/usr/bin/env python3
\"\"\"
Quick run script for generating reports
\"\"\"

import argparse
import sys
from excel_report_generator import ExcelReportGenerator

def main():
    parser = argparse.ArgumentParser(description='Generate Excel Reports')
    parser.add_argument('--type', choices=['daily', 'weekly'], default='daily',
                       help='Type of report to generate')
    parser.add_argument('--config', default='report_config.json',
                       help='Configuration file path')
    
    args = parser.parse_args()
    
    try:
        generator = ExcelReportGenerator(args.config)
        report_path = generator.generate_report(args.type)
        print(f"✅ Report generated: {report_path}")
    except Exception as e:
        print(f"❌ Error generating report: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()
"""
    
    with open("run_report.py", "w") as f:
        f.write(run_script)
    
    # Make executable on Unix systems
    if os.name != 'nt':
        os.chmod("run_report.py", 0o755)
    
    print("✅ Run script created: run_report.py")

def create_gitignore():
    """Create .gitignore file"""
    print("📝 Creating .gitignore...")
    
    gitignore_content = """# Python
__pycache__/
*.py[cod]
*$py.class
*.so
.Python
build/
develop-eggs/
dist/
downloads/
eggs/
.eggs/
lib/
lib64/
parts/
sdist/
var/
wheels/
*.egg-info/
.installed.cfg
*.egg
MANIFEST

# Virtual Environment
venv/
env/
ENV/

# IDE
.vscode/
.idea/
*.swp
*.swo
*~

# OS
.DS_Store
Thumbs.db

# Project specific
reports/
logs/
*.xlsx
*.xls
!sample_*.xlsx

# Config files with sensitive data
config_production.json
.env
"""
    
    with open(".gitignore", "w") as f:
        f.write(gitignore_content)
    print("✅ .gitignore file created")

def main():
    """Main setup function"""
    print("🔧 Setting up Automated Excel Report Generator...")
    print("=" * 50)
    
    # Check Python version
    if sys.version_info < (3, 7):
        print("❌ Python 3.7 or higher is required")
        sys.exit(1)
    
    # Run setup steps
    install_requirements()
    create_directories()
    create_config_file()
    create_sample_data()
    create_run_script()
    create_gitignore()
    
    print("\\n🎉 Setup completed successfully!")
    print("\\n📋 Next steps:")
    print("1. Edit 'report_config.json' with your company details")
    print("2. Run 'python excel_report_generator.py' to test")
    print("3. Run 'python run_report.py --type daily' for quick reports")
    print("4. Check the 'reports' folder for generated files")
    print("\\n🚀 Happy reporting!")

if __name__ == "__main__":
    main()
