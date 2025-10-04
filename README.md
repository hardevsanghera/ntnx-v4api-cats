# 🏷️ Nutanix v4 APIs: Working with Categories via Microsoft Excel

> **Automate VM category management using PowerShell, Python, and Microsoft Excel with Nutanix v4 APIs**

## 📋 Overview

This project provides a collection of scripts to manage VM categories in Nutanix environments through REST API calls. It combines PowerShell automation with Excel integration to create an educational workflow for understanding Nutanix v4 API interactions.

## 🎯 Objectives

- **Educational Focus**: Demonstrate REST API usage with Nutanix v4 APIs
- **Excel Integration**: Leverage Microsoft Excel as an automation interface
- **Workflow Breakdown**: Separate scripts to reduce the learning curve
- **Practical Implementation**: Real-world category management scenarios

## 🏗️ Architecture

- **Target Environment**: Nutanix clusters with AHV hypervisor
- **API Version**: Nutanix v4 REST APIs  
- **Primary Language**: PowerShell 7
- **Secondary Language**: Python (for reliable POST operations)
- **Interface**: Microsoft Excel via COM automation

## ⚠️ Prerequisites

- ✅ **Microsoft Excel** installed (COM interop required)
- ✅ **PowerShell 7** (pwsh)
- ✅ **Python 3.x** (for update operations)
- ✅ **Nutanix Prism Central** access
- ⚠️ **Security Note**: Scripts use plain-text passwords - modify for production use

## 📚 Documentation

📖 **Educational Resource**: [`files/educate.pdf`](files/educate.pdf) - REST and APIs fundamentals

## 🚀 Quick Start

### Step 1: Configuration
```powershell
# Edit configuration file with your Prism Central details
notepad files\vars.txt
```

### Step 2: Data Collection
```powershell
# Collect VM information
.\list_vms.ps1

# Collect category definitions  
.\list_categories.ps1
```

### Step 3: Build Workbook
```powershell
# Generate Excel workbook with VM-category mappings
.\build_workbook.ps1
```

### Step 4: Excel Operations
1. Open `VMsToUpdate.xlsx` in Microsoft Excel
2. Navigate to the **"ToUpdate"** sheet
3. Add entries with:
   - **VM Name**
   - **VM extID** 
   - **Categories** to associate
4. **Save and close** the workbook
Screenshot, the "ToUpdate" sheet when first opened:
<img src="files/excel1.png" alt="REST Slide" width="500">
Screenshot, now with VM, extID, Update Categories
<img src="files/excel2.png" alt="REST Slide" width="500">

### Step 5: Validate category update parameters
```powershell
# Validate parameters
.\update_vm_categories.ps1
```

### Step 6: Review Results
Open `VMsToUpdate.xlsx` to examine the status of the parameter validations.
Screenshot, the "ToUpdate" sheet with validated parameters:
<img src="files/excel3.png" alt="REST Slide" width="500">

### Step 7: Apply Updates
```python
# Execute category updates via PYTHON
python update_vm_categories_for_vm.py
```

### Step 8: Review Results
Open `VMsToUpdate.xlsx` to examine the status of category associations.
Screenshot, the "ToUpdate" sheet with status of the VM update:
<img src="files/excel4.png" alt="REST Slide" width="500">

## 📁 Project Structure

```
ntnx-v4api-cats/
├── 📄 README.md                      # This documentation
├── 📄 list_vms.ps1                   # VM discovery script
├── 📄 list_categories.ps1            # Category enumeration script  
├── 📄 build_workbook.ps1             # Excel workbook generator
├── 📄 update_vm_categories.ps1       # PowerShell update script
├── 🐍 update_categories_for_vm.py    # Python update script
├── 📂 files/
│   ├── 📄 vars.txt                   # Configuration file
│   ├── 📄 requirements.txt           # Python dependencies
│   ├── 📊 VMsToUpdate_SKEL.xlsx      # Excel template
│   └── 📖 educate.pdf               # Educational documentation
└── 📂 scratch/                       # Output directory
    ├── 📄 vm_list.json               # VM data export
    ├── 📄 categories.json            # Category definitions
    └── 📊 cat_map.xlsx               # Category mappings
```

## 🔧 Technical Notes

- **PowerShell Limitation**: Inconsistent POST call behavior led to Python implementation for updates
- **COM Integration**: Excel automation requires local Microsoft Office installation
- **Output Locations**: All generated files are saved to `scratch/` directory
- **API Approach**: Educational focus prioritizes clarity over optimization

## 👨‍💻 Author

**Hardev Sanghera** - [hardev@nutanix.com](mailto:hardev@nutanix.com)

*October 2025*

---

## 🤝 Contributing

This project is designed for educational purposes. Feel free to:
- 🍴 Fork the repository
- 🔧 Customize scripts for your environment  
- 📝 Improve documentation
- 🛡️ Enhance security implementations

## 📄 License

This project is provided "AS IS" for educational purposes. Use at your own risk.
