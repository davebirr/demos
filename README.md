# M365 Lab Demos Repository

This repository contains PowerShell scripts and sample data for setting up Microsoft 365 hands-on lab environments.

## 📁 Repository Structure

```
├── Scripts/
│   ├── ConditionalAccess/     # Conditional Access policy scripts
│   ├── Intune/               # Microsoft Intune configuration scripts
│   ├── Lab-Setup/            # General lab setup and configuration scripts
│   ├── UserManagement/       # User and object import/management scripts
│   └── Vendor/               # Third-party vendor scripts (Datto, Huntress, etc.)
├── SampleData/               # Sample CSV files for demos and testing
├── Reports/                  # Generated reports and documentation
├── upload-users.ps1          # Legacy user upload script (see UserManagement folder for organized version)
└── README.md                 # This file
```

## 🚀 Key Scripts

### User Management
- **`Import_m365bUsers.ps1`** - Imports demo users into M365 tenants from CSV
- **`Import_m365bObjects.ps1`** - Imports users, groups, and contacts into on-premises AD

### Lab Setup
- **`AssignM365Licenses.ps1`** - Automates license assignment for demo users
- **`create_teams_newtenant.ps1`** - Sets up Teams for new tenant demos
- **`create_w365.ps1`** - Configures Windows 365 Cloud PC environments

### Security & Compliance
- **`Baseline-ConditionalAccessPolicies.ps1`** - Creates recommended baseline CA policies
- **`setup-intune.ps1`** - Imports baseline Intune configurations for compliance and device management

## 📋 Prerequisites

- PowerShell 5.1 or PowerShell 7+
- Microsoft Graph PowerShell modules
- Azure AD PowerShell module (for legacy scripts)
- Appropriate admin permissions in target M365 tenants

## 🔧 Usage

1. **For User Import:**
   ```powershell
   # Edit the CSV file with your demo users
   # Run the import script
   .\Scripts\UserManagement\Import_m365bUsers.ps1
   ```

2. **For Lab Setup:**
   ```powershell
   # Connect to your M365 tenant first
   Connect-MgGraph -Scopes "Directory.ReadWrite.All"
   
   # Run desired setup scripts
   .\Scripts\Lab-Setup\AssignM365Licenses.ps1
   ```

## 📊 Sample Data

The `SampleData/` folder contains CSV files with sample user data for demos:
- `m365bUsers.csv` - Sample users with various departments and roles

## ⚠️ Important Notes

1. **Security**: Never commit actual tenant IDs, API keys, or other sensitive information
2. **Testing**: Always test scripts in a demo/dev tenant before production use
3. **Permissions**: Ensure you have appropriate admin rights before running scripts
4. **Backup**: Consider backing up existing configurations before making changes

## 🤝 Contributing

When contributing new scripts:
1. Follow PowerShell best practices
2. Include parameter validation and error handling
3. Add comprehensive comments and help documentation
4. Test thoroughly in isolated environments
5. Remove any sensitive information before committing

## 📝 License

These scripts are provided as-is for educational and demonstration purposes.

---

*Last updated: October 2025*
