![](https://media0.giphy.com/media/5DQdk5oZzNgGc/giphy.gif?cid=ecf05e47230c90092821aca8e99bb73074786bad69cb798b&rid=giphy.gif)
# Office 365: Inventory by tenant

This script allows you to quickly inventory all your partner tenants and its users into a CSV file containing useful information such as:
* **UserPrincipalName**
* **DisplayName**
* **FirstName** and **LastName**
* **Department**
* **Licenses** (in friendly names)
* **WhenCreated**

## ⚠️ Requirements
* **MSOnline** - [Connect-MsolService (MSOnline) | Microsoft Docs](https://docs.microsoft.com/en-us/powershell/module/msonline/connect-msolservice?view=azureadps-1.0)
* **PowerShell 5.1**
  * Might work with previous versions, but is confirmed working on 5.1.

## ⚙️ Installation

1. Download (or copy + paste) *code.ps1* to your local computer.
2. Open it with PowerShell ISE, or edit it accordingly and run it directly with PowerShell.
3. Edit the following:
   1. **$PathToStoreCsv** 
      - full path to folder to store everything in, remember a \ on the end!
   2. **$SortCsvByDepartment** (optional) 
      - $true/$false whether to create different csv for each deparment found
   3. **$GetOnlyLicensedUsers** (optional) 
      - $true/$false whether to get licensed users only or all users
   4. **$GetSpecificTenantByDomainName** (optional) 
      - full name of a verified domain, only retrieves users for that tenant
4. Press *Run Script* (F5) and enter your partner account details.
5. Watch the magic work!

## 🏷️ Features
* Automatically retrieves tenants and its users and outputs all information mentioned above to a .CSV
* Automatically sorts users by department and creates/appends a .CSV for each department with
* Inventory all tenants or one specific tenant with
* Inventory only licensed or all users

