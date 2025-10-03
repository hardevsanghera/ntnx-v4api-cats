# ntnx-v4api-cats
Use Windows Powershell (and 1 python script) to add/update Categories to a VM(s) via Microsoft Excel.
The target environment is Nutanix clusters with the AHV hypervisor using REST with 
the v4 APIs.

You will need Microsoft Excel installed on your workstation as the COM thingy is used.

The objective here is to "educate" on the use of REST calls and Nutanix v4 APIs with the use of
Microsoft Excel to aid as an automation hack.  It's not necessarily best practice and the workflows
can certainly be consolidated, however the separate script approach does break down the "curve
of understanding".

The original goal was to create Powershell 7 scripts only, however I could not get
Powershell to consistently make successful POST calls, in Python I could, so that's why there's a single
Python script.


**NOTE: These scripts use plain-text passwords, feel free to change that! ***

Educational document re: REST and APIs <code_dir>\files\educate.pdf

Usage:
    1. Edit <code_dir>\files\vars.txt to point to your Prism Central (PC), note plain-text password
    2. Run <code_dir>\list_vms.ps1 (writes output to console and <code_dir>\scratch)
    3. Run <code_dir>\list_categories.ps1 (writes output to console and <code_dir>\scratch)
    4. Run <code_dir>\build_workbook.ps1 (writes output to console and <code_dir>\scratch)
    5. Run <code_dir>\build_workbook.ps1 (writes output to console and <code_dir>\scratch), notably VMsToUpdate.xlsx
    6. Open the <code_dir>\VMsToUpdate.xlsx with Excel and then populate the "ToUpdate" sheet to add:
        VM Name, VM extID, Categories to associate with the VM
       Save and CLOSE the workbook
    7. Run <code_dir>\update_vm_categories_for_vm.py
       This will write output to console and update <code_dir>\VMsToUpdate.xlsx with the status of the category associations.
    8. Open <code_dir>\VMsToUpdate.xlsx and examine

hardev@nutanix.com
Oct '25
