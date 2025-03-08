# Creating Family Trees using MS Access
![Family Tree Image](/FamilyTree_Sample.png "This is how the family tree looks!")
![Family Table Image](/FamilyTree_Table.png "This is how the Family table looks.")

This method is only compatible with MS Access 2007 because the ActiveX control, Microsoft TreeView Control, is not available in later MS Office versions. Furthermore, Microsoft Windows versions after Windows 7 lack the MSCOMCTL.OCX file needed to run the TreeView Control. Therefore, you must obtain this file before installing MS Office 2007.

## Getting MSCOMCTL.OCX File
1. Go to this [link from Microsoft](https://www.microsoft.com/en-us/download/details.aspx?id=10019) and download Microsoft Visual Basic 6.0 Common Controls.
2. Alternatively, you can download MSCOMCTL.OCX from [GeekPage](https://thegeekpage.com/wp-content/uploads/2020/05/mscomctl.zip).
3. Right click and extract downloaded file using WinZip or any other extracting applications.
4. Copy MSCOMCTL.OCX file.
5. Navigate to the below path in C drive,
   - If your Computer is 64 bit- C:\Windows\SysWOW64
   - If your Computer is 32 bit- C:\Windows\System32
6. Paste the file.

## Working on MS Access 2007
01. Install MS Office 2007.
02. Open MS Access 2007. Create a database named "Fam."
03. Create a table named "Family." This table will have 3 columns named ID, NameID and ParentID with data types as Number, Text and Number, respectively.
04. Populate the table with your family data.
    - The progenitor (or first ancestor) will have the ParentID of 0. This will be considered as the root node.
    - Ensure that your ParentID values in the table correctly refer to existing ID values in the same table to create valid parent-child relationships
05. Create a new form named "frmFamilyTree" in Design View.
06. In Form Design View, make sure the "Control Toolbox" is visible (usually on the left side, or go to "View" > "Toolbox"). Click on the "More Controls" button in the Toolbox (it looks like a hammer and wrench). In the "Insert ActiveX Control" dialog, scroll down and select "Microsoft TreeView Control, version 6.0 (SP6)". Click "OK".

07. Click on your form to place the TreeView control.
    - Important: Check the name of your TreeView control. Select the TreeView, and in the "Property Sheet" (if not visible, go to "View" > "Property Sheet"), look at the "Name" property under the "Other" tab.  If it's not "TreeView0", replace Me.TreeView0 in the VBA code with the actual name of your TreeView control.

08. Right click on TreeViewControl area on your form.
09. Go to TreeView Object > Properties. Change LineStyle to 1-tvwRootLines.
10. In Form Design View, click on "View" > "Code" (or press Alt + F11 to open the VBA editor directly).
11. In the VBA editor, paste the code provided in the "Form_frmFamilyTree.cls" file of this repository.
12. Save your form. Switch to Form View to run the form. The TreeView should now be populated with your family tree data.

---