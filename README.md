# Creating Family Trees using MS Access
 Do note that this method can only be used in MS Access 2007 because the ActiveX control named Microsoft TreeView Control has been discontinued in later versions of MS Office. The versions of Microsoft Windows released after Windows 7 do not contain the required MSCOMCTL.OCX file required to run Microsoft TreeView Control. So, we need to get them before installing MS Office 2007.

## Getting MSCOMCTL.OCX File
1. Go to this [link from Microsoft](https://www.microsoft.com/en-us/download/details.aspx?id=10019) and download Microsoft Visual Basic 6.0 Common Controls.
2. Click on Download to download on your PC.
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
04. Populate your table with some arbitary data.
05. Create a new form named "frmFamilyTree" in Design View.
06. In Form Design View, make sure the "Control Toolbox" is visible (usually on the left side, or go to "View" > "Toolbox"). Click on the "More Controls" button in the Toolbox (it looks like a hammer and wrench). In the "Insert ActiveX Control" dialog, scroll down and select "Microsoft TreeView Control, version 6.0 (SP6)". Click "OK".

07. Click on your form to place the TreeView control.
    - Important: Check the name of your TreeView control. Select the TreeView, and in the "Property Sheet" (if not visible, go to "View" > "Property Sheet"), look at the "Name" property under the "Other" tab.  If it's not "TreeView0", replace Me.TreeView0 in the VBA code with the actual name of your TreeView control.

08. Right click on TreeViewControl area on your form.
09. Go to TreeView Object > Properties. Change LineStyle to 1-tvwRootLines.
10. In Form Design View, click on "View" > "Code" (or press Alt + F11 to open the VBA editor directly).
11. In the VBA editor, paste the code provided in the "Form_frmFamilyTree" file of this repository.
12. Save your form. Switch to Form View to run the form. The TreeView should now be populated with your family tree data.