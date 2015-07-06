Sub update()
    'Remarks: The code included in ' ' means outdated and replaced by others
    '         Command after ' is my own comment
    
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    Set report_book = ActiveWorkbook  'Set is needed as ActiveWorkbook is a Object
    Set report_console_sheet = ActiveSheet
    Set report_member_sheet = Worksheets("Member_Profile")  'Need update here once changed the worksheet name
    Set report_trans_sheet = Worksheets("Raw_data")
    Set report_report_sheet = Worksheets("Report")
    'Windows("Members.xlsx").Activate'
    
    
    ChDir ActiveWorkbook.Path  'Set the default location for import file

    'Get the Filename of Member.xlsx
    member_file = Application.GetOpenFilename(fileFilter:="xlsx files (*.xlsx),*.xlsx", Title:="Import the Members.xlsx")
    If member_file = False Then
        Exit Sub
    End If
    Workbooks.Open (member_file)
    Set member_book = ActiveWorkbook  'Avoid hardcode again
    Set member_sheet = ActiveSheet
    
    members_num = member_sheet.Range("A1").End(xlDown).Row - 1  'The address of member_sheet.Range("A1").End(xlDown) is $A$11
    'Range("A2:F11").Select'
    Range("A2:F" & members_num).Select  'As the number of members varies with month
    Selection.Copy
    
    'Windows("Reporting Template.xlsm").Activate'
    report_book.Activate 'Avoid hardcode
    'Sheets("Member_profile").Select'
    report_member_sheet.Select
    Range("A2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    'Get the Filename of Transactions.csv
    transaction_file = Application.GetOpenFilename(fileFilter:="csv files (*.csv),*.csv", Title:="Import the Transactions.csv")
    If transaction_file = False Then
        Exit Sub
    End If
    Workbooks.Open (transaction_file)
    Set transaction_book = ActiveWorkbook  'Avoid hardcode again
    Set transaction_sheet = ActiveSheet
    'Windows("Transactions.csv").Activate'
    
    trans_num = transaction_sheet.Range("A1").End(xlDown).Row - 1
    
    'Range("A2:F31").Select'
    Range("A2:F" & trans_num).Select  'As the number of transaction varies with month
    Selection.Copy
    
    'Windows("Reporting Template.xlsm").Activate'
    report_book.Activate
    report_trans_sheet.Select
    Range("A2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

    Range("G2:L2").Select
    Application.CutCopyMode = False
    'Selection.AutoFill destination:=Range("G2:L31"), Type:=xlFillDefault'
    Selection.AutoFill destination:=Range("G2:L" & trans_num), Type:=xlFillDefault
    'Range("G2:L31").Select'
    
    
    'Windows("Members.xlsx").Activate'
    'ActiveWindow.Close'
    member_book.Close (False)  'The params, False tells excel do not save the workbook
    
    'Windows("Transactions.csv").Activate'
    'ActiveWindow.Close'
    transaction_book.Close (False)
    
    'Sheets("Report").Select'
    report_report_sheet.Select
    
    report_member_sheet.Visible = False  'Hide the worksheets of Raw_data, Members_profile
    report_trans_sheet.Visible = False
    report_console_sheet.Visible = False
    
    'ActiveWorkbook.Save'
    report_book.Save
    
    MsgBox "Success! The report is generated!"
        
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True  'Turn on the function after running this macro
End Sub

