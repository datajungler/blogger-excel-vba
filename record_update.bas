Sub recorded_update()
    Windows("Members.xlsx").Activate
    Range("A2:F11").Select
    Selection.Copy
    
    Windows("Reporting Template.xlsm").Activate
    Sheets("Member_profile").Select
    Range("A2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Windows("Transactions.csv").Activate
    Range("A2:F31").Select
    Application.CutCopyMode = False
    Selection.Copy
    
    Windows("Reporting Template.xlsm").Activate
    Sheets("Raw_data").Select
    Range("A2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

    Range("G2:L2").Select
    Application.CutCopyMode = False
    Selection.AutoFill Destination:=Range("G2:L31"), Type:=xlFillDefault
    Range("G2:L31").Select
    
    
    Windows("Members.xlsx").Activate
    ActiveWindow.Close
    Windows("Transactions.csv").Activate
    ActiveWindow.Close
    
    Sheets("Report").Select
    ActiveWorkbook.Save
End Sub

