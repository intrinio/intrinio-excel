Attribute VB_Name = "IntrinioUtilities"
Option Explicit

Public Sub IntrinioAPIKeys()
    frmIntrinioAPIKeys.Show
End Sub

Public Sub IntrinioFixLinks()
    Dim Sht As Worksheet
    Application.ScreenUpdating = False
    
    For Each Sht In Worksheets
        Sht.Cells.Replace _
            What:="'*\Intrinio_Excel_Addin.xlam'!", _
            Replacement:="", _
            LookAt:=xlPart, _
            SearchOrder:=xlByRows, _
            MatchCase:=False
    Next Sht
    
    For Each Sht In Worksheets
        Sht.Cells.Replace _
            What:="'*Intrinio_Excel_Addin.xlam'!", _
            Replacement:="", _
            LookAt:=xlPart, _
            SearchOrder:=xlByRows, _
            MatchCase:=False
    Next Sht
    
    For Each Sht In Worksheets
        Sht.Cells.Replace _
            What:="'*\Intrinio_Excel_Addin-32bit.xlam'!", _
            Replacement:="", _
            LookAt:=xlPart, _
            SearchOrder:=xlByRows, _
            MatchCase:=False
    Next Sht
    
    For Each Sht In Worksheets
        Sht.Cells.Replace _
            What:="'*Intrinio_Excel_Addin-32bit.xlam'!", _
            Replacement:="", _
            LookAt:=xlPart, _
            SearchOrder:=xlByRows, _
            MatchCase:=False
    Next Sht
    Application.Run "ResetFindReplace"
    Application.ScreenUpdating = True
End Sub

Public Sub IntrinioUnlink()
    Dim ws As Worksheet
    Dim Ans As Variant
    Dim fileSaveName As Variant
    Dim wbName As String
    Dim Msg As String
    Dim fileName As String
    Dim i As Integer
    Dim r As Range
    
    Application.EnableCancelKey = xlDisabled
    Application.Calculation = xlCalculationManual
    
    wbName = ActiveWorkbook.Name
    wbName = Replace(wbName, ".xlsm", "")
    wbName = Replace(wbName, ".xlsx", "")

    Msg = "After unlinking " & wbName & " from the Intrinio Add-in, you will lose the ability to pull up-to-date information into this workbook. " _
            & "However, unlinking the workbook will allow you to share this workbook with people who may not have the Intrinio Excel Add-in installed. " + vbNewLine + vbNewLine _
            & "This change cannot be reversed - therefore, you will be prompted to save as a new unlinked workbook. " + vbNewLine + vbNewLine _
            & "Do you wish to continue and unlink " + wbName + " from the Intrinio Add-in?"

    Ans = MsgBox(Msg, vbYesNo, "Unlink Intrinio Excel Add-in?")
     
    Select Case Ans
              
    Case vbYes
        fileSaveName = Application.GetSaveAsFilename( _
            InitialFileName:=wbName & " - UNLINKED", _
            fileFilter:="Excel Workbook (*.xlsx), *.xlsx")

        If TypeName(fileSaveName) <> "Boolean" Then
            Application.DisplayAlerts = False
            ActiveWorkbook.Save
            
            For i = 1 To Sheets.Count
            On Error Resume Next
            For Each r In Sheets(i).UsedRange.SpecialCells(xlCellTypeFormulas)
            If r.Formula Like "*Intrinio*" Then r.Value = r.Value
            Next r
            Next i
            On Error GoTo 0

            ActiveWorkbook.SaveAs fileName:=fileSaveName, FileFormat:=xlOpenXMLWorkbook
            Application.DisplayAlerts = True
        End If
    
    End Select
    Application.EnableCancelKey = xlInterrupt
    Application.Calculation = xlCalculationAutomatic
End Sub

Public Sub ResetFindReplace()
   'Resets the find/replace dialog box options
   Dim r As Range

   On Error Resume Next

   Set r = Cells.Find(What:="", _
   LookIn:=xlFormulas, _
   SearchOrder:=xlRows, _
   LookAt:=xlPart, _
   MatchCase:=False)

   On Error GoTo 0

   'Reset the defaults

   On Error Resume Next

   Set r = Cells.Find(What:="", _
   LookIn:=xlFormulas, _
   SearchOrder:=xlRows, _
   LookAt:=xlPart, _
   MatchCase:=False)

   On Error GoTo 0
End Sub

