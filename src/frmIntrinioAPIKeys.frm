VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmIntrinioAPIKeys 
   Caption         =   "Intrinio API Key"
   ClientHeight    =   3450
   ClientLeft      =   45
   ClientTop       =   -1905
   ClientWidth     =   9030.001
   OleObjectBlob   =   "frmIntrinioAPIKeys.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmIntrinioAPIKeys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private oCol As Collection

Private Sub cmdUpdate_Click()
    Dim File_Num As Long
    Dim sOutFolder As String, sOutFile As String
    Dim IntrinioUsername As String
    
    IntrinioUsername = VBA.Trim(txtUserAPIKey.Value)
    
    If IntrinioUsername = "" Then
        IntrinioUsername = "<INTRINIO_USER_API_KEY>"
    End If
    
    On Error Resume Next
    sOutFolder = ThisWorkbook.path

    On Error GoTo 0
    File_Num = FreeFile
    With ActiveSheet
        'Specify the output filename without destroying the original value
        sOutFile = "Intrinio_API_Keys"
        'Specify the correct output folder and the output file name
        Open sOutFolder & Application.PathSeparator & VBA.Trim(sOutFile) & ".txt" For Output As #File_Num
        Print #1, IntrinioUsername
        Close #File_Num
    End With
    Call IntrinioInitialize
    Unload Me
End Sub

Private Sub lblAPIKeys_Click()
    Dim Url As String
    Url = "https://intrinio.com/tutorial/excel"
    ActiveWorkbook.FollowHyperlink Url
End Sub



Private Sub pasteUserAPIKey_Click()
    Dim ClipBoard As MSForms.DataObject, txt As String
    Set ClipBoard = New MSForms.DataObject
    ClipBoard.GetFromClipboard
    txt = ClipBoard.GetText
    
    If txt <> "" Then
        txtUserAPIKey.Value = txt
    End If
End Sub


Private Sub UserForm_Initialize()
    Dim File_Num As Long
    Dim sInFolder As String, sInFile As String
    Dim i As Integer
    Dim textline As String
    Dim lLength As Integer
    Dim bString As Integer
    Dim IntrinioUsername As String
    Dim IntrinioAPIKeysExists As Boolean

    On Error Resume Next
    sInFolder = ThisWorkbook.path
    
    sInFile = "Intrinio_API_Keys"
    IntrinioAPIKeysExists = FileOrDirExists(sInFolder & Application.PathSeparator & VBA.Trim(sInFile) & ".txt")
    
    If IntrinioAPIKeysExists = True Then
        File_Num = FreeFile
        With ActiveSheet
            'Specify the correct output folder and the output file name
            Open sInFolder & Application.PathSeparator & VBA.Trim(sInFile) & ".txt" For Input As #File_Num
            i = 1
            Do Until EOF(1)
                Line Input #1, textline
                IntrinioUsername = textline
            Loop
    
            Close #File_Num
            If IntrinioUsername <> "<INTRINIO_USER_API_KEY>" Then
                txtUserAPIKey.Value = IntrinioUsername
                cmdUpdate.Caption = "UPDATE"
            Else
                txtUserAPIKey.Value = ""
                cmdUpdate.Caption = "START"
            End If
        End With
    Else
        txtUserAPIKey.Value = ""
        cmdUpdate.Caption = "START"
    End If
    #If Mac Then
    #Else
        Dim oCCPClass As ClssCutCopyPaste
        
        Set oCol = New Collection
        
        Dim oCtl As Control
        For Each oCtl In Me.Controls
            If TypeOf oCtl Is MSForms.TextBox Then
                Set oCCPClass = New ClssCutCopyPaste
               Set oCCPClass.TxtBox = oCtl
                oCol.Add oCCPClass
            End If
        Next
    #End If
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    #If Mac Then
    #Else
        If CloseMode = vbFormControlMenu Then
            Dim File_Num As Long
            Dim sOutFolder As String, sOutFile As String
            Dim IntrinioUsername As String
            Dim sInFolder As String, sInFile As String, textline As String
            Dim i As Integer, lLength As Integer, bString As Integer
            
            sInFolder = ThisWorkbook.path
            
            sInFile = "Intrinio_API_Keys"
    
            File_Num = FreeFile
            Open sInFolder & Application.PathSeparator & VBA.Trim(sInFile) & ".txt" For Input As #File_Num
            i = 1
            Do Until EOF(1)
                Line Input #1, textline
                IntrinioUsername = textline
            Loop
            
            Close #File_Num
            
            If IntrinioUsername <> "" Then
                Unload Me
            Else
                IntrinioUsername = "<INTRINIO_USER_API_KEY>"
                
                On Error Resume Next
                sOutFolder = ThisWorkbook.path
            
                On Error GoTo 0
                File_Num = FreeFile
                With ActiveSheet
                    'Specify the output filename without destroying the original value
                    sOutFile = "Intrinio_API_Keys"
                    'Specify the correct output folder and the output file name
                    Open sOutFolder & Application.PathSeparator & VBA.Trim(sOutFile) & ".txt" For Output As #File_Num
                    Print #1, IntrinioUsername
                    Close #File_Num
                End With
                Call IntrinioInitialize
                Unload Me
            End If
        End If
    #End If
End Sub

Private Sub UserForm_Terminate()
    #If Mac Then
    
    #Else
        Set oCol = Nothing
    #End If
End Sub

Private Function FileOrDirExists(PathName As String) As Boolean
     
    Dim iTemp As Integer
     
    On Error Resume Next
    iTemp = GetAttr(PathName)

    Select Case Err.Number
    Case Is = 0
        FileOrDirExists = True
    Case Else
        FileOrDirExists = False
    End Select

    On Error GoTo 0
End Function
