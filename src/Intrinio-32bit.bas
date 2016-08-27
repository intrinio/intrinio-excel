Attribute VB_Name = "Intrinio"
Option Explicit

Public CompanyDic As New Dictionary
Public SecuritiesDic As New Dictionary
Public BankDic As New Dictionary
Public DataPointDic As New Dictionary
Public HistoricalPricesDic As New Dictionary
Public HistoricalDataDic As New Dictionary
Public NewsDic As New Dictionary
Public FundamentalsDic As New Dictionary
Public StandardizedFinancialsDic As New Dictionary
Public StandardizedTagsDic As New Dictionary
Public ReportedFundamentalsDic As New Dictionary
Public ReportedFinancialsDic As New Dictionary
Public ReportedTagsDic As New Dictionary
Public BankFundamentalsDic As New Dictionary
Public BankFinancialsDic As New Dictionary
Public BankTagsDic As New Dictionary

Private CompanySuccessDic As New Dictionary
Public DataPointRequestedTags As New Dictionary

Private iCredentials As New Dictionary
Private iVersion As New Dictionary
Private LoginFailure As Boolean
Private UpdatePrompt As Boolean

Private APICallsAtLimit As Boolean

Public Const BaseUrl = "https://api.intrinio.com"
Public Const Intrinio_Addin_Version = "2.6.0"

Public Sub IntrinioInitialize()

    Dim File_Num As Long
    Dim sInFolder As String, sInFile As String
    Dim i As Integer
    Dim textline As String
    Dim lLength As Integer
    Dim bString As Integer
    Dim IntrinioUsername As String
    Dim IntrinioPassword As String
    Dim initialize As Boolean
    Dim login_answer As Integer
    
    On Error GoTo ErrorHandler
    
    #If Win32 Or Win64 Then
        Call DescribeIntrinioDataPoint
        Call DescribeIntrinioHistoricalPrices
        Call DescribeIntrinioHistoricalData
        Call DescribeIntrinioNews
        Call DescribeIntrinioStandardizedFundamentals
        Call DescribeIntrinioStandardizedTags
        Call DescribeIntrinioStandardizedFinancials
        Call DescribeIntrinioFundamentals
        Call DescribeIntrinioTags
        Call DescribeIntrinioFinancials
        Call DescribeIntrinioReportedFundamentals
        Call DescribeIntrinioReportedTags
        Call DescribeIntrinioReportedFinancials
        Call DescribeIntrinioBankFundamentals
        Call DescribeIntrinioBankTags
        Call DescribeIntrinioBankFinancials
        Call IntrinioRibbon
    #End If
    
    If iCredentials.Exists("username") = True Or iCredentials.Exists("password") = True Then
        iCredentials.RemoveAll
    End If

    On Error Resume Next
    sInFolder = ThisWorkbook.path

    Dim IntrinioKeysExist As Boolean
    
    sInFile = "Intrinio_API_Keys"
    IntrinioKeysExist = FileOrDirExists(sInFolder & Application.PathSeparator & VBA.Trim(sInFile) & ".txt")
    
    If IntrinioKeysExist = True Then
        File_Num = FreeFile
        With ActiveSheet
            Open sInFolder & Application.PathSeparator & VBA.Trim(sInFile) & ".txt" For Input As #File_Num
            Do Until EOF(1)
                Line Input #1, textline
                lLength = Len(textline)
                bString = InStr(textline, ":")
                IntrinioUsername = VBA.Left(textline, bString - 1)
                IntrinioPassword = VBA.Right(textline, lLength - bString)
                If IntrinioUsername <> "<INTRINIO_USER_API_KEY>" Or IntrinioPassword <> "<INTRINIO_COLLABORATOR_KEY>" Then
                    iCredentials.Add "username", IntrinioUsername
                    iCredentials.Add "password", IntrinioPassword
                    initialize = False
                Else
                    If LoginFailure = False And APICallsAtLimit = False Then
                        initialize = True
                        LoginFailure = True
                        login_answer = MsgBox("Unable to authenticate with the Intrinio API", , "Unable to Login")
                    Else
                        iCredentials.Add "username", IntrinioUsername
                        iCredentials.Add "password", IntrinioPassword
                    End If
                End If
            Loop
    
            Close #File_Num
        End With
    Else
        initialize = True
    End If
    
    If initialize = True Then
        frmIntrinioAPIKeys.Show
    End If
    
    Dim version As String
    Dim web_url As String
    Dim IntrinioRespCode As String
    Dim answer As Integer
    
    IntrinioRespCode = IntrinioAddinVersion("status_code")
    
    If IntrinioRespCode = "" Then
        IntrinioRespCode = 401
    End If

    If IntrinioRespCode = 200 Then
        version = IntrinioAddinVersion("version")
        If version <> "429" Or version <> "403" Then
            If Intrinio_Addin_Version = version Then
                LoginFailure = False
                APICallsAtLimit = False
            Else
                If UpdatePrompt = False Then
                    answer = MsgBox("Version " & version & " of the Intrinio Excel Addin is available for download!" & vbNewLine & "Current version: " & Intrinio_Addin_Version _
                        & vbNewLine & "Would you like to install it now?", vbYesNo, "Update Intrinio Excel Add-in")
        
                    If answer = vbYes Then
                        #If Mac Then
                            web_url = IntrinioAddinVersion("mac_download_url")
                        #ElseIf Win32 Or Win64 Then
                            web_url = IntrinioAddinVersion("windows_download_url")
                        #Else
                            web_url = IntrinioAddinVersion("download_url")
                        #End If
                        ActiveWorkbook.FollowHyperlink Address:=web_url
                    End If
                    LoginFailure = False
                    APICallsAtLimit = False
                    UpdatePrompt = True
                Else
                    LoginFailure = False
                    APICallsAtLimit = False
                    UpdatePrompt = True
                    
                End If
            End If
        End If
    Else
        LoginFailure = True
        login_answer = MsgBox("Unable to authenticate with the Intrinio API", , "Unable to Login")
    End If

    Application.CalculateFull
ExitHere:
    Exit Sub
ErrorHandler:
    MsgBox "Unable to connect to the Intrinio API"
End Sub

Private Function IntrinioCompanies(ticker As String, Item As String)
    On Error GoTo ErrorHandler
    
    ticker = VBA.UCase(ticker)
    
    If ticker <> "" And LoginFailure = False And APICallsAtLimit = False Then
        If CompanyDic.Exists(ticker) = False Then
            Dim IntrinioClient As New WebClient
            IntrinioClient.BaseUrl = BaseUrl
            If iCredentials.Exists("username") = False Or iCredentials.Exists("password") = False Or iCredentials("username") = Empty Or iCredentials("password") = Empty Then
                Call IntrinioInitialize
            End If
            
            Dim inUsername As String
            Dim inPassword As String
            inUsername = iCredentials("username")
            inPassword = iCredentials("password")
            Dim Auth As New HttpBasicAuthenticator
            Auth.Setup _
                Username:=inUsername, _
                Password:=inPassword
            Set IntrinioClient.Authenticator = Auth

            Dim Request As New WebRequest
            Request.Resource = "companies/verify"
            Request.Method = HttpGet
            Request.Format = Json
            Request.AddQuerystringParam "ticker", ticker
            
            Dim Response As WebResponse
            Set Response = IntrinioClient.Execute(Request)

            If Response.StatusCode = Ok Then
                If Response.Data Is Nothing Then
                    IntrinioCompanies = ""
                Else
                    CompanyDic.Add ticker, Response.Data
                    If IsNull(CompanyDic(ticker)(Item)) Then
                        IntrinioCompanies = ""
                    Else
                        IntrinioCompanies = CompanyDic(ticker)(Item)
                    End If
                End If
            ElseIf Response.StatusCode = 429 Then
                APICallsAtLimit = True
                IntrinioCompanies = "429"
            ElseIf Response.StatusCode = 403 Then
                IntrinioCompanies = "403"
            Else
                IntrinioCompanies = ""
            End If
            
        ElseIf CompanyDic.Exists(ticker) = True Then
            If IsNull(CompanyDic(ticker)(Item)) Then
                IntrinioCompanies = ""
            Else
                IntrinioCompanies = CompanyDic(ticker)(Item)
            End If
        End If
    Else
        If APICallsAtLimit = True Then
            If CompanyDic.Exists(ticker) = True Then
                If IsNull(CompanyDic(ticker)(Item)) Then
                    IntrinioCompanies = "Plan Limit Reached"
                Else
                    IntrinioCompanies = CompanyDic(ticker)(Item)
                End If
            Else
                IntrinioCompanies = CompanyDic(ticker)(Item)
            End If
            
        ElseIf LoginFailure = True Then
            IntrinioCompanies = "Invalid API Keys"
        Else
            IntrinioCompanies = ""
        End If
    End If
ExitHere:
    Exit Function
ErrorHandler:
    IntrinioCompanies = ""
End Function

Private Function IntrinioSecurities(ticker As String, Item As String)
    On Error GoTo ErrorHandler
    
    ticker = VBA.UCase(ticker)
    
    If ticker <> "" And LoginFailure = False And APICallsAtLimit = False Then
        If SecuritiesDic.Exists(ticker) = False Then
            Dim IntrinioClient As New WebClient
            IntrinioClient.BaseUrl = BaseUrl
            If iCredentials.Exists("username") = False Or iCredentials.Exists("password") = False Or iCredentials("username") = Empty Or iCredentials("password") = Empty Then
                Call IntrinioInitialize
            End If
            
            Dim inUsername As String
            Dim inPassword As String
            inUsername = iCredentials("username")
            inPassword = iCredentials("password")
            Dim Auth As New HttpBasicAuthenticator
            Auth.Setup _
                Username:=inUsername, _
                Password:=inPassword
            Set IntrinioClient.Authenticator = Auth

            Dim Request As New WebRequest
            Request.Resource = "securities/verify"
            Request.Method = HttpGet
            Request.Format = Json
            Request.AddQuerystringParam "ticker", ticker
            
            Dim Response As WebResponse
            Set Response = IntrinioClient.Execute(Request)

            If Response.StatusCode = Ok Then
                If Response.Data Is Nothing Then
                    IntrinioSecurities = ""
                Else
                    SecuritiesDic.Add ticker, Response.Data
                    If IsNull(SecuritiesDic(ticker)(Item)) Then
                        IntrinioSecurities = ""
                    Else
                        IntrinioSecurities = SecuritiesDic(ticker)(Item)
                    End If
                End If
            ElseIf Response.StatusCode = 429 Then
                APICallsAtLimit = True
                IntrinioSecurities = "429"
            ElseIf Response.StatusCode = 403 Then
                IntrinioSecurities = "403"
            Else
                IntrinioSecurities = ""
            End If
            
        ElseIf SecuritiesDic.Exists(ticker) = True Then
            If IsNull(SecuritiesDic(ticker)(Item)) Then
                IntrinioSecurities = ""
            Else
                IntrinioSecurities = SecuritiesDic(ticker)(Item)
            End If
        End If
    Else
        If APICallsAtLimit = True Then
            If SecuritiesDic.Exists(ticker) = True Then
                If IsNull(SecuritiesDic(ticker)(Item)) Then
                    IntrinioSecurities = "Plan Limit Reached"
                Else
                    IntrinioSecurities = SecuritiesDic(ticker)(Item)
                End If
            Else
                IntrinioSecurities = SecuritiesDic(ticker)(Item)
            End If
        ElseIf LoginFailure = True Then
            IntrinioSecurities = "Invalid API Keys"
        Else
            IntrinioSecurities = ""
        End If
    End If
ExitHere:
    Exit Function
ErrorHandler:
    IntrinioSecurities = ""
End Function


Private Function IntrinioBanks(identifier As String, Item As String)
    On Error GoTo ErrorHandler
        
    If identifier <> "" And LoginFailure = False And APICallsAtLimit = False Then
        If BankDic.Exists(identifier) = False Then
            Dim IntrinioClient As New WebClient
            IntrinioClient.BaseUrl = BaseUrl
            If iCredentials.Exists("username") = False Or iCredentials.Exists("password") = False Or iCredentials("username") = Empty Or iCredentials("password") = Empty Then
                Call IntrinioInitialize
            End If
            
            Dim inUsername As String
            Dim inPassword As String
            inUsername = iCredentials("username")
            inPassword = iCredentials("password")
            Dim Auth As New HttpBasicAuthenticator
            Auth.Setup _
                Username:=inUsername, _
                Password:=inPassword
            Set IntrinioClient.Authenticator = Auth

            Dim Request As New WebRequest
            Request.Resource = "banks/verify"
            Request.Method = HttpGet
            Request.Format = Json
            Request.AddQuerystringParam "identifier", identifier
            
            Dim Response As WebResponse
            Set Response = IntrinioClient.Execute(Request)
            Debug.Print Response.Content
            If Response.StatusCode = Ok Then
                If Response.Data Is Nothing Then
                    IntrinioBanks = ""
                Else
                    BankDic.Add Response.Data("identifier"), Response.Data
                    If IsNull(BankDic(identifier)(Item)) Then
                        IntrinioBanks = ""
                    Else
                        IntrinioBanks = BankDic(identifier)(Item)
                    End If
                End If
            ElseIf Response.StatusCode = 429 Then
                APICallsAtLimit = True
                IntrinioBanks = "429"
            ElseIf Response.StatusCode = 403 Then
                IntrinioBanks = "403"
            Else
                IntrinioBanks = ""
            End If
            
        ElseIf BankDic.Exists(identifier) = True Then
            If IsNull(BankDic(identifier)(Item)) Then
                IntrinioBanks = ""
            Else
                IntrinioBanks = BankDic(identifier)(Item)
            End If
        End If
    Else
        If APICallsAtLimit = True Then
            If BankDic.Exists(identifier) = True Then
                If IsNull(BankDic(identifier)(Item)) Then
                    IntrinioBanks = "Plan Limit Reached"
                Else
                    IntrinioBanks = BankDic(identifier)(Item)
                End If
            Else
                IntrinioBanks = BankDic(identifier)(Item)
            End If
        ElseIf LoginFailure = True Then
            IntrinioBanks = "Invalid API Keys"
        Else
            IntrinioBanks = ""
        End If
    End If
ExitHere:
    Exit Function
ErrorHandler:
    IntrinioBanks = ""
End Function

Sub DescribeIntrinioDataPoint()
    Dim FuncName As String
    Dim FuncDesc As String
    Dim Category As String
    Dim ArgDesc(1 To 2) As String
    
    FuncName = "IntrinioDataPoint"
    FuncDesc = "Returns a data point for a company based on a selected tag"
    Category = "Intrinio"
    ArgDesc(1) = "The unique identifier for the data set (i.e. 'AAPL, FRED.GDP, DMD.ERP')"
    ArgDesc(2) = "The data point tag requested for the company (i.e. 'current_yr_ave_eps_est' returns the current fiscal year average Wall Street consensus EPS estimate)"
    
    Application.MacroOptions Macro:=FuncName, _
       Description:=FuncDesc, _
       Category:=Category, _
       ArgumentDescriptions:=ArgDesc
    
    FuncName = "IDP"
    Application.MacroOptions Macro:=FuncName, _
       Description:=FuncDesc, _
       Category:=Category, _
       ArgumentDescriptions:=ArgDesc
End Sub

Public Function IDP(identifier As String, Item As String)
    IDP = IntrinioDataPoint(identifier, Item)
End Function

Public Function IntrinioDataPoint(identifier As String, Item As String)
Attribute IntrinioDataPoint.VB_Description = "Returns a data point for a company based on a selected tag"
Attribute IntrinioDataPoint.VB_ProcData.VB_Invoke_Func = " \n19"
    Dim Key As String
    Dim api_ticker As String
    Dim coFailure As Boolean
    Dim dValue As Double
    Dim exchange_pos As Integer
    Dim index_pos As Integer
    
    On Error GoTo ErrorHandler
    
    exchange_pos = InStr(identifier, ":")
    index_pos = InStr(identifier, "$")
    identifier = VBA.UCase(identifier)
    
    If identifier <> "" And LoginFailure = False Then
        If CompanySuccessDic.Exists(identifier) = False Then
            If identifier = "" Then
                coFailure = False
            ElseIf VBA.Left(identifier, 5) = "FRED." Then
                CompanySuccessDic.Add identifier, False
                coFailure = CompanySuccessDic(identifier)
            ElseIf identifier = "DMD.ERP" Then
                CompanySuccessDic.Add identifier, False
                coFailure = CompanySuccessDic(identifier)
            ElseIf exchange_pos > 0 Then
                CompanySuccessDic.Add identifier, False
                coFailure = CompanySuccessDic(identifier)
            ElseIf index_pos = 1 Then
                CompanySuccessDic.Add identifier, False
                coFailure = CompanySuccessDic(identifier)
            Else
                api_ticker = IntrinioCompanies(identifier, "ticker")
                If api_ticker = identifier Then
                    CompanySuccessDic.Add identifier, False
                    coFailure = CompanySuccessDic(identifier)
                Else
                    api_ticker = IntrinioSecurities(identifier, "ticker")
                    If api_ticker = identifier Then
                        CompanySuccessDic.Add identifier, False
                        coFailure = CompanySuccessDic(identifier)
                    Else
                        api_ticker = IntrinioBanks(identifier, "identifier")
                        If api_ticker = identifier Then
                            CompanySuccessDic.Add identifier, False
                            coFailure = CompanySuccessDic(identifier)
                        Else
                            CompanySuccessDic.Add identifier, True
                            coFailure = CompanySuccessDic(identifier)
                        End If
                        
                    End If
                End If
                
            End If
        Else
            If APICallsAtLimit = False Then
                coFailure = CompanySuccessDic(identifier)
            Else
                coFailure = False
            End If
        End If
    End If

    If identifier <> "" And LoginFailure = False And APICallsAtLimit = False And coFailure = False Then
        Key = identifier & "_" & Item
        If DataPointDic.Exists(Key) = False Then
            Dim tags As String
            Dim requestItemCount As Integer
            #If Win32 Or Win64 Then

                Dim DPRT_idenftifier As String
                DPRT_idenftifier = identifier & "_DPRT"
                If DataPointRequestedTags.Exists(DPRT_idenftifier) = True Then
                    tags = DataPointRequestedTags(DPRT_idenftifier)
                    DataPointRequestedTags.Remove DPRT_idenftifier

                Else
                    Call FindAllDataPointTags

                    tags = DataPointRequestedTags(DPRT_idenftifier)
                    DataPointRequestedTags.Remove DPRT_idenftifier
                End If

                If tags = "" Then
                    tags = Item
                    requestItemCount = 0
                Else
                    Dim replaced As String
                    Dim requestedTagsContain As String
    
                    If VBA.Left(identifier, 5) = "FRED." Then
                        tags = Item
                        requestItemCount = 0
                    ElseIf identifier = "DMD.ERP" Then
                        tags = Item
                        requestItemCount = 0
                    Else
                        requestedTagsContain = VBA.InStr(1, tags, Item)
                        If requestedTagsContain = 0 Then
                            tags = tags & "," & Item
                        End If
                        
                        replaced = Replace(tags, ",", "")
                        requestItemCount = VBA.Len(tags) - VBA.Len(replaced)
                    End If
                End If
            #Else
                tags = Item
                requestItemCount = 0
            #End If
            
            Dim IntrinioClient As New WebClient
            IntrinioClient.BaseUrl = BaseUrl
            If iCredentials.Exists("username") = False Or iCredentials.Exists("password") = False Or iCredentials("username") = Empty Or iCredentials("password") = Empty Then
                Call IntrinioInitialize
            End If
            
            Dim inUsername As String
            Dim inPassword As String
            inUsername = iCredentials("username")
            inPassword = iCredentials("password")
            Dim Auth As New HttpBasicAuthenticator
            Auth.Setup _
                Username:=inUsername, _
                Password:=inPassword
            Set IntrinioClient.Authenticator = Auth
            
            Dim Request As New WebRequest
            Request.Resource = "data_point"
            Request.Method = HttpGet
            Request.Format = Json
            Request.AddQuerystringParam "identifier", identifier
            Request.AddQuerystringParam "item", tags
            
            Dim Response As WebResponse
            Set Response = IntrinioClient.Execute(Request)

            If Response.StatusCode = Ok Then
                If requestItemCount = 0 Then
                    DataPointDic.Add Key, Response.Data("value")
                ElseIf requestItemCount > 0 Then
                    Dim X As Variant
                    Dim long_key As String
                    Dim sValue As String
                    Dim nValue As Double
                    Dim rTag As String
                    Dim nKey As String

                    long_key = identifier & "_long_key"
                    
                    If DataPointDic.Exists(long_key) = True Then
                        DataPointDic.Remove (long_key)
                    End If
                    DataPointDic.Add long_key, Response.Data("data")
                    
                    For Each X In DataPointDic(long_key)
                        
                        rTag = X("item")
                        nKey = identifier & "_" & rTag
                        If DataPointDic.Exists(nKey) = True Then
                            DataPointDic.Remove (nKey)
                        End If
                        If IsNumeric(X("value")) = True Then
                            nValue = X("value")
                            DataPointDic.Add nKey, nValue
                        Else
                            sValue = X("value")
                            DataPointDic.Add nKey, sValue
                        End If
                        
                    Next
                    DataPointDic.Remove (long_key)
                End If

                If IsNull(DataPointDic(Key)) Then
                    IntrinioDataPoint = ""
                Else
                    If IsNumeric(DataPointDic(Key)) = True Then
                        dValue = DataPointDic(Key)
                        IntrinioDataPoint = dValue
                    Else
                        IntrinioDataPoint = DataPointDic(Key)
                    End If
                End If
            ElseIf Response.StatusCode = 429 Then
                APICallsAtLimit = True
                IntrinioDataPoint = "Plan Limit Reached"
            ElseIf Response.StatusCode = 403 Then
                IntrinioDataPoint = "Visit Intrinio.com to Subscribe"
            Else
                IntrinioDataPoint = ""
            End If
        ElseIf DataPointDic.Exists(Key) = True Then
            If IsNumeric(DataPointDic(Key)) = True Then
                dValue = DataPointDic(Key)
                IntrinioDataPoint = dValue
            ElseIf IsNull(DataPointDic(Key)) Then
                IntrinioDataPoint = ""
            Else
                IntrinioDataPoint = DataPointDic(Key)
            End If
        End If
    Else
        If APICallsAtLimit = True Then
            Key = identifier & "_" & Item
            If DataPointDic.Exists(Key) = True Then
                If IsNumeric(DataPointDic(Key)) = True Then
                    dValue = DataPointDic(Key)
                    IntrinioDataPoint = dValue
                ElseIf IsNull(DataPointDic(Key)) Then
                    IntrinioDataPoint = ""
                Else
                    IntrinioDataPoint = DataPointDic(Key)
                End If
            Else
                IntrinioDataPoint = "Plan Limit Reached"
            End If
        ElseIf LoginFailure = True Then
            IntrinioDataPoint = "Invalid API Keys"
        ElseIf coFailure = True Then
            IntrinioDataPoint = "Invalid Identifier"
        Else
            IntrinioDataPoint = ""
        End If
    End If
ExitHere:
    Exit Function
ErrorHandler:
    IntrinioDataPoint = ""
End Function

Private Sub FindAllDataPointTags()
    Dim fnd As String, FirstFound As String
    Dim FoundCell As Range, rng As Range
    Dim myRange As Range, LastCell As Range
    Dim book As Workbook
    
    On Error GoTo ErrorHandler
    
    Dim Sht As Worksheet
    Dim curSheet As String
    
    curSheet = ActiveSheet.Name
    
    For Each book In Workbooks
        For Each Sht In book.Worksheets
            fnd = "IntrinioDataPoint"
        
            Set myRange = Sht.UsedRange
            Set LastCell = myRange.Cells(myRange.Cells.Count)
            Set FoundCell = myRange.Find(What:=fnd, After:=LastCell)
            If Not FoundCell Is Nothing Then
                FirstFound = FoundCell.Address
                Set rng = FoundCell
                
                On Error Resume Next
                Do Until FoundCell Is Nothing
                    Set FoundCell = myRange.Find(What:=fnd, After:=FoundCell)
                    
                    Dim getFormula As String
                    Dim identifier As String
                    Dim ticker As String, dpFunctionName As String
                    
                    Dim a As Integer, b As Integer, c As Integer, d As Integer
                    
                    dpFunctionName = "IntrinioDataPoint"
                    getFormula = FoundCell.Formula
                    
                    a = VBA.InStr(1, getFormula, dpFunctionName, vbTextCompare)
                    
                    If a > 0 Then
                        b = VBA.InStr(a, getFormula, "(", vbTextCompare)
                        c = VBA.InStr(a, getFormula, ",", vbTextCompare)
                        d = VBA.InStr(a, getFormula, ")", vbTextCompare)
                        
                        ticker = VBA.Mid(getFormula, b + 1, c - b - 1)
                        
                        Dim valid_address As String

                        valid_address = Workbooks(book.Name).Sheets(Sht.Name).Range(ticker).Address(External:=True)
                        
                        If ValidAddress(valid_address) Then
                            Dim sheet_ticker As String
                            sheet_ticker = Range(valid_address)
                            If sheet_ticker = "" Then
                                ticker = Sheets(Sht.Name).Range(ticker)
                            Else
                                ticker = sheet_ticker
                            End If
                        Else
                            ticker = VBA.Mid(getFormula, b + 2, c - b - 3)
                        End If
                                            
                        If ticker <> "" Then
                            Dim new_tag As String
    
                            Dim DQ As String
                            DQ = Chr(34)
                            
                            new_tag = VBA.Mid(getFormula, c + 1, d - c - 1)
                            
                            Dim valid_address2 As String
                            valid_address2 = Workbooks(book.Name).Sheets(Sht.Name).Range(new_tag).Address(External:=True)
                            
                            If ValidAddress(valid_address2) Then
                                new_tag = Range(valid_address2)
                            Else
                                new_tag = VBA.Mid(getFormula, c + 2, d - c - 3)
                                new_tag = Replace(new_tag, DQ, "")
                            End If
                            
                            Dim DPRT_idenftifier As String
                            DPRT_idenftifier = ticker & "_DPRT"
                    
                            Dim tags As String
                            
                            tags = ""
                            
                            If DataPointRequestedTags.Exists(DPRT_idenftifier) = True Then
                                tags = DataPointRequestedTags(DPRT_idenftifier)
                            End If
                            
                            Dim Key As String
                            Key = ticker & "_" & new_tag
                            
                            If DataPointDic.Exists(Key) = False Then
                                If tags = "" Then
                                    If DataPointRequestedTags.Exists(DPRT_idenftifier) = True Then
                                        DataPointRequestedTags.Remove DPRT_idenftifier
                                    End If
                                    DataPointRequestedTags.Add DPRT_idenftifier, new_tag
                                Else
                                    Dim pos As Integer

                                    pos = InStr(tags, new_tag)
                                    
                                    If pos = 0 Then
                                        Dim replaced As String
                                        Dim requestItemCount As Integer
                                        
                                        replaced = Replace(tags, ",", "")
                                        requestItemCount = VBA.Len(tags) - VBA.Len(replaced)
                                        
                                        If requestItemCount < 50 Then
                                            If DataPointRequestedTags.Exists(DPRT_idenftifier) = True Then
                                                DataPointRequestedTags.Remove DPRT_idenftifier
                                            End If
                                            DataPointRequestedTags.Add DPRT_idenftifier, tags & "," & new_tag
                                        End If
                                    End If
                                End If
                            End If
                        End If
            
                        If FoundCell.Address = FirstFound Then Exit Do
                    End If
                    
                Loop
            
            End If
        Next Sht
    Next book
    Sheets(curSheet).Select
    Application.Run "ResetFindReplace"
Exit Sub

'Error Handler
ErrorHandler:

End Sub

Sub DescribeIntrinioHistoricalPrices()
    Dim FuncName As String
    Dim FuncDesc As String
    Dim Category As String
    Dim ArgDesc(1 To 6) As String
    
    FuncName = "IntrinioHistoricalPrices"
    FuncDesc = "Returns a historical price data point for a company based on the sequence number"
    Category = "Intrinio"
    ArgDesc(1) = "The company's ticker symbol (i.e. 'AMZN')"
    ArgDesc(2) = "The data tag for the historical stock price"
    ArgDesc(3) = "The sequence order of the fundamental from newest to oldest (0..last available)"
    ArgDesc(4) = "(Optional) The earliest date for the historical stock prices series"
    ArgDesc(5) = "(Optional) The latest date for the historical stock prices series"
    ArgDesc(6) = "(Optional) The frequency of price data (daily,weekly,monthly,quarterly,yearly)"
    
    Application.MacroOptions Macro:=FuncName, _
       Description:=FuncDesc, _
       Category:=Category, _
       ArgumentDescriptions:=ArgDesc
End Sub

Public Function IntrinioHistoricalPrices(ticker As String, Item As String, sequence As Integer, Optional start_date As String = "", Optional end_date As String = "", Optional frequency As String = "daily")
Attribute IntrinioHistoricalPrices.VB_Description = "Returns a historical price data point for a company based on the sequence number"
Attribute IntrinioHistoricalPrices.VB_ProcData.VB_Invoke_Func = " \n19"
    Dim Key As String
    Dim api_ticker As String
    Dim coFailure As Boolean
    Dim index_pos As Integer
    
    On Error GoTo ErrorHandler
    
    index_pos = InStr(ticker, "$")
    ticker = VBA.UCase(ticker)
    
    If ticker <> "" And LoginFailure = False Then
        If CompanySuccessDic.Exists(ticker) = False Then
            api_ticker = IntrinioCompanies(ticker, "ticker")
            If api_ticker = ticker Then
                CompanySuccessDic.Add ticker, False
                coFailure = CompanySuccessDic(ticker)
            ElseIf index_pos = 1 Then
                CompanySuccessDic.Add ticker, False
                coFailure = CompanySuccessDic(ticker)
            Else
                api_ticker = IntrinioSecurities(ticker, "ticker")
                If api_ticker = ticker Then
                    CompanySuccessDic.Add ticker, False
                    coFailure = CompanySuccessDic(ticker)
                Else
                    CompanySuccessDic.Add ticker, True
                    coFailure = CompanySuccessDic(ticker)
                End If
            End If
        Else
            If APICallsAtLimit = False Then
                coFailure = CompanySuccessDic(ticker)
            Else
                coFailure = False
            End If
        End If
    End If
    
    If ticker <> "" And Item <> "" And LoginFailure = False And APICallsAtLimit = False And coFailure = False Then
        Key = ticker & "_" & start_date & "_" & end_date & "_" & frequency
        If HistoricalPricesDic.Exists(Key) = False Then
            Dim IntrinioClient As New WebClient
            IntrinioClient.BaseUrl = BaseUrl
            If iCredentials.Exists("username") = False Or iCredentials.Exists("password") = False Or iCredentials("username") = Empty Or iCredentials("password") = Empty Then
                Call IntrinioInitialize
            End If

            Dim inUsername As String
            Dim inPassword As String
            inUsername = iCredentials("username")
            inPassword = iCredentials("password")
            Dim Auth As New HttpBasicAuthenticator
            Auth.Setup _
                Username:=inUsername, _
                Password:=inPassword
            Set IntrinioClient.Authenticator = Auth
             
            Dim Request As New WebRequest
            Request.Resource = "prices"
            Request.Method = HttpGet
            Request.Format = Json
            Request.AddQuerystringParam "ticker", ticker
            If start_date <> "" Then
                Request.AddQuerystringParam "start_date", start_date
            End If
            If end_date <> "" Then
                Request.AddQuerystringParam "end_date", end_date
            End If
            If frequency <> "" Then
                Request.AddQuerystringParam "frequency", frequency
            End If
            
            Dim Response As WebResponse
            Set Response = IntrinioClient.Execute(Request)
            
            If Response.StatusCode = Ok Then
                HistoricalPricesDic.Add Key, Response.Data("data")
                IntrinioHistoricalPrices = HistoricalPricesDic(Key)(sequence + 1)(Item)
                If IntrinioHistoricalPrices = Empty Then
                    IntrinioHistoricalPrices = ""
                End If
                If Item = "date" Then
                    IntrinioHistoricalPrices = VBA.DateValue(IntrinioHistoricalPrices)
                Else
                    IntrinioHistoricalPrices = VBA.Round(IntrinioHistoricalPrices * 1, 4)
                End If
            ElseIf Response.StatusCode = 429 Then
                APICallsAtLimit = True
                IntrinioHistoricalPrices = "Plan Limit Reached"
            ElseIf Response.StatusCode = 403 Then
                IntrinioHistoricalPrices = "Visit Intrinio.com to Subscribe"
            Else
                IntrinioHistoricalPrices = ""
            End If
        ElseIf HistoricalPricesDic.Exists(Key) = True Then
            IntrinioHistoricalPrices = HistoricalPricesDic(Key)(sequence + 1)(Item)
            If IntrinioHistoricalPrices = Empty Then
                IntrinioHistoricalPrices = ""
            End If
            If Item = "date" Then
                IntrinioHistoricalPrices = VBA.DateValue(IntrinioHistoricalPrices)
            Else
                IntrinioHistoricalPrices = VBA.Round(IntrinioHistoricalPrices * 1, 4)
            End If
            
        End If

    Else
        If ticker = "" Then
            IntrinioHistoricalPrices = ""
        ElseIf Item = "" Then
            IntrinioHistoricalPrices = ""
        ElseIf APICallsAtLimit = True Then
            Key = ticker & "_" & start_date & "_" & end_date & "_" & frequency
            If HistoricalPricesDic.Exists(Key) = True Then
                IntrinioHistoricalPrices = HistoricalPricesDic(Key)(sequence + 1)(Item)
                If IntrinioHistoricalPrices = Empty Then
                    IntrinioHistoricalPrices = ""
                End If
            Else
                IntrinioHistoricalPrices = "Plan Limit Reached"
            End If
        ElseIf LoginFailure = True Then
            IntrinioHistoricalPrices = "Invalid API Keys"
        ElseIf coFailure = True Then
            IntrinioHistoricalPrices = "Invalid Ticker Symbol"
        Else
            IntrinioHistoricalPrices = ""
        End If
    End If
ExitHere:
    Exit Function
ErrorHandler:
    IntrinioHistoricalPrices = ""
End Function

Sub DescribeIntrinioHistoricalData()
    Dim FuncName As String
    Dim FuncDesc As String
    Dim Category As String
    Dim ArgDesc(1 To 8) As String
    
    FuncName = "IntrinioHistoricalData"
    FuncDesc = "Returns a historical data point for a company based on the sequence number"
    Category = "Intrinio"
    ArgDesc(1) = "The company's ticker symbol (i.e. 'AMZN')"
    ArgDesc(2) = "The data tag for the historical data"
    ArgDesc(3) = "The sequence order of the data point from newest to oldest (0..last available)"
    ArgDesc(4) = "(Optional) The earliest date for the historical data series"
    ArgDesc(5) = "(Optional) The latest date for the historical data series"
    ArgDesc(6) = "(Optional) The frequency of data series (daily,weekly,monthly,quarterly,yearly)"
    ArgDesc(7) = "(Optional) The data type for the series - fiscal period type for standardized financials & stat for SIC indices (TTM,QTR,FY or sum, mean, media, etc.)"
    ArgDesc(8) = "(Optional) By default, show_date is false, meaning that it will show the return value - if true, it will return the date associated with the data point."
    
    Application.MacroOptions Macro:=FuncName, _
       Description:=FuncDesc, _
       Category:=Category, _
       ArgumentDescriptions:=ArgDesc
       
    FuncName = "IHD"
    Application.MacroOptions Macro:=FuncName, _
       Description:=FuncDesc, _
       Category:=Category, _
       ArgumentDescriptions:=ArgDesc
End Sub

Public Function IntrinioHistoricalData(ticker As String, Item As String, sequence As Integer, Optional start_date As String = "", Optional end_date As String = "", Optional frequency As String = "", Optional data_type As String = "", Optional show_date As Boolean = False)
    
    Dim str_start_date As String
    Dim str_end_date As String
    Dim str_frequency As String
    Dim str_data_type As String
    Dim bol_show_date As Boolean

    str_start_date = start_date
    str_end_date = end_date
    str_frequency = frequency
    str_data_type = data_type
    bol_show_date = show_date

    IntrinioHistoricalData = IHD(ticker, Item, sequence, str_start_date, str_end_date, str_frequency, str_data_type, bol_show_date)
End Function

Public Function IHD(ticker As String, Item As String, sequence As Integer, Optional start_date As String = "", Optional end_date As String = "", Optional frequency As String = "", Optional data_type As String = "", Optional show_date As Boolean = False)
    Dim Key As String
    Dim api_ticker As String
    Dim coFailure As Boolean
    Dim index_pos As Integer
    
    On Error GoTo ErrorHandler

    index_pos = InStr(ticker, "$")
    ticker = VBA.UCase(ticker)
    
    If ticker <> "" And LoginFailure = False And APICallsAtLimit = False Then
        If CompanySuccessDic.Exists(ticker) = False Then
            api_ticker = IntrinioCompanies(ticker, "ticker")
            If api_ticker = ticker Then
                CompanySuccessDic.Add ticker, False
                coFailure = CompanySuccessDic(ticker)
            ElseIf index_pos = 1 Then
                CompanySuccessDic.Add ticker, False
                coFailure = CompanySuccessDic(ticker)
            Else
                api_ticker = IntrinioSecurities(ticker, "ticker")
                If api_ticker = ticker Then
                    CompanySuccessDic.Add ticker, False
                    coFailure = CompanySuccessDic(ticker)
                Else
                    CompanySuccessDic.Add ticker, True
                    coFailure = CompanySuccessDic(ticker)
                End If
            End If
        Else
            If APICallsAtLimit = False Then
                coFailure = CompanySuccessDic(ticker)
            Else
                coFailure = False
            End If
        End If
    End If
    
    If ticker <> "" And Item <> "" And LoginFailure = False And APICallsAtLimit = False And coFailure = False Then
        Key = ticker & "_" & Item & "_" & start_date & "_" & end_date & "_" & frequency & "_" & data_type
        If HistoricalDataDic.Exists(Key) = False Then
            Dim IntrinioClient As New WebClient
            IntrinioClient.BaseUrl = BaseUrl
            If iCredentials.Exists("username") = False Or iCredentials.Exists("password") = False Or iCredentials("username") = Empty Or iCredentials("password") = Empty Then
                Call IntrinioInitialize
            End If

            Dim inUsername As String
            Dim inPassword As String
            inUsername = iCredentials("username")
            inPassword = iCredentials("password")
            Dim Auth As New HttpBasicAuthenticator
            Auth.Setup _
                Username:=inUsername, _
                Password:=inPassword
            Set IntrinioClient.Authenticator = Auth
             
            Dim Request As New WebRequest
            Request.Resource = "historical_data"
            Request.Method = HttpGet
            Request.Format = Json
            Request.AddQuerystringParam "ticker", ticker
            Request.AddQuerystringParam "item", Item
            If start_date <> "" Then
                Request.AddQuerystringParam "start_date", start_date
            End If
            If end_date <> "" Then
                Request.AddQuerystringParam "end_date", end_date
            End If
            If frequency <> "" Then
                Request.AddQuerystringParam "frequency", frequency
            End If
            If data_type <> "" Then
                Request.AddQuerystringParam "type", data_type
            End If
            
            Dim Response As WebResponse
            Set Response = IntrinioClient.Execute(Request)
            
            If Response.StatusCode = Ok Then
                HistoricalDataDic.Add Key, Response.Data("data")
                If show_date = True Then
                    IHD = HistoricalDataDic(Key)(sequence + 1)("date")
                    IHD = VBA.DateValue(IHD)
                ElseIf show_date = False Then
                    IHD = HistoricalDataDic(Key)(sequence + 1)("value")
                    IHD = VBA.Round(IHD * 1, 4)
                End If
                If IHD = Empty Then
                    IHD = ""
                End If
            ElseIf Response.StatusCode = 429 Then
                APICallsAtLimit = True
                IHD = "Plan Limit Reached"
            ElseIf Response.StatusCode = 403 Then
                IHD = "Visit Intrinio.com to Subscribe"
            Else
                IHD = ""
            End If
        ElseIf HistoricalDataDic.Exists(Key) = True Then
            If show_date = True Then
                IHD = HistoricalDataDic(Key)(sequence + 1)("date")
                IHD = VBA.DateValue(IHD)
            ElseIf show_date = False Then
                IHD = HistoricalDataDic(Key)(sequence + 1)("value")
                IHD = VBA.Round(IHD * 1, 4)
            End If
            If IHD = Empty Then
                IHD = ""
            End If
        End If

    Else
        If ticker = "" Then
            IHD = ""
        ElseIf Item = "" Then
            IHD = ""
        ElseIf APICallsAtLimit = True Then
            Key = ticker & "_" & Item & "_" & start_date & "_" & end_date & "_" & frequency & "_" & data_type
            If HistoricalDataDic.Exists(Key) = True Then
                If show_date = True Then
                    IHD = HistoricalDataDic(Key)(sequence + 1)("date")
                    IHD = VBA.DateValue(IHD)
                ElseIf show_date = False Then
                    IHD = HistoricalDataDic(Key)(sequence + 1)("value")
                    IHD = VBA.Round(IHD * 1, 4)
                End If
                If IHD = Empty Then
                    IHD = ""
                End If
            Else
                IHD = "Plan Limit Reached"
            End If
        ElseIf LoginFailure = True Then
            IHD = "Invalid API Keys"
        ElseIf coFailure = True Then
            IHD = "Invalid Ticker Symbol"
        Else
            IHD = ""
        End If
    End If
ExitHere:
    Exit Function
ErrorHandler:
    IHD = ""
End Function

Sub DescribeIntrinioNews()
    Dim FuncName As String
    Dim FuncDesc As String
    Dim Category As String
    Dim ArgDesc(1 To 6) As String
    
    FuncName = "IntrinioNews"
    FuncDesc = "Returns a historical price data point for a company based on the sequence number"
    Category = "Intrinio"
    ArgDesc(1) = "The company's ticker symbol or combination of ticker symbols (i.e. 'AMZN' or 'AMZN,AAPL,GOOGL')"
    ArgDesc(2) = "The return value selected (i.e. 'title', 'publication_date', 'url', 'summary'"
    ArgDesc(3) = "The sequence order of the article from the feed, newest to oldest (0..last available)"
    
    Application.MacroOptions Macro:=FuncName, _
       Description:=FuncDesc, _
       Category:=Category, _
       ArgumentDescriptions:=ArgDesc
End Sub
Public Function IntrinioNews(ticker As String, Item As String, sequence As Integer)
Attribute IntrinioNews.VB_Description = "Returns a historical price data point for a company based on the sequence number"
Attribute IntrinioNews.VB_ProcData.VB_Invoke_Func = " \n19"
    Dim Key As String
    Dim api_ticker As String
    Dim coFailure As Boolean
    Dim index_pos As Integer
    
    On Error GoTo ErrorHandler
    
    index_pos = InStr(ticker, "$")
    ticker = VBA.UCase(ticker)
    
    If ticker <> "" And LoginFailure = False Then
        If CompanySuccessDic.Exists(ticker) = False Then
            api_ticker = IntrinioCompanies(ticker, "ticker")
            If api_ticker = ticker Then
                CompanySuccessDic.Add ticker, False
                coFailure = CompanySuccessDic(ticker)
            ElseIf index_pos = 1 Then
                CompanySuccessDic.Add ticker, False
                coFailure = CompanySuccessDic(ticker)
            Else
                api_ticker = IntrinioSecurities(ticker, "ticker")
                If api_ticker = ticker Then
                    CompanySuccessDic.Add ticker, False
                    coFailure = CompanySuccessDic(ticker)
                Else
                    CompanySuccessDic.Add ticker, True
                    coFailure = CompanySuccessDic(ticker)
                End If
            End If
        Else
            If APICallsAtLimit = False Then
                coFailure = CompanySuccessDic(ticker)
            Else
                coFailure = False
            End If
        End If
    End If
    
    If ticker <> "" And Item <> "" And LoginFailure = False And APICallsAtLimit = False And coFailure = False Then
        Key = ticker
        If NewsDic.Exists(Key) = False Then
            Dim IntrinioClient As New WebClient
            IntrinioClient.BaseUrl = BaseUrl
            If iCredentials.Exists("username") = False Or iCredentials.Exists("password") = False Or iCredentials("username") = Empty Or iCredentials("password") = Empty Then
                Call IntrinioInitialize
            End If

            Dim inUsername As String
            Dim inPassword As String
            inUsername = iCredentials("username")
            inPassword = iCredentials("password")
            Dim Auth As New HttpBasicAuthenticator
            Auth.Setup _
                Username:=inUsername, _
                Password:=inPassword
            Set IntrinioClient.Authenticator = Auth
             
            Dim Request As New WebRequest
            Request.Resource = "news"
            Request.Method = HttpGet
            Request.Format = Json
            Request.AddQuerystringParam "ticker", ticker
            
            Dim Response As WebResponse
            Set Response = IntrinioClient.Execute(Request)
            
            If Response.StatusCode = Ok Then
                NewsDic.Add Key, Response.Data("data")
                IntrinioNews = NewsDic(Key)(sequence + 1)(Item)
                If IntrinioNews = Empty Then
                    IntrinioNews = ""
                End If
            ElseIf Response.StatusCode = 429 Then
                APICallsAtLimit = True
                IntrinioNews = "Plan Limit Reached"
            ElseIf Response.StatusCode = 403 Then
                IntrinioNews = "Visit Intrinio.com to Subscribe"
            Else
                IntrinioNews = ""
            End If
        ElseIf NewsDic.Exists(Key) = True Then
            IntrinioNews = NewsDic(Key)(sequence + 1)(Item)
            If IntrinioNews = Empty Then
                IntrinioNews = ""
            End If
        End If

    Else
        If ticker = "" Then
            IntrinioNews = ""
        ElseIf Item = "" Then
            IntrinioNews = ""
        ElseIf APICallsAtLimit = True Then
            Key = ticker
            If NewsDic.Exists(Key) = True Then
                IntrinioNews = NewsDic(Key)(sequence + 1)(Item)
                If IntrinioNews = Empty Then
                    IntrinioNews = ""
                End If
            Else
                IntrinioNews = "Plan Limit Reached"
            End If
            
        ElseIf LoginFailure = True Then
            IntrinioNews = "Invalid API Keys"
        ElseIf coFailure = True Then
            IntrinioNews = "Invalid Ticker Symbol"
        Else
            IntrinioNews = ""
        End If
    End If
ExitHere:
    Exit Function
ErrorHandler:
    IntrinioNews = ""
End Function

Sub DescribeIntrinioStandardizedFundamentals()
    Dim FuncName As String
    Dim FuncDesc As String
    Dim Category As String
    Dim ArgDesc(1 To 5) As String
    
    FuncName = "IntrinioStandardizedFundamentals"
    FuncDesc = "Returns a standardized financial statement fundamental based on a period type and sequence number selected."
    Category = "Intrinio"
    ArgDesc(1) = "The company's ticker symbol (i.e. 'AMZN')"
    ArgDesc(2) = "The financial statement selected ('income_statement','balance_sheet','cash_flow_statement','calculations')"
    ArgDesc(3) = "The period type ('FY','QTR','TTM','YTD')"
    ArgDesc(4) = "The sequence order of the fundamental from newest to oldest (0..last available)"
    ArgDesc(5) = "The item you are selecting (i.e. 'fiscal_year' returns 2014, 'fiscal_period' returns 'FY', 'end_date' returns the last date of the period, 'start_date' returns the beginning of the period)"
    
    Application.MacroOptions Macro:=FuncName, _
       Description:=FuncDesc, _
       Category:=Category, _
       ArgumentDescriptions:=ArgDesc
End Sub

Sub DescribeIntrinioFundamentals()
    Dim FuncName As String
    Dim FuncDesc As String
    Dim Category As String
    Dim ArgDesc(1 To 5) As String
    
    FuncName = "IntrinioFundamentals"
    FuncDesc = "Returns a standardized financial statement fundamental based on a period type and sequence number selected."
    Category = "Intrinio"
    ArgDesc(1) = "The company's ticker symbol (i.e. 'AMZN')"
    ArgDesc(2) = "The financial statement selected ('income_statement','balance_sheet','cash_flow_statement','calculations')"
    ArgDesc(3) = "The period type ('FY','QTR','TTM','YTD')"
    ArgDesc(4) = "The sequence order of the fundamental from newest to oldest (0..last available)"
    ArgDesc(5) = "The item you are selecting (i.e. 'fiscal_year' returns 2014, 'fiscal_period' returns 'FY', 'end_date' returns the last date of the period, 'start_date' returns the beginning of the period)"
    
    Application.MacroOptions Macro:=FuncName, _
       Description:=FuncDesc, _
       Category:=Category, _
       ArgumentDescriptions:=ArgDesc
End Sub

Public Function IntrinioStandardizedFundamentals(ticker As String, _
                           statement As String, _
                           period_type As String, _
                           sequence As Integer, _
                           Item As String)
Attribute IntrinioStandardizedFundamentals.VB_Description = "Returns a standardized financial statement fundamental based on a period type and sequence number selected."
Attribute IntrinioStandardizedFundamentals.VB_ProcData.VB_Invoke_Func = " \n19"
    Dim Key As String
    Dim api_ticker As String
    Dim coFailure As Boolean
    
    On Error GoTo ErrorHandler
    
    ticker = VBA.UCase(ticker)
    
    If ticker <> "" And LoginFailure = False Then
        If CompanySuccessDic.Exists(ticker) = False Then
            api_ticker = IntrinioCompanies(ticker, "ticker")
            If api_ticker = ticker Then
                CompanySuccessDic.Add ticker, False
                coFailure = CompanySuccessDic(ticker)
            Else
                api_ticker = IntrinioSecurities(ticker, "ticker")
                If api_ticker = ticker Then
                    CompanySuccessDic.Add ticker, False
                    coFailure = CompanySuccessDic(ticker)
                Else
                    CompanySuccessDic.Add ticker, True
                    coFailure = CompanySuccessDic(ticker)
                End If
            End If
        Else
            If APICallsAtLimit = False Then
                coFailure = CompanySuccessDic(ticker)
            Else
                coFailure = False
            End If
        End If
    End If
    
    If ticker <> "" And statement <> "" And period_type <> "" And LoginFailure = False And APICallsAtLimit = False And coFailure = False Then
        Key = ticker & "_" & statement & "_" & period_type
        
        If FundamentalsDic.Exists(Key) = False Then
            Dim IntrinioClient As New WebClient
            IntrinioClient.BaseUrl = BaseUrl
            If iCredentials.Exists("username") = False Or iCredentials.Exists("password") = False Or iCredentials("username") = Empty Or iCredentials("password") = Empty Then
                Call IntrinioInitialize
            End If
            
            Dim inUsername As String
            Dim inPassword As String
            inUsername = iCredentials("username")
            inPassword = iCredentials("password")
            Dim Auth As New HttpBasicAuthenticator
            Auth.Setup _
                Username:=inUsername, _
                Password:=inPassword
            Set IntrinioClient.Authenticator = Auth
            
            Dim Request As New WebRequest
            Request.Resource = "fundamentals/standardized"
            Request.Method = HttpGet
            Request.Format = Json
            Request.AddQuerystringParam "ticker", ticker
            Request.AddQuerystringParam "statement", statement
            Request.AddQuerystringParam "type", period_type
            
            Dim Response As WebResponse
            Set Response = IntrinioClient.Execute(Request)
            If Response.StatusCode = Ok Then
                FundamentalsDic.Add Key, Response.Data("data")
                IntrinioStandardizedFundamentals = FundamentalsDic(Key)(sequence + 1)(Item)
            ElseIf Response.StatusCode = 429 Then
                APICallsAtLimit = True
                IntrinioStandardizedFundamentals = "Plan Limit Reached"
            ElseIf Response.StatusCode = 403 Then
                IntrinioStandardizedFundamentals = "Visit Intrinio.com to Subscribe"
            Else
                IntrinioStandardizedFundamentals = ""
            End If
        ElseIf FundamentalsDic.Exists(Key) = True Then
            IntrinioStandardizedFundamentals = FundamentalsDic(Key)(sequence + 1)(Item)
        End If
    Else
        If APICallsAtLimit = True Then
            Key = ticker & "_" & statement & "_" & period_type
            If FundamentalsDic.Exists(Key) = True Then
                IntrinioStandardizedFundamentals = FundamentalsDic(Key)(sequence + 1)(Item)
            Else
                IntrinioStandardizedFundamentals = "Plan Limit Reached"
            End If
        ElseIf LoginFailure = True Then
            IntrinioStandardizedFundamentals = "Invalid API Keys"
        ElseIf coFailure = True Then
            IntrinioStandardizedFundamentals = "Invalid Ticker Symbol"
        Else
            IntrinioStandardizedFundamentals = ""
        End If
    End If
ExitHere:
    Exit Function
ErrorHandler:
    IntrinioStandardizedFundamentals = ""
End Function

Public Function IntrinioFundamentals(ticker As String, _
                           statement As String, _
                           period_type As String, _
                           sequence As Integer, _
                           Item As String)
Attribute IntrinioFundamentals.VB_Description = "Returns a standardized financial statement fundamental based on a period type and sequence number selected."
Attribute IntrinioFundamentals.VB_ProcData.VB_Invoke_Func = " \n19"
    IntrinioFundamentals = IntrinioStandardizedFundamentals(ticker, statement, period_type, sequence, Item)
End Function

Sub DescribeIntrinioStandardizedTags()
    Dim FuncName As String
    Dim FuncDesc As String
    Dim Category As String
    Dim ArgDesc(1 To 4) As String
    
    FuncName = "IntrinioStandardizedTags"
    FuncDesc = "Returns a standardized tag for a selected company and financial statement, by selecting a specific tag based on the sequence number selected."
    Category = "Intrinio"
    ArgDesc(1) = "The company's ticker symbol (i.e. 'CSCO')"
    ArgDesc(2) = "The financial statement selected ('income_statement','balance_sheet','cash_flow_statement','calculations')"
    ArgDesc(3) = "The sequence order of the tag from first to last (0..last available)"
    ArgDesc(4) = "The item you are selecting (i.e. 'name' returns the human readable name, 'tag' returns the standardized tag, 'balance' returns debit or credit, 'unit' returns the units for the tag)"
    
    Application.MacroOptions Macro:=FuncName, _
       Description:=FuncDesc, _
       Category:=Category, _
       ArgumentDescriptions:=ArgDesc
End Sub

Sub DescribeIntrinioTags()
    Dim FuncName As String
    Dim FuncDesc As String
    Dim Category As String
    Dim ArgDesc(1 To 4) As String
    
    FuncName = "IntrinioTags"
    FuncDesc = "Returns a standardized tag for a selected company and financial statement, by selecting a specific tag based on the sequence number selected."
    Category = "Intrinio"
    ArgDesc(1) = "The company's ticker symbol (i.e. 'CSCO')"
    ArgDesc(2) = "The financial statement selected ('income_statement','balance_sheet','cash_flow_statement','calculations')"
    ArgDesc(3) = "The sequence order of the tag from first to last (0..last available)"
    ArgDesc(4) = "The item you are selecting (i.e. 'name' returns the human readable name, 'tag' returns the standardized tag, 'balance' returns debit or credit, 'unit' returns the units for the tag)"
    
    Application.MacroOptions Macro:=FuncName, _
       Description:=FuncDesc, _
       Category:=Category, _
       ArgumentDescriptions:=ArgDesc
End Sub

Public Function IntrinioStandardizedTags(ticker As String, _
                           statement As String, _
                           sequence As Integer, _
                           Item As String)
Attribute IntrinioStandardizedTags.VB_Description = "Returns a standardized tag for a selected company and financial statement, by selecting a specific tag based on the sequence number selected."
Attribute IntrinioStandardizedTags.VB_ProcData.VB_Invoke_Func = " \n19"
    Dim Key As String
    Dim api_ticker As String
    Dim coFailure As Boolean
    
    On Error GoTo ErrorHandler
    
    ticker = VBA.UCase(ticker)
    
    If ticker <> "" And LoginFailure = False Then
        If CompanySuccessDic.Exists(ticker) = False Then
            api_ticker = IntrinioCompanies(ticker, "ticker")
            If api_ticker = ticker Then
                CompanySuccessDic.Add ticker, False
                coFailure = CompanySuccessDic(ticker)
            Else
                api_ticker = IntrinioSecurities(ticker, "ticker")
                If api_ticker = ticker Then
                    CompanySuccessDic.Add ticker, False
                    coFailure = CompanySuccessDic(ticker)
                Else
                    CompanySuccessDic.Add ticker, True
                    coFailure = CompanySuccessDic(ticker)
                End If
            End If
        Else
            If APICallsAtLimit = False Then
                coFailure = CompanySuccessDic(ticker)
            Else
                coFailure = False
            End If
        End If
    End If
    
    If ticker <> "" And statement <> "" And LoginFailure = False And APICallsAtLimit = False And coFailure = False Then
        Key = ticker & "_" & statement
        
        If StandardizedTagsDic.Exists(Key) = False Then
            Dim IntrinioClient As New WebClient
            IntrinioClient.BaseUrl = BaseUrl
            If iCredentials.Exists("username") = False Or iCredentials.Exists("password") = False Or iCredentials("username") = Empty Or iCredentials("password") = Empty Then
                Call IntrinioInitialize
            End If
            
            Dim inUsername As String
            Dim inPassword As String
            inUsername = iCredentials("username")
            inPassword = iCredentials("password")
            Dim Auth As New HttpBasicAuthenticator
            Auth.Setup _
                Username:=inUsername, _
                Password:=inPassword
            Set IntrinioClient.Authenticator = Auth
            
            Dim Request As New WebRequest
            Request.Resource = "tags/standardized"
            Request.Method = HttpGet
            Request.Format = Json
            Request.AddQuerystringParam "ticker", ticker
            Request.AddQuerystringParam "statement", statement
            
            Dim Response As WebResponse
            Set Response = IntrinioClient.Execute(Request)
            If Response.StatusCode = Ok Then
                StandardizedTagsDic.Add Key, Response.Data("data")
                IntrinioStandardizedTags = StandardizedTagsDic(Key)(sequence + 1)(Item)
            ElseIf Response.StatusCode = 429 Then
                APICallsAtLimit = True
                IntrinioStandardizedTags = "Plan Limit Reached"
            ElseIf Response.StatusCode = 403 Then
                IntrinioStandardizedTags = "Visit Intrinio.com to Subscribe"
            Else
                IntrinioStandardizedTags = ""
            End If
        ElseIf StandardizedTagsDic.Exists(Key) = True Then
            IntrinioStandardizedTags = StandardizedTagsDic(Key)(sequence + 1)(Item)
        End If
    Else
        If APICallsAtLimit = True Then
            Key = ticker & "_" & statement
            
            If StandardizedTagsDic.Exists(Key) = True Then
            IntrinioStandardizedTags = StandardizedTagsDic(Key)(sequence + 1)(Item)
            Else
                IntrinioStandardizedTags = "Plan Limit Reached"
            End If
            
        ElseIf LoginFailure = True Then
            IntrinioStandardizedTags = "Invalid API Keys"
        ElseIf coFailure = True Then
            IntrinioStandardizedTags = "Invalid Ticker Symbol"
        Else
            IntrinioStandardizedTags = ""
        End If
    End If
ExitHere:
    Exit Function
ErrorHandler:
    IntrinioStandardizedTags = ""
End Function

Public Function IntrinioTags(ticker As String, _
                           statement As String, _
                           sequence As Integer, _
                           Item As String)
Attribute IntrinioTags.VB_Description = "Returns a standardized tag for a selected company and financial statement, by selecting a specific tag based on the sequence number selected."
Attribute IntrinioTags.VB_ProcData.VB_Invoke_Func = " \n19"
    
    IntrinioTags = IntrinioStandardizedTags(ticker, statement, sequence, Item)
                           
End Function

Sub DescribeIntrinioStandardizedFinancials()
    Dim FuncName As String
    Dim FuncDesc As String
    Dim Category As String
    Dim ArgDesc(1 To 6) As String
    
    FuncName = "IntrinioStandardizedFinancials"
    FuncDesc = "Returns a historical standardized financial statement data point for a company, based on the tag, fiscal year and fiscal period."
    Category = "Intrinio"
    ArgDesc(1) = "The company's ticker symbol (i.e. 'ORCL')"
    ArgDesc(2) = "The financial statement selected ('income_statement','balance_sheet','cash_flow_statement','calculations')"
    ArgDesc(3) = "The selected fiscal year for the chosen statement (i.e. 2014, 2013, 2012, etc.)"
    ArgDesc(4) = "The selected fiscal period for the chosen statement ('FY', 'Q1', 'Q2', 'Q3', 'Q1TTM', 'Q2TTM', 'Q3TTM', 'Q2YTD', 'Q3YTD')"
    ArgDesc(5) = "The selected tag contained within the statement (i.e. 'totalrevenue', 'netppe', 'totalequity', 'purchaseofplantpropertyandequipment', 'netchangeincash')"
    ArgDesc(6) = "(Optional) Round the value (blank or 'A' for actuals, 'K' for thousands, 'M' for millions, 'B' for billions)"
    
    Application.MacroOptions Macro:=FuncName, _
       Description:=FuncDesc, _
       Category:=Category, _
       ArgumentDescriptions:=ArgDesc
End Sub

Sub DescribeIntrinioFinancials()
    Dim FuncName As String
    Dim FuncDesc As String
    Dim Category As String
    Dim ArgDesc(1 To 6) As String
    
    FuncName = "IntrinioFinancials"
    FuncDesc = "Returns a historical standardized financial statement data point for a company, based on the tag, fiscal year and fiscal period."
    Category = "Intrinio"
    ArgDesc(1) = "The company's ticker symbol (i.e. 'ORCL')"
    ArgDesc(2) = "The financial statement selected ('income_statement','balance_sheet','cash_flow_statement','calculations')"
    ArgDesc(3) = "The selected fiscal year for the chosen statement (i.e. 2014, 2013, 2012, etc.)"
    ArgDesc(4) = "The selected fiscal period for the chosen statement ('FY', 'Q1', 'Q2', 'Q3', 'Q1TTM', 'Q2TTM', 'Q3TTM', 'Q2YTD', 'Q3YTD')"
    ArgDesc(5) = "The selected tag contained within the statement (i.e. 'totalrevenue', 'netppe', 'totalequity', 'purchaseofplantpropertyandequipment', 'netchangeincash')"
    ArgDesc(6) = "(Optional) Round the value (blank or 'A' for actuals, 'K' for thousands, 'M' for millions, 'B' for billions)"

    Application.MacroOptions Macro:=FuncName, _
       Description:=FuncDesc, _
       Category:=Category, _
       ArgumentDescriptions:=ArgDesc
End Sub

Public Function IntrinioStandardizedFinancials(ticker As String, _
                           statement As String, _
                           fiscal_year As Integer, _
                           fiscal_period As String, _
                           tag As String, _
                           Optional rounding As String = "A")
Attribute IntrinioStandardizedFinancials.VB_Description = "Returns a historical standardized financial statement data point for a company, based on the tag, fiscal year and fiscal period."
Attribute IntrinioStandardizedFinancials.VB_ProcData.VB_Invoke_Func = " \n19"
                           
    Dim Key As String
    Dim eKey As String
    Dim nKey As String
    Dim X As Variant
    Dim rTag As String
    Dim rValue As Double
    Dim sValue As String
    Dim Value As Double
    Dim Rounder As Double
    Dim api_ticker As String
    Dim coFailure As Boolean
    Dim fundamental_sequence As Integer
    Dim fundamental_type As String
    
    On Error GoTo ErrorHandler
    
    ticker = VBA.UCase(ticker)
    
    If ticker <> "" And LoginFailure = False Then
        If CompanySuccessDic.Exists(ticker) = False Then
            api_ticker = IntrinioCompanies(ticker, "ticker")
            If api_ticker = ticker Then
                CompanySuccessDic.Add ticker, False
                coFailure = CompanySuccessDic(ticker)
            Else
                api_ticker = IntrinioSecurities(ticker, "ticker")
                If api_ticker = ticker Then
                    CompanySuccessDic.Add ticker, False
                    coFailure = CompanySuccessDic(ticker)
                Else
                    CompanySuccessDic.Add ticker, True
                    coFailure = CompanySuccessDic(ticker)
                End If
            End If
        Else
            If APICallsAtLimit = False Then
                coFailure = CompanySuccessDic(ticker)
            Else
                coFailure = False
            End If
        End If
    End If
    
    
    If ticker <> "" And statement <> "" And LoginFailure = False And APICallsAtLimit = False And coFailure = False Then
        If fiscal_year < 1900 Then
            fundamental_type = fiscal_period
            fundamental_sequence = fiscal_year
            fiscal_year = IntrinioStandardizedFundamentals(ticker, statement, fundamental_type, fundamental_sequence, "fiscal_year")
            fiscal_period = IntrinioStandardizedFundamentals(ticker, statement, fundamental_type, fundamental_sequence, "fiscal_period")
        End If
    End If
    
    If ticker <> "" And statement <> "" And fiscal_year <> 0 And fiscal_period <> "" And LoginFailure = False And APICallsAtLimit = False And coFailure = False Then
        

        Key = ticker & "_" & statement & "_" & fiscal_year & "_" & fiscal_period
        
        If StandardizedFinancialsDic.Exists(Key) = False Then
            Dim IntrinioClient As New WebClient
            IntrinioClient.BaseUrl = BaseUrl
            If iCredentials.Exists("username") = False Or iCredentials.Exists("password") = False Or iCredentials("username") = Empty Or iCredentials("password") = Empty Then
                Call IntrinioInitialize
            End If
            
            Dim inUsername As String
            Dim inPassword As String
            inUsername = iCredentials("username")
            inPassword = iCredentials("password")
            Dim Auth As New HttpBasicAuthenticator
            Auth.Setup _
                Username:=inUsername, _
                Password:=inPassword
            Set IntrinioClient.Authenticator = Auth
            
            Dim Request As New WebRequest
            Dim last_page As Integer
            Dim is_last_page As Boolean
            Dim page As Integer
            Dim Response As WebResponse
            
            page = 1
            
            Do Until is_last_page = True
                Request.Resource = "financials/standardized"
                Request.Method = HttpGet
                Request.Format = Json
                Request.AddQuerystringParam "ticker", ticker
                Request.AddQuerystringParam "statement", statement
                Request.AddQuerystringParam "fiscal_year", fiscal_year
                Request.AddQuerystringParam "fiscal_period", fiscal_period
                Request.AddQuerystringParam "page_size", 400
                Request.AddQuerystringParam "page_number", page
                
                Set Response = IntrinioClient.Execute(Request)

                If Response.StatusCode = Ok Then
                    If Response.Content <> "" Then
                        last_page = Response.Data("total_pages")
                        If last_page > 0 Then
                            If last_page = page Then
                                is_last_page = True
                            Else
                                is_last_page = False
                                page = page + 1
                            End If
                            
                            If StandardizedFinancialsDic.Exists(Key) = False Then
                                StandardizedFinancialsDic.Add Key, Response.Data("data")
                            ElseIf StandardizedFinancialsDic.Exists(Key) = True Then
                                StandardizedFinancialsDic.Remove (Key)
                                StandardizedFinancialsDic.Add Key, Response.Data("data")
                            End If
                            
                            For Each X In StandardizedFinancialsDic(Key)
                                rTag = X("tag")
                                sValue = X("value")
                                nKey = ticker & "_" & statement & "_" & fiscal_year & "_" & fiscal_period & "_" & rTag
                                If StandardizedFinancialsDic.Exists(nKey) = True Then
                                    StandardizedFinancialsDic.Remove (nKey)
                                End If
                                StandardizedFinancialsDic.Add nKey, sValue
                            Next
                        Else
                            is_last_page = True
                            If StandardizedFinancialsDic.Exists(Key) = False Then
                                StandardizedFinancialsDic.Add Key, Response.Data("data")
                            ElseIf StandardizedFinancialsDic.Exists(Key) = True Then
                                StandardizedFinancialsDic.Remove (Key)
                                StandardizedFinancialsDic.Add Key, Response.Data("data")
                            End If
                        End If
                    Else
                        is_last_page = True
                        If StandardizedFinancialsDic.Exists(Key) = False Then
                            StandardizedFinancialsDic.Add Key, Empty
                        ElseIf StandardizedFinancialsDic.Exists(Key) = True Then
                            StandardizedFinancialsDic.Remove (Key)
                            StandardizedFinancialsDic.Add Key, Empty
                        End If
                    End If
                Else
                    is_last_page = True
                    If Response.StatusCode = 429 Then
                        APICallsAtLimit = True
                        IntrinioStandardizedFinancials = "Plan Limit Reached"
                    ElseIf Response.StatusCode = 403 Then
                        IntrinioStandardizedFinancials = "Visit Intrinio.com to Subscribe"
                    Else
                        IntrinioStandardizedFinancials = ""
                    End If
                End If
            Loop
                            
            eKey = ticker & "_" & statement & "_" & fiscal_year & "_" & fiscal_period & "_" & tag

            If StandardizedFinancialsDic.Exists(Key) = True Then

                If StandardizedFinancialsDic(Key) Is Not Empty Then
                    
                    If IsNumeric(StandardizedFinancialsDic(eKey)) = True Then
                        Value = StandardizedFinancialsDic(eKey)
                        If rounding = "K" Then
                            Rounder = 1000
                        ElseIf rounding = "M" Then
                            Rounder = 1000000
                        ElseIf rounding = "B" Then
                            Rounder = 1000000000
                        Else
                            Rounder = 1
                        End If
                    
                        IntrinioStandardizedFinancials = Value / Rounder
                    Else
                        IntrinioStandardizedFinancials = StandardizedFinancialsDic(eKey)
                    End If
                Else
                    IntrinioStandardizedFinancials = ""
                End If
            Else
                IntrinioStandardizedFinancials = ""
            End If
        ElseIf StandardizedFinancialsDic.Exists(Key) = True Then
            eKey = ticker & "_" & statement & "_" & fiscal_year & "_" & fiscal_period & "_" & tag
            
            If IsNumeric(StandardizedFinancialsDic(eKey)) = True Then
                Value = StandardizedFinancialsDic(eKey)
                
                If rounding = "K" Then
                    Rounder = 1000
                ElseIf rounding = "M" Then
                    Rounder = 1000000
                ElseIf rounding = "B" Then
                    Rounder = 1000000000
                Else
                    Rounder = 1
                End If
                
                IntrinioStandardizedFinancials = Value / Rounder
            Else
                IntrinioStandardizedFinancials = StandardizedFinancialsDic(eKey)
            End If
        End If
    Else
        If APICallsAtLimit = True Then
            IntrinioStandardizedFinancials = "Plan Limit Reached"
        ElseIf LoginFailure = True Then
            IntrinioStandardizedFinancials = "Invalid API Keys"
        ElseIf coFailure = True Then
            IntrinioStandardizedFinancials = "Invalid Ticker Symbol"
        Else
            IntrinioStandardizedFinancials = ""
        End If
    End If
ExitHere:
    Exit Function
ErrorHandler:
    IntrinioStandardizedFinancials = ""
    Resume Next
End Function

Public Function IntrinioFinancials(ticker As String, _
                           statement As String, _
                           fiscal_year As Integer, _
                           fiscal_period As String, _
                           tag As String, _
                           Optional rounding As String = "A")
Attribute IntrinioFinancials.VB_Description = "Returns a historical standardized financial statement data point for a company, based on the tag, fiscal year and fiscal period."
Attribute IntrinioFinancials.VB_ProcData.VB_Invoke_Func = " \n19"
                           
    IntrinioFinancials = IntrinioStandardizedFinancials(ticker, statement, fiscal_year, fiscal_period, tag, rounding)
                           
End Function

Sub DescribeIntrinioReportedFundamentals()
    Dim FuncName As String
    Dim FuncDesc As String
    Dim Category As String
    Dim ArgDesc(1 To 5) As String
    
    FuncName = "IntrinioReportedFundamentals"
    FuncDesc = "Returns a historical as reported financial statement fundamental based on the period type selected"
    Category = "Intrinio"
    ArgDesc(1) = "The company's ticker symbol (i.e. 'AMZN')"
    ArgDesc(2) = "The financial statement selected ('income_statement','balance_sheet','cash_flow_statement')"
    ArgDesc(3) = "The period type ('FY','QTR')"
    ArgDesc(4) = "The sequence order of the fundamental from newest to oldest (0..last available)"
    ArgDesc(5) = "The item you are selecting (i.e. 'fiscal_year' returns 2014, 'fiscal_period' returns 'FY', 'end_date' returns the last date of the period)"
    
    Application.MacroOptions Macro:=FuncName, _
       Description:=FuncDesc, _
       Category:=Category, _
       ArgumentDescriptions:=ArgDesc
End Sub

Public Function IntrinioReportedFundamentals(ticker As String, _
                           statement As String, _
                           period_type As String, _
                           sequence As Integer, _
                           Item As String)
Attribute IntrinioReportedFundamentals.VB_Description = "Returns a historical as reported financial statement fundamental based on the period type selected"
Attribute IntrinioReportedFundamentals.VB_ProcData.VB_Invoke_Func = " \n19"
    Dim Key As String
    Dim api_ticker As String
    Dim coFailure As Boolean
    
    On Error GoTo ErrorHandler
    
    ticker = VBA.UCase(ticker)
    
    If ticker <> "" And LoginFailure = False Then
        If CompanySuccessDic.Exists(ticker) = False Then
            api_ticker = IntrinioCompanies(ticker, "ticker")
            If api_ticker = ticker Then
                CompanySuccessDic.Add ticker, False
                coFailure = CompanySuccessDic(ticker)
            Else
                api_ticker = IntrinioSecurities(ticker, "ticker")
                If api_ticker = ticker Then
                    CompanySuccessDic.Add ticker, False
                    coFailure = CompanySuccessDic(ticker)
                Else
                    CompanySuccessDic.Add ticker, True
                    coFailure = CompanySuccessDic(ticker)
                End If
            End If
        Else
            If APICallsAtLimit = False Then
                coFailure = CompanySuccessDic(ticker)
            Else
                coFailure = False
            End If
        End If
    End If
    
    If ticker <> "" And statement <> "" And period_type <> "" And LoginFailure = False And APICallsAtLimit = False And coFailure = False Then
        Key = ticker & "_" & statement & "_" & period_type
        
        If ReportedFundamentalsDic.Exists(Key) = False Then
            Dim IntrinioClient As New WebClient
            IntrinioClient.BaseUrl = BaseUrl
            If iCredentials.Exists("username") = False Or iCredentials.Exists("password") = False Or iCredentials("username") = Empty Or iCredentials("password") = Empty Then
                Call IntrinioInitialize
            End If
            
            Dim inUsername As String
            Dim inPassword As String
            inUsername = iCredentials("username")
            inPassword = iCredentials("password")
            Dim Auth As New HttpBasicAuthenticator
            Auth.Setup _
                Username:=inUsername, _
                Password:=inPassword
            Set IntrinioClient.Authenticator = Auth
            
            Dim Request As New WebRequest
            Request.Resource = "fundamentals/reported"
            Request.Method = HttpGet
            Request.Format = Json
            Request.AddQuerystringParam "ticker", ticker
            Request.AddQuerystringParam "statement", statement
            Request.AddQuerystringParam "type", period_type
            
            Dim Response As WebResponse
            Set Response = IntrinioClient.Execute(Request)
            If Response.StatusCode = Ok Then
                ReportedFundamentalsDic.Add Key, Response.Data("data")
                IntrinioReportedFundamentals = ReportedFundamentalsDic(Key)(sequence + 1)(Item)
            ElseIf Response.StatusCode = 429 Then
                APICallsAtLimit = True
                IntrinioReportedFundamentals = "Plan Limit Reached"
            ElseIf Response.StatusCode = 403 Then
                IntrinioReportedFundamentals = "Visit Intrinio.com to Subscribe"
            Else
                IntrinioReportedFundamentals = ""
            End If
        ElseIf ReportedFundamentalsDic.Exists(Key) = True Then
            IntrinioReportedFundamentals = ReportedFundamentalsDic(Key)(sequence + 1)(Item)
        End If
    Else
        If APICallsAtLimit = True Then
            Key = ticker & "_" & statement & "_" & period_type
            If ReportedFundamentalsDic.Exists(Key) = True Then
                IntrinioReportedFundamentals = ReportedFundamentalsDic(Key)(sequence + 1)(Item)
            Else
                IntrinioReportedFundamentals = "Plan Limit Reached"
            End If
        ElseIf LoginFailure = True Then
            IntrinioReportedFundamentals = "Invalid API Keys"
        ElseIf coFailure = True Then
            IntrinioReportedFundamentals = "Invalid Ticker Symbol"
        Else
            IntrinioReportedFundamentals = ""
        End If
    End If
ExitHere:
    Exit Function
ErrorHandler:
    IntrinioReportedFundamentals = ""
End Function

Sub DescribeIntrinioReportedTags()
    Dim FuncName As String
    Dim FuncDesc As String
    Dim Category As String
    Dim ArgDesc(1 To 6) As String
    
    FuncName = "IntrinioReportedTags"
    FuncDesc = "Returns the as reported XBRL tags and labels for a given ticker, statement, and date or fiscal period."
    Category = "Intrinio"
    ArgDesc(1) = "The company's ticker symbol (i.e. 'IBM')"
    ArgDesc(2) = "The financial statement selected ('income_statement','balance_sheet','cash_flow_statement')"
    ArgDesc(3) = "The selected fiscal year for the chosen statement (i.e. 2014, 2013, 2012, etc.)"
    ArgDesc(4) = "The selected fiscal period for the chosen statement ('FY', 'Q1', 'Q2', 'Q3', 'Q2YTD', 'Q3YTD')"
    ArgDesc(5) = "The sequence order of the fundamental from newest to oldest (0..last available)"
    ArgDesc(6) = "The return value selected ('name', 'tag', 'domain_tag', 'balance', 'unit', 'abstract', 'depth')"
    
    Application.MacroOptions Macro:=FuncName, _
       Description:=FuncDesc, _
       Category:=Category, _
       ArgumentDescriptions:=ArgDesc

End Sub

Public Function IntrinioReportedTags(ticker As String, _
                           statement As String, _
                           fiscal_year As Integer, _
                           fiscal_period As String, _
                           sequence As Integer, _
                           Item As String)
Attribute IntrinioReportedTags.VB_Description = "Returns the as reported XBRL tags and labels for a given ticker, statement, and date or fiscal period."
Attribute IntrinioReportedTags.VB_ProcData.VB_Invoke_Func = " \n19"
    Dim Key As String
    Dim api_ticker As String
    Dim coFailure As Boolean
    Dim fundamental_sequence As Integer
    Dim fundamental_type As String
    Dim last_page As Integer

    On Error GoTo ErrorHandler
    
    ticker = VBA.UCase(ticker)
    
    If ticker <> "" And LoginFailure = False Then
        If CompanySuccessDic.Exists(ticker) = False Then
            api_ticker = IntrinioCompanies(ticker, "ticker")
            If api_ticker = ticker Then
                CompanySuccessDic.Add ticker, False
                coFailure = CompanySuccessDic(ticker)
            Else
                api_ticker = IntrinioSecurities(ticker, "ticker")
                If api_ticker = ticker Then
                    CompanySuccessDic.Add ticker, False
                    coFailure = CompanySuccessDic(ticker)
                Else
                    CompanySuccessDic.Add ticker, True
                    coFailure = CompanySuccessDic(ticker)
                End If
            End If
        Else
            If APICallsAtLimit = False Then
                coFailure = CompanySuccessDic(ticker)
            Else
                coFailure = False
            End If
        End If
    End If

    If ticker <> "" And statement <> "" And LoginFailure = False And APICallsAtLimit = False And coFailure = False Then
        If fiscal_year < 1900 Then
            fundamental_type = fiscal_period
            fundamental_sequence = fiscal_year
            fiscal_year = IntrinioReportedFundamentals(ticker, statement, fundamental_type, fundamental_sequence, "fiscal_year")
            fiscal_period = IntrinioReportedFundamentals(ticker, statement, fundamental_type, fundamental_sequence, "fiscal_period")
        End If
    End If
    
    If ticker <> "" And statement <> "" And fiscal_year <> 0 And fiscal_period <> "" And LoginFailure = False And APICallsAtLimit = False And coFailure = False Then
        Key = ticker & "_" & statement & "_" & fiscal_year & "_" & fiscal_period
        
        If ReportedTagsDic.Exists(Key) = False Then
            Dim IntrinioClient As New WebClient
            IntrinioClient.BaseUrl = BaseUrl
            If iCredentials.Exists("username") = False Or iCredentials.Exists("password") = False Or iCredentials("username") = Empty Or iCredentials("password") = Empty Then
                Call IntrinioInitialize
            End If
            
            Dim inUsername As String
            Dim inPassword As String
            inUsername = iCredentials("username")
            inPassword = iCredentials("password")
            Dim Auth As New HttpBasicAuthenticator
            Auth.Setup _
                Username:=inUsername, _
                Password:=inPassword
            Set IntrinioClient.Authenticator = Auth
            
            Dim Request As New WebRequest
            Request.Resource = "tags/reported"
            Request.Method = HttpGet
            Request.Format = Json
            Request.AddQuerystringParam "ticker", ticker
            Request.AddQuerystringParam "statement", statement
            Request.AddQuerystringParam "fiscal_year", fiscal_year
            Request.AddQuerystringParam "fiscal_period", fiscal_period
            
            Dim Response As WebResponse
            Set Response = IntrinioClient.Execute(Request)

            If Response.StatusCode = Ok Then
                last_page = Response.Data("total_pages")
                If last_page > 0 Then
                    ReportedTagsDic.Add Key, Response.Data("data")
                    If Item = "domain_tag" Then
                        If IsNull(ReportedTagsDic(Key)(sequence + 1)(Item)) Then
                            IntrinioReportedTags = ""
                        Else
                            IntrinioReportedTags = ReportedTagsDic(Key)(sequence + 1)(Item)
                        End If
                    Else
                        IntrinioReportedTags = ReportedTagsDic(Key)(sequence + 1)(Item)
                    End If
                Else
                    IntrinioReportedTags = ""
                    ReportedTagsDic.Add Key, Response.Data("data")
                End If
            ElseIf Response.StatusCode = 429 Then
                APICallsAtLimit = True
                IntrinioReportedTags = "Plan Limit Reached"
            ElseIf Response.StatusCode = 403 Then
                IntrinioReportedTags = "Visit Intrinio.com to Subscribe"
            Else
                IntrinioReportedTags = ""
            End If
        ElseIf ReportedTagsDic.Exists(Key) = True Then
            If Item = "domain_tag" Then
                If IsNull(ReportedTagsDic(Key)(sequence + 1)(Item)) Then
                    IntrinioReportedTags = ""
                Else
                    IntrinioReportedTags = ReportedTagsDic(Key)(sequence + 1)(Item)
                End If
            Else
                IntrinioReportedTags = ReportedTagsDic(Key)(sequence + 1)(Item)
            End If
        End If
    Else
        If APICallsAtLimit = True Then
            Key = ticker & "_" & statement & "_" & fiscal_year & "_" & fiscal_period
            If ReportedTagsDic.Exists(Key) = True Then
                If Item = "domain_tag" Then
                    If IsNull(ReportedTagsDic(Key)(sequence + 1)(Item)) Then
                        IntrinioReportedTags = ""
                    Else
                        IntrinioReportedTags = ReportedTagsDic(Key)(sequence + 1)(Item)
                    End If
                Else
                    IntrinioReportedTags = ReportedTagsDic(Key)(sequence + 1)(Item)
                End If
            Else
                IntrinioReportedTags = "Plan Limit Reached"
            End If
        ElseIf LoginFailure = True Then
            IntrinioReportedTags = "Invalid API Keys"
        ElseIf coFailure = True Then
            IntrinioReportedTags = "Invalid Ticker Symbol"
        Else
            IntrinioReportedTags = ""
        End If
    End If
    
ExitHere:
    Exit Function
ErrorHandler:
    IntrinioReportedTags = ""
End Function

Sub DescribeIntrinioReportedFinancials()
    Dim FuncName As String
    Dim FuncDesc As String
    Dim Category As String
    Dim ArgDesc(1 To 6) As String
    
    FuncName = "IntrinioReportedFinancials"
    FuncDesc = "Returns a historical as reported financial statement data point for a company, based on the xbrl tag, domain tag, fiscal year and fiscal period."
    Category = "Intrinio"
    ArgDesc(1) = "The company's ticker symbol (i.e. 'INTC')"
    ArgDesc(2) = "The financial statement selected ('income_statement','balance_sheet','cash_flow_statement')"
    ArgDesc(3) = "The selected fiscal year for the chosen statement (i.e. 2014, 2013, 2012, etc.)"
    ArgDesc(4) = "The selected fiscal period for the chosen statement ('FY', 'Q1', 'Q2', 'Q3', 'Q2YTD', 'Q3YTD')"
    ArgDesc(5) = "The specified XBRL tag from the as reported statement"
    ArgDesc(6) = "(Optional) The specified domain XBRL tag, associated with certain data points on the financial statements that have a dimension associated with the data point"
    
    Application.MacroOptions Macro:=FuncName, _
       Description:=FuncDesc, _
       Category:=Category, _
       ArgumentDescriptions:=ArgDesc
End Sub

Public Function IntrinioReportedFinancials(ticker As String, _
                           statement As String, _
                           fiscal_year As Integer, _
                           fiscal_period As String, _
                           xbrl_tag As String, _
                           Optional domain_tag As String = "")
Attribute IntrinioReportedFinancials.VB_Description = "Returns a historical as reported financial statement data point for a company, based on the xbrl tag, domain tag, fiscal year and fiscal period."
Attribute IntrinioReportedFinancials.VB_ProcData.VB_Invoke_Func = " \n19"
    Dim Key As String
    Dim eKey As String
    Dim nKey As String
    Dim X As Variant
    Dim rXBRLTag As String
    Dim rDomainTag As String
    Dim rValue As Double
    Dim api_ticker As String
    Dim coFailure As Boolean
    Dim fundamental_sequence As Integer
    Dim fundamental_type As String
    Dim last_page As Integer
    
    On Error GoTo ErrorHandler
    
    ticker = VBA.UCase(ticker)

    If ticker <> "" And LoginFailure = False Then
        If CompanySuccessDic.Exists(ticker) = False Then
            api_ticker = IntrinioCompanies(ticker, "ticker")
            If api_ticker = ticker Then
                CompanySuccessDic.Add ticker, False
                coFailure = CompanySuccessDic(ticker)
            Else
                api_ticker = IntrinioSecurities(ticker, "ticker")
                If api_ticker = ticker Then
                    CompanySuccessDic.Add ticker, False
                    coFailure = CompanySuccessDic(ticker)
                Else
                    CompanySuccessDic.Add ticker, True
                    coFailure = CompanySuccessDic(ticker)
                End If
            End If
        Else
            If APICallsAtLimit = False Then
                coFailure = CompanySuccessDic(ticker)
            Else
                coFailure = False
            End If
        End If
    End If
    
    If ticker <> "" And statement <> "" And LoginFailure = False And APICallsAtLimit = False And coFailure = False Then
        If fiscal_year < 1900 Then
            fundamental_type = fiscal_period
            fundamental_sequence = fiscal_year
            fiscal_year = IntrinioReportedFundamentals(ticker, statement, fundamental_type, fundamental_sequence, "fiscal_year")
            fiscal_period = IntrinioReportedFundamentals(ticker, statement, fundamental_type, fundamental_sequence, "fiscal_period")
        End If
    End If
    
    If ticker <> "" And statement <> "" And fiscal_year <> 0 And fiscal_period <> "" And LoginFailure = False And APICallsAtLimit = False And coFailure = False Then
        
        Key = ticker & "_" & statement & "_" & fiscal_year & "_" & fiscal_period
        
        If ReportedFinancialsDic.Exists(Key) = False Then
            Dim IntrinioClient As New WebClient
            IntrinioClient.BaseUrl = BaseUrl
            If iCredentials.Exists("username") = False Or iCredentials.Exists("password") = False Or iCredentials("username") = Empty Or iCredentials("password") = Empty Then
                Call IntrinioInitialize
            End If
            
            Dim inUsername As String
            Dim inPassword As String
            inUsername = iCredentials("username")
            inPassword = iCredentials("password")
            Dim Auth As New HttpBasicAuthenticator
            Auth.Setup _
                Username:=inUsername, _
                Password:=inPassword
            Set IntrinioClient.Authenticator = Auth
            
            Dim Request As New WebRequest
            Request.Resource = "financials/reported"
            Request.Method = HttpGet
            Request.Format = Json
            Request.AddQuerystringParam "ticker", ticker
            Request.AddQuerystringParam "statement", statement
            Request.AddQuerystringParam "fiscal_year", fiscal_year
            Request.AddQuerystringParam "fiscal_period", fiscal_period
            
            Dim Response As WebResponse
            Set Response = IntrinioClient.Execute(Request)
            
            If Response.StatusCode = Ok Then
                last_page = Response.Data("total_pages")
                If last_page > 0 Then
                    ReportedFinancialsDic.Add Key, Response.Data("data")
                    For Each X In ReportedFinancialsDic(Key)
                        rXBRLTag = X("xbrl_tag")
                        If X("domain_tag") <> "null" Then
                            rDomainTag = X("domain_tag")
                            nKey = ticker & "_" & statement & "_" & fiscal_year & "_" & fiscal_period & "_" & rXBRLTag & "_" & rDomainTag
                        Else
                            nKey = ticker & "_" & statement & "_" & fiscal_year & "_" & fiscal_period & "_" & rXBRLTag
                        End If
                        rValue = X("value")
                        ReportedFinancialsDic.Add nKey, rValue
                    Next
                    eKey = ticker & "_" & statement & "_" & fiscal_year & "_" & fiscal_period & "_" & xbrl_tag & domain_tag
                    If VBA.Right(xbrl_tag, 8) = "Abstract" Then
                        IntrinioReportedFinancials = ""
                    ElseIf xbrl_tag = "" Then
                        IntrinioReportedFinancials = ""
                    Else
                        IntrinioReportedFinancials = ReportedFinancialsDic(eKey)
                    End If
                Else
                    ReportedFinancialsDic.Add Key, Response.Data("data")
                End If
            ElseIf Response.StatusCode = 429 Then
                APICallsAtLimit = True
                IntrinioReportedFinancials = "Plan Limit Reached"
            ElseIf Response.StatusCode = 403 Then
                IntrinioReportedFinancials = "Visit Intrinio.com to Subscribe"
            Else
                IntrinioReportedFinancials = ""
            End If
        ElseIf ReportedFinancialsDic.Exists(Key) = True Then
            If domain_tag = "" Then
                eKey = ticker & "_" & statement & "_" & fiscal_year & "_" & fiscal_period & "_" & xbrl_tag
            Else
                eKey = ticker & "_" & statement & "_" & fiscal_year & "_" & fiscal_period & "_" & xbrl_tag & "_" & domain_tag
            End If
            If VBA.Right(xbrl_tag, 8) = "Abstract" Then
                IntrinioReportedFinancials = ""
            ElseIf xbrl_tag = "" Then
                IntrinioReportedFinancials = ""
            Else
                IntrinioReportedFinancials = ReportedFinancialsDic(eKey)
            End If
        End If
    Else
        If LoginFailure = True Then
            Key = ticker & "_" & statement & "_" & fiscal_year & "_" & fiscal_period
            If ReportedFinancialsDic.Exists(Key) = True Then
                If domain_tag = "" Then
                    eKey = ticker & "_" & statement & "_" & fiscal_year & "_" & fiscal_period & "_" & xbrl_tag
                Else
                    eKey = ticker & "_" & statement & "_" & fiscal_year & "_" & fiscal_period & "_" & xbrl_tag & "_" & domain_tag
                End If
                If VBA.Right(xbrl_tag, 8) = "Abstract" Then
                    IntrinioReportedFinancials = ""
                ElseIf xbrl_tag = "" Then
                    IntrinioReportedFinancials = ""
                Else
                    IntrinioReportedFinancials = ReportedFinancialsDic(eKey)
                End If
            Else
                IntrinioReportedFinancials = "Invalid API Keys"
            End If
        ElseIf APICallsAtLimit = True Then
            IntrinioReportedFinancials = "Plan Limit Reached"
        ElseIf coFailure = True Then
            IntrinioReportedFinancials = "Invalid Ticker Symbol"
        Else
            IntrinioReportedFinancials = ""
        End If
    End If
ExitHere:
    Exit Function
ErrorHandler:
    IntrinioReportedFinancials = ""
    Resume Next
End Function

Sub DescribeIntrinioBankFundamentals()
    Dim FuncName As String
    Dim FuncDesc As String
    Dim Category As String
    Dim ArgDesc(1 To 5) As String
    
    FuncName = "IntrinioBankFundamentals"
    FuncDesc = "Returns a banks financial statement fundamental based on a period type and sequence number selected."
    Category = "Intrinio"
    ArgDesc(1) = "The company's identifier (i.e. ticker symbol 'JPM' or RSSD ID '361354')"
    ArgDesc(2) = "The financial statement selected ('RI')"
    ArgDesc(3) = "The period type ('FY','QTR','YTD')"
    ArgDesc(4) = "The sequence order of the fundamental from newest to oldest (0..last available)"
    ArgDesc(5) = "The item you are selecting (i.e. 'fiscal_year' returns 2014, 'fiscal_period' returns 'FY', 'end_date' returns the last date of the period, 'start_date' returns the beginning of the period)"
    
    Application.MacroOptions Macro:=FuncName, _
       Description:=FuncDesc, _
       Category:=Category, _
       ArgumentDescriptions:=ArgDesc
End Sub


Public Function IntrinioBankFundamentals(identifier As String, _
                           statement As String, _
                           period_type As String, _
                           sequence As Integer, _
                           Item As String)
Attribute IntrinioBankFundamentals.VB_Description = "Returns a banks financial statement fundamental based on a period type and sequence number selected."
Attribute IntrinioBankFundamentals.VB_ProcData.VB_Invoke_Func = " \n19"
    Dim Key As String
    Dim api_identifier As String
    Dim coFailure As Boolean
    
    On Error GoTo ErrorHandler
    
    If identifier <> "" And LoginFailure = False Then
        If CompanySuccessDic.Exists(identifier) = False Then
            api_identifier = IntrinioBanks(identifier, "identifier")
            If api_identifier = identifier Then
                CompanySuccessDic.Add identifier, False
                coFailure = CompanySuccessDic(identifier)
            Else
                CompanySuccessDic.Add identifier, True
                coFailure = CompanySuccessDic(identifier)
            End If
        Else
            If APICallsAtLimit = False Then
                coFailure = CompanySuccessDic(identifier)
            Else
                coFailure = False
            End If
        End If
    End If
    
    If identifier <> "" And statement <> "" And period_type <> "" And LoginFailure = False And APICallsAtLimit = False And coFailure = False Then
        Key = identifier & "_" & statement & "_" & period_type
        
        If BankFundamentalsDic.Exists(Key) = False Then
            Dim IntrinioClient As New WebClient
            IntrinioClient.BaseUrl = BaseUrl
            If iCredentials.Exists("username") = False Or iCredentials.Exists("password") = False Or iCredentials("username") = Empty Or iCredentials("password") = Empty Then
                Call IntrinioInitialize
            End If
            
            Dim inUsername As String
            Dim inPassword As String
            inUsername = iCredentials("username")
            inPassword = iCredentials("password")
            Dim Auth As New HttpBasicAuthenticator
            Auth.Setup _
                Username:=inUsername, _
                Password:=inPassword
            Set IntrinioClient.Authenticator = Auth
            
            Dim Request As New WebRequest
            Request.Resource = "fundamentals/banks"
            Request.Method = HttpGet
            Request.Format = Json
            Request.AddQuerystringParam "identifier", identifier
            Request.AddQuerystringParam "statement", statement
            Request.AddQuerystringParam "type", period_type
            
            Dim Response As WebResponse
            Set Response = IntrinioClient.Execute(Request)
            
            If Response.StatusCode = Ok Then
                BankFundamentalsDic.Add Key, Response.Data("data")
                IntrinioBankFundamentals = BankFundamentalsDic(Key)(sequence + 1)(Item)
            ElseIf Response.StatusCode = 429 Then
                APICallsAtLimit = True
                IntrinioBankFundamentals = "Plan Limit Reached"
            ElseIf Response.StatusCode = 403 Then
                IntrinioBankFundamentals = "Visit Intrinio.com to Subscribe"
            Else
                IntrinioBankFundamentals = ""
            End If
        ElseIf BankFundamentalsDic.Exists(Key) = True Then
            IntrinioBankFundamentals = BankFundamentalsDic(Key)(sequence + 1)(Item)
        End If
    Else
        If APICallsAtLimit = True Then
            Key = identifier & "_" & statement & "_" & period_type
            If BankFundamentalsDic.Exists(Key) = True Then
                IntrinioBankFundamentals = BankFundamentalsDic(Key)(sequence + 1)(Item)
            Else
                IntrinioBankFundamentals = "Plan Limit Reached"
            End If
        ElseIf LoginFailure = True Then
            IntrinioBankFundamentals = "Invalid API Keys"
        ElseIf coFailure = True Then
            IntrinioBankFundamentals = "Invalid identifier Symbol"
        Else
            IntrinioBankFundamentals = ""
        End If
    End If
ExitHere:
    Exit Function
ErrorHandler:
    IntrinioBankFundamentals = ""
End Function


Sub DescribeIntrinioBankTags()
    Dim FuncName As String
    Dim FuncDesc As String
    Dim Category As String
    Dim ArgDesc(1 To 4) As String
    
    FuncName = "IntrinioBankTags"
    FuncDesc = "Returns a bank tag for a selected bank and financial statement, by selecting a specific tag based on the sequence number selected."
    Category = "Intrinio"
    ArgDesc(1) = "The banks's Identifier (i.e. ticker symbol 'JPM' or RSSD ID '361354')"
    ArgDesc(2) = "The financial statement selected"
    ArgDesc(3) = "The sequence order of the tag from first to last (0..last available)"
    ArgDesc(4) = "The item you are selecting (i.e. 'name' returns the human readable name, 'tag' returns the Bank tag, 'balance' returns debit or credit, 'unit' returns the units for the tag)"
    
    Application.MacroOptions Macro:=FuncName, _
       Description:=FuncDesc, _
       Category:=Category, _
       ArgumentDescriptions:=ArgDesc
End Sub


Public Function IntrinioBankTags(identifier As String, _
                           statement As String, _
                           sequence As Integer, _
                           Item As String)
Attribute IntrinioBankTags.VB_Description = "Returns a bank tag for a selected bank and financial statement, by selecting a specific tag based on the sequence number selected."
Attribute IntrinioBankTags.VB_ProcData.VB_Invoke_Func = " \n19"
    Dim Key As String
    Dim api_identifier As String
    Dim coFailure As Boolean
    
    On Error GoTo ErrorHandler
    
    identifier = VBA.UCase(identifier)
    
    If identifier <> "" And LoginFailure = False Then
        If CompanySuccessDic.Exists(identifier) = False Then
            api_identifier = IntrinioBanks(identifier, "identifier")
            If api_identifier = identifier Then
                CompanySuccessDic.Add identifier, False
                coFailure = CompanySuccessDic(identifier)
            Else
                CompanySuccessDic.Add identifier, True
                coFailure = CompanySuccessDic(identifier)
            End If
        Else
            If APICallsAtLimit = False Then
                coFailure = CompanySuccessDic(identifier)
            Else
                coFailure = False
            End If
        End If
    End If
    
    If identifier <> "" And statement <> "" And LoginFailure = False And APICallsAtLimit = False And coFailure = False Then
        Key = identifier & "_" & statement
        
        If BankTagsDic.Exists(Key) = False Then
            Dim IntrinioClient As New WebClient
            IntrinioClient.BaseUrl = BaseUrl
            If iCredentials.Exists("username") = False Or iCredentials.Exists("password") = False Or iCredentials("username") = Empty Or iCredentials("password") = Empty Then
                Call IntrinioInitialize
            End If
            
            Dim inUsername As String
            Dim inPassword As String
            inUsername = iCredentials("username")
            inPassword = iCredentials("password")
            Dim Auth As New HttpBasicAuthenticator
            Auth.Setup _
                Username:=inUsername, _
                Password:=inPassword
            Set IntrinioClient.Authenticator = Auth
            
            Dim Request As New WebRequest
            Request.Resource = "tags/banks"
            Request.Method = HttpGet
            Request.Format = Json
            Request.AddQuerystringParam "identifier", identifier
            Request.AddQuerystringParam "statement", statement
            
            Dim Response As WebResponse
            Set Response = IntrinioClient.Execute(Request)

            If Response.StatusCode = Ok Then
                BankTagsDic.Add Key, Response.Data("data")
                IntrinioBankTags = BankTagsDic(Key)(sequence + 1)(Item)
            ElseIf Response.StatusCode = 429 Then
                APICallsAtLimit = True
                IntrinioBankTags = "Plan Limit Reached"
            ElseIf Response.StatusCode = 403 Then
                IntrinioBankTags = "Visit Intrinio.com to Subscribe"
            Else
                IntrinioBankTags = ""
            End If
        ElseIf BankTagsDic.Exists(Key) = True Then
            IntrinioBankTags = BankTagsDic(Key)(sequence + 1)(Item)
        End If
    Else
        If APICallsAtLimit = True Then
            Key = identifier & "_" & statement
            
            If BankTagsDic.Exists(Key) = True Then
            IntrinioBankTags = BankTagsDic(Key)(sequence + 1)(Item)
            Else
                IntrinioBankTags = "Plan Limit Reached"
            End If
            
        ElseIf LoginFailure = True Then
            IntrinioBankTags = "Invalid API Keys"
        ElseIf coFailure = True Then
            IntrinioBankTags = "Invalid identifier Symbol"
        Else
            IntrinioBankTags = ""
        End If
    End If
ExitHere:
    Exit Function
ErrorHandler:
    IntrinioBankTags = ""
End Function


Sub DescribeIntrinioBankFinancials()
    Dim FuncName As String
    Dim FuncDesc As String
    Dim Category As String
    Dim ArgDesc(1 To 6) As String
    
    FuncName = "IntrinioBankFinancials"
    FuncDesc = "Returns historical financial statement data point for a bank, based on the tag, fiscal year and fiscal period."
    Category = "Intrinio"
    ArgDesc(1) = "The company's identifier (i.e. ticker symbol 'JPM' or RSSD ID '749635')"
    ArgDesc(2) = "The financial statement selected"
    ArgDesc(3) = "The selected fiscal year for the chosen statement (i.e. 2014, 2013, 2012, etc.)"
    ArgDesc(4) = "The selected fiscal period for the chosen statement ('FY', 'Q1', 'Q2', 'Q3', 'Q2YTD', 'Q3YTD')"
    ArgDesc(5) = "The selected tag contained within the statement"
    ArgDesc(6) = "(Optional) Round the value (blank or 'A' for actuals, 'K' for thousands, 'M' for millions, 'B' for billions)"
    
    Application.MacroOptions Macro:=FuncName, _
       Description:=FuncDesc, _
       Category:=Category, _
       ArgumentDescriptions:=ArgDesc
End Sub


Public Function IntrinioBankFinancials(identifier As String, _
                           statement As String, _
                           fiscal_year As Integer, _
                           fiscal_period As String, _
                           tag As String, _
                           Optional rounding As String = "A")
Attribute IntrinioBankFinancials.VB_Description = "Returns historical financial statement data point for a bank, based on the tag, fiscal year and fiscal period."
Attribute IntrinioBankFinancials.VB_ProcData.VB_Invoke_Func = " \n19"
                           
    Dim Key As String
    Dim eKey As String
    Dim nKey As String
    Dim X As Variant
    Dim rTag As String
    Dim rValue As Double
    Dim sValue As String
    Dim Value As Double
    Dim Rounder As Double
    Dim api_identifier As String
    Dim coFailure As Boolean
    Dim fundamental_sequence As Integer
    Dim fundamental_type As String
    
    On Error GoTo ErrorHandler
    
    identifier = VBA.UCase(identifier)
    
    If identifier <> "" And LoginFailure = False Then
        If CompanySuccessDic.Exists(identifier) = False Then
            api_identifier = IntrinioBanks(identifier, "identifier")
            If api_identifier = identifier Then
                CompanySuccessDic.Add identifier, False
                coFailure = CompanySuccessDic(identifier)
            Else
                api_identifier = IntrinioSecurities(identifier, "identifier")
                If api_identifier = identifier Then
                    CompanySuccessDic.Add identifier, False
                    coFailure = CompanySuccessDic(identifier)
                Else
                    CompanySuccessDic.Add identifier, True
                    coFailure = CompanySuccessDic(identifier)
                End If
            End If
        Else
            If APICallsAtLimit = False Then
                coFailure = CompanySuccessDic(identifier)
            Else
                coFailure = False
            End If
        End If
    End If
    
    
    If identifier <> "" And statement <> "" And LoginFailure = False And APICallsAtLimit = False And coFailure = False Then
        If fiscal_year < 1900 Then
            fundamental_type = fiscal_period
            fundamental_sequence = fiscal_year
            fiscal_year = IntrinioBankFundamentals(identifier, statement, fundamental_type, fundamental_sequence, "fiscal_year")
            fiscal_period = IntrinioBankFundamentals(identifier, statement, fundamental_type, fundamental_sequence, "fiscal_period")
        End If
    End If
    
    If identifier <> "" And statement <> "" And fiscal_year <> 0 And fiscal_period <> "" And LoginFailure = False And APICallsAtLimit = False And coFailure = False Then
        

        Key = identifier & "_" & statement & "_" & fiscal_year & "_" & fiscal_period
        
        If BankFinancialsDic.Exists(Key) = False Then
            Dim IntrinioClient As New WebClient
            IntrinioClient.BaseUrl = BaseUrl
            If iCredentials.Exists("username") = False Or iCredentials.Exists("password") = False Or iCredentials("username") = Empty Or iCredentials("password") = Empty Then
                Call IntrinioInitialize
            End If
            
            Dim inUsername As String
            Dim inPassword As String
            inUsername = iCredentials("username")
            inPassword = iCredentials("password")
            Dim Auth As New HttpBasicAuthenticator
            Auth.Setup _
                Username:=inUsername, _
                Password:=inPassword
            Set IntrinioClient.Authenticator = Auth
            
            Dim Request As New WebRequest
            Dim last_page As Integer
            Dim is_last_page As Boolean
            Dim page As Integer
            Dim Response As WebResponse
            
            page = 1
            
            Do Until is_last_page = True
                Request.Resource = "financials/banks"
                Request.Method = HttpGet
                Request.Format = Json
                Request.AddQuerystringParam "identifier", identifier
                Request.AddQuerystringParam "statement", statement
                Request.AddQuerystringParam "fiscal_year", fiscal_year
                Request.AddQuerystringParam "fiscal_period", fiscal_period
                
                Set Response = IntrinioClient.Execute(Request)

                If Response.StatusCode = Ok Then
                    If Response.Content <> "" Then
                        last_page = Response.Data("total_pages")
                        If last_page > 0 Then
                            If last_page = page Then
                                is_last_page = True
                            Else
                                is_last_page = False
                                page = page + 1
                            End If
                            
                            If BankFinancialsDic.Exists(Key) = False Then
                                BankFinancialsDic.Add Key, Response.Data("data")
                            ElseIf BankFinancialsDic.Exists(Key) = True Then
                                BankFinancialsDic.Remove (Key)
                                BankFinancialsDic.Add Key, Response.Data("data")
                            End If
                            
                            For Each X In BankFinancialsDic(Key)
                                rTag = X("tag")
                                sValue = X("value")
                                nKey = identifier & "_" & statement & "_" & fiscal_year & "_" & fiscal_period & "_" & rTag
                                If BankFinancialsDic.Exists(nKey) = True Then
                                    BankFinancialsDic.Remove (nKey)
                                End If
                                BankFinancialsDic.Add nKey, sValue
                            Next
                        Else
                            is_last_page = True
                            If BankFinancialsDic.Exists(Key) = False Then
                                BankFinancialsDic.Add Key, Response.Data("data")
                            ElseIf BankFinancialsDic.Exists(Key) = True Then
                                BankFinancialsDic.Remove (Key)
                                BankFinancialsDic.Add Key, Response.Data("data")
                            End If
                        End If
                    Else
                        is_last_page = True
                        If BankFinancialsDic.Exists(Key) = False Then
                            BankFinancialsDic.Add Key, Empty
                        ElseIf BankFinancialsDic.Exists(Key) = True Then
                            BankFinancialsDic.Remove (Key)
                            BankFinancialsDic.Add Key, Empty
                        End If
                    End If
                Else
                    is_last_page = True
                    If Response.StatusCode = 429 Then
                        APICallsAtLimit = True
                        IntrinioBankFinancials = "Plan Limit Reached"
                    ElseIf Response.StatusCode = 403 Then
                        IntrinioBankFinancials = "Visit Intrinio.com to Subscribe"
                    Else
                        IntrinioBankFinancials = ""
                    End If
                End If
            Loop
                            
            eKey = identifier & "_" & statement & "_" & fiscal_year & "_" & fiscal_period & "_" & tag

            If BankFinancialsDic.Exists(Key) = True Then

                If BankFinancialsDic(Key) Is Not Empty Then
                    
                    If IsNumeric(BankFinancialsDic(eKey)) = True Then
                        Value = BankFinancialsDic(eKey)
                        If rounding = "K" Then
                            Rounder = 1000
                        ElseIf rounding = "M" Then
                            Rounder = 1000000
                        ElseIf rounding = "B" Then
                            Rounder = 1000000000
                        Else
                            Rounder = 1
                        End If
                    
                        IntrinioBankFinancials = Value / Rounder
                    Else
                        IntrinioBankFinancials = BankFinancialsDic(eKey)
                    End If
                Else
                    IntrinioBankFinancials = ""
                End If
            Else
                IntrinioBankFinancials = ""
            End If
        ElseIf BankFinancialsDic.Exists(Key) = True Then
            eKey = identifier & "_" & statement & "_" & fiscal_year & "_" & fiscal_period & "_" & tag
            
            If IsNumeric(BankFinancialsDic(eKey)) = True Then
                Value = BankFinancialsDic(eKey)
                
                If rounding = "K" Then
                    Rounder = 1000
                ElseIf rounding = "M" Then
                    Rounder = 1000000
                ElseIf rounding = "B" Then
                    Rounder = 1000000000
                Else
                    Rounder = 1
                End If
                
                IntrinioBankFinancials = Value / Rounder
            Else
                IntrinioBankFinancials = BankFinancialsDic(eKey)
            End If
        End If
    Else
        If APICallsAtLimit = True Then
            IntrinioBankFinancials = "Plan Limit Reached"
        ElseIf LoginFailure = True Then
            IntrinioBankFinancials = "Invalid API Keys"
        ElseIf coFailure = True Then
            IntrinioBankFinancials = "Invalid Identifier"
        Else
            IntrinioBankFinancials = ""
        End If
    End If
ExitHere:
    Exit Function
ErrorHandler:
    IntrinioBankFinancials = ""
    Resume Next
End Function

Private Function IntrinioAddinVersion(Item As String)
    Dim IntrinioClient As New WebClient
    IntrinioClient.BaseUrl = BaseUrl
    
    If iCredentials.Exists("username") = False Or iCredentials.Exists("password") = False Or iCredentials("username") = Empty Or iCredentials("password") = Empty Then
        Call IntrinioInitialize
    End If
    
    Dim inUsername As String
    Dim inPassword As String
    inUsername = iCredentials("username")
    inPassword = iCredentials("password")
    Dim Auth As New HttpBasicAuthenticator
    Auth.Setup _
        Username:=inUsername, _
        Password:=inPassword
    Set IntrinioClient.Authenticator = Auth
    
    Dim Request As New WebRequest
    Request.Resource = "excel"
    Request.Method = HttpGet
    Request.Format = Json
    
    Dim Response As WebResponse
    Set Response = IntrinioClient.Execute(Request)

    If Response.StatusCode = Ok Then
        APICallsAtLimit = False
        
        If iVersion.Exists("Version") = True Then
            iVersion.RemoveAll
        End If
        iVersion.Add "Version", Response.Data("version")
        iVersion.Add "Change_Log", Response.Data("change_log")
        iVersion.Add "Download_URL", Response.Data("download_url")
        iVersion.Add "Windows_Download_URL", Response.Data("windows_download_url")
        iVersion.Add "Mac_Download_URL", Response.Data("mac_download_url")
        If Item = "status_code" Then
            IntrinioAddinVersion = Response.StatusCode
        Else
            IntrinioAddinVersion = Response.Data(Item)
        End If
        
    ElseIf Response.StatusCode = 429 Then
        APICallsAtLimit = True
        If iVersion.Exists("Version") = True Then
            If Item = "status_code" Then
                IntrinioAddinVersion = Response.StatusCode
            Else
                IntrinioAddinVersion = iVersion(Item)
            End If
        Else
            IntrinioAddinVersion = Response.StatusCode
        End If
    ElseIf Response.StatusCode = 403 Then
        IntrinioAddinVersion = Response.StatusCode
    Else
        IntrinioAddinVersion = ""
    End If
End Function

Sub IntrinioRibbon()

    Dim hFile As Long
    Dim path As String, fileName As String, ribbonXML As String, appdata As String
    Dim FilePath As String, CurrentXML As String
    Dim ribbonXMLHeader As String, ribbonXMLContent As String, ribbonXMLNew As String, ribbonXMLFooter As String
    Dim XPos1 As Long, XPos2 As Long, XPos3 As Long, XPos4 As Long, XPos5 As Long, X As Long
    Dim refresh As Boolean
  
    On Error GoTo ErrorHandler
    
    #If Win32 Or Win64 Then
        refresh = False
        
        hFile = FreeFile
        appdata = Environ("LOCALAPPDATA")
        path = appdata & "\Microsoft\Office\"
        fileName = "Excel.officeUI"
        
        FilePath = path & fileName
    
        If FileOrDirExists(FilePath) Then
            X = FreeFile
            Open FilePath For Input As #X
            Input #X, CurrentXML
            Close #X
            
            XPos1 = InStr(1, CurrentXML, "<mso:tabs>")
            If XPos1 > 0 Then
                XPos2 = InStr(XPos1, CurrentXML, "intrinioTab")
                XPos3 = XPos2 - 14
                XPos4 = Len(CurrentXML)
                XPos5 = InStr(XPos3, CurrentXML, "</mso:tab>")
                ribbonXMLHeader = Left$(CurrentXML, XPos3)
                ribbonXMLFooter = Right$(CurrentXML, XPos4 - XPos5 - 9)
                
                ribbonXML = ribbonXMLHeader + "<mso:tab id=" & Chr(34) & "intrinioTab" & Chr(34) & " label=" & Chr(34) & "Intrinio" & Chr(34) & " insertBeforeQ=" & Chr(34) & "mso:TabFormat" & Chr(34) & ">"
                ribbonXML = ribbonXML + "<mso:group id=" & Chr(34) & "intrinioGroup" & Chr(34) & " label=" & Chr(34) & "Intrinio Add-in" & Chr(34) & " autoScale=" & Chr(34) & "true" & Chr(34) & ">"
                ribbonXML = ribbonXML + "<mso:button id=" & Chr(34) & "apiKeys" & Chr(34) & " label=" & Chr(34) & "API Keys" & Chr(34) & " "
                ribbonXML = ribbonXML + "imageMso=" & Chr(34) & "DatabaseMakeMdeFile" & Chr(34) & " onAction=" & Chr(34) & "IntrinioAPIKeys" & Chr(34) & "/>"
                ribbonXML = ribbonXML + "<mso:button id=" & Chr(34) & "refreshData" & Chr(34) & " label=" & Chr(34) & "Refresh Data" & Chr(34) & " "
                ribbonXML = ribbonXML + "imageMso=" & Chr(34) & "RecurrenceEdit" & Chr(34) & " onAction=" & Chr(34) & "IntrinioRefresh" & Chr(34) & "/>"
                ribbonXML = ribbonXML + "<mso:button id=" & Chr(34) & "openTemplate" & Chr(34) & " label=" & Chr(34) & "Intrinio Templates" & Chr(34) & " "
                ribbonXML = ribbonXML + "imageMso=" & Chr(34) & "FilePublishExcelServices" & Chr(34) & " onAction=" & Chr(34) & "IntrinioTemplates" & Chr(34) & "/>"
                ribbonXML = ribbonXML + "<mso:button id=" & Chr(34) & "unlinkAddin" & Chr(34) & " label=" & Chr(34) & "Unlink Add-in" & Chr(34) & " "
                ribbonXML = ribbonXML + "imageMso=" & Chr(34) & "DatabaseObjectDependencies" & Chr(34) & " onAction=" & Chr(34) & "IntrinioUnlink" & Chr(34) & "/>"
                ribbonXML = ribbonXML + "<mso:button id=" & Chr(34) & "checkUpdate" & Chr(34) & " label=" & Chr(34) & "Check for Update" & Chr(34) & " "
                ribbonXML = ribbonXML + "imageMso=" & Chr(34) & "PageMenu" & Chr(34) & " onAction=" & Chr(34) & "IntrinioUpdate" & Chr(34) & "/>"
                ribbonXML = ribbonXML + "<mso:button id=" & Chr(34) & "helpMe" & Chr(34) & " label=" & Chr(34) & "Help" & Chr(34) & " "
                ribbonXML = ribbonXML + "imageMso=" & Chr(34) & "FunctionsLogicalInsertGallery" & Chr(34) & " onAction=" & Chr(34) & "IntrinioHelp" & Chr(34) & "/>"
                ribbonXML = ribbonXML + "</mso:group>"
                ribbonXML = ribbonXML + "</mso:tab>"
                ribbonXML = ribbonXML + ribbonXMLFooter
                
                Open path & fileName For Output Access Write As hFile
                Print #hFile, ribbonXML
                Close hFile
            Else
                Close hFile
                refresh = True
            End If
        Else
            refresh = True
            
        End If
        
        If refresh = True Then
            ribbonXML = "<mso:customUI xmlns:mso=" & Chr(34) & "http://schemas.microsoft.com/office/2009/07/customui" & Chr(34) & ">"
            ribbonXML = ribbonXML + "<mso:ribbon>"
            ribbonXML = ribbonXML + "<mso:qat></mso:qat>"
            ribbonXML = ribbonXML + "<mso:tabs>"
            ribbonXML = ribbonXML + "<mso:tab id=" & Chr(34) & "intrinioTab" & Chr(34) & " label=" & Chr(34) & "Intrinio" & Chr(34) & " insertBeforeQ=" & Chr(34) & "mso:TabFormat" & Chr(34) & ">"
            ribbonXML = ribbonXML + "<mso:group id=" & Chr(34) & "intrinioGroup" & Chr(34) & " label=" & Chr(34) & "Intrinio Add-in" & Chr(34) & " autoScale=" & Chr(34) & "true" & Chr(34) & ">"
            ribbonXML = ribbonXML + "<mso:button id=" & Chr(34) & "apiKeys" & Chr(34) & " label=" & Chr(34) & "API Keys" & Chr(34) & " "
            ribbonXML = ribbonXML + "imageMso=" & Chr(34) & "DatabaseMakeMdeFile" & Chr(34) & " onAction=" & Chr(34) & "IntrinioAPIKeys" & Chr(34) & "/>"
            ribbonXML = ribbonXML + "<mso:button id=" & Chr(34) & "refreshData" & Chr(34) & " label=" & Chr(34) & "Refresh Data" & Chr(34) & " "
            ribbonXML = ribbonXML + "imageMso=" & Chr(34) & "RecurrenceEdit" & Chr(34) & " onAction=" & Chr(34) & "IntrinioRefresh" & Chr(34) & "/>"
            ribbonXML = ribbonXML + "<mso:button id=" & Chr(34) & "openTemplate" & Chr(34) & " label=" & Chr(34) & "Intrinio Templates" & Chr(34) & " "
            ribbonXML = ribbonXML + "imageMso=" & Chr(34) & "FilePublishExcelServices" & Chr(34) & " onAction=" & Chr(34) & "IntrinioTemplates" & Chr(34) & "/>"
            ribbonXML = ribbonXML + "<mso:button id=" & Chr(34) & "unlinkAddin" & Chr(34) & " label=" & Chr(34) & "Unlink Add-in" & Chr(34) & " "
            ribbonXML = ribbonXML + "imageMso=" & Chr(34) & "DatabaseObjectDependencies" & Chr(34) & " onAction=" & Chr(34) & "IntrinioUnlink" & Chr(34) & "/>"
            ribbonXML = ribbonXML + "<mso:button id=" & Chr(34) & "checkUpdate" & Chr(34) & " label=" & Chr(34) & "Check for Update" & Chr(34) & " "
            ribbonXML = ribbonXML + "imageMso=" & Chr(34) & "PageMenu" & Chr(34) & " onAction=" & Chr(34) & "IntrinioUpdate" & Chr(34) & "/>"
            ribbonXML = ribbonXML + "<mso:button id=" & Chr(34) & "helpMe" & Chr(34) & " label=" & Chr(34) & "Help" & Chr(34) & " "
            ribbonXML = ribbonXML + "imageMso=" & Chr(34) & "FunctionsLogicalInsertGallery" & Chr(34) & " onAction=" & Chr(34) & "IntrinioHelp" & Chr(34) & "/>"
            ribbonXML = ribbonXML + "</mso:group>"
            ribbonXML = ribbonXML + "</mso:tab>"
            ribbonXML = ribbonXML + "</mso:tabs>"
            ribbonXML = ribbonXML + "</mso:ribbon>"
            ribbonXML = ribbonXML + "</mso:customUI>"
            
            Open path & fileName For Output Access Write As hFile
            Print #hFile, ribbonXML
            Close hFile
            
        End If
    #End If
ExitHere:
    Exit Sub
ErrorHandler:
    #If Win32 Or Win64 Then
        hFile = FreeFile
        appdata = Environ("LOCALAPPDATA")
        path = appdata & "\Microsoft\Office\"
        fileName = "Excel.officeUI"
        
        FilePath = path & fileName
    
        If FileOrDirExists(FilePath) Then
            SetAttr FilePath, vbNormal
            Kill FilePath
            
            ribbonXML = "<mso:customUI xmlns:mso=" & Chr(34) & "http://schemas.microsoft.com/office/2009/07/customui" & Chr(34) & ">"
            ribbonXML = ribbonXML + "<mso:ribbon>"
            ribbonXML = ribbonXML + "<mso:qat></mso:qat>"
            ribbonXML = ribbonXML + "<mso:tabs>"
            ribbonXML = ribbonXML + "<mso:tab id=" & Chr(34) & "intrinioTab" & Chr(34) & " label=" & Chr(34) & "Intrinio" & Chr(34) & " insertBeforeQ=" & Chr(34) & "mso:TabFormat" & Chr(34) & ">"
            ribbonXML = ribbonXML + "<mso:group id=" & Chr(34) & "intrinioGroup" & Chr(34) & " label=" & Chr(34) & "Intrinio Add-in" & Chr(34) & " autoScale=" & Chr(34) & "true" & Chr(34) & ">"
            ribbonXML = ribbonXML + "<mso:button id=" & Chr(34) & "apiKeys" & Chr(34) & " label=" & Chr(34) & "API Keys" & Chr(34) & " "
            ribbonXML = ribbonXML + "imageMso=" & Chr(34) & "DatabaseMakeMdeFile" & Chr(34) & " onAction=" & Chr(34) & "IntrinioAPIKeys" & Chr(34) & "/>"
            ribbonXML = ribbonXML + "<mso:button id=" & Chr(34) & "refreshData" & Chr(34) & " label=" & Chr(34) & "Refresh Data" & Chr(34) & " "
            ribbonXML = ribbonXML + "imageMso=" & Chr(34) & "RecurrenceEdit" & Chr(34) & " onAction=" & Chr(34) & "IntrinioRefresh" & Chr(34) & "/>"
            ribbonXML = ribbonXML + "<mso:button id=" & Chr(34) & "openTemplate" & Chr(34) & " label=" & Chr(34) & "Intrinio Templates" & Chr(34) & " "
            ribbonXML = ribbonXML + "imageMso=" & Chr(34) & "FilePublishExcelServices" & Chr(34) & " onAction=" & Chr(34) & "IntrinioTemplates" & Chr(34) & "/>"
            ribbonXML = ribbonXML + "<mso:button id=" & Chr(34) & "unlinkAddin" & Chr(34) & " label=" & Chr(34) & "Unlink Add-in" & Chr(34) & " "
            ribbonXML = ribbonXML + "imageMso=" & Chr(34) & "DatabaseObjectDependencies" & Chr(34) & " onAction=" & Chr(34) & "IntrinioUnlink" & Chr(34) & "/>"
            ribbonXML = ribbonXML + "<mso:button id=" & Chr(34) & "checkUpdate" & Chr(34) & " label=" & Chr(34) & "Check for Update" & Chr(34) & " "
            ribbonXML = ribbonXML + "imageMso=" & Chr(34) & "PageMenu" & Chr(34) & " onAction=" & Chr(34) & "IntrinioUpdate" & Chr(34) & "/>"
            ribbonXML = ribbonXML + "<mso:button id=" & Chr(34) & "helpMe" & Chr(34) & " label=" & Chr(34) & "Help" & Chr(34) & " "
            ribbonXML = ribbonXML + "imageMso=" & Chr(34) & "FunctionsLogicalInsertGallery" & Chr(34) & " onAction=" & Chr(34) & "IntrinioHelp" & Chr(34) & "/>"
            ribbonXML = ribbonXML + "</mso:group>"
            ribbonXML = ribbonXML + "</mso:tab>"
            ribbonXML = ribbonXML + "</mso:tabs>"
            ribbonXML = ribbonXML + "</mso:ribbon>"
            ribbonXML = ribbonXML + "</mso:customUI>"
            
            Open path & fileName For Output Access Write As hFile
            Print #hFile, ribbonXML
            Close hFile
        Else
            ribbonXML = "<mso:customUI xmlns:mso=" & Chr(34) & "http://schemas.microsoft.com/office/2009/07/customui" & Chr(34) & ">"
            ribbonXML = ribbonXML + "<mso:ribbon>"
            ribbonXML = ribbonXML + "<mso:qat></mso:qat>"
            ribbonXML = ribbonXML + "<mso:tabs>"
            ribbonXML = ribbonXML + "<mso:tab id=" & Chr(34) & "intrinioTab" & Chr(34) & " label=" & Chr(34) & "Intrinio" & Chr(34) & " insertBeforeQ=" & Chr(34) & "mso:TabFormat" & Chr(34) & ">"
            ribbonXML = ribbonXML + "<mso:group id=" & Chr(34) & "intrinioGroup" & Chr(34) & " label=" & Chr(34) & "Intrinio Add-in" & Chr(34) & " autoScale=" & Chr(34) & "true" & Chr(34) & ">"
            ribbonXML = ribbonXML + "<mso:button id=" & Chr(34) & "apiKeys" & Chr(34) & " label=" & Chr(34) & "API Keys" & Chr(34) & " "
            ribbonXML = ribbonXML + "imageMso=" & Chr(34) & "DatabaseMakeMdeFile" & Chr(34) & " onAction=" & Chr(34) & "IntrinioAPIKeys" & Chr(34) & "/>"
            ribbonXML = ribbonXML + "<mso:button id=" & Chr(34) & "refreshData" & Chr(34) & " label=" & Chr(34) & "Refresh Data" & Chr(34) & " "
            ribbonXML = ribbonXML + "imageMso=" & Chr(34) & "RecurrenceEdit" & Chr(34) & " onAction=" & Chr(34) & "IntrinioRefresh" & Chr(34) & "/>"
            ribbonXML = ribbonXML + "<mso:button id=" & Chr(34) & "openTemplate" & Chr(34) & " label=" & Chr(34) & "Intrinio Templates" & Chr(34) & " "
            ribbonXML = ribbonXML + "imageMso=" & Chr(34) & "FilePublishExcelServices" & Chr(34) & " onAction=" & Chr(34) & "IntrinioTemplates" & Chr(34) & "/>"
            ribbonXML = ribbonXML + "<mso:button id=" & Chr(34) & "unlinkAddin" & Chr(34) & " label=" & Chr(34) & "Unlink Add-in" & Chr(34) & " "
            ribbonXML = ribbonXML + "imageMso=" & Chr(34) & "DatabaseObjectDependencies" & Chr(34) & " onAction=" & Chr(34) & "IntrinioUnlink" & Chr(34) & "/>"
            ribbonXML = ribbonXML + "<mso:button id=" & Chr(34) & "checkUpdate" & Chr(34) & " label=" & Chr(34) & "Check for Update" & Chr(34) & " "
            ribbonXML = ribbonXML + "imageMso=" & Chr(34) & "PageMenu" & Chr(34) & " onAction=" & Chr(34) & "IntrinioUpdate" & Chr(34) & "/>"
            ribbonXML = ribbonXML + "<mso:button id=" & Chr(34) & "helpMe" & Chr(34) & " label=" & Chr(34) & "Help" & Chr(34) & " "
            ribbonXML = ribbonXML + "imageMso=" & Chr(34) & "FunctionsLogicalInsertGallery" & Chr(34) & " onAction=" & Chr(34) & "IntrinioHelp" & Chr(34) & "/>"
            ribbonXML = ribbonXML + "</mso:group>"
            ribbonXML = ribbonXML + "</mso:tab>"
            ribbonXML = ribbonXML + "</mso:tabs>"
            ribbonXML = ribbonXML + "</mso:ribbon>"
            ribbonXML = ribbonXML + "</mso:customUI>"
            
            Open path & fileName For Output Access Write As hFile
            Print #hFile, ribbonXML
            Close hFile
        End If
    #End If
    Exit Sub
End Sub

Sub IntrinioResetRibbon()
    Dim hFile As Long
    Dim path As String, fileName As String, appdata As String, FilePath As String
    #If Win32 Or Win64 Then
        hFile = FreeFile
        appdata = Environ("LOCALAPPDATA")
        path = appdata & "\Microsoft\Office\"
        fileName = "Excel.officeUI"
        
        FilePath = path & fileName
    
        If FileOrDirExists(FilePath) Then
            SetAttr FilePath, vbNormal
            Kill FilePath
            
            Call IntrinioRibbon
        Else
            Call IntrinioRibbon
        End If
    #End If
End Sub

Public Sub IntrinioRefresh()
    Dim status As String
    
    LoginFailure = False
    APICallsAtLimit = False
    status = IntrinioAddinVersion("status_code")
    
    If status = "200" Then
        Call IntrinioFixLinks
    
        CompanySuccessDic.RemoveAll
        DataPointRequestedTags.RemoveAll
        
        CompanyDic.RemoveAll
        SecuritiesDic.RemoveAll
        BankDic.RemoveAll
        DataPointDic.RemoveAll
        HistoricalPricesDic.RemoveAll
        HistoricalDataDic.RemoveAll
        NewsDic.RemoveAll
        FundamentalsDic.RemoveAll
        StandardizedFinancialsDic.RemoveAll
        StandardizedTagsDic.RemoveAll
        ReportedFundamentalsDic.RemoveAll
        ReportedFinancialsDic.RemoveAll
        ReportedTagsDic.RemoveAll
        BankFundamentalsDic.RemoveAll
        BankFinancialsDic.RemoveAll
        BankTagsDic.RemoveAll
        
        APICallsAtLimit = False
        LoginFailure = False
        
        Application.CalculateFull
    ElseIf status = "429" Then
        Application.CalculateFull
    ElseIf Response.StatusCode = 403 Then
        Application.CalculateFull
    End If
End Sub

Public Sub IntrinioHelp()
    Dim Url As String
    Url = "http://community.intrinio.com/docs/excel-add-in/"
    ActiveWorkbook.FollowHyperlink Url
End Sub

Public Sub IntrinioUpdate()
    Dim version As String
    Dim web_url As String
    Dim IntrinioRespCode As String
    Dim answer As Integer
    
    UpdatePrompt = True
    
    IntrinioRespCode = IntrinioAddinVersion("status_code")
    
    If IntrinioRespCode = "" Then
        IntrinioRespCode = 401
    End If
    
    If IntrinioRespCode = 200 Then
        version = IntrinioAddinVersion("version")
        If Intrinio_Addin_Version = version Then
            answer = MsgBox("There are no updates available! " & _
                    "Version " & Intrinio_Addin_Version & " is the most recent release of the Intrinio Excel Add-in", vbOKOnly, "Update Intrinio Excel Add-in")
        Else
            answer = MsgBox("Version " & version & " of the Intrinio Excel Addin is available for download!" & vbNewLine & "Current version: " & Intrinio_Addin_Version _
                & vbNewLine & "Would you like to install it now?", vbYesNo, "Update Intrinio Excel Add-in")

            If answer = vbYes Then
                #If Mac Then
                    web_url = IntrinioAddinVersion("mac_download_url")
                #ElseIf Win32 Or Win64 Then
                    web_url = IntrinioAddinVersion("windows_download_url")
                #Else
                    web_url = IntrinioAddinVersion("download_url")
                #End If
                ActiveWorkbook.FollowHyperlink Address:=web_url
            End If
        End If
    ElseIf IntrinioRespCode = 429 Then
        answer = MsgBox("Unable to check for update at this time.", vbOKOnly, "Update Intrinio Excel Add-in")
    Else
        answer = MsgBox("Unable to connect to the Intrinio API.", vbOKOnly, "Update Intrinio Excel Add-in")
    End If
End Sub

Public Sub IntrinioTemplates()
    
    Dim fileName
    Dim ActSheet As Worksheet
    Dim ActBook As Workbook
    Dim CurrentFile As String
    Dim NewFileType As String
    Dim NewFile As String
    Dim selectedTemplateName As String, userprofile As String, userprofilepath As String

    #If Win32 Or Win64 Then
        ChDir Intrinio_Excel_Addin_Path & "\Templates"
        
        fileName = Application.GetOpenFilename("Intrinio Templates (*.xlsm),*.xlsm")
        
        If fileName <> "" And fileName <> "False" Then
            Application.AskToUpdateLinks = False
            Application.DisplayAlerts = False
     
            selectedTemplateName = Dir(fileName)
    
            NewFileType = "Excel Macro-Enabled Workbook (*.xlsm), *.xlsm,"
            
            userprofile = Environ("USERPROFILE") & "\Documents\Intrinio"
            userprofilepath = FileOrDirExists(userprofile)
            
            If userprofilepath = "False" Then
                userprofile = Environ("USERPROFILE") & "\My Documents\Intrinio"
                userprofilepath = FileOrDirExists(userprofile)
                If userprofilepath = "False" Then
                    userprofile = Environ("USERPROFILE")
                End If
            End If

            ChDir userprofile
            
            NewFile = Application.GetSaveAsFilename( _
                InitialFileName:=selectedTemplateName, _
                fileFilter:=NewFileType)
     
            If NewFile <> "" And NewFile <> "False" Then
                Application.Workbooks.Open fileName:=fileName
            
                ActiveWorkbook.SaveAs fileName:=NewFile, _
                    FileFormat:=xlOpenXMLWorkbookMacroEnabled, _
                    Password:="", _
                    WriteResPassword:="", _
                    ReadOnlyRecommended:=False, _
                    CreateBackup:=False
            
                
            End If
            Application.AskToUpdateLinks = True
            Application.DisplayAlerts = True
        End If
    #End If
    
End Sub

'******************************************************************************
'Helper Macros
'******************************************************************************
Private Function Intrinio_Excel_Addin_Path()
    
   Dim oAddIn As AddIn
   Dim zMsg   As String
   Dim add_in_path As String

   For Each oAddIn In AddIns
      If oAddIn.Name = "Intrinio_Excel_Addin.xlam" Then
        add_in_path = oAddIn.path
      End If
      
   Next oAddIn
   Intrinio_Excel_Addin_Path = add_in_path
   
End Function

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

Private Function ValidAddress(strAddress As String) As Boolean
    Dim r As Range
    On Error Resume Next
    Set r = Range(strAddress)
    If Not r Is Nothing Then ValidAddress = True
End Function

