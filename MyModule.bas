Attribute VB_Name = "MyModule"
Declare Function GetKeyboardLayout Lib "user32" (ByVal dwLayout As Long) As Long
Declare Function ActivateKeyboardLayout Lib "user32" (ByVal HKL As Long, ByVal flags As Long) As Long
Global ConnectString As String
Global empNo As Integer
Global ClientNo As Long
Global systemConfigration  As New Configration

Global OperationEmpStr As String
Global MaintTYpeEmpStr As String
Global PaymentEmpStr As String


Global LoadForm As Boolean
Global ServerName As String
Global UID As String
Global PWD As String
Global DataBase As String
Global DatabaseYear As Integer
Global Const EnglishLayout = 67699721
Global IDBill As Double
Global idCallNo As Double
Global ProgramTitle As String

Type ExternalReparationByTeamType
    FromDate As Variant
    TillDate As Variant
    TeamNo As Variant
    PaymenrTypeMethod As Variant
End Type


Type ExternalReparationType
    FromDate As Variant
    TillDate As Variant
    PaymenrTypeMethod As Variant
End Type

Type TeamInfo
    TeamNo As Integer
    TeamName As String
    LeaderEmpNo As Integer
    LeaderFullName As String
    AssistantEmpNo As Integer
    AssistantFullName As String
End Type

Type TreansferProvinceStockType
    ProvinceStock As ADODB.Recordset
    FactoryStock As ADODB.Recordset
End Type
'
'Type CustomerRecType
'        CustomerNo As Integer
'        CustomerName As String
'        CustomerHomePhone As String
'End Type

Type TTlType
    Count As Integer
    HallQtySum As Double
    FactoryQtySum As Double
    
End Type

Type statisticsType
      SumSelectedValues As Double
      CountSelectedValues As Double
      SumUnSelectedValues As Double
      CountunSelectedValues As Double
End Type

Type PordFAmilyType

    ProdFamNo As Integer
    ProdFamName As String
End Type

Type MntCallRecTYpe
    CompNo As Integer
    CallNo  As Double
    CallDateTime As String
    ModNo As Integer
    cliNo As Double
    CallDEscription As String
    CallStatus As Integer
    Notes As String
    Defindname As String
    CallReceiver As String
    CallReceiverEmpNo As Integer
    PaymentTYpeId As Integer
End Type

Type CustomerRecType
    CompNo As Integer
    adhamNo As Double
    AdhamName As String
    AdhamPhon As String
    Workphone As String
    MobilePhone As String
    AdhamAdress As String
    Zone As Integer
    Defindname As String
    Email As String
    Notes As String
End Type

Type SearchMaintCallRecType
    CustomerId As Double
    HomePhone As String
    Workphone As String
    MobilePhone As String
    ZoneNo As Integer
    Address As String
    FromDate As String
    TillDate As String
    FromTime As String
    tillTime As String
    Via  As String
    ProductFamilyNo As Integer
    repNo As Integer
    Note As String
End Type

Type MvMaintPaymentRecTYpe
    SerByyear As Integer
    BillNo As Double
    Billdate As String
    FixBillDate As String
    OperationType As Integer
    MaintType As Integer
    PaymentTYpeId As Integer
    DestinationId As Integer
    clientId As Double
    Class As Integer
    ModNo As Integer
    ModelQty  As Integer
    FeesDescription As String
    OtherFeesQty As Integer
    OtherFeesPrice As Double
'    FeesTYpeId As Integer
'    FeesQty As Integer
'    FeesPriceType As Integer
'    FeesAmount As Double
    IsFixed As Integer
    IsTransfered As Integer
    Roe As Double
End Type

Type MvMaintPaymentRecTypeDetails
    Id As Double
    BillNo As Double
    stkno As String
    discount As Double
    Qty As Double
    PriceTYpe As Integer
    Price As Double
    DestinationStoreId As Double
    Class As Integer
    OperationType As Integer
End Type

Type MvWshopType
    Ser As Integer
    Id As Integer
    stkno As String
    Qty As Double
    Date As String
End Type

Type CustomerInformationTypeRec
    CallNo As Integer
    AdhamName As String
    AdhamPhon As String
    CallDateTime As String
    RepDate As String
    ProdFamNameA As String
    CallDEscription As String
    Description As String
    Notes As String
    TeamName As String
    RepPrice As String
    CountRec As Integer
End Type

Type SearchRecType
    OperationId As Integer
    TYpeId As Integer
    PaymentId As Integer
    clientId As Double
    ClientType As Integer
    ModNo As Integer
    stkno  As String
    FeesId As Integer
    OperationFromBillNo As String
    OperationTillBillno As String
    FromBillNo As Double
    TillBillNo As Double
    OperationFromDate As String
    OperationTillDate As String
    FromDate As String
    TillDate As String
    
    
    OperationFixFromDate As String
    OperationFixTillDate As String
    FixFromDate As String
    FixTillDate As String
    
    
    OperationFromAmount As String
    OperationTillAmount As String
    FromAmount As Double
    TillAmount As Double
    OperationFromFees As String
    OperationTillFees As String
    FromFees As Double
    TillFees As Double
    OperationFromTotal As String
    OperationTillTotal As String
    FromTotal As Double
    TillTotal As Double
    
    
    Voption As Integer
    DestinationId As Integer
End Type


Type CustomerType
    CustomerId As Double
    customerName As String
    CustomerPhoneNBR As String
End Type

Type ModelListItem
    Id As Double
    ModNo As Integer
    stkno As String
End Type

Type StkRelatedItem
    Id As Double
    stkno As String
    StkRelatedNo As String
    Qty As Double
End Type


Type HallsType
   Id As Double
   Date As String
   Amount As Double
   HallId As Integer
   ClassId As Integer
   Notes As String
End Type

Type MvStockType ' Stmov
    ByanId As Double
    stkId As Double
    Strid As Integer
    MovDate As String
    Doctype As Integer
    DocNum As Double
    Qty As Double
    QtyType As Integer
    WshopId As Integer
    SecondaryTYpeId As Integer
    OrderNo As String
    Correspondence As String
    PackingList As String
    CountryNo As Integer
End Type

Type MovBetweenTowStoreType ' Transfer Data Between Tow Stores
    IdTarget As Double
    IdDestination As Double
    ByanId As Double
    MovDate As String
    stkId As Double
    StrTarget As Integer
    StrDestination As Integer
    Doctype As Integer
    DocNum As Double
    QtyTarget As Double
    QtyTypeTarget As Integer
    QtyDestination As Double
    QtyTypeDestination As Integer
End Type




Global CustomerInformationRec As CustomerInformationTypeRec
Global CustomerRec As CustomerType

Global clientId As Double
Global ClientName As String
Global ClientPhoneNBr As String

Global GByanId As Double
Global StrNo As String
Global ByanType As String
Global SelectedStr As String
Global FormOk As Boolean
Global customerName As String
Global customerNumber As Long
Global searchClientIsAllow As Boolean


Sub ColorRow(Row As Integer, Color As Long, FlexGrid As VSFlexGrid)
With FlexGrid
    For i = 1 To .Cols - 2
        .Col = i
        .Row = Row
        .CellBackColor = Color
    Next
End With
End Sub

Function Gettag(empNo As Integer, TagId As Integer) As Boolean
On Error GoTo ErrorHandler
Dim rs As New ADODB.Recordset
    sqlText = "Select * from comaintpermission Where empno = " & empNo & " and TagId=" & TagId
    Set rs = de.con.Execute(sqlText)
    If rs.RecordCount > 0 Then
        Gettag = True
    Else
        Gettag = False
    End If
Exit Function
ErrorHandler:
Gettag = False
End Function


Function GetFullName(empNo As Integer) As String
On Error GoTo ErrorHandler
Dim rsEmployee  As New ADODB.Recordset
sqlText = "Select FullName From empFullName Where Empno=" & empNo
Set rsEmployee = de.con.Execute(sqlText)
If rsEmployee.RecordCount > 0 Then
    GetFullName = rsEmployee!FullName
Else
    GetFullName = ""
End If

Exit Function
ErrorHandler:
GetFullName = ""
MsgBox Err.Description
End Function


Sub SetColWidths(ByVal ColNo As Integer, FlexGrid As VSFlexGrid)
    With FlexGrid
        .AutoSize (ColNo)
    End With
End Sub

Sub MoveCursor(KeyCode As Integer, FlexGrid As VSFlexGrid)
On Error Resume Next
If Not FlexGrid.Visible Then Exit Sub

With FlexGrid
    If KeyCode = vbKeyDown Then
        .Row = .Row + 1
    ElseIf KeyCode = vbKeyUp And .Row <> 1 Then
        .Row = .Row - 1
    End If
If Not .RowIsVisible(.Row) Then
    .TopRow = .Row
End If
.Col = 0
.ColSel = .Cols - 1
End With
End Sub

Sub ChangeToEnglish()
    If GetKeyboardLayout(0) <> EnglishLayout Then ActivateKeyboardLayout 0, 0
End Sub

Sub ChangeToArabic()
    If GetKeyboardLayout(0) = EnglishLayout Then ActivateKeyboardLayout 0, 0
End Sub

Function LikeExpression(Expr As String, Optional MultiLeter As String = "%", Optional OneLetter As String = "_") As String
    Dim X As String
    X = Replace(Expr, "Ç", "*Ç*")
    X = Replace(X, "Ã", "*Ã*")
    X = Replace(X, "Å", "*Å*")
    X = Replace(X, "Â", "*Â*")
    X = Replace(X, "ì", "*ì*")
    X = Replace(X, "í", "*í*")
    
    X = Replace(X, "*Ç*", "[ÇÃÅÂì]")
    X = Replace(X, "*Ã*", "[ÇÃÅÂ]")
    X = Replace(X, "*Å*", "[ÇÃÅÂ]")
    X = Replace(X, "*Â*", "[ÇÃÅÂ]")
    X = Replace(X, "*ì*", "[Çìí]")
    X = Replace(X, "*í*", "[ìí]")
    
    X = Replace(X, MultiLeter, "!" & MultiLeter)
    X = Replace(X, OneLetter, "!" & OneLetter)
    'Replace each space with "_%"
    X = Replace(X, " ", OneLetter & MultiLeter)
    X = Replace(MultiLeter & X & MultiLeter, MultiLeter & MultiLeter, MultiLeter)
    X = "'" & X & "'"
    If InStr(1, X, "!" & MultiLeter) > 0 Or InStr(1, X, "!" & OneLetter) > 0 Then
        X = X & " ESCAPE '!'"
    End If
    LikeExpression = X
End Function

Sub ReadIniFile(FileName As String, Delimiter As String)
    Dim Fnum As Integer, XX() As String, FileStr As String
    Fnum = FreeFile()
    Open FileName For Input As #Fnum
    FileStr = Input(LOF(Fnum), Fnum)
    XX() = Split(FileStr, Delimiter)
    For i = LBound(XX) To UBound(XX)
        If i = 0 Then
            ServerName = XX(i)
        ElseIf i = 1 Then
            DataBase = XX(i)
        ElseIf i = 2 Then
            DatabaseYear = XX(i)
       End If
    Next
End Sub

Function ConvertControlDate(ByVal StrDate As String) As String
Dim str1 As String
If IsDate(StrDate) Then
    ConvertControlDate = Right("00" + Mid(StrDate, 4, 2), 2) + "/" + Right("00" + Mid(StrDate, 1, 2), 2) + "/" + Right("0000" + Mid(StrDate, 7, 4), 4)
Else
    ConvertControlDate = "01/01/1900"
End If
End Function

Function ConvertSqlDate(DateStr As String) As String
If IsDate(DateStr) Then
    ConvertSqlDate = Right("00" + Mid(DateStr, 1, 2), 2) + "/" + Right("00" + Mid(DateStr, 4, 2), 2) + "/" + Right("0000" + Mid(DateStr, 7, 4), 4)
Else
    ConvertSqlDate = "__/__/____"
End If
End Function
Function DeleteRow(Grid As VSFlexGrid, Vrow As Integer, Col As Integer, Table As String, Id As String, Optional ByVal Op) As Boolean
On Error GoTo ErrorHandler
With Grid
If IsMissing(Op) Then
    sqlText = "Delete From " & Table & " Where " & Id & " = " & .TextMatrix(Vrow, Col)
    de.con.Execute (sqlText)
ElseIf Op = 2 Then
    sqlText = "Delete From " & Table & " Where " & Id & " = '" & .TextMatrix(Vrow, Col) & "'"
    de.con.Execute (sqlText)
End If
End With
DeleteRow = True
Exit Function
ErrorHandler:
DeleteRow = False
MsgBox (Err.Description)
End Function

'Function GetPass() As String
'    Dim CnnPass As New ADODB.Connection
'    Dim CnnTest As New ADODB.Connection
'    Dim CmdTime As New ADODB.Command
'    Dim rsCmdTime As New ADODB.Recordset
'
'    Dim StrHashPass As String
'    Dim str1 As String, str2 As String, str3 As String, str4 As String, str5 As String
'    Const StrSymbols As String = "~!@#$%^&*?><~!@#$%^&*?><~!@#"
'    Const StrCapLettres As String = "QWERTYUIOPASDFGHJKLZXCVBNMQWERTYUIOQWERT"
'    Dim StrSmlLettres As String
'    Const StrDigits As String = "0123456789346127589465827346342598069"
'    Dim NumDay As Integer, NumMonth As Integer, NumHour As Integer
'
'    CnnPass.ConnectionString = "Provider=SQLOLEDB.1;Password=usertime;Persist Security Info=True;User ID=Usertime;Data Source=MAINSERVER;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;Workstation ID=HUSAM;Use Encryption for Data=False;Tag with column collation when possible=False"
'    CnnPass.Open
'    CmdTime.ActiveConnection = CnnPass
'    CmdTime.CommandText = " Select GetDate() as CurrTime "
'
'
'    StrSmlLettres = LCase(StrCapLettres)
'    Dim CurrDate As Date
'    If rsCmdTime.State = adStateOpen Then rsCmdTime.Close
'    Set rsCmdTime = CmdTime.Execute()
'
'    CurrDate = rsCmdTime.Fields(0).Value
'    NumDay = Day(CurrDate)
'    NumMonth = Month(CurrDate)
'    NumHour = Hour(CurrDate)
'
'    str1 = Mid(StrSymbols, NumMonth, 4)
'    str2 = Mid(StrCapLettres, NumDay, 4)
'    str3 = Mid(StrSmlLettres, NumDay + 4, 4)
'    str4 = Mid(StrDigits, NumDay, 4)
'    str5 = Mid(StrSymbols, NumHour, 4)
'    StrHashPass = ""
'    For I = 1 To 4
'        StrHashPass = StrHashPass & Mid(str1, I, 1)
'        StrHashPass = StrHashPass & Mid(str2, I, 1)
'        StrHashPass = StrHashPass & Mid(str3, I, 1)
'        StrHashPass = StrHashPass & Mid(str4, I, 1)
'        StrHashPass = StrHashPass & Mid(str5, I, 1)
'    Next I
'
'    CnnPass.Close
'    On Error Resume Next
'    If CnnTest.State <> adStateClosed Then CnnTest.Close
'    CnnTest.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=True;Data Source=MAINSERVER;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;Workstation ID=HUSAM;Use Encryption for Data=False;Tag with column collation when possible=False"
'    CnnTest.Open , "user1", StrHashPass
'   ' MsgBox (CnnTest.State = 1)
'    CnnTest.Close
'    GetPass = StrHashPass
'End Function

Function ConnectName(vdate As String) As String
    systemConfigration.Password = GetPass()
    If vdate = "" Then
        ConnectName = "Odbc;Uid=" & systemConfigration.UserId & ";Pwd=" & systemConfigration.Password & ";Dsn=" & systemConfigration.DSN & ";Database=" & systemConfigration.DatabaseName & ";"
    Else

        ConnectName = "Odbc;Uid=" & systemConfigration.UserId & ";Pwd=" & systemConfigration.Password & ";Dsn=" & GetOdbcName(vdate) & ";Database=" & systemConfigration.DatabaseName & ";"
    End If
End Function

Function GetOdbcName(vdate As String) As String
    Dim rs As New ADODB.Recordset
    If Not IsDate(vdate) Then vyear = Year(Now) Else vyear = Year(vdate)
    sqlText = "Select OdbcName From master.dbo.CoDatabaseName Where Year=" & vyear & " and Class=0"
    Set rs = de.con.Execute(sqlText)
    GetOdbcName = rs!OdbcName & ""
End Function
'Function ConnectName() As String
'    ConnectName = "Odbc;Uid=user1;Pwd=" & GetPass1 & ";Dsn=ss;Database=" & DataBase & ";"
'End Function

Function GetDatabaseName(vdate As String) As String
On Error GoTo ErrorHandler
    Dim rs As New ADODB.Recordset
    If vdate = "01/01/1900" Then vdate = "01/01/" & LTrim(RTrim(Str(Year(Now))))
    If Not IsDate(vdate) Then vyear = Year(Now) Else vyear = Year(vdate)
    sqlText = "Select DataBaseName From master.dbo.CoDatabaseName Where Year=" & vyear
    Set rs = de.con.Execute(sqlText)
    GetDatabaseName = rs!DatabaseName & ""
Exit Function
ErrorHandler:
MsgBox Err.Description
End Function
Function NewRec() As Double
Dim RsMax As New ADODB.Recordset
sqlText = "Select Isnull(Max(ByanId),0) As MaxByanId From Stmov"
Set RsMax = de.con.Execute(sqlText)

If RsMax!maxByanId = 0 Then
    NewRec = 1
Else
    NewRec = RsMax!maxByanId + 1
End If
End Function
Function GetStkId(stkno As String) As Double
On Error GoTo ErrorHandler
Dim rs As New ADODB.Recordset
sqlText = "Select Id From CoStock Where Stkno='" & stkno & "'"
Set rs = de.con.Execute(sqlText)
If rs.RecordCount > 0 Then
   GetStkId = rs!Id
Else
    GetStkId = -1
End If
Exit Function
ErrorHandler:
GetStkId = -1
End Function

Function GetStrId(StrNo As Integer) As Integer
On Error GoTo ErrorHandler
Dim rs As New ADODB.Recordset
sqlText = "Select Id From namestr Where strno=" & StrNo
Set rs = de.con.Execute(sqlText)
If rs.RecordCount > 0 Then
   GetStrId = rs!Id
Else
    GetStrId = -1
End If
Exit Function
ErrorHandler:
GetStrId = -1
End Function

Function GetBalance(stkno As String) As Double
On Error GoTo ErrorHandler
Dim rs As New ADODB.Recordset
sqlText = "Select FnlQnt From GetBalanceQry Where StkNo='" & stkno & "'"
Set rs = de.con.Execute(sqlText)
If rs.RecordCount > 0 Then
    GetBalance = rs!fnlqnt
Else
    GetBalance = 0
End If

Exit Function
ErrorHandler:
GetBalance = 0
End Function

'Function ConvertControlDate(ByVal StrDate As String) As String
'Dim Str1 As String
'If IsDate(StrDate) Then
'    ConvertControlDate = Right("00" + Mid(StrDate, 4, 2), 2) + "/" + Right("00" + Mid(StrDate, 1, 2), 2) + "/" + Right("0000" + Mid(StrDate, 7, 4), 4)
'Else
'    ConvertControlDate = "01/01/1900"
'End If
'End Function
'
'Function ConvertSqlDate(dateStr As String) As String
'If IsDate(dateStr) Then
'    ConvertSqlDate = Right("00" + Mid(dateStr, 1, 2), 2) + "/" + Right("00" + Mid(dateStr, 4, 2), 2) + "/" + Right("0000" + Mid(dateStr, 7, 4), 4)
'Else
'    ConvertSqlDate = "__/__/____"
'End If
'End Function

Sub ParseXml()
    Dim doc As New MSXML2.DOMDocument
    Dim success As Boolean
    doc.async = False
    doc.validateOnParse = True
    
 success = doc.Load(App.Path & "\init.xml")
    If Not success Then
        MsgBox "there is not configration file"
    Else
    Dim nodeList As MSXML2.IXMLDOMNodeList
    Set nodeList = doc.selectNodes("/Application/Keys")
    If Not nodeList Is Nothing Then
    Dim node As MSXML2.IXMLDOMNode
     Dim name As String
     Dim Value As String
    
     For Each node In nodeList
        Select Case node.selectSingleNode("@name").Text
         Case "SystemName":
            systemConfigration.SystemName = node.selectSingleNode("@value").Text
         Case "ServerName":
            systemConfigration.ServerName = node.selectSingleNode("@value").Text
         Case "DatabaseName":
            systemConfigration.DatabaseName = node.selectSingleNode("@value").Text
         Case "UserName":
            systemConfigration.UserId = node.selectSingleNode("@value").Text
         Case "PasswordMethod":
            systemConfigration.PasswordMethod = node.selectSingleNode("@value").Text
         Case "Year":
            systemConfigration.Year = node.selectSingleNode("@value").Text
        Case "Version":
            systemConfigration.Version = node.selectSingleNode("@value").Text
         Case "MainStoreNo":
            systemConfigration.MainStoreNo = node.selectSingleNode("@value").Text
         Case "DatabaseDestination":
            systemConfigration.DatabaseDestination = node.selectSingleNode("@value").Text
         Case "Dsn":
            systemConfigration.DSN = node.selectSingleNode("@value").Text
          Case "SystemUserName":
            systemConfigration.SystemUserName = node.selectSingleNode("@value").Text
        Case "SystemPassword":
            systemConfigration.SystemPassword = node.selectSingleNode("@value").Text
        Case "HafezHallsDatabaseDestination":
            systemConfigration.HafezHallsDatabaseDestination = node.selectSingleNode("@value").Text
        
         End Select
     Next node
    End If
    End If
End Sub

