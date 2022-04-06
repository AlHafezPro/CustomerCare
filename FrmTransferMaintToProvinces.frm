VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FrmTransferMaintToProvinces 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " —ÕÌ· »Ì«‰«  Œœ„… «·„” Â·ﬂ"
   ClientHeight    =   1875
   ClientLeft      =   45
   ClientTop       =   720
   ClientWidth     =   13320
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   1875
   ScaleWidth      =   13320
   Begin MSComctlLib.ProgressBar ProgressBar1 
      DragMode        =   1  'Automatic
      Height          =   435
      Left            =   30
      TabIndex        =   10
      Top             =   1860
      Width           =   13245
      _ExtentX        =   23363
      _ExtentY        =   767
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.CommandButton Command7 
      Caption         =   "«·⁄«∆·« "
      Height          =   375
      Left            =   3840
      TabIndex        =   2
      Top             =   1470
      Width           =   1845
   End
   Begin VB.CommandButton Command6 
      Caption         =   "≈‘⁄«—«  «·„Ê«œ"
      Height          =   375
      Left            =   9555
      TabIndex        =   5
      Top             =   1470
      Width           =   1845
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4440
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "«·„ÊœÌ·« "
      Height          =   375
      Left            =   30
      TabIndex        =   0
      Top             =   1470
      Width           =   1845
   End
   Begin VB.CommandButton Command2 
      Caption         =   "«·„Ê«œ"
      Height          =   375
      Left            =   1935
      TabIndex        =   1
      Top             =   1470
      Width           =   1845
   End
   Begin VB.Frame Frame1 
      Caption         =   "«·„⁄·Ê„«  «·≈”«”Ì…"
      ForeColor       =   &H00800000&
      Height          =   1455
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   13275
      Begin VB.OptionButton Option1 
         Caption         =   " ’œÌ— «·„⁄·Ê„«  «·√”«”Ì…"
         DragMode        =   1  'Automatic
         ForeColor       =   &H000000FF&
         Height          =   375
         Index           =   0
         Left            =   480
         TabIndex        =   9
         Top             =   360
         Value           =   -1  'True
         Width           =   3255
      End
      Begin VB.OptionButton Option1 
         Caption         =   "≈” Ì—«œ «·„⁄·Ê„«  «·√”«”Ì…"
         ForeColor       =   &H000000FF&
         Height          =   375
         Index           =   1
         Left            =   480
         TabIndex        =   8
         Top             =   840
         Visible         =   0   'False
         Width           =   3255
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "—»ÿ «·⁄«∆·Â „⁄ «·„ÊœÌ·"
      Height          =   375
      Left            =   5745
      TabIndex        =   3
      Top             =   1470
      Width           =   1845
   End
   Begin VB.CommandButton Command4 
      Caption         =   "«·√⁄ÿ«·"
      Height          =   375
      Left            =   7650
      TabIndex        =   4
      Top             =   1470
      Width           =   1845
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Œ—ÊÃ"
      Height          =   375
      Left            =   11460
      TabIndex        =   6
      Top             =   1470
      Width           =   1845
   End
   Begin VB.Menu File 
      Caption         =   "«·—»ÿ „⁄ »Ì«‰«  «·„»Ì⁄« "
      Begin VB.Menu mnuimport 
         Caption         =   "«” Ì—«œ «·„Êœ»·«  „‰ „·› «·„»Ì⁄« "
      End
   End
End
Attribute VB_Name = "FrmTransferMaintToProvinces"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub Export(Vindex As Integer)
On Error Resume Next
Dim rs As New ADODB.Recordset
Select Case Vindex
    Case 1 ' AdhamModels
        sqlText = "Select ModNo, AccNo, RetAccNo, Symbol, Name, GrpNo, FamNo, DealPrice, DistPrice, ConsPrice, DealDisc, DistDisc, ModYear, ProdKind, InventPoint, ItemNo From adhammodels"
        Set rs = de.con.Execute(sqlText)
        Kill "E:\MainData\AdhamModels.txt"
        rs.Save "E:\MainData\AdhamModels.txt", adPersistADTG
        rs.Close
        MsgBox " „ Õ›Ÿ „⁄·Ê„«  «·„ÊœÌ·«  ⁄·Ï «·„”«—" & " " & "E:\MainData\AdhamModels.txt"""
    Case 2 'Pieces
'        sqltext = "select stkno , stkname , dealpriceafterdiscount as dealprice , DistPriceafterdiscount distprice  , CliPriceafterdiscount cliprice , discount , accno from hafez2012.dbo.costock"
        sqlText = "select Id , PieceStockNo stkno , PieceName stkname ,  CliPrice dealPrice , CliPrice , CliPrice DistPrice , CliPriceafterdiscount dealpriceafterdiscount  ,  CliPriceafterdiscount DistPriceafterdiscount ,  CliPriceafterdiscount , discount , accno from pieces where CliPrice!=0"
        Set rs = de.con.Execute(sqlText)
        Kill "E:\MainData\Pieces.txt"
        rs.Save "E:\MainData\Pieces.txt", adPersistADTG
        rs.Close
        MsgBox " „ Õ›Ÿ „⁄·Ê„«  «·√—ﬁ«„ «·„Œ“‰Ì… ⁄·Ï «·„”«—" & " " & "E:\MaintData\Pieces.txt"
    Case 3 'MntPiecesFamJoin
        sqlText = "Select PieceStockNo, famNo From MntPiecesFamJoin"
        Set rs = de.con.Execute(sqlText)
        Kill "E:\MainData\MntPiecesFamJoin.txt"
        rs.Save "E:\MainData\MntPiecesFamJoin.txt", adPersistADTG
        rs.Close
        MsgBox " „ Õ›Ÿ „⁄·Ê„«  «·„ÊœÌ·«  ⁄·Ï «·„”«—" & " " & "E:\MaintData\MntPiecesFamJoin.txt"
    Case 4 'coReparationType
        sqlText = "Select RepTypeNo, RepTypeDescription, RepTypeTime, RepTypePrice, PanneNo, RepClassNo, Notes From coReparationType"
        Set rs = de.con.Execute(sqlText)
        Kill "E:\MainData\coReparationType.txt"
        rs.Save "E:\MainData\coReparationType.txt", adPersistADTG
        rs.Close
        MsgBox " „ Õ›Ÿ «·√⁄ÿ«· ⁄·Ï «·„”«—" & " " & "E:\MaintData\coReparationType.txt"
    Case 6 'adhamproductfamily
        sqlText = "select ProdFamNo, ProdFamName, ProdFamNameA, ProdFamOrd, ProductivityCall from adhamproductfamily"
        Set rs = de.con.Execute(sqlText)
        Kill "E:\MainData\AdhamProductFamily.txt"
        rs.Save "E:\MainData\AdhamProductFamily.txt", adPersistADTG
        rs.Close
        MsgBox " „ Õ›Ÿ «·⁄«∆·«  ⁄·Ï «·„”«—" & " " & "E:\MaintData\AdhamProductFamily.txt"
End Select
'    sqltext = "SELECT a.* FROM OPENROWSET('SQLOLEDB','" & ServerName & "';'" & UID & "';'" & PWD & "','SELECT * FROM hafezgeneral.dbo.tt4 ') AS a"
End Sub
Function foundItems(NotificationNo, PieceStockNo, MvtDate As String) As Boolean
On Error GoTo ErrorHandler
Dim rs As New ADODB.Recordset
sqlText = "Select count(*) as CountREc From piecesinstock where NotificationNo=" & NotificationNo & " And PieceStockNo='" & PieceStockNo & "' and MvtDate='" & ConvertControlDate(MvtDate) & "'"
Set rs = de.con.Execute(sqlText)
If rs!CountRec = 0 Then
    foundItems = False
Else
    foundItems = True
End If
Exit Function
ErrorHandler:
foundItems = False
End Function

Private Sub Import(Vindex As Integer)
On Error GoTo ErrorHandler
Dim ICount As Double
Dim sqlText As String
Dim FileName  As String

Dim rsMaint As New ADODB.Recordset
CommonDialog1.Filter = "*.txt"
CommonDialog1.ShowOpen
If CommonDialog1.FileName = "" Then
    MsgBox " ·„ Ì „  ÕœÌœ ≈”„ «·„·›", vbExclamation
    Exit Sub
End If
rsMaint.Open CommonDialog1.FileName, , , , adCmdFile


If rsMaint.RecordCount > 0 Then
    rsMaint.MoveFirst
Else
    MsgBox " ·«ÌÊÃœ »Ì«‰«  ÷„‰ «·„·›", vbExclamation
    Exit Sub
End If
ProgressBar1.Visible = True
ProgressBar1.Min = 1
ProgressBar1.Max = rsMaint.RecordCount + 1
ProgressBar1.Value = 1

Select Case Vindex
    Case 1 'AdhamModels
        sqlText = "Delete From AdhamModels"
        de.con.Execute (sqlText)
    Case 2 'Pieces
        sqlText = "Delete From Pieces"
        de.con.Execute (sqlText)
    Case 3 'MntPiecesFamJoin
        sqlText = "Delete From MntPiecesFamJoin"
        de.con.Execute (sqlText)
    Case 4 'coReparationType
        sqlText = "Delete From coReparationType"
        de.con.Execute (sqlText)
    Case 6
        sqlText = "Delete From adhamproductfamily"
        de.con.Execute (sqlText)
End Select

While Not rsMaint.EOF
Select Case Vindex
    Case 1 'AdhamModels
            sqlText = "Insert Into dbo.AdhamModels (ModNo, AccNo, RetAccNo, Symbol, Name, GrpNo, FamNo, InventPoint, ItemNo) Values(" & rsMaint!ModNo & ",'" & rsMaint!AccNo & "','" & rsMaint!RetAccNo & "','" & rsMaint!Symbol & "','" & rsMaint!name & "'," & rsMaint!GrpNo & "," & rsMaint!FamNo & "," & Abs(rsMaint!InventPoint) & "," & rsMaint!ItemNo & ")"
            de.con.Execute (sqlText)
            ICount = ICount + 1
    Case 2 'Pieces
            sqlText = "insert into Pieces(PieceName, PieceStockNo, qty, CliPrice, DistPrice, DealPrice, discount , AccNo)VAlues('" & Replace(rsMaint!StkName, "'", "''") & "','" & rsMaint!stkno & "',5000," & IIf(IsNull(rsMaint!CliPrice), 0, rsMaint!CliPrice) & "," & IIf(IsNull(rsMaint!DistPrice), 0, rsMaint!DistPrice) & "," & IIf(IsNull(rsMaint!DealPrice), 0, rsMaint!DealPrice) & "," & IIf(IsNull(rsMaint!discount), 0, rsMaint!discount) & ",'" & rsMaint!AccNo & "')"
            de.con.Execute (sqlText)
            ICount = ICount + 1
    Case 3 'Models
            sqlText = "insert into MntPiecesFamJoin(PieceStockNo, famNo)Values('" & rsMaint!PieceStockNo & "'," & FamNo & ")"
            ICount = ICount + 1
    Case 4 'coReparationType
            sqlText = "insert into dbo.coReparationType(RepTypeNo, RepTypeDescription)Values('" & rsMaint!RepTypeNo & "','" & rsMaint!RepTypeDescription & "')"
            de.con.Execute (sqlText)
            ICount = ICount + 1
    Case 5 'piecesinstock
            If Not foundItems(rsMaint!NotificationNo, rsMaint!PieceStockNo, rsMaint!MvtDate) Then
                sqlText = "Insert Into dbo.piecesinstock(NotificationNo, PieceStockNo, MvtDate, Qty, OpKind,FromMainCenter)Values(" & rsMaint!NotificationNo & ",'" & rsMaint!PieceStockNo & "','" & ConvertControlDate(rsMaint!MvtDate) & "'," & rsMaint!Qty & "," & rsMaint!OpKind & ",2)"
                de.con.Execute (sqlText)
                ICount = ICount + 1
            End If
    Case 6
            sqlText = "insert into dbo.adhamproductfamily(ProdFamNo, ProdFamName, ProdFamNameA, ProdFamOrd )Values(" & rsMaint!ProdFamNo & ",'" & rsMaint!ProdFamName & "','" & rsMaint!ProdFamNameA & "'," & IIf(IsNull(rsMaint!ProdFamOrd), 0, rsMaint!ProdFamOrd) & ")"
            de.con.Execute (sqlText)
            ICount = ICount + 1
        
End Select
    ProgressBar1.Value = ProgressBar1.Value + 1
    rsMaint.MoveNext
Wend
MsgBox " „ ≈÷«›… " & ICount & " ⁄·Ï ﬁ«⁄œ… «·»Ì«‰« "
ProgressBar1.Visible = False
Exit Sub
ErrorHandler:
MsgBox Err.Description
End Sub


Private Sub Command1_Click()
If Option1(0).Value Then
    Export (1)
Else
    Import (1)
End If
End Sub

Private Sub Command2_Click()
If Option1(0).Value Then
    Export (2)
Else
    Import (2)
End If
End Sub

Private Sub Command3_Click()
If Option1(0).Value Then
    Export (3)
Else
    Import (3)
End If
End Sub

Private Sub Command4_Click()
If Option1(0).Value Then
    Export (4)
Else
    Import (4)
End If
End Sub

Private Sub Command5_Click()
Unload Me
End Sub

Private Sub Command6_Click()
Import (5)
End Sub

Private Sub Command7_Click()
If Option1(0).Value Then
    Export (6)
Else
    Import (6)
End If
End Sub

Private Sub Form_Load()
init
End Sub

Sub init()
    top = 0
    left = 0
End Sub

Function ImortModelFromSales() As Boolean
On Error GoTo ErrorHandler
ImortModelFromSales = False
sqlText = "Dbo.Transfer_ModelsToAdham"
de.con.Execute (sqlText)
ImortModelFromSales = True

Exit Function
ErrorHandler:
MsgBox Err.Description
ImortModelFromSales = False

Exit Function
End Function
Private Sub mnuimport_Click()
If ImortModelFromSales() Then

    MsgBox " „ «” Ì—«œ «·„ÊœÌ·«  „‰ „·› «·„»Ì⁄« "
End If
End Sub
