VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmReturnStockBalance 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " œÊÌ— «·«—’œÂ ··«—’œÂ «·„⁄·„Â"
   ClientHeight    =   1740
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   1740
   ScaleWidth      =   4575
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7890
      Top             =   60
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   38
      ImageHeight     =   38
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   38
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReturnStockBalance.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReturnStockBalance.frx":26D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReturnStockBalance.frx":5533
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReturnStockBalance.frx":7CD8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReturnStockBalance.frx":A5D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReturnStockBalance.frx":CAF9
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReturnStockBalance.frx":F2B1
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReturnStockBalance.frx":11CC5
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReturnStockBalance.frx":14617
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReturnStockBalance.frx":1738B
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReturnStockBalance.frx":19BDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReturnStockBalance.frx":1CA81
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReturnStockBalance.frx":1F7DB
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReturnStockBalance.frx":2217B
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReturnStockBalance.frx":250DD
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReturnStockBalance.frx":27B05
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReturnStockBalance.frx":2A4BF
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReturnStockBalance.frx":2CDF0
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReturnStockBalance.frx":2F6F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReturnStockBalance.frx":3215E
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReturnStockBalance.frx":3510F
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReturnStockBalance.frx":37A38
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReturnStockBalance.frx":3A36D
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReturnStockBalance.frx":3CAAD
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReturnStockBalance.frx":3F41C
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReturnStockBalance.frx":41CA4
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReturnStockBalance.frx":43EB3
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReturnStockBalance.frx":46810
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReturnStockBalance.frx":4903A
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReturnStockBalance.frx":4BAEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReturnStockBalance.frx":4E52A
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReturnStockBalance.frx":51040
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReturnStockBalance.frx":53FDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReturnStockBalance.frx":56DA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReturnStockBalance.frx":59A20
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReturnStockBalance.frx":5C8B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReturnStockBalance.frx":5F4F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReturnStockBalance.frx":6209F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   690
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   1217
      ButtonWidth     =   1191
      ButtonHeight    =   1164
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   14
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin MSMask.MaskEdBox TxtFirstDate 
      Height          =   375
      Left            =   1590
      TabIndex        =   0
      Top             =   1200
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "99/99/9999"
      PromptChar      =   "_"
   End
   Begin VB.Label LDatabaseDestination 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   780
      Width           =   2655
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "”‰Â «· —ÕÌ·"
      Height          =   195
      Left            =   3540
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   1200
      Width           =   810
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "ﬁ«⁄œÂ «·»Ì«‰«  «·Âœ›"
      Height          =   195
      Left            =   2940
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   780
      Width           =   1410
   End
End
Attribute VB_Name = "FrmReturnStockBalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim frmOption_ As FrmPieces
Private Sub Form_Load()
init
End Sub

Sub init()
Dim rsDatabases As New ADODB.Recordset
Dim sqlText As String

TxtFirstDate.Text = "01/01/" & Right("0000" + Str(Year(Now) + 1), 4)
LDatabaseDestination.Caption = systemConfigration.DatabaseDestination

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
    Case 1
       If TransferBalances Then
            MsgBox " „  —ŒÌ· «·«—’œÂ »‰Ã«Õ", vbInformation, " —ÕÌ· «·«—’œÂ"
        Else
            MsgBox " ÌÊÃœ Õÿ√ ›Ì  —ÕÌ· «·«—’œÂ", vbCritical, " ‰»ÌÂ"
        End If
    Case 3
        Unload Me
End Select
End Sub

Function GetMaxByanId() As Double
On Error GoTo errorhandler
Dim RsMax As New ADODB.Recordset
sqlText = "Select isnull(Max(ByanId),0) as MaxByanId From " & systemConfigration.DatabaseDestination & ".dbo.Stmov"
Set RsMax = de.con.Execute(sqlText)
If RsMax!maxByanId = 0 Then
    GetMaxByanId = 1
Else
    GetMaxByanId = RsMax!maxByanId + 1
End If
Exit Function
errorhandler:
GetMaxByanId = -1
MsgBox Err.Description
End Function

Function TransferBalances() As Boolean
On Error GoTo errorhandler
Dim sqlText As String
Dim rsCount As New ADODB.Recordset
Dim transactionNbr As Double
sqlText = ""
Screen.MousePointer = vbHourglass

    transactionNbr = de.con.BeginTrans
    sqlText = sqlText & " delete from " & systemConfigration.DatabaseDestination & ".dbo.stmov where doctype=0 "
    
    sqlText = sqlText & "insert into " & systemConfigration.DatabaseDestination & ".dbo.Stmov(ByanId, StkId,  StrId, MovDate,  DocType,  Qty, QtyType, EmpNo)"

    sqlText = sqlText & "select " & GetMaxByanId & " , stkid , strid , '" & ConvertControlDate(TxtFirstDate.Text) & "' ,0, sum(qty) , 0 ," & empNo & "    from "
    sqlText = sqlText & "("
    sqlText = sqlText & "select strid , stkid , fnlqnt qty  from stkinf where fnlqnt !=0 "
    sqlText = sqlText & "Union All "
    sqlText = sqlText & "Select  " & GetStrId(systemConfigration.MainStoreNo) & " , c1.id  ,case when operationtype=1 then qty else -qty end qty from "
    
    sqlText = sqlText & "MvMaintPayments m1 inner join "
    sqlText = sqlText & "MvMaintPaymentsDetails m2 on m1.BillNo = m2.BillNo inner join "
    sqlText = sqlText & "costock c1 on m2.stkno = c1.stkno collate arabic_ci_as "
    
    sqlText = sqlText & "Where IsFixed = 0 "
    sqlText = sqlText & ")t1 group by strid , stkid "
    de.con.Execute (sqlText)
    
    sqlText = "select count(*) as countrec from " & systemConfigration.DatabaseDestination & ".dbo.stkinf where fnlqnt <0 "
    Set rsCount = de.con.Execute(sqlText)
    If rsCount!CountRec >= 1 Then
        de.con.RollbackTrans
    End If
    de.con.CommitTrans

TransferBalances = True
Screen.MousePointer = vbDefault

Exit Function
errorhandler:
MsgBox Err.Description
TransferBalances = False
Screen.MousePointer = vbDefault
If transactionNbr > 0 Then de.con.RollbackTrans
End Function


