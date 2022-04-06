VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmCustomerCareMeasuring 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ﬁÌ«” —÷Ï «·“»«∆‰"
   ClientHeight    =   6765
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   16575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   6765
   ScaleWidth      =   16575
   Begin Threed.SSFrame SSFrame1 
      Height          =   645
      Left            =   30
      TabIndex        =   4
      Top             =   6060
      Width           =   16515
      _ExtentX        =   29131
      _ExtentY        =   1138
      _Version        =   131074
      Begin VB.Label Lcount 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   13560
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   180
         Width           =   1455
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "⁄œœ «·«’·«Õ« "
         Height          =   285
         Left            =   15450
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   180
         Width           =   975
      End
   End
   Begin Crystal.CrystalReport cr1 
      Left            =   5010
      Top             =   2640
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSMask.MaskEdBox TxtDate 
      Height          =   315
      Left            =   13620
      TabIndex        =   1
      Top             =   750
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "99/99/9999"
      PromptChar      =   "_"
   End
   Begin VSFlex8Ctl.VSFlexGrid flexGrid 
      Height          =   4935
      Left            =   30
      TabIndex        =   2
      Top             =   1080
      Width           =   16515
      _cx             =   29131
      _cy             =   8705
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   -1  'True
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   5
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   -1  'True
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   810
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
            Picture         =   "FrmCustomerCareMeasuring.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCustomerCareMeasuring.frx":26D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCustomerCareMeasuring.frx":4E7B
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCustomerCareMeasuring.frx":7777
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCustomerCareMeasuring.frx":A126
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCustomerCareMeasuring.frx":C64B
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCustomerCareMeasuring.frx":EE03
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCustomerCareMeasuring.frx":11817
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCustomerCareMeasuring.frx":14169
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCustomerCareMeasuring.frx":16EDD
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCustomerCareMeasuring.frx":1972C
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCustomerCareMeasuring.frx":1C5D3
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCustomerCareMeasuring.frx":1F32D
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCustomerCareMeasuring.frx":21CCD
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCustomerCareMeasuring.frx":24C2F
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCustomerCareMeasuring.frx":27A8C
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCustomerCareMeasuring.frx":2A4B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCustomerCareMeasuring.frx":2CE6E
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCustomerCareMeasuring.frx":2F79F
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCustomerCareMeasuring.frx":323DD
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCustomerCareMeasuring.frx":34CE3
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCustomerCareMeasuring.frx":3774B
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCustomerCareMeasuring.frx":3A6FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCustomerCareMeasuring.frx":3D025
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCustomerCareMeasuring.frx":3F95A
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCustomerCareMeasuring.frx":42509
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCustomerCareMeasuring.frx":44C49
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCustomerCareMeasuring.frx":475B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCustomerCareMeasuring.frx":49E40
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCustomerCareMeasuring.frx":4C04F
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCustomerCareMeasuring.frx":4E9AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCustomerCareMeasuring.frx":511D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCustomerCareMeasuring.frx":53C8A
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCustomerCareMeasuring.frx":566C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCustomerCareMeasuring.frx":591DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCustomerCareMeasuring.frx":5C176
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCustomerCareMeasuring.frx":5EF42
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCustomerCareMeasuring.frx":61BBC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   690
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   16575
      _ExtentX        =   29236
      _ExtentY        =   1217
      ButtonWidth     =   1191
      ButtonHeight    =   1164
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   16
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   25
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   15
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   " «—ÌŒ ≈œŒ«· «·«’·«Õ"
      Height          =   195
      Index           =   0
      Left            =   15120
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   780
      Width           =   1365
   End
End
Attribute VB_Name = "FrmCustomerCareMeasuring"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const ColIsSmsSend = 1
Const ColMobilePhone = 2
Const ColCallNo = 3
Const colCountRec = 4
Const coladhamName = 5
Const ColAdhamPhon = 6
Const ColCallDatetime = 7

Const colRepDate = 8
Const ColRegestDate = 9


Const colProdFamNameA = 10
Const ColCallDEscription = 11
Const colDescription = 12
Const colNotes = 13
Const ColTeamName = 14
Const colRepPrice = 15



Sub FillFormating(ByVal i As Integer)
If i = 1 Then
    fs = "|>" + "≈—”«· SMS"
    fs = fs + "|>" + "—ﬁ„ «·„Ê»«Ì·/‘ﬂ«ÊÏ"
    fs = fs + "|>" + "—ﬁ„ «·‘ﬂÊÏ"
    fs = fs + "|>" + "⁄œœ „—«  «·“Ì«—… "
    fs = fs + "|>" + "≈”„ «·“»Ê‰ "
    fs = fs + "|>" + "—ﬁ„ «·Â« ›"
    fs = fs + "|>" + " «—ÌŒ «·‘ﬂÊÏ"
    fs = fs + "|>" + " «—ÌŒ «·≈’·«Õ"
    fs = fs + "|>" + " «—ÌŒ «·«œŒ«·"
    fs = fs + "|>" + "«·„‰ Ã"
    fs = fs + "|>" + "«·‘ﬂÊÏ"
    fs = fs + "|>" + "‘—Õ"
    fs = fs + "|>" + "„·«ÕŸ« "
    fs = fs + "|>" + "«·Ê—‘… «·„‰›–Â"
    fs = fs + "|>" + "«·ﬂ·›…"
    
    With flexGrid
        .FormatString = fs
        .Cols = 16
        .ColWidth(ColCallNo) = 0
        SetColWidths ColMobilePhone, flexGrid
        SetColWidths coladhamName, flexGrid
        SetColWidths ColAdhamPhon, flexGrid
        SetColWidths ColCallDatetime, flexGrid
        SetColWidths colRepDate, flexGrid
        SetColWidths ColRegestDate, flexGrid
        SetColWidths colProdFamNameA, flexGrid
        SetColWidths ColCallDEscription, flexGrid
        .ColWidth(colDescription) = 0
        SetColWidths colNotes, flexGrid
        SetColWidths ColTeamName, flexGrid
        SetColWidths colRepPrice, flexGrid
        SetColWidths colCountRec, flexGrid
        SetColWidths ColIsSmsSend, flexGrid
        .ColDataType(ColIsSmsSend) = flexDTBoolean
        
    End With
End If
End Sub

'Sub SetColWidths(ByVal ColNo As Integer, FlexGrid As VSFlexGrid)
'    With FlexGrid
'        .AutoSize (ColNo)
'    End With
'End Sub


Private Sub FlexGrid_DblClick()
'    With flexGrid
'        CustomerInformationRec.AdhamName = .TextMatrix(.Row, ColAdhamName)
'        CustomerInformationRec.AdhamPhon = .TextMatrix(.Row, ColAdhamPhon)
'
'        CustomerInformationRec.CallDatetime = .TextMatrix(.Row, ColCallDatetime)
'        CustomerInformationRec.RepDate = .TextMatrix(.Row, ColRepDate)
'        CustomerInformationRec.CallDEscription = .TextMatrix(.Row, ColCallDEscription)
'        CustomerInformationRec.Description = .TextMatrix(.Row, ColNotes)
'         CustomerInformationRec.TeamName = .TextMatrix(.Row, ColTeamName)
'        CustomerInformationRec.RepPrice = .TextMatrix(.Row, ColRepPrice)
'        CustomerInformationRec.CountRec = .TextMatrix(.Row, colCountRec)
'
'
'    End With
'    FrmCustomerSastisfaction.Show 1
    
End Sub

Sub FillGrid(vdate As String)
On Error GoTo ErrorHandler
Dim sqlText As String
Dim rs As New ADODB.Recordset
'sqlText = "Exec sp_GetCalls '" & ConvertControlDate(vdate) & "'"
'de.con.Execute (sqlText)

sqlText = "select IsSmsSend, MobilePhone , CallNo, CountRec , adhamname, adhamphon, Convert(varchar(10),CallDateTime,102)CallDateTime  , RepDate , RegestDate ,  ProdFamNameA,"
sqlText = sqlText & "CallDescription, Description, Notes, TeamName, RepPrice   "
sqlText = sqlText & " from  t_CallsDate Order By TeamNo"
Set rs = de.con.Execute(sqlText)

Set flexGrid.DataSource = rs
FillFormating 1
Exit Sub
ErrorHandler:
MsgBox Err.Description
End Sub
Sub init()
top = 0
left = 0
flexGrid.Rows = 1
FillFormating 1
End Sub

Private Sub Form_Load()
init
End Sub
Sub PrintRep()
On Error GoTo ErrorHandler
Dim sqlText As String
With cr1
    .Connect = ConnectName("")
    .ReportFileName = App.Path & "\Reports\RepCustomerCaseMessure.rpt"
     sqlText = "SELECT RepDate, adhamname, adhamphon, Notes, TeamName, RepPrice FROM t_CallsDate order by RegestDate  ,teamno"
    .SQLQuery = sqlText
    .DiscardSavedData = True
    
    .WindowState = crptMaximized
    .Action = 1
End With
Exit Sub
ErrorHandler:
MsgBox Err.Description
End Sub

Function ExecProcedure() As Boolean
On Error GoTo ErrorHandler
Dim sqlText As String
If IsDate(TxtDate.Text) Then
    sqlText = "Exec sp_GetCalls '" & ConvertControlDate(TxtDate.Text) & "','" & systemConfigration.DatabaseName & "'"
    de.con.Execute (sqlText)
Else
    MsgBox "«· «—ÌŒ €Ì— ’ÕÌÕ «Ê ·„ Ì „ ≈œŒ«· «· «—ÌŒ", vbExclamation, " «—ÌŒ €Ì— ’ÕÌÕ"
    ExecProcedure = False
    Exit Function
End If
ExecProcedure = True
Exit Function
ErrorHandler:
ExecProcedure = False
MsgBox Err.Description
End Function

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
    Case 1
        If ExecProcedure Then
            FillGrid (TxtDate.Text)
            Lcount.Caption = flexGrid.Rows - 1
        End If
    Case 2 ' PrintData
        If ExecProcedure Then
            PrintRep
        End If
    Case 4
        Unload Me
End Select

End Sub
