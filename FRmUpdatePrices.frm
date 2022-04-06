VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FRmUpdatePrices 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ⁄œÌ· «·«”⁄«—"
   ClientHeight    =   7305
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   10995
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   7305
   ScaleWidth      =   10995
   Begin Threed.SSFrame SSFrame1 
      Height          =   585
      Left            =   60
      TabIndex        =   10
      Top             =   1680
      Width           =   10905
      _ExtentX        =   19235
      _ExtentY        =   1032
      _Version        =   131074
      Begin VB.TextBox TxtSearch 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   6600
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   120
         Width           =   2655
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "»ÕÀ"
         Height          =   195
         Left            =   10530
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   90
         Width           =   285
      End
   End
   Begin VB.CheckBox ChkCloseTYpe 
      Alignment       =   1  'Right Justify
      Caption         =   "«· ﬁ—Ì» ··«œ‰Ï"
      Height          =   285
      Left            =   6750
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1350
      Width           =   1275
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   4545
      Left            =   30
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2310
      Width           =   10905
      _cx             =   19235
      _cy             =   8017
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
      ScrollTips      =   0   'False
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
   Begin VB.TextBox TxtClose 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   8070
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   1290
      Width           =   1215
   End
   Begin VB.TextBox TxtPercentage 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   8070
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   960
      Width           =   1215
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   510
      Top             =   780
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   38
      ImageHeight     =   38
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   39
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRmUpdatePrices.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRmUpdatePrices.frx":26D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRmUpdatePrices.frx":5568
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRmUpdatePrices.frx":83C5
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRmUpdatePrices.frx":AB6A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRmUpdatePrices.frx":D466
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRmUpdatePrices.frx":F98B
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRmUpdatePrices.frx":12143
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRmUpdatePrices.frx":14B57
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRmUpdatePrices.frx":174A9
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRmUpdatePrices.frx":1A21D
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRmUpdatePrices.frx":1CA6C
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRmUpdatePrices.frx":1F913
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRmUpdatePrices.frx":2266D
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRmUpdatePrices.frx":2500D
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRmUpdatePrices.frx":27F6F
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRmUpdatePrices.frx":2A997
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRmUpdatePrices.frx":2D351
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRmUpdatePrices.frx":2FC82
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRmUpdatePrices.frx":32588
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRmUpdatePrices.frx":34FF0
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRmUpdatePrices.frx":37FA1
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRmUpdatePrices.frx":3A8CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRmUpdatePrices.frx":3D1FF
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRmUpdatePrices.frx":3F93F
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRmUpdatePrices.frx":422AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRmUpdatePrices.frx":44B36
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRmUpdatePrices.frx":46D45
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRmUpdatePrices.frx":496A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRmUpdatePrices.frx":4BECC
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRmUpdatePrices.frx":4E980
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRmUpdatePrices.frx":513BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRmUpdatePrices.frx":53ED2
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRmUpdatePrices.frx":56E6C
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRmUpdatePrices.frx":59C38
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRmUpdatePrices.frx":5C8B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRmUpdatePrices.frx":5F744
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRmUpdatePrices.frx":62382
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRmUpdatePrices.frx":64F31
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   690
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   10995
      _ExtentX        =   19394
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
            Object.ToolTipText     =   "«· ⁄œÌ· ⁄·Ï «·‘«‘Â "
            ImageIndex      =   15
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "«· ⁄œÌ· ⁄·Ï ﬁ«⁄œÂ «·»Ì«‰« "
            ImageIndex      =   39
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   3
         EndProperty
      EndProperty
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   435
         Left            =   2190
         TabIndex        =   7
         Top             =   60
         Visible         =   0   'False
         Width           =   6045
         _ExtentX        =   10663
         _ExtentY        =   767
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
   End
   Begin VB.Label LCount 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   8430
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   6930
      Width           =   1725
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "⁄œœ «·„Ê«œ "
      Height          =   195
      Left            =   10260
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   6930
      Width           =   720
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "·«ﬁ—» "
      Height          =   195
      Left            =   10560
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   1320
      Width           =   420
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "«·‰”»Â"
      Height          =   195
      Left            =   10515
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   960
      Width           =   420
   End
   Begin VB.Menu mnuchk 
      Caption         =   "chk"
      Begin VB.Menu mnuselect 
         Caption         =   " ÕœÌœ"
      End
      Begin VB.Menu mnucancel 
         Caption         =   "«·€«¡ «· ÕœÌœ"
      End
   End
End
Attribute VB_Name = "FRmUpdatePrices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const ColChk = 1
Const ColStkId = 2
Const ColStkNo = 3
Const ColStkName = 4
Const ColStkCurrentPrice = 5
'Const ColStkOldPrice = 5
Const ColStkUpdatePrice = 6


Sub FillFormating(ByVal i As Integer, flexGrid As VSFlexGrid)
If i = 1 Then
    fs = "|>" + ""
    fs = fs + "|>" + "StkId"
    fs = fs + "|>" + "«·—ﬁ„ «·„Œ“‰Ì"
    fs = fs + "|>" + "«·‘—Õ"
    fs = fs + "|>" + "«·”⁄— «·Õ«·Ì"
'    Fs = Fs + "|>" + "«·”⁄— «·ﬁœÌ„"
    fs = fs + "|>" + "«·”⁄— «·„⁄œ·"
    With flexGrid
        .FormatString = fs
        .Cols = 7
        SetColWidths ColChk, flexGrid
        .ColDataType(ColChk) = flexDTBoolean
        .ColWidth(ColStkId) = 0
        SetColWidths ColStkNo, flexGrid
        SetColWidths ColStkName, flexGrid
        SetColWidths ColStkCurrentPrice, flexGrid
'        SetColWidths ColStkOldPrice, flexGrid
        SetColWidths ColStkUpdatePrice, flexGrid
    End With
End If
End Sub

Sub SetColWidths(ByVal ColNo As Integer, flexGrid As VSFlexGrid)
    With flexGrid
        .AutoSize (ColNo)
    End With
End Sub

Sub UpdateGrid()
On Error GoTo ErrorHandler

Dim Percentage As Double, ToClose As Integer, CloseType As Integer, SearchData As String
SearchData = TxtSearch.Text
If TxtPercentage.Text <> "" And TxtClose.Text <> "" Then
    Percentage = TxtPercentage.Text
    ToClose = TxtClose.Text
    CloseType = IIf(ChkCloseTYpe.Value = Checked, 1, 2)
'    Sqltext = "Select Id , StkNo , StkName , Cliprice , oldCliPrice , dbo.GetUpdatePrice(Id ," & Percentage & "," & ToClose & "," & CloseType & ")  CurrentPrice From CoStock Where CliPrice is not null"
    sqlText = "Select  chk , Id , StkNo , Ltrim(rtrim(StkName))StkName , Cliprice  , dbo.GetUpdatePrice(Id ," & Percentage & "," & ToClose & "," & CloseType & ")  CurrentPrice From CoStock Where CliPrice is not null"
    sqlText = sqlText & " and StkNo like " & LikeExpression(SearchData)
    Set rs = de.con.Execute(sqlText)
    Set Grid.DataSource = rs
    FillFormating 1, Grid
    LCount.Caption = Grid.Rows - 1
End If
Exit Sub
ErrorHandler:
MsgBox Err.Description
End Sub
Sub UpdateInDatabase()
On Error GoTo ErrorHandler

If MsgBox("Â· «‰  „ √ﬂœ „‰  ⁄œÌ· «·«”⁄«—«·„⁄·„Â", vbYesNo + vbDefaultButton2, " ⁄œÌ· «·«”⁄«—") = vbYes Then
    de.con.BeginTrans
    ProgressBar1.Min = 1
    ProgressBar1.Max = Grid.Rows
    ProgressBar1.Visible = True
    
    With Grid
        For i = 1 To .Rows - 1
                If Abs(Val(.TextMatrix(i, ColChk))) = 1 Then
                    .TextMatrix(i, ColChk) = 0
                    sqlText = "update CoStock set oldcliprice = CliPrice Where Id=" & .TextMatrix(i, ColStkId)
                    de.con.Execute (sqlText)
                    sqlText = "update CoStock set CliPrice = " & .TextMatrix(i, ColStkUpdatePrice) & "  Where id=" & .TextMatrix(i, ColStkId)
                    de.con.Execute (sqlText)
                End If
                ProgressBar1.Value = i
        Next
    ProgressBar1.Visible = False
    End With
    sqlText = "update CoStock set CHK=0 where ISNULL(CHK,0)=1"
    de.con.Execute (sqlText)
    de.con.CommitTrans
End If
Exit Sub
ErrorHandler:
   de.con.RollbackTrans
MsgBox Err.Description
End Sub

Private Sub ChkCloseTYpe_Click()
If ChkCloseTYpe.Value Then
    ChkCloseTYpe.Caption = "«· ﬁ—Ì» ··√⁄·Ï"
Else
    ChkCloseTYpe.Caption = "«· ﬁ—Ì» ··«œ‰Ï"
End If
UpdateGrid
End Sub

Private Sub Form_Load()
init
End Sub

Sub init()
On Error GoTo ErrorHandler
Dim rs As New ADODB.Recordset
top = 0
left = 0
'Sqltext = "Select Id , StkNo , StkName , Cliprice , oldCliPrice , null CurrentPrice From CoStock Where CliPrice is not null"
sqlText = "Select Chk,Id , StkNo , ltrim(rtrim(StkName)) as StkName , Cliprice  , null CurrentPrice From CoStock Where CliPrice is not null"
Set rs = de.con.Execute(sqlText)

Set Grid.DataSource = rs
FillFormating 1, Grid
LCount.Caption = rs.RecordCount
Grid.Editable = flexEDKbdMouse
Exit Sub
ErrorHandler:
MsgBox Err.Description
End Sub

Private Sub Grid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrorHandler
With Grid
    sqlText = "update costock set chk=" & Abs(.TextMatrix(Row, ColChk)) & " where id=" & .TextMatrix(Row, ColStkId)
    de.con.Execute (sqlText)
End With
Exit Sub
ErrorHandler:
MsgBox Err.Description
End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, cancel As Boolean)
If Col <> ColChk Then cancel = True
End Sub

Private Sub Grid_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrorHandler
    If Button And vbRightButton Then
        PopupMenu mnuchk
    End If
Exit Sub
ErrorHandler:
MsgBox Err.Description
End Sub
Private Sub mnuCancel_Click()
With Grid
    If .Row >= .RowSel Then
            FirstRow = .Row
            LastRow = .RowSel
        Else
            FirstRow = .RowSel
            LastRow = .Row
        End If
        For i = FirstRow To LastRow Step -1
            Vrow = i
            .TextMatrix(Vrow, ColChk) = 0
            sqlText = "update costock set chk=" & Abs(.TextMatrix(Vrow, ColChk)) & " where id=" & .TextMatrix(Vrow, ColStkId)
            de.con.Execute (sqlText)
        Next
End With
End Sub

Private Sub mnuSelect_Click()
With Grid
    If .Row >= .RowSel Then
            FirstRow = .Row
            LastRow = .RowSel
        Else
            FirstRow = .RowSel
            LastRow = .Row
        End If
        For i = FirstRow To LastRow Step -1
            Vrow = i
            .TextMatrix(Vrow, ColChk) = 1
            sqlText = "update costock set chk=" & Abs(.TextMatrix(Vrow, ColChk)) & " where id=" & .TextMatrix(Vrow, ColStkId)
            de.con.Execute (sqlText)
        Next
End With
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
    Case 1
        UpdateGrid
    Case 2
        UpdateInDatabase
    Case 4
        Unload Me
End Select
End Sub

Private Sub TxtClose_Change()
UpdateGrid
End Sub

Private Sub TxtPercentage_Change()
UpdateGrid
End Sub

Private Sub TxtSearch_Change()
UpdateGrid
End Sub
