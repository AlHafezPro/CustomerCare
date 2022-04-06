VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmMvItem 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Õ—ﬂ… „«œ…"
   ClientHeight    =   6780
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11430
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   6780
   ScaleWidth      =   11430
   Begin Crystal.CrystalReport cr1 
      Left            =   1860
      Top             =   1050
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   5175
      Left            =   0
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1590
      Visible         =   0   'False
      Width           =   5715
      _cx             =   10081
      _cy             =   9128
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
      BackColor       =   16777215
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   16777215
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
      ScrollTrack     =   0   'False
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
      ExplorerBar     =   0
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
   Begin VSFlex8Ctl.VSFlexGrid FlexGrid 
      Height          =   5175
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1590
      Width           =   11415
      _cx             =   20135
      _cy             =   9128
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
   Begin Threed.SSFrame SSFrame1 
      Height          =   885
      Left            =   0
      TabIndex        =   2
      Top             =   690
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   1561
      _Version        =   131074
      Begin VB.TextBox TxtStrNo 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   7470
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   90
         Width           =   2805
      End
      Begin VB.TextBox TxtType 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   4200
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   90
         Width           =   1995
      End
      Begin VB.TextBox TxtStkNo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   8550
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   480
         Width           =   1725
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "—ﬁ„ „Œ“‰Ì"
         Height          =   285
         Left            =   10410
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   540
         Width           =   975
      End
      Begin Threed.SSCommand CmdSearch 
         Height          =   345
         Left            =   7050
         TabIndex        =   9
         Top             =   90
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   609
         _Version        =   131074
         PictureFrames   =   1
         Picture         =   "FrmMvItem.frx":0000
      End
      Begin Threed.SSCommand CmdType 
         Height          =   345
         Left            =   3750
         TabIndex        =   8
         Top             =   90
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   609
         _Version        =   131074
         PictureFrames   =   1
         Picture         =   "FrmMvItem.frx":0114
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "‰Ê⁄ «·Õ—ﬂ…"
         Height          =   195
         Index           =   2
         Left            =   6210
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   90
         Width           =   780
      End
      Begin VB.Label LStkName 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   4230
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   480
         Width           =   4305
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "«·„” Êœ⁄"
         Height          =   195
         Left            =   10680
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   60
         Width           =   630
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3600
      Top             =   90
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
            Picture         =   "FrmMvItem.frx":0228
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMvItem.frx":28FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMvItem.frx":575B
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMvItem.frx":7F00
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMvItem.frx":A7FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMvItem.frx":CD21
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMvItem.frx":F4D9
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMvItem.frx":11EED
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMvItem.frx":1483F
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMvItem.frx":175B3
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMvItem.frx":19E02
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMvItem.frx":1CCA9
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMvItem.frx":1FA03
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMvItem.frx":223A3
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMvItem.frx":25305
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMvItem.frx":27D2D
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMvItem.frx":2A6E7
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMvItem.frx":2D018
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMvItem.frx":2F91E
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMvItem.frx":32386
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMvItem.frx":35337
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMvItem.frx":37C60
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMvItem.frx":3A595
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMvItem.frx":3CCD5
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMvItem.frx":3F644
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMvItem.frx":41ECC
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMvItem.frx":440DB
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMvItem.frx":46A38
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMvItem.frx":49262
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMvItem.frx":4BD16
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMvItem.frx":4E752
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMvItem.frx":51268
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMvItem.frx":54202
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMvItem.frx":56FCE
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMvItem.frx":59C48
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMvItem.frx":5CADA
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMvItem.frx":5F718
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMvItem.frx":622C7
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
      Width           =   11430
      _ExtentX        =   20161
      _ExtentY        =   1217
      ButtonWidth     =   1191
      ButtonHeight    =   1164
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   37
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   15
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmMvItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const ColStkId = 1
Const ColStkNo = 2
Const ColStkName = 3
Const ColStrid = 4
Const ColStrNo = 5
Const ColStrName = 6
Const ColQtyIn = 7
Const ColQtyOut = 8
Const ColCurrentBalance = 9
Const ColFnlQnt = 10
Const ColMovDate = 11

Const ColStkId_1 = 1
Const ColStkno_1 = 2
Const ColStkname_1 = 3

Const colModNo = 1
Const ColModSymbol = 2
Const ColModName = 3

Dim Flag As Boolean, FGrid As Boolean

Sub PrintData()
With cr1
    .Connect = ConnectName("")
    .ReportFileName = App.Path + "\Reports\RepCoItemsMov.Rpt"
    sqlText = FillSqltext
    de.con.Execute (sqlText)
    .SQLQuery = "select id, byanid, StkId, StkNo, StkName, StrId, StrNo, StrName, QtyIn, QtyOut, CurrentBalance , doctype, qtytype, FnlQnt, Movdate, [in] ,[out],unitNo , UnitName ,Correspondence , Docnum , TypeName,qty  from t_StmovQry stmovQry order by StkNo , StrNo ,movdate,byanid "
    .DiscardSavedData = True
    .WindowState = crptMaximized
    .Action = 1
End With
End Sub

Sub FillList(sqlText As String, Field1 As String, Field2 As String, List As VSFlexGrid, ByVal Switch As Boolean)
    Set rs = de.con.Execute(sqlText)
    If rs.RecordCount > 0 Then
        Set List.DataSource = rs
        FillFormatVSFlex List, Switch
        List.Row = 1
        List.Col = 1
        List.ColSel = List.Cols - 1
        List.Visible = True
        TxtStkNo.SetFocus
    Else
        List.Text = ""
        List.Visible = False
        TxtStkNo.SetFocus
    End If
End Sub

Sub FillFormatVSFlex(FlexGrid As VSFlexGrid, ByVal Switch As Boolean)
If Switch Then
    fs = "|ModNo"
    fs = fs + "|<" + "—„“ «·„ÊœÌ·"
    fs = fs + "|<" + "≈”„ «·„ÊœÌ·"
    With FlexGrid
        .Visible = False
        .FormatString = fs
            .ColWidth(colModNo) = 0
            SetColWidths ColModSymbol, FlexGrid
            SetColWidths ColModName, FlexGrid
            .Visible = True
    End With
Else
    fs = "|ID"
    fs = fs + "|<" + "«·—ﬁ„ «·„Œ“‰Ì"
    fs = fs + "|<" + "«·≈”„"
    With FlexGrid
        .Visible = False
        .FormatString = fs
            .ColWidth(ColStkId) = 0
            SetColWidths ColStkno_1, FlexGrid
            SetColWidths ColStkname_1, FlexGrid
            .Visible = True
    End With
End If

 End Sub

Sub FillActiveControl(List As VSFlexGrid)
    With List
        If ActiveControl.Text <> "" Then
            If Not ActiveControl.DataChanged Then Exit Sub
            Flag = False
            ActiveControl.Text = IIf(.Visible = False, "", .TextMatrix(.Row, ColStkno_1))
            LStkName.Caption = IIf(.Visible = False, "", .TextMatrix(.Row, ColStkname_1))
            Flag = True
            ActiveControl.Tag = IIf(.Visible = False, "", .TextMatrix(.Row, ColStkId_1))
        Else
            ActiveControl.Text = ""
            ActiveControl.Tag = ""
            LStkName.Caption = ""
        End If
        List.Visible = False
        ActiveControl.DataChanged = False
    End With
End Sub
Sub MoveCursor(KeyCode As Integer)
On Error Resume Next
With Grid
    If KeyCode = vbKeyDown Then
        .Row = .Row + 1
    ElseIf KeyCode = vbKeyUp Then
        .Row = .Row - 1
    End If
If Not .RowIsVisible(.Row) Then
    .TopRow = .Row
End If
.Col = 0
.ColSel = .Cols - 1
End With
End Sub

Sub FillFormatString()
    fs = "|>" + "StkId"
    fs = fs + "|>" + "«·—ﬁ„"
    fs = fs + "|>" + "«·≈”„"
    fs = fs + "|>" + "StrId"
    fs = fs + "|>" + "«·„” Êœ⁄"
    fs = fs + "|>" + "«·≈”„"
    fs = fs + "|>" + "«·œ«Œ·"
    fs = fs + "|>" + "«·Œ«—Ã"
    fs = fs + "|>" + "«·—’Ìœ «··ÕŸÌ"

    fs = fs + "|>" + "«·—’Ìœ"
    fs = fs + "|>" + "«· «—ÌŒ"
    With FlexGrid
        .Cols = 12
        .FormatString = fs
        .ColWidth(ColStkId) = 0
        .ColWidth(ColStrid) = 0
        SetColWidths ColStkNo, FlexGrid
        SetColWidths ColStkName, FlexGrid
        SetColWidths ColStrNo, FlexGrid
        SetColWidths ColStrName, FlexGrid
        SetColWidths ColQtyIn, FlexGrid
        SetColWidths ColQtyOut, FlexGrid
        SetColWidths ColCurrentBalance, FlexGrid
        
        SetColWidths ColFnlQnt, FlexGrid
        SetColWidths ColMovDate, FlexGrid
   End With
End Sub
Sub SetColWidths(ColNo As Integer, FlexGrid As VSFlexGrid)
    Dim i, J, s, w
    With FlexGrid
            s = 0
            For i = 0 To .Rows - 1
                w = TextWidth(.TextMatrix(i, ColNo))
                If w > s Then s = w
            Next i
            .ColWidth(ColNo) = s + 100
    End With
End Sub

Function Fillitems() As String
Dim Str As String
Str = ""
With FrmChooseItems.Grid
    For i = 0 To .Rows - 1
        Str = Str + "," + .TextMatrix(i, 1)
    Next
End With
Fillitems = Mid(Str, 2)
End Function

Function FillSqltext() As String
Dim sqlText As String
'Sqltext = "Select StkId , StkNo , StkName , StrId , StrNo , StrName , Case When QtyType=0 then Qty else 0 end QtyIn , Case When QtyType=1 then Qty else 0 end QtyOut ,FnlQnt ,  Convert(Varchar(10),MovDate,102) Movdate , [in],[out] From StmovQry Where ByanId <> -1"
Dim StrNo As String


If TxtStrNo.Text <> "" Then
    StrNo = Replace(TxtStrNo.Text, "-", ",")
    'Sqltext = Sqltext & " And Strno in (" & Replace(TxtStrNo.Text, "-", ",") & ")"
Else
StrNo = Strids
    'Sqltext = Sqltext & " And Strid in (" & Strids & ")"
End If
If TxtType.Text <> "" Then
End If

'If ComboStr.BoundText <> "" Then
'    Sqltext = Sqltext & " And StrId=" & ComboStr.BoundText
'End If
Dim DocTypeStr As String
If TxtType.Text <> "" Then
   DocTypeStr = Replace(TxtType.Text, "-", ",")
   ' Sqltext = Sqltext & " And DocType in(" & Replace(TxtType.Text, "-", ",") & ")"
End If

Dim stkId As Double
stkId = TxtStkNo.Tag
'If TxtStkNo.Text <> "" Then
'    If ChkMod_Stk.Value Then
'        Sqltext = Sqltext & " And StkId =" & TxtStkNo.Tag & " and StkidTYpe=1"
'    Else
'        Sqltext = Sqltext & " And StkId =" & TxtStkNo.Tag & " and StkidTYpe=0"
'    End If
'End If

'If chkChoose.Value Then
'    sTRItems = Fillitems
'    If LTrim(RTrim(sTRItems)) <> "" Then
'        Sqltext = Sqltext & " And stkid in(" & sTRItems & ")"
'    End If
'End If
'If ComboByanType.BoundText <> "" Then
'    Sqltext = Sqltext & " And DocType=" & ComboByanType.BoundText
'End If
'Sqltext = Sqltext & " Order by byanid"
'FillSqltext = Sqltext

sqlText = "exec sp_items_Mov " & stkId & ",'" & StrNo & "','" & DocTypeStr & "'"
FillSqltext = sqlText
End Function

Sub SearchRec()
On Error GoTo ErrorHandler
Dim rs As New ADODB.Recordset
Dim sqlText As String
sqlText = FillSqltext
de.con.Execute (sqlText)
sqlText = "select StkId , StkNo , StkName  , StrId , StrNo , StrName  , QtyIn , QtyOut ,CurrentBalance , FnlQnt ,  Movdate  from t_stmovQry Order BY id,movdate,byanid"
Set rs = de.con.Execute(sqlText)
Set FlexGrid.DataSource = rs
FillFormatString
Exit Sub
ErrorHandler:
MsgBox Err.Description
End Sub
Sub init()
'    Dim RSStr As New ADODB.Recordset
'    Sqltext = "Select Id , StrNo , StrName From NameStr where id in(" & Strids & ") Order By StrNo"
'    Set RSStr = de.con.Execute(Sqltext)
'    If RSStr.RecordCount > 0 Then
'        Set ComboStr.RowSource = RSStr
'        ComboStr.ListField = "StrName"
'        ComboStr.BoundColumn = "Id"
'        RSStr.MoveFirst
'        ComboStr.BoundText = RSStr!Id
'        Dim RsByanType As New ADODB.Recordset
'        Sqltext = "Select TypeId , TypeName  From CoByanType  Order BY TypeId "
'        Set RsByanType = de.con.Execute(Sqltext)
'        Set ComboByanType.RowSource = RsByanType
'        ComboByanType.ListField = "TypeName"
'        ComboByanType.BoundColumn = "TypeId"
'    End If
FGrid = False
FillFormatString
FlexGrid.Rows = 1
top = 0
left = 0
FGrid = True
End Sub

'Private Sub chkChoose_Click(Value As Integer)
'If chkChoose.Value Then
'    FrmChooseItems.Show 1
'End If
'End Sub

Private Sub ChkMod_Stk_Click()
If ChkMod_Stk.Value Then
    ChkMod_Stk.Caption = "„ÊœÌ·"
Else
    ChkMod_Stk.Caption = "—ﬁ„ „Œ“‰Ì"
End If
End Sub

Private Sub CmdSearch_Click()
FrmChoose.Show 1
TxtStrNo = StrNo
End Sub

Private Sub CmdType_Click()
FrmChooseTypes.Show 1
TxtType.Text = ByanType
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{Tab}"
    SendKeys "{Home}+{End}"
End If
End Sub

Private Sub Form_Load()
init
End Sub

'Private Sub Grid_RowColChange()
'If FGrid Then
'    With Grid
'        If ChkMod_Stk.Value Then
'            Flag = False
'            TxtStkNo.Tag = .TextMatrix(.Row, colModNo)
'            TxtStkNo.Text = .TextMatrix(.Row, ColModSymbol)
'            LStkName.Caption = .TextMatrix(.Row, ColModName)
'            Flag = True
'        End If
'    End With
'End If
'End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
    Case 1
        PrintData
    Case 3
        SearchRec
    Case 5
       Unload Me
End Select
End Sub

Private Sub TxtStkNo_Change()
If Flag Then
    Dim sqlText As String
    If Trim(TxtStkNo.Text) = "" Then
        TxtStkNo.Tag = ""
        Grid.Visible = False
        Exit Sub
    End If
'    If ChkMod_Stk.Value Then
'        FGrid = False
'        Sqltext = "select Top 15 ModNo Id , Symbol StkNo , Name  StkName from dbo.models where Symbol like " & LikeExpression(TxtStkNo.Text) & " Or Name Like " & LikeExpression(TxtStkNo.Text)
'        FillList Sqltext, "Id", "StkNo", Grid, ChkMod_Stk.Value
'        FGrid = True
'    Else
        sqlText = "Select top 15 Id , StkNo , StkName From CoStock  where StkNo like " & LikeExpression(TxtStkNo.Text) & " Order By len(ltrim(rtrim(StkNo))) , ltrim(rtrim(StkNo))"
        FillList sqlText, "Id", "StkNo", Grid, 0
'    End If
End If
End Sub

Private Sub TxtStkNo_KeyDown(KeyCode As Integer, Shift As Integer)
    Flag = False
    If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then MoveCursor KeyCode
    Flag = True
End Sub
Function DataOk(stkno As String) As Boolean
Dim RsFind As New ADODB.Recordset
sqlText = "Select StkNo From CoStock Where StkNo ='" & stkno & "'"
Set RsFind = de.con.Execute(sqlText)
If RsFind.RecordCount > 0 Then
   DataOk = True
Else
    DataOk = False
End If
End Function
Function GetStkId(stkno As String, Strid As String) As Double
    Dim rs As New ADODB.Recordset
If Strid <> "" Then
    sqlText = "Select c1.Id From CoStock c1 inner join Stkinf s1 on c1.Id = s1.StkId  And Strno in (" & Strid & ") Where c1.StkNo ='" & stkno & "'"
Else
    sqlText = "Select c1.Id From CoStock c1 inner join Stkinf s1 on c1.Id = s1.StkId Where c1.StkNo ='" & stkno & "'"
End If
    Set rs = de.con.Execute(sqlText)
    If rs.RecordCount > 0 Then
        GetStkId = rs!Id
    Else
        GetStkId = 0
    End If
End Function
Function GetStkName(stkId As Double) As String
Dim RsStkName As New ADODB.Recordset
    sqlText = "Select StkName From CoStock Where Id=" & stkId
    Set RsStkName = de.con.Execute(sqlText)
    If RsStkName.RecordCount > 0 Then
        GetStkName = RsStkName!StkName
    Else
        GetStkName = ""
    End If
End Function
Function GetBalance(stkId As Double, Strid As Double) As Double
Dim RsBalance As New ADODB.Recordset
    sqlText = "Select FnlQnt From Stkinf Where Stkid=" & stkId & " And StrId=" & Strid
    Set RsBalance = de.con.Execute(sqlText)
    If RsBalance.RecordCount > 0 Then GetBalance = IIf(IsNull(RsBalance!fnlqnt), 0, RsBalance!fnlqnt)
End Function
Private Sub txtStkNo_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'            FillActiveControl Grid
'            SendKeys "{home}+{end}"
'    End If
On Error GoTo ErrorHandler
If KeyAscii = 13 Then
'        If ChkMod_Stk.Value Then
'            FGrid = False
'            TxtStkNo.Tag = Grid.TextMatrix(Grid.Row, colModNo)
'            TxtStkNo.Text = Grid.TextMatrix(Grid.Row, ColModSymbol)
'            LStkName.Caption = Grid.TextMatrix(Grid.Row, ColModName)
'            FGrid = True
'            Grid.Visible = False
'        Else
            Grid.Visible = False
            If DataOk(TxtStkNo.Text) Then
                TxtStkNo.Tag = GetStkId(TxtStkNo.Text, Replace(TxtStrNo.Text, "-", ","))
                If TxtStkNo.Tag <> 0 Then
                    LStkName.Caption = GetStkName(TxtStkNo.Tag)
                Else
                    LStkName.Caption = ""
                    MsgBox "«·„«œ… €Ì— „⁄—›… ›Ì «·„” Êœ⁄", vbInformation, " ‰»ÌÂ"
                    SendKeys "{home}+{end}"
                    Exit Sub
                End If
            Else
                LStkName.Caption = ""
                SendKeys "{home}+{end}"
                MsgBox "«·„«œ… €Ì— „ÊÃÊœ… ÷„‰ ﬁ«∆„… «·√—ﬁ«„ «·„Œ“‰Ì…", vbExclamation, " ‰»ÌÂ"
                Exit Sub
            End If
'        End If
    End If
Exit Sub
ErrorHandler:
Grid.Visible = False
MsgBox Err.Description
End Sub

