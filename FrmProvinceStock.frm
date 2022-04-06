VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FrmProvinceStock 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ãÞÇÑäå ãÓÊæÏÚ ÇáãÍÇÝÙÇÊ  ÈÇÔÛÇÑÇÊ ÇáÇÓÊáÇã"
   ClientHeight    =   7425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13770
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   7425
   ScaleWidth      =   13770
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   3195
      Left            =   5280
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   2370
      Visible         =   0   'False
      Width           =   2445
      _cx             =   4313
      _cy             =   5636
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
   Begin Threed.SSFrame SSFrame1 
      Height          =   615
      Left            =   30
      TabIndex        =   3
      Top             =   6750
      Width           =   13695
      _ExtentX        =   24156
      _ExtentY        =   1085
      _Version        =   131074
      Begin VB.Label LSumQtyDiff 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   840
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   60
         Width           =   1425
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "ãÌãæÚ ÇáÝÑÞ Èíä ÇáÕÇáå æ ÇáãÚãá"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   2
         Left            =   2355
         TabIndex        =   11
         Top             =   90
         Width           =   2190
      End
      Begin VB.Label LSumQtyFactory 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   4710
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   90
         Width           =   1425
      End
      Begin VB.Label LSumQtyHall 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   8430
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   60
         Width           =   1425
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "ãÌãæÚ ßãíÇÊ ÇáãÚãá"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   1
         Left            =   6210
         TabIndex        =   7
         Top             =   90
         Width           =   1395
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "ãÌãæÚ ßãíÇÊ ÇáÕÇáå"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   0
         Left            =   9960
         TabIndex        =   6
         Top             =   90
         Width           =   1365
      End
      Begin VB.Label LCountHall 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   11640
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   90
         Width           =   1035
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "ÚÏÏ ÇáÓÌáÇÊ"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   10
         Left            =   12750
         TabIndex        =   4
         Top             =   60
         Width           =   855
      End
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   6660
      Top             =   3150
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1680
      Top             =   780
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
            Picture         =   "FrmProvinceStock.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProvinceStock.frx":26D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProvinceStock.frx":5533
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProvinceStock.frx":7CD8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProvinceStock.frx":A5D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProvinceStock.frx":CAF9
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProvinceStock.frx":F2B1
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProvinceStock.frx":11CC5
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProvinceStock.frx":14617
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProvinceStock.frx":1738B
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProvinceStock.frx":19BDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProvinceStock.frx":1CA81
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProvinceStock.frx":1F7DB
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProvinceStock.frx":2217B
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProvinceStock.frx":250DD
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProvinceStock.frx":27B05
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProvinceStock.frx":2A4BF
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProvinceStock.frx":2CDF0
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProvinceStock.frx":2F6F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProvinceStock.frx":3215E
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProvinceStock.frx":3510F
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProvinceStock.frx":37A38
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProvinceStock.frx":3A36D
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProvinceStock.frx":3CAAD
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProvinceStock.frx":3F41C
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProvinceStock.frx":41CA4
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProvinceStock.frx":43EB3
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProvinceStock.frx":46810
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProvinceStock.frx":4903A
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProvinceStock.frx":4BAEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProvinceStock.frx":4E52A
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProvinceStock.frx":51040
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProvinceStock.frx":53FDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProvinceStock.frx":56DA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProvinceStock.frx":59A20
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProvinceStock.frx":5C8B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProvinceStock.frx":5F4F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProvinceStock.frx":6209F
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
      Width           =   13770
      _ExtentX        =   24289
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
            Object.ToolTipText     =   "ÇÓÊÚÑÇÖ ÇáãáÝÇÊ"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   2
         EndProperty
      EndProperty
      Begin Threed.SSFrame SSFrame2 
         Height          =   615
         Left            =   9270
         TabIndex        =   13
         Top             =   0
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   1085
         _Version        =   131074
         Begin VB.TextBox TxtClientName 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0E0FF&
            Height          =   375
            Left            =   90
            RightToLeft     =   -1  'True
            TabIndex        =   0
            Top             =   150
            Width           =   3555
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "ÇáÕÇáå"
            Height          =   315
            Left            =   3660
            RightToLeft     =   -1  'True
            TabIndex        =   14
            Top             =   180
            Width           =   705
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   525
         Left            =   9300
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   90
         Width           =   2325
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid FlexStock 
      Height          =   6015
      Left            =   30
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   720
      Width           =   13695
      _cx             =   24156
      _cy             =   10610
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
End
Attribute VB_Name = "FrmProvinceStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Pos As Integer, RecNum   As Integer
Dim Ok As Boolean, Flag As Boolean

Const ColNo = 1
Const ColName = 2



Const ColNotificationNo = 1
Const ColStkNo = 2
Const ColStkName = 3

Const ColHallQty = 4
Const ColHallDate = 5

Const ColFactoryQty = 6
Const ColFactoryDate = 7

Const ColDiff = 8

Dim tracnsferStock As TreansferProvinceStockType



Sub ColorRow(Row As Integer, Color As Long)
With FlexHall
    For i = 1 To .Cols - 1
        .Col = i
        .Row = Row
        .CellBackColor = Color
    Next
End With
End Sub



Sub FillFormating(ByVal i As Integer, flexGrid As VSFlexGrid)

If i = 1 Then
    fs = "|>" + "ÑÞã ÇáÇÔÚÇÑ"
    fs = fs + "|>" + "ÇáÑÞã ÇáãÎÒäí"
    fs = fs + "|>" + "ÇáÔÑÍ"
    fs = fs + "|>" + "Çáßãíå ãä ÈÑäÇãÌ ÇáÕÇáå"
    fs = fs + "|>" + "ÇáÊÇÑíÎ ãä ÈÑäÇãÌ ÇáÕÇáå"
    fs = fs + "|>" + "Çáßãíå ãä ÈÑäÇãÌ ÇáãÚãá"
    fs = fs + "|>" + "ÇáÊÇÑíÎ ãä ÈÑäÇãÌ ÇáãÚãá"
    fs = fs + "|>" + "ÇáÝÑÞ"
   
    
    With flexGrid
        .FormatString = fs
        .Cols = 9
        SetColWidths ColNotificationNo, flexGrid
        SetColWidths ColStkNo, flexGrid

        SetColWidths ColStkName, flexGrid
        SetColWidths ColHallQty, flexGrid
        SetColWidths ColHallDate, flexGrid
        SetColWidths ColFactoryQty, flexGrid
        SetColWidths ColFactoryDate, flexGrid
        SetColWidths ColDiff, flexGrid


End With
ElseIf i = 2 Then
    fs = "|>" + ""
    fs = fs + "|>" + ""

    With Grid
        .FormatString = fs
        .Cols = 3
        .ColWidth(ColNo) = 0
        SetColWidths ColName, Grid
    End With
End If
End Sub

Sub SetColWidths(ByVal ColNo As Integer, flexGrid As VSFlexGrid)
    With flexGrid
        .AutoSize (ColNo)
    End With
End Sub


Sub init()
top = 0
left = 0
Ok = True
FlexStock.Rows = 1
FillFormating 1, FlexStock


End Sub

Sub SearchData(NotificationNo As Double)
On Error GoTo ErrorHandler

If NotificationNo = 0 Then
    MsgBox "áã íÊã ÇÏÎÇá ÑÞã ÇáÇÔÚÇÑ", vbCritical, "ÅÏÎÇá ÑÞã ÇáÇÔÚÇÑ"
    Exit Sub
End If


Dim i As Integer
With FlexHall

For i = 1 To .Rows - 1


    If NotificationNo = .TextMatrix(i, ColNotificationNo) Then
        
        .RowData(i) = 1
        ColorRow i, &HFFFFC0
    Else
         .RowData(i) = 0
        ColorRow i, vbWhite
    End If
Next
End With
Exit Sub
ErrorHandler:
MsgBox Err.Description




End Sub



Private Sub Form_Load()
init
End Sub

Function GetStkName(ByVal stkno As String) As String
On Error GoTo ErrorHandler
Dim rs As New ADODB.Recordset
sqlText = "Select StkName From CoStock Where StkNo = '" & stkno & "'"
Set rs = de.con.Execute(sqlText)
GetStkName = Trim(rs!StkName)
Exit Function
ErrorHandler:
GetStkName = ""
End Function
Function GetHallsQty(rs As ADODB.Recordset, NotificationNo As Double, stkno As String) As Double


On Error GoTo ErrorHandler

GetHallsQty = 0

rs.Filter = "NotificationNo=" & NotificationNo & " and PieceStockNo='" & stkno & "'"
If rs.RecordCount > 0 Then
    GetHallsQty = rs!Qty
End If
Exit Function
ErrorHandler:
GetHallsQty = 0
MsgBox Err.Description
End Function


Function GetFactoryQty(NotificationNo As Double, stkno As String) As Double


On Error GoTo ErrorHandler
Dim rs As New ADODB.Recordset
GetFactoryQty = 0

sqlText = "select isnull(m2.qty,0) Qty from MvMaintPayments m1 inner join  MvMaintPaymentsDetails m2 on m1.BillNo = m2.BillNo  where m1.billno = " & NotificationNo & "  and StkNo ='" & stkno & "'"
Set rs = de.con.Execute(sqlText)

If rs.RecordCount > 0 Then
    GetFactoryQty = rs!Qty
End If
Exit Function
ErrorHandler:
GetFactoryQty = 0
MsgBox Err.Description
End Function
Function GetFactoryDate(NotificationNo As Double, stkno As String) As Date


On Error GoTo ErrorHandler
Dim rs As New ADODB.Recordset
GetFactoryDate = vbNull

sqlText = "select Billdate  from MvMaintPayments m1 inner join  MvMaintPaymentsDetails m2 on m1.BillNo = m2.BillNo  where m1.billno = " & NotificationNo & "  and StkNo ='" & stkno & "'"
Set rs = de.con.Execute(sqlText)

If rs.RecordCount > 0 Then
    GetFactoryDate = rs!Billdate
End If
Exit Function
ErrorHandler:
GetFactoryDate = vbNull
MsgBox Err.Description
End Function
Sub FillProvinceInformation(flexGrid As VSFlexGrid, Vrow As Integer, rsProvinceStock As ADODB.Recordset)
If rsProvinceStock.RecordCount = 0 Then Exit Sub

With flexGrid
    rsProvinceStock.MoveFirst
    While Not rsProvinceStock.EOF
        If LTrim(RTrim(rsProvinceStock!NotificationNo)) = RTrim(LTrim(.TextMatrix(Vrow, ColNotificationNo))) _
            And LTrim(RTrim(rsProvinceStock!PieceStockNo)) = RTrim(LTrim(.TextMatrix(Vrow, ColStkNo))) Then
             
            .TextMatrix(Vrow, ColHallQty) = rsProvinceStock!Qty
            .TextMatrix(Vrow, ColHallDate) = rsProvinceStock!MvtDate
            Exit Sub
        End If
        rsProvinceStock.MoveNext
    Wend
End With
End Sub
Friend Sub FillGrid(transferProvinceStock As TreansferProvinceStockType, flexGrid As VSFlexGrid)
On Error GoTo ErrorHandler

Dim Vrow As Integer
flexGrid.Rows = 1
If transferProvinceStock.FactoryStock.RecordCount = 0 Then Exit Sub


transferProvinceStock.FactoryStock.MoveFirst


While Not transferProvinceStock.FactoryStock.EOF
    With flexGrid
        .AddItem ""
        Vrow = .Rows - 1
        .TextMatrix(Vrow, ColNotificationNo) = transferProvinceStock.FactoryStock!NotificationNo
        .TextMatrix(Vrow, ColStkNo) = transferProvinceStock.FactoryStock!PieceStockNo
        .TextMatrix(Vrow, ColStkName) = transferProvinceStock.FactoryStock!StkName
        .TextMatrix(Vrow, ColFactoryQty) = transferProvinceStock.FactoryStock!Qty
        FillProvinceInformation flexGrid, Vrow, transferProvinceStock.ProvinceStock
        
        .TextMatrix(Vrow, ColFactoryDate) = transferProvinceStock.FactoryStock!MvtDate
        .TextMatrix(Vrow, ColDiff) = Val(.TextMatrix(Vrow, ColFactoryQty)) - Val(.TextMatrix(Vrow, ColHallQty))
    
    
    
    End With
    transferProvinceStock.FactoryStock.MoveNext
Wend

Exit Sub
ErrorHandler:
MsgBox Err.Description
End Sub
 Friend Function GetTTl(flexGrid As VSFlexGrid) As TTlType
    Dim ttlrec As TTlType
    
    
    ttlrec.Count = flexGrid.Aggregate(flexSTCount, flexGrid.FixedRows, ColStkNo, flexGrid.Rows - 1, ColStkNo)
    ttlrec.HallQtySum = flexGrid.Aggregate(flexSTSum, flexGrid.FixedRows, ColHallQty, flexGrid.Rows - 1, ColHallQty)
    ttlrec.FactoryQtySum = flexGrid.Aggregate(flexSTSum, flexGrid.FixedRows, ColFactoryQty, flexGrid.Rows - 1, ColFactoryQty)
    GetTTl = ttlrec
 End Function

Function GetProvinceStock() As ADODB.Recordset
On Error GoTo ErrorHandler
Dim rs As New ADODB.Recordset
Set GetProvinceStock = rs
        CD.Filter = "*.txt"
        CD.ShowOpen
        
        rs.Open CD.FileName, , , , adCmdFile
Set GetProvinceStock = rs
Exit Function

ErrorHandler:
Set GetProvinceStock = rs
MsgBox Err.Description

End Function

Function GetFactoryStock(ClientNo As Double, rsProvinceStock As ADODB.Recordset) As ADODB.Recordset
On Error GoTo ErrorHandler
Dim rs As New ADODB.Recordset
Dim MinDate As Date, MaxDate As Date
Set GetFactoryStock = rs

If rsProvinceStock.RecordCount = 0 Then Exit Function

MinDate = Empty
MaxDate = Empty
rsProvinceStock.MoveFirst
MinDate = rsProvinceStock!MvtDate
MaxDate = rsProvinceStock!MvtDate

While Not rsProvinceStock.EOF

    If rsProvinceStock!MvtDate < MinDate Then MinDate = rsProvinceStock!MvtDate
    If rsProvinceStock!MvtDate > MaxDate Then MaxDate = rsProvinceStock!MvtDate
    rsProvinceStock.MoveNext
Wend

    sqlText = " Select m1.billno NotificationNo, m2.stkno PieceStockNo , Ltrim(rtrim(StkName))StkName ,  Qty , billdate MvtDate   from "
    sqlText = sqlText & "MvMaintPayments m1 inner join "
    sqlText = sqlText & " MvMaintPaymentsDetails m2 on m1.BillNo = m2.BillNo inner join "
    sqlText = sqlText & " costock c1 on m2.stkno = c1.stkno collate arabic_ci_as "
    sqlText = sqlText & " Where  m1.ClientId = " & ClientNo & " and BillDate >='" & ConvertControlDate(MinDate) & "' and BillDate<='" & ConvertControlDate(MaxDate) & "'"

    sqlText = sqlText & "  order by m2.stkno "
    Set rs = de.con.Execute(sqlText)
    Set GetFactoryStock = rs
Exit Function

ErrorHandler:
Set GetFactoryStock = rs
MsgBox Err.Description
End Function
Friend Function GetTransferProvinceStock(ClientNo As Double) As TreansferProvinceStockType
On Error GoTo ErrorHandler

Dim transferStock As TreansferProvinceStockType
GetTransferProvinceStock = transferStock
Set transferStock.ProvinceStock = GetProvinceStock()

Set transferStock.FactoryStock = GetFactoryStock(ClientNo, transferStock.ProvinceStock)

 GetTransferProvinceStock = transferStock
Exit Function
ErrorHandler:
 GetTransferProvinceStock = transferStock
End Function
Sub GetData()
On Error GoTo ErrorHandler
Dim treansferProvinceStock As TreansferProvinceStockType
Dim ttlrec As TTlType

If Val(TxtClientName.Tag) = 0 Then
    MsgBox "áã íÊã ÊÍÏíÏ ÇáÕÇáå", vbCritical, "ÊÍÏíÏ ÇáÕÇáå"
    Exit Sub
End If

treansferProvinceStock = GetTransferProvinceStock(TxtClientName.Tag)




FillGrid treansferProvinceStock, FlexStock
FillFormating 1, FlexStock
ttlrec = GetTTl(FlexStock)
LCountHall.Caption = ttlrec.Count
LSumQtyHall.Caption = ttlrec.HallQtySum
LSumQtyFactory.Caption = ttlrec.FactoryQtySum
LSumQtyDiff.Caption = ttlrec.HallQtySum - ttlrec.FactoryQtySum



Exit Sub
ErrorHandler:
MsgBox Err.Description
End Sub

Sub FilLabels(flexGrid As VSFlexGrid)


End Sub

Private Sub Grid_RowColChange()
On Error GoTo ErrorHandler
If Flag Then
    Ok = False
    With Grid
       Select Case Pos
        Case 2
            TxtClientName.Tag = .TextMatrix(.Row, ColNo)
            TxtClientName.Text = .TextMatrix(.Row, ColName)
       End Select
    End With
    Ok = True
End If
Exit Sub
ErrorHandler:
MsgBox Err.Description

End Sub



Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
    Case 1
        GetData
    Case 3
        Unload Me
End Select
End Sub
Sub ChangeCursor(ByVal X As Integer)
If X = 2 Then
    With TxtClientName
       Grid.top = SSFrame2.top + .top + .Height
       Grid.left = SSFrame2.left + .left
       Grid.Width = .Width
    End With
End If
End Sub
Sub MoveCursor(KeyCode As Integer, flexGrid As VSFlexGrid)

On Error Resume Next
If Not flexGrid.Visible Then Exit Sub
With flexGrid
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
Private Sub txtClientName_Change()
On Error GoTo ErrorHandler
Dim RsSearch As New ADODB.Recordset
If TxtClientName.Text = "" Then
    TxtClientName.Tag = 0
    Grid.Visible = False
    Exit Sub
End If
If Ok Then
    Flag = False
    sqlText = "Select Top 10 [ClientId] , [ClientName]    From ClientQry Where class=4 and (ClientName like" & LikeExpression(TxtClientName.Text) & " or ClientPhoneNBr like " & LikeExpression(TxtClientName.Text) & ")"
    Set RsSearch = de.con.Execute(sqlText)
    If RsSearch.RecordCount > 0 Then
        Set Grid.DataSource = RsSearch
        Grid.Row = 0
        FillFormating 2, Grid
        ChangeCursor 2
        Grid.Visible = True
    Else
        TxtClientName.Tag = 0
        Grid.Visible = False
    End If
    Flag = True
End If
Exit Sub
ErrorHandler:
MsgBox Err.Description
End Sub


Private Sub TxtClientName_GotFocus()
ChangeToArabic
Pos = 2
End Sub

Private Sub txtClientName_KeyDown(KeyCode As Integer, Shift As Integer)
    Flag = True
    If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then MoveCursor KeyCode, Grid
    Flag = True
End Sub

Private Sub txtClientName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Grid.Visible Then
        Ok = False
        TxtClientName.Tag = Grid.TextMatrix(IIf(Grid.Row = 0, 1, Grid.Row), ColNo)
        TxtClientName.Text = Grid.TextMatrix(IIf(Grid.Row = 0, 1, Grid.Row), ColName)
        Ok = True
    ElseIf Grid.Visible = False And TxtClientName.Text <> "" And Val(TxtClientName.Tag) <> 0 Then
        Exit Sub
    ElseIf Grid.Visible = False And TxtClientName.Text <> "" And Val(TxtClientName.Tag) = 0 Then
        Ok = False
        TxtClientName.Tag = 0
        Ok = True
    Else
        Ok = False
        TxtClientName.Tag = 0
        TxtClientName.Text = ""
        Ok = True
    End If
    Grid.Visible = False
End If

End Sub

