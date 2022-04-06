VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmPrintByans 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ÿ»«⁄… «·»Ì«‰« "
   ClientHeight    =   8625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9735
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   8625
   ScaleWidth      =   9735
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   5175
      Left            =   4530
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   3060
      Visible         =   0   'False
      Width           =   4245
      _cx             =   7488
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
   Begin VB.TextBox TxtStkNo 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   4590
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   2700
      Width           =   4185
   End
   Begin VB.TextBox Txtrecite 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   7200
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   2310
      Width           =   1575
   End
   Begin VB.TextBox TxtType 
      Alignment       =   1  'Right Justify
      Height          =   330
      Left            =   5040
      Locked          =   -1  'True
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   1500
      Width           =   3735
   End
   Begin VSFlex8Ctl.VSFlexGrid FlexGrid 
      Height          =   5145
      Left            =   60
      TabIndex        =   6
      Top             =   3420
      Width           =   9675
      _cx             =   17066
      _cy             =   9075
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
   Begin VB.TextBox TxtByanNo 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   7200
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   1920
      Width           =   1575
   End
   Begin Crystal.CrystalReport cr1 
      Left            =   390
      Top             =   930
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   330
      Top             =   1410
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
            Picture         =   "FrmPrintByans.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrintByans.frx":26D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrintByans.frx":5533
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrintByans.frx":7CD8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrintByans.frx":A5D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrintByans.frx":CAF9
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrintByans.frx":F2B1
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrintByans.frx":11CC5
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrintByans.frx":14617
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrintByans.frx":1738B
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrintByans.frx":19BDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrintByans.frx":1CA81
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrintByans.frx":1F7DB
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrintByans.frx":2217B
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrintByans.frx":250DD
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrintByans.frx":27B05
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrintByans.frx":2A4BF
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrintByans.frx":2CDF0
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrintByans.frx":2F6F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrintByans.frx":3215E
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrintByans.frx":3510F
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrintByans.frx":37A38
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrintByans.frx":3A36D
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrintByans.frx":3CAAD
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrintByans.frx":3F41C
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrintByans.frx":41CA4
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrintByans.frx":43EB3
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrintByans.frx":46810
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrintByans.frx":4903A
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrintByans.frx":4BAEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrintByans.frx":4E52A
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrintByans.frx":51040
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrintByans.frx":53FDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrintByans.frx":56DA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrintByans.frx":59A20
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrintByans.frx":5C8B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrintByans.frx":5F4F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrintByans.frx":6209F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   690
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   1217
      ButtonWidth     =   1191
      ButtonHeight    =   1164
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   37
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Tag             =   "1"
                  Text            =   "ÿ»«⁄… «·Ê«ÃÂ…"
               EndProperty
            EndProperty
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
            ImageIndex      =   34
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin MSDataListLib.DataCombo ComboStr 
      Height          =   315
      Left            =   5040
      TabIndex        =   2
      Top             =   1140
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSMask.MaskEdBox TxtFromDate 
      Height          =   315
      Left            =   7680
      TabIndex        =   0
      Top             =   720
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "99/99/9999"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox TxtToDate 
      Height          =   315
      Left            =   5070
      TabIndex        =   1
      Top             =   720
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "99/99/9999"
      PromptChar      =   "_"
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
      Height          =   315
      Left            =   30
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   2730
      Width           =   4515
   End
   Begin VB.Label LStkNo 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "—ﬁ„ «·„«œÂ"
      Height          =   195
      Left            =   9015
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   2760
      Width           =   675
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "—ﬁ„ «·≈‘⁄«—"
      Height          =   195
      Index           =   5
      Left            =   8910
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   2340
      Width           =   780
   End
   Begin Threed.SSCommand CmdType 
      Height          =   345
      Left            =   4590
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1500
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   609
      _Version        =   131074
      PictureFrames   =   1
      Picture         =   "FrmPrintByans.frx":64A4E
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "≈·Ï  «—ÌŒ"
      Height          =   195
      Index           =   4
      Left            =   6360
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   750
      Width           =   675
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "—ﬁ„ «·»Ì«‰"
      Height          =   195
      Index           =   3
      Left            =   9030
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   2010
      Width           =   660
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "‰Ê⁄ «·Õ—ﬂ…"
      Height          =   195
      Index           =   2
      Left            =   8940
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   1590
      Width           =   750
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "«·„” Êœ⁄"
      Height          =   195
      Index           =   1
      Left            =   9060
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   1170
      Width           =   630
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   " „‰  «—ÌŒ "
      Height          =   195
      Index           =   0
      Left            =   9000
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   750
      Width           =   690
   End
   Begin VB.Menu mnufile 
      Caption         =   "file"
      Visible         =   0   'False
      Begin VB.Menu printOrder 
         Caption         =   "ÿ»«⁄… √„— ’—›"
      End
   End
End
Attribute VB_Name = "FrmPrintByans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const ColChk = 1
Const ColByanId = 2
Const ColStrid = 3
Const ColStrNo = 4
Const ColDocNum = 5
Const ColByanDate = 6

Const ColDocCount = 7
Const ColDocType = 8
Const ColTypeName = 9


Const ColStkId_1 = 1
Const ColStkno_1 = 2
Const ColStkname_1 = 3

Const ColId = 1
Const ColNo = 2
Const ColName = 3



Dim ChkCorrespondence As Boolean, Ok As Boolean, Flag As Boolean
Dim Pos As Integer
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

Sub FillGridCombos(flexGrid As VSFlexGrid, ByVal Col As Integer)
Dim RsCountry As New ADODB.Recordset
Dim Lst As String
    sqlText = "Select CountryNo, CountryName  From Hafez2000dev.dbo.iocoCountry "
    Set RsCountry = de.con.Execute(sqlText)
    If RsCountry.RecordCount > 0 Then
        With flexGrid
            Lst = .BuildComboList(RsCountry, "CountryName", "CountryNo", vbYellow)
            .ColComboList(Col) = Lst
        End With
    Else
        With flexGrid
            .Rows = 1
        End With
    End If
End Sub

Function ViewCorrespondenceCol(ByVal empNo As Integer) As Boolean
On Error GoTo ErrorHandler
Dim RsCorrespondence As New ADODB.Recordset
sqlText = "Select ChkCorrespondence From StockUsers Where empno =" & empNo
Set RsCorrespondence = de.con.Execute(sqlText)
If RsCorrespondence.RecordCount > 0 Then
    ViewCorrespondenceCol = RsCorrespondence!ChkCorrespondence
Else
    ViewCorrespondenceCol = False
End If
Exit Function
ErrorHandler:
ViewCorrespondenceCol = False
MsgBox Err.Description
End Function

Sub FillCombos()
Dim RsSTr As New ADODB.Recordset
sqlText = "Select Id , StrNo , StrName From NameStr order by strno "
Set RsSTr = de.con.Execute(sqlText)
If RsSTr.RecordCount > 0 Then
    Set ComboStr.RowSource = RsSTr
    ComboStr.listField = "StrName"
    ComboStr.BoundColumn = "Id"
    ComboStr.BoundText = SelectedStr
End If


'Dim RsCountry  As New ADODB.Recordset
'Sqltext = "Select CountryNo, CountryName  From Hafez2000dev.dbo.iocoCountry "
'Set RsCountry = de.con.Execute(Sqltext)
'
'
'Set ComboCountry.RowSource = RsCountry
'
'ComboCountry.ListField = "CountryName"
'ComboCountry.BoundColumn = "CountryNo"
End Sub

Sub FillFormatString()
    fs = "|>" + " "
    fs = fs + "|>" + "—ﬁ„ «·»Ì«‰"
    fs = fs + "|>" + "Strid"
    fs = fs + "|>" + "—ﬁ„ «·„” Êœ⁄"
    fs = fs + "|>" + "—ﬁ„ «·Õ—ﬂ…"
    fs = fs + "|>" + " «—ÌŒ «·»Ì«‰"
    fs = fs + "|>" + "⁄œœ «·Õ—ﬂ« "
    fs = fs + "|>" + "‰Ê⁄ «·Õ—ﬂ…"
    fs = fs + "|>" + "‰Ê⁄ «·Õ—ﬂ…"
    With flexGrid
        .FormatString = fs
        .Cols = 10
        .ColWidth(ColChk) = 300
        
        SetColWidths ColByanId, flexGrid
        .ColWidth(ColStrid) = 0
            SetColWidths ColStrNo, flexGrid
        SetColWidths ColDocNum, flexGrid
        SetColWidths ColByanDate, flexGrid
        .ColWidth(ColDocType) = 0
        SetColWidths ColTypeName, flexGrid
        SetColWidths ColDocCount, flexGrid
   End With
End Sub

Sub SetColWidths(ColNo As Integer, flexGrid As VSFlexGrid)
    Dim i, J, s, w
    With flexGrid
    .AutoSize (ColNo)
'            s = 0
'            For i = 0 To .Rows - 1
'                w = TextWidth(.TextMatrix(i, ColNo))
'                If w > s Then s = w
'            Next i
'            .ColWidth(ColNo) = s + 100
    End With
End Sub

Sub init()
  '  ChkCorrespondence = ViewCorrespondenceCol(EmpNo)
    top = 0
    left = 0
    Flag = False
    Ok = True
    Dim rsInit As New ADODB.Recordset
    FillCombos
'    top = (Screen.Width / 2) - Me.Width
'    left = (Screen.Height / 2)
    TxtFromDate.Text = Format(Date, "dd/mm/yyyy")
    TxtToDate.Text = "31/12/" & Trim(Str(Year(Date)))
    flexGrid.Editable = flexEDKbdMouse
'    Set Rsinit = de.con.Execute(SearchRec)
'    Set FlexGrid.DataSource = Rsinit
    flexGrid.Rows = 1
    FillFormatString
    flexGrid.ColDataType(ColChk) = flexDTBoolean
    FillFormatString
    flexGrid.Editable = flexEDKbdMouse
    Flag = True
    
End Sub

Function SearchRec(Optional i) As String
On Error GoTo ErrorHandler
Dim sqlText As String

If IsMissing(i) Then
    sqlText = "select 0 Chk , ByanId , s1.StrId , StrNo , DocNum , Convert(Varchar(10),MovDate,103)MovDate , Count(*) , DocType , TypeName  from StmovQry s1 Where ByanId <> 0"
Else
    sqlText = "Select StmovQry.Id , StmovQry.ByanId , StmovQry.Stkid  , StmovQry.StkNo , StmovQry.StkName , StmovQry.StrId , StmovQry.StrNo , StmovQry.StrName , StmovQry.MovDate , StmovQry.DocType , StmovQry.Qty , StmovQry.QtyType , StmovQry.FnlQnt ,  StmovQry.Correspondence, CountryNo   From StmovQry Where  StmovQry.ByanId in  (" & ByansSelect(1) & ") AND StmovQry.STRID IN (" & ByansSelect(2) & ")"
End If

If IsDate(TxtFromDate.Text) Then
    sqlText = sqlText & " and MOVDATE >='" & ConvertControlDate(TxtFromDate.Text) & "'"
End If

If IsDate(TxtToDate.Text) Then
    sqlText = sqlText & " and MOVDATE<='" & ConvertControlDate(TxtToDate.Text) & "'"
End If
If ComboStr.BoundText <> "" Then
    sqlText = sqlText & " And StrId=" & ComboStr.BoundText
Else
    sqlText = sqlText & " And StrId <>-1"
End If
If TxtType.Text <> "" Then
    sqlText = sqlText & " and DocType in (" & Replace(TxtType.Text, "-", ",") & ")"
End If
'If ComboByanType.BoundText <> "" Then
'    Sqltext = Sqltext & " and DocType=" & ComboByanType.BoundText
'End If
If IsNumeric(TxtByanNo.Text) Then
    sqlText = sqlText & " and ByanId =" & TxtByanNo.Text
End If
If IsNumeric(Txtrecite.Text) Then
    sqlText = sqlText & " and docnum =" & Txtrecite.Text
End If
If IsNumeric(TxtStkNo.Tag) Then
    sqlText = sqlText & " and stkid=" & TxtStkNo.Tag
End If


If IsMissing(i) Then
    sqlText = sqlText & " Group By ByanId ,s1.StrId , StrNo , DocNum , MovDate , docTYpe , TypeName  Order by MovDate"
Else
    sqlText = sqlText & " Order by Byanid , StrNo , ltrim(rtrim(StkNo))"
End If
SearchRec = sqlText
Exit Function
ErrorHandler:
MsgBox Err.Description

End Function

Private Sub CmdType_Click()
FrmChooseTypes.Show 1
TxtType.Text = ByanType
End Sub

Private Sub FlexGrid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
With flexGrid
'    Sqltext = "Update Stmov Set Correspondence='" & .TextMatrix(Row, colCorrespondence) & "',DocNum=" & Val(.TextMatrix(Row, ColDocNum)) & " Where ByanId=" & .TextMatrix(Row, ColByanId)
   sqlText = "Update Stmov Set Correspondence='" & .TextMatrix(Row, colCorrespondence) & "',CountryNo=" & Val(.Cell(flexTextFlat, Row, ColCountryNo, Row, ColCountryNo)) & " Where ByanId=" & .TextMatrix(Row, ColByanId)
    de.con.Execute (sqlText)
End With
End Sub

Private Sub FlexGrid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, cancel As Boolean)
'If Col <> ColChk And Col <> colCorrespondence And Col <> ColDocNum Then Cancel = True
If Col <> ColChk And Col <> colCorrespondence And Col <> ColCountryNo Then cancel = True
End Sub

Private Sub FlexGrid_DblClick()
If FormOk Then
    With flexGrid
        GByanId = .TextMatrix(.Row, ColByanId)
        Unload Me
    End With
End If
End Sub

Private Sub FlexGrid_KeyDown(KeyCode As Integer, Shift As Integer)
With flexGrid
    If KeyCode = vbKeyEscape And .Col = ColCountryNo Then
        .TextMatrix(.Row, ColCountryNo) = ""
        sqlText = "Update Stmov Set CountryNo=" & Val(.Cell(flexTextFlat, .Row, ColCountryNo, .Row, ColCountryNo)) & " Where ByanId=" & .TextMatrix(.Row, ColByanId)
        de.con.Execute (sqlText)
    End If
End With
End Sub

Private Sub FlexGrid_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrorHandler
    If flexGrid.TextMatrix(flexGrid.Row, ColDocType) <> 6 Then Exit Sub
    If Button And vbRightButton Then
        PopupMenu mnufile
    End If
Exit Sub
ErrorHandler:
MsgBox Err.Description
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape And ActiveControl.name <> "FlexGrid" Then
    GByanId = 0
    Unload Me
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If ActiveControl.name <> "FlexGrid" Then
        SendKeys "{tab}"
        SendKeys "{home}+{End}"
    End If
End If
End Sub

Private Sub Form_Load()
init
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
Function ByansSelect(Vindex As Integer) As String
Dim ByanStr As String
ByanStr = ""
With flexGrid
    For i = 1 To .Rows - 1
        If .TextMatrix(i, ColChk) = -1 Then
            Select Case Vindex
                Case 1 ' BYNID
                    ByanStr = ByanStr & "," & .TextMatrix(i, ColByanId)
                Case 2 ' STR
                     ByanStr = ByanStr & "," & .TextMatrix(i, ColStrid)
            End Select
        End If
    Next
End With
If ByanStr = "" Then
    ByansSelect = "0"
Else
    ByansSelect = Mid(ByanStr, 2)
End If
End Function


Sub PrintData(i As Integer)
On Error GoTo ErrorHandler
With cr1
    .Connect = ConnectName("")
    Select Case i
        Case 1:
            .ReportFileName = App.Path + "\Reports\ByanRep.rpt"
            .SQLQuery = SearchRec(1)
        Case 2:
            .ReportFileName = App.Path + "\Reports\RepInterface.rpt"
            .SQLQuery = SearchRec
    End Select
    .DiscardSavedData = True
    .WindowState = crptMaximized
    .Action = 1
End With
Exit Sub
ErrorHandler:
MsgBox Err.Description
End Sub

Private Sub Grid_RowColChange()
If Flag Then
    Ok = False
    With Grid
       Select Case Pos
        Case 3
            TxtStkNo.Tag = .TextMatrix(.Row, ColId)
            TxtStkNo.Text = .TextMatrix(.Row, ColNo)
            LStkName.Caption = .TextMatrix(.Row, ColName)
       End Select
    End With
    Ok = True
End If

End Sub

Private Sub printOrder_Click()
With cr1
    sqlText = "Select OrderId, OrderDate, ByanId, Model, WsNum, WsName, StrNo, StrName, StkNo, StkName, Qty, Amount, Multi, Product, Balance, AllBalance,PrevOrderId , PrevMulti, PrevOrderDate, PrevProduct FROM  MvPaymentOrderQry Where  OrderId =" & flexGrid.TextMatrix(flexGrid.Row, ColDocNum)
    .SQLQuery = sqlText
    .ReportFileName = App.Path & "\Reports\RepPaymentOrder_3.rpt"
    .Connect = ConnectName("")
    .DiscardSavedData = True
    .WindowState = crptMaximized
    .Action = 1
End With
End Sub

Function ExcelSearch() As String
On Error GoTo ErrorHandler

sqlText = "Select StmovQry.Id , StmovQry.ByanId , StmovQry.Stkid  , StmovQry.StkNo , StmovQry.StkName , StmovQry.StrId , StmovQry.StrNo , StmovQry.StrName , StmovQry.MovDate , StmovQry.DocNum ,  StmovQry.DocType , StmovQry.Qty , StmovQry.QtyType , StmovQry.FnlQnt   From StmovQry  Where  StmovQry.ByanId in  (" & ByansSelect(1) & ") AND StmovQry.STRID IN (" & ByansSelect(2) & ")"
If IsDate(TxtFromDate.Text) Then
    sqlText = sqlText & " and Convert(Varchar(10),movdate,101) >='" & Format(TxtFromDate.Text, "mm/dd/yyyy") & "'"
End If

If IsDate(TxtToDate.Text) Then
    sqlText = sqlText & " and Convert(Varchar(10),movdate,101) <='" & Format(TxtToDate.Text, "mm/dd/yyyy") & "'"
End If
If ComboStr.BoundText <> "" Then
    sqlText = sqlText & " And StrId=" & ComboStr.BoundText
Else
    sqlText = sqlText & " And StrId IN ( " & Strids & ")"
End If
If TxtType.Text <> "" Then
    sqlText = sqlText & " and DocType in (" & Replace(TxtType.Text, "-", ",") & ")"
End If
If IsNumeric(TxtByanNo.Text) Then
    sqlText = sqlText & " and ByanId =" & TxtByanNo.Text
End If

If IsNumeric(Txtrecite.Text) Then
    sqlText = sqlText & " and docnum =" & Txtrecite.Text
End If

ExcelSearch = sqlText
Exit Function
ErrorHandler:
ExcelSearch = ""
End Function

Private Sub ExportToExcel()

Dim rs As New ADODB.Recordset

Dim objXL As Excel.Application
Dim objWB As Excel.Workbook
Dim objWS As Excel.Worksheet
Dim r As Long
Dim c As Long
Set objXL = New Excel.Application
Set objWB = objXL.Workbooks.Add
Set objWS = objWB.Worksheets(1)


Set rs = de.con.Execute(ExcelSearch)
If rs.RecordCount = 0 Then Exit Sub
With objWS
With rs
    For c = 0 To rs.fields.Count - 1
        objWS.Cells(1, c + 1) = rs.fields(c).name
    Next
    rs.MoveFirst
    For r = 0 To rs.RecordCount - 1
        For c = 0 To rs.fields.Count - 1
            objWS.Cells(r + 2, c + 1) = rs.fields(c)
        Next
        rs.MoveNext
    Next
End With
'.Cells.Columns.AutoFit
End With
objXL.Visible = True
Set objWS = Nothing
Set objWB = Nothing
Set objXL = Nothing
End Sub



Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim RsSearch  As New ADODB.Recordset
Select Case Button.Index
    Case 1
        PrintData 1
    Case 3
        Set RsSearch = de.con.Execute(SearchRec)
        Set flexGrid.DataSource = RsSearch
        FillFormatString
        flexGrid.ColDataType(ColChk) = flexDTBoolean
        FillGridCombos flexGrid, ColCountryNo
        If RsSearch.RecordCount = 0 Then
            MsgBox "·«ÌÊÃœ Õ—ﬂ« ", vbInformation + vbQuestion, "—”«·…  Õ–Ì—"
        End If
    Case 5
        ExportToExcel
    Case 7
        Unload Me
End Select
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Select Case ButtonMenu.Tag
    Case 1
        PrintData 2
End Select
End Sub
Sub FillFormatVSFlex(flexGrid As VSFlexGrid)
    fs = "|ID"
    fs = fs + "|>" + "«·—ﬁ„ «·„Œ“‰Ì"
    fs = fs + "|>" + "«·≈”„"
    With flexGrid
        .Visible = False
        .FormatString = fs
            .ColWidth(ColStkId_1) = 0
            SetColWidths ColStkno_1, flexGrid
            SetColWidths ColStkname_1, flexGrid
            .Visible = True
    End With
End Sub

Private Sub TxtStkNo_Change()
On Error GoTo ErrorHandler
Dim RsSearch As New ADODB.Recordset
If TxtStkNo.Text = "" Then
    TxtStkNo.Tag = "0"
    Grid.Visible = False
    Exit Sub
End If

If Ok Then
    Flag = False
    sqlText = "Select top 15 Id , ltrim(rtrim(StkNo))StkNo , ltrim(rtrim(StkName))StkName  from CoStock Where StkName Like" & LikeExpression(TxtStkNo.Text) & " or StkNo like '" & TxtStkNo.Text & "%'"
    Set RsSearch = de.con.Execute(sqlText)
    If RsSearch.RecordCount > 0 Then
        Set Grid.DataSource = RsSearch
        Grid.Row = 0
        FillFormatVSFlex Grid
        'ChangeCursor 3
        Grid.Visible = True
    Else
        TxtStkNo.Tag = "0"
        Grid.Visible = False
    End If
    Flag = True
End If
Exit Sub
ErrorHandler:
MsgBox Err.Description

End Sub
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

Private Sub TxtStkNo_GotFocus()
Pos = 3
End Sub

Private Sub TxtStkNo_KeyDown(KeyCode As Integer, Shift As Integer)
    Flag = True
    If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then MoveCursor KeyCode
    Flag = True
End Sub

Private Sub txtStkNo_KeyPress(KeyAscii As Integer)
On Error GoTo ErrorHandler
If KeyAscii = 13 Then
        Grid.Visible = False
        If DataOk(TxtStkNo.Text) Then
            TxtStkNo.Tag = GetStkId(TxtStkNo.Text)
            If TxtStkNo.Tag <> 0 Then
                LStkName.Caption = GetStkName(TxtStkNo.Tag)
            Else
                LStkName.Caption = ""
                MsgBox "«·„«œ… €Ì— „⁄—›… ›Ì «·„” Êœ⁄", vbInformation, " ‰»ÌÂ"
                SendKeys "{home}+{end}"
                Exit Sub
            End If
        Else
            Ok = False
            TxtStkNo.Text = Grid.TextMatrix(IIf(Grid.Row = 0, 1, Grid.Row), ColNo)
            Ok = True
            TxtStkNo.Tag = GetStkId(TxtStkNo.Text)
            If TxtStkNo.Tag <> 0 Then
                LStkName.Caption = GetStkName(TxtStkNo.Tag)
            Else
                LStkName.Caption = ""
                MsgBox "«·„«œ… €Ì— „⁄—›… ›Ì «·„” Êœ⁄", vbInformation, " ‰»ÌÂ"
                SendKeys "{home}+{end}"
                Exit Sub
            End If
        End If
        SendKeys "{home}+{end}"
    End If
Exit Sub
ErrorHandler:
Grid.Visible = False
MsgBox Err.Description

End Sub
