VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmSearch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«·»ÕÀ ⁄‰ ‘ﬂÊÏ"
   ClientHeight    =   7170
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11355
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   7170
   ScaleWidth      =   11355
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   1365
      Left            =   60
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   8730
      Visible         =   0   'False
      Width           =   10155
      _cx             =   17912
      _cy             =   2408
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
   Begin Threed.SSFrame SSFrame3 
      Height          =   495
      Left            =   30
      TabIndex        =   30
      Top             =   6630
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   873
      _Version        =   131074
   End
   Begin VSFlex8Ctl.VSFlexGrid FlexGrid 
      Height          =   2805
      Left            =   30
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   3780
      Width           =   11295
      _cx             =   19923
      _cy             =   4948
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
   Begin Threed.SSFrame SSFrame1 
      Height          =   1545
      Left            =   30
      TabIndex        =   16
      Top             =   720
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   2725
      _Version        =   131074
      ForeColor       =   255
      Caption         =   "„⁄·Ê„«  «·“»Ê‰"
      Alignment       =   1
      Begin VB.TextBox TxtAddress 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   60
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   1110
         Width           =   5355
      End
      Begin VB.TextBox TxtWorkPhone 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   60
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   660
         Width           =   2055
      End
      Begin VB.TextBox TxtMobilePhone 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   3360
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   720
         Width           =   2055
      End
      Begin VB.TextBox TxtCustomerName 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   60
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   270
         Width           =   9825
      End
      Begin VB.TextBox TxtZoneName 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   7830
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   1140
         Width           =   2055
      End
      Begin VB.TextBox TxtHomePhone 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   7830
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   750
         Width           =   2055
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "—ﬁ„ „ﬂ«‰ «·⁄„·"
         Height          =   195
         Index           =   6
         Left            =   2280
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   750
         Width           =   1035
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "«·⁄‰Ê«‰"
         Height          =   195
         Index           =   8
         Left            =   6060
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   1140
         Width           =   480
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "«·„‰ÿﬁ…"
         Height          =   195
         Index           =   7
         Left            =   10650
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   1170
         Width           =   480
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "—ﬁ„ «·„Ê»«Ì·"
         Height          =   195
         Index           =   5
         Left            =   5760
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   750
         Width           =   825
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "—ﬁ„ «·Â« › «·À«» "
         Height          =   195
         Index           =   4
         Left            =   9960
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   720
         Width           =   1155
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "≈”„ «·“»Ê‰"
         Height          =   195
         Index           =   3
         Left            =   10380
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   330
         Width           =   735
      End
   End
   Begin Threed.SSFrame SSFrame2 
      Height          =   1425
      Left            =   30
      TabIndex        =   17
      Top             =   2310
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   2514
      _Version        =   131074
      ForeColor       =   255
      Caption         =   "„⁄·Ê„«  «·‘ﬂÊÏ"
      Alignment       =   1
      Begin VB.TextBox TxtCallVia 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3420
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   690
         Width           =   7065
      End
      Begin VB.TextBox TxtREpName 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3420
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   1050
         Width           =   3045
      End
      Begin VB.TextBox TxtREpNo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   6480
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   1050
         Width           =   885
      End
      Begin MSMask.MaskEdBox TxtFromDate 
         Height          =   345
         Left            =   9360
         TabIndex        =   6
         Top             =   300
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   609
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "99/99/9999"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtTillDate 
         Height          =   345
         Left            =   7470
         TabIndex        =   7
         Top             =   300
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   609
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "99/99/9999"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtTilltime 
         Height          =   345
         Left            =   3420
         TabIndex        =   9
         Top             =   300
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   609
         _Version        =   393216
         MaxLength       =   5
         Mask            =   "99:99"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtFromTime 
         Height          =   345
         Left            =   5400
         TabIndex        =   8
         Top             =   300
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   609
         _Version        =   393216
         MaxLength       =   5
         Mask            =   "99:99"
         PromptChar      =   "_"
      End
      Begin MSDataListLib.DataCombo ComboProductFamily 
         Height          =   315
         Left            =   7380
         TabIndex        =   11
         Top             =   1050
         Width           =   3105
         _ExtentX        =   5477
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin Threed.SSCheck chk 
         Height          =   375
         Left            =   60
         TabIndex        =   14
         Top             =   990
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   661
         _Version        =   131074
         Caption         =   "«·€Ì— „ÿ»Ê⁄…"
         Alignment       =   1
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "«·⁄«∆·…"
         Height          =   195
         Index           =   6
         Left            =   10650
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   1020
         Width           =   420
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "«·„⁄—›…"
         Height          =   195
         Index           =   2
         Left            =   10650
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   630
         Width           =   495
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "≈·Ï «·”«⁄…"
         Height          =   195
         Index           =   1
         Left            =   6630
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   270
         Width           =   765
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "„‰ «·”«⁄…"
         Height          =   195
         Index           =   1
         Left            =   4620
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   270
         Width           =   690
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "≈·Ï  «—ÌŒ"
         Height          =   195
         Index           =   0
         Left            =   8610
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   270
         Width           =   675
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "„‰  «—ÌŒ"
         Height          =   195
         Index           =   0
         Left            =   10545
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   270
         Width           =   600
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4080
      Top             =   30
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
            Picture         =   "FrmSearch.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSearch.frx":26D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSearch.frx":5568
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSearch.frx":83C5
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSearch.frx":AB6A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSearch.frx":D466
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSearch.frx":F98B
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSearch.frx":12143
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSearch.frx":14B57
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSearch.frx":174A9
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSearch.frx":1A21D
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSearch.frx":1CA6C
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSearch.frx":1F913
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSearch.frx":2266D
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSearch.frx":2500D
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSearch.frx":27F6F
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSearch.frx":2A997
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSearch.frx":2D351
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSearch.frx":2FC82
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSearch.frx":32588
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSearch.frx":34FF0
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSearch.frx":37FA1
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSearch.frx":3A8CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSearch.frx":3D1FF
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSearch.frx":3F93F
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSearch.frx":422AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSearch.frx":44B36
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSearch.frx":46D45
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSearch.frx":496A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSearch.frx":4BECC
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSearch.frx":4E980
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSearch.frx":513BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSearch.frx":53ED2
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSearch.frx":56E6C
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSearch.frx":59C38
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSearch.frx":5C8B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSearch.frx":5F744
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSearch.frx":62382
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSearch.frx":64F31
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   690
      Left            =   0
      TabIndex        =   32
      Top             =   0
      Width           =   11355
      _ExtentX        =   20029
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
            ImageIndex      =   38
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   16
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Const ColCallNo = 1
    Const colModNo = 2
    Const ColProdFamName = 3
    Const ColCallStatus = 4
    Const ColCallDEscription = 5
    Const ColCallDate = 6
    Const ColCallTime = 7
    Const ColCliNo = 8
    Const ColCliName = 9
    Const ColAdhamPhon = 10
    Const ColMobilePhone = 11
    Const ColWorkPhone = 12
    Const ColZone = 13
    Const ColZoneName = 14
    Const ColAdhamAdress = 15
    Const ColVia = 16
    Const ColViaName = 17
    Const ColNotes = 18
    
    
    
    Const ColCustomerId = 1
    Const ColCustomerName = 2
    Const ColCustomerHomephone = 3
    Const ColCustomerMobilePhone = 4
    Const ColCustomerWorkPhone = 5
    Const ColCustomerAddress = 6
    
    
    Const ColViaNo_1 = 1
    Const ColViaName_1 = 2
    
    
    Const ColZoneNo_1 = 1
    Const ColZoneName_1 = 2
    
    
    Dim Flag As Boolean
    Dim OkRepNo As Boolean
    Dim SearchRec As SearchMaintCallRecType
    


Sub FillFormating(flexGrid As VSFlexGrid, i As Integer)
If i = 1 Then
    Fs = "|>" & "—ﬁ„ «·‘ﬂÊÏ"
    Fs = Fs + "|>" & "—ﬁ„ «·⁄«∆·…"
    Fs = Fs + "|>" & "≈”„ «·⁄«∆·…"
    Fs = Fs + "|>" & "„ÿ»Ê⁄"
    Fs = Fs + "|>" & "«·‘—Õ"
    Fs = Fs + "|>" & " «—ÌŒ «·‘ﬂÊÏ"
    Fs = Fs + "|>" & "Êﬁ  «·‘ﬂÊÏ"
    Fs = Fs + "|>" & "—ﬁ„ «·“»Ê‰"
    Fs = Fs + "|>" & "≈”„ «·“»Ê‰"
    Fs = Fs + "|>" & "«·Â« › «·À«» "
    Fs = Fs + "|>" & "«·„Ê»«Ì·"
    Fs = Fs + "|>" & "—ﬁ„ «·⁄„·"
    Fs = Fs + "|>" & "—ﬁ„ «·„‰ÿﬁ…"
    Fs = Fs + "|>" & "≈”„ «·„‰ÿﬁ…"
    Fs = Fs + "|>" & "«·⁄‰Ê«‰"
    Fs = Fs + "|>" & "—ﬁ„ «·„⁄—›…"
    Fs = Fs + "|>" & "≈”„ «·„⁄—›…"
    Fs = Fs + "|>" & "„·«ÕŸ« "
    With flexGrid
        .Visible = False
        .Cols = 19
        .FormatString = Fs
        
    SetColWidths ColCallNo, flexGrid
    SetColWidths ColProdFamName, flexGrid
    SetColWidths ColCallStatus, flexGrid
    SetColWidths ColCallDEscription, flexGrid
    SetColWidths ColCallDate, flexGrid
    SetColWidths ColCallTime, flexGrid
    SetColWidths ColCliName, flexGrid
    SetColWidths ColAdhamPhon, flexGrid
    SetColWidths ColMobilePhone, flexGrid
    SetColWidths ColWorkPhone, flexGrid
    SetColWidths ColZoneName, flexGrid
    SetColWidths ColAdhamAdress, flexGrid
    SetColWidths ColViaName, flexGrid
    SetColWidths ColNotes, flexGrid
    
    .ColWidth(colModNo) = 0
    .ColWidth(ColCliNo) = 0
    .ColWidth(ColZone) = 0
    .ColWidth(ColVia) = 0
    .Visible = True
    End With
ElseIf i = 2 Then
    Fs = "|>" & "—ﬁ„ «·“»Ê‰"
    Fs = Fs + "|>" & "≈”„ «·“»Ê‰"
    Fs = Fs + "|>" & "—ﬁ„ «·„‰“·"
    Fs = Fs + "|>" & "—ﬁ„ «·„Ê»«Ì·"
    Fs = Fs + "|>" & "—ﬁ„ «·⁄„·"
    Fs = Fs + "|>" & "«·⁄‰Ê«‰"
    With flexGrid
        .Visible = False
        .Cols = 7
        .FormatString = Fs
        SetColWidths ColCustomerId, flexGrid
        SetColWidths ColCustomerName, flexGrid
        SetColWidths ColCustomerHomephone, flexGrid
        SetColWidths ColCustomerMobilePhone, flexGrid
        SetColWidths ColCustomerWorkPhone, flexGrid
        SetColWidths ColCustomerAddress, flexGrid
        .Visible = True
    End With
ElseIf i = 4 Then
    Fs = "|>" & "—ﬁ„ «·„⁄—›…"
    Fs = Fs + "|>" & "≈”„ «·„⁄—›…"
    With flexGrid
        .Visible = False
        .Cols = 3
        .FormatString = Fs
        .ColWidth(ColViaNo_1) = 0
        SetColWidths ColViaName_1, flexGrid
        .Visible = True
    End With
ElseIf i = 5 Then
    Fs = "|>" & "—ﬁ„ «·„‰ÿﬁ…"
    Fs = Fs + "|>" & "≈”„ «·„‰ÿﬁ…"
    With flexGrid
        .Visible = False
        .Cols = 3
        .FormatString = Fs
        .ColWidth(ColZoneNo_1) = 0
        SetColWidths ColZoneName_1, flexGrid
        .Visible = True
    End With
End If
End Sub

Sub FillList(sqltext As String, Field1 As String, Field2 As String, List As VSFlexGrid, ByVal Switch As Integer)
    
    Set rs = de.con.Execute(sqltext)
    If rs.RecordCount > 0 Then
        Set List.DataSource = rs
        FillFormating List, Switch
        List.Row = 1
        List.Col = 1
        List.ColSel = List.Cols - 1
        List.Visible = True
    Else
        List.Rows = 1
        ActiveControl.Tag = 0
        List.Visible = False

    End If
End Sub
Sub SetColWidths(ByVal ColNo As Integer, flexGrid As VSFlexGrid)
    With flexGrid
        .AutoSize ColNo
    End With
End Sub

Sub MoveCursor(KeyCode As Integer, Grid As VSFlexGrid)
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
Sub FillCombos()
    Dim rs As New ADODB.Recordset
    sqltext = "Select ProdFamNo  , ProdFamNamea From AdhamProductFamily "
    Set rs = de.con.Execute(sqltext)
    Set ComboProductFamily.RowSource = rs
    ComboProductFamily.ListField = "ProdFamNamea"
    ComboProductFamily.BoundColumn = "ProdFamNo"
End Sub


Sub MoveGrid(VTop As Integer, VLeft As Integer, VWidth As Integer, Grid As VSFlexGrid)
With Grid
    .Top = VTop
    .Width = VWidth
    .Left = VLeft
End With
End Sub
Sub Init()
Top = 0
Left = 0
Flag = True
FillCombos
OkRepNo = True
FillFormating flexGrid, 1
flexGrid.Rows = 1
End Sub

Private Sub Chk_Click(Value As Integer)
If Not Chk.Value Then
    Chk.Caption = "«·€Ì— „ÿ»Ê⁄…"
Else
    Chk.Caption = "«·„ÿ»Ê⁄…"
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SendKeys "{tab}"
   SendKeys "{home}+{end}"
End If
End Sub

Private Sub Form_Load()
Init
End Sub
Function FillStructure() As Boolean

On Error GoTo errorhandler
With SearchRec
    .Address = TxtAddress.Text
    .CustomerId = Val(TxtCustomerName.Tag)
    If IsDate(TxtFromDate.Text) Then
        .FromDate = ConvertControlDate(TxtFromDate.Text)
    Else
        .FromDate = ""
    End If
    
    If IsDate(TxtTillDate.Text) Then
        .TillDate = ConvertControlDate(TxtTillDate.Text)
    Else
        .TillDate = ""
    End If
    
    .FromTime = TxtFromTime.Text
    .tillTime = TxtTilltime.Text
    .HomePhone = TxtHomePhone.Text
    .MobilePhone = TxtMobilePhone.Text
    .Workphone = TxtWorkPhone.Text
    .ProductFamilyNo = Val(ComboProductFamily.BoundText)
    .RepNo = Val(txtRepNo.Text)
    .Note = TxtREpName.Text
    .Via = TxtCallVia.Text
    .ZoneNo = Val(TxtZoneName.Tag)
    FillStructure = True
End With
Exit Function
errorhandler:
MsgBox Err.Description
End Function
Sub SearchData()
Dim rs As New ADODB.Recordset
Dim sqltext As String
If FillStructure Then
    With SearchRec
        sqltext = "Exec Sp_search_Maint_Calls " & .CustomerId & ",'" & .HomePhone & "','" & .Workphone & "','" & .MobilePhone & "'," & .ZoneNo & ",'" & .Address & "','" & .FromDate & "','" & .TillDate & "','" & .FromTime & "','" & .tillTime & "','" & .Via & "'," & .ProductFamilyNo & "," & .RepNo & ",'" & .Note & "'"
        de.con.Execute (sqltext)
        sqltext = "select CallNo, ModNo, ProdFamName, CallStatus, CallDescription, convert(varchar(10),CallDatetime,103) CallDate, convert(varchar(5),CallDatetime,108) CAlltime ,  cliNo, CliName, AdhamPhon, MobilePhone, WorkPhone, Zone, ZoneName, adhamadress, Via, ViaName, notes from t_Search_Calls"
        Set rs = de.con.Execute(sqltext)
        Set flexGrid.DataSource = rs
        FillFormating flexGrid, 1
    End With
End If

End Sub
Sub PrintData()

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
    Case 1
        PrintData
    Case 3
        SearchData
    Case 5
        Unload Me
End Select
End Sub

Private Sub TxtCallVia_Change()
If Flag Then
    Dim sqltext As String
    If Trim(TxtCallVia.Text) = "" Then
         TxtCallVia.Tag = 0
        Grid.Visible = False
        Exit Sub
        End If
        sqltext = "select top 10 accNo , AccName From MaintCallVia Where AccName like " & LikeExpression(TxtCallVia.Text)
        FillList sqltext, "AccNo", "AccName", Grid, 4
        MoveGrid SSFrame2.Top + TxtCallVia.Top + TxtCallVia.Height, SSFrame2.Left + TxtCallVia.Left, TxtCallVia.Width, Grid
        End If
End Sub

Private Sub TxtCallVia_KeyDown(KeyCode As Integer, Shift As Integer)
    Flag = False
    If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then MoveCursor KeyCode, Grid
    Flag = True
End Sub

Private Sub TxtCallVia_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
With Grid
    If .Visible Then
        TxtCallVia.Tag = .TextMatrix(.Row, ColViaNo_1)
        Flag = False
        TxtCallVia.Text = .TextMatrix(.Row, ColViaName_1)
        Flag = True
        .Visible = False
    Else
        Flag = False
        TxtCallVia.Tag = 0
        TxtCallVia.Text = ""
        Flag = True
    End If
End With
End If

End Sub

Private Sub TxtCustomerName_Change()
If Flag Then
    Dim sqltext As String
    If Trim(TxtCustomerName.Text) = "" Then
        TxtCustomerName.Tag = 0
        Grid.Visible = False
        Exit Sub
        End If
        sqltext = "Select top 10 [adhamno] , [AdhamName] , [AdhamPhon] , [MobilePhone] , [WorkPhone] ,  [adhamadress]  From adhamview7 Where   AdhamName like" & LikeExpression(TxtCustomerName.Text)
        FillList sqltext, "AdhamNo", "AdhamName", Grid, 2
        MoveGrid SSFrame1.Top + TxtCustomerName.Top + TxtCustomerName.Height, SSFrame1.Left + TxtCustomerName.Left, TxtCustomerName.Width, Grid
End If
End Sub

Private Sub TxtCustomerName_KeyDown(KeyCode As Integer, Shift As Integer)
    Flag = False
    If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then MoveCursor KeyCode, Grid
    Flag = True
End Sub

Function GetCustomerName(CustomerId As Double) As String
Dim rs As New ADODB.Recordset
sqltext = "Select adhamName from AdhamView7 Where AdhamNo=" & CustomerId
Set rs = de.con.Execute(sqltext)
GetCustomerName = rs!AdhamName
End Function



Private Sub TxtCustomerName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then

With Grid
    If .Visible Then
        TxtCustomerName.Tag = .TextMatrix(.Row, ColCustomerId)
        Flag = False
        TxtCustomerName.Text = .TextMatrix(.Row, ColCustomerName)
        Flag = True
        .Visible = False
    ElseIf Val(TxtCustomerName.Tag) <> 0 And .Visible = False Then
        Flag = False
        TxtCustomerName.Text = GetCustomerName(TxtCustomer.Tag)
        Flag = True
    Else
        Flag = False
        TxtCustomerName.Tag = 0
        TxtCustomerName.Text = ""
        Flag = True
    End If
End With
End If

End Sub
Function GetRepName(ByVal CallNo As Integer) As String
Dim rs As New ADODB.Recordset

On Error GoTo errorhandler
sqltext = "Select Callid , Calldescription From mntcallscode m1 inner join adhamproductfamily a1 on m1.prodfamno = a1.prodfamno Where Callid = " & CallNo & "  and m1.ProdFamNo =" & Val(ComboProductFamily.BoundText)
Set rs = de.con.Execute(sqltext)

If rs.RecordCount > 0 Then
    GetRepName = rs!CallDEscription
Else
    GetRepName = ""
End If
Exit Function
errorhandler:
GetRepName = ""
End Function

Private Sub TxtREpNo_Change()
If OkRepNo Then
    TxtREpName = GetRepName(Val(txtRepNo.Text))
End If

End Sub

Private Sub TxtZoneName_Change()
If Flag Then
    Dim sqltext As String
    If Trim(TxtZoneName.Text) = "" Then
         TxtZoneName.Tag = 0
        Grid.Visible = False
        Exit Sub
        End If
        sqltext = "Select ZoneNo , ZoneName From CoZone Where ZoneName like " & LikeExpression(TxtZoneName.Text)
        FillList sqltext, "ZoneNo", "ZoneName", Grid, 5
        MoveGrid SSFrame1.Top + TxtZoneName.Top + TxtZoneName.Height, TxtZoneName.Left + SSFrame1.Left, TxtZoneName.Width, Grid
End If
End Sub

Private Sub TxtZoneName_KeyDown(KeyCode As Integer, Shift As Integer)
    Flag = False
    If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then MoveCursor KeyCode, Grid
    Flag = True
End Sub

Private Sub TxtZoneName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
With Grid
    If .Visible Then
        TxtZoneName.Tag = .TextMatrix(.Row, ColZoneNo_1)
        Flag = False
        TxtZoneName.Text = .TextMatrix(.Row, ColZoneName_1)
        Flag = True
        .Visible = False
    ElseIf Val(TxtZoneName.Tag) = 0 Then
        Flag = False
        TxtZoneName.Tag = 0
        TxtZoneName.Text = ""
        Flag = True
    End If
End With
End If
End Sub
