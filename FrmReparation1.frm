VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmReparation1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«·«’·«Õ« "
   ClientHeight    =   8880
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11685
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   8880
   ScaleWidth      =   11685
   Begin VB.TextBox TxtEngineBarcode 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   2850
      RightToLeft     =   -1  'True
      TabIndex        =   88
      Top             =   4650
      Width           =   1935
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      DragMode        =   1  'Automatic
      Height          =   2685
      Left            =   7200
      TabIndex        =   61
      Top             =   6360
      Visible         =   0   'False
      Width           =   2415
      _cx             =   4260
      _cy             =   4736
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
   Begin VB.TextBox TxtProductName 
      Alignment       =   1  'Right Justify
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
      Height          =   450
      Left            =   9900
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   2310
      Width           =   1710
   End
   Begin VB.TextBox txtFindCallNo 
      Alignment       =   1  'Right Justify
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
      Height          =   450
      Left            =   10650
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   300
      Width           =   1020
   End
   Begin VB.TextBox TxtStkQty 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   930
      RightToLeft     =   -1  'True
      TabIndex        =   26
      Top             =   4650
      Width           =   855
   End
   Begin VB.TextBox TxtStkName 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   7320
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Top             =   4650
      Width           =   4365
   End
   Begin VB.TextBox txtRepPrice 
      Height          =   360
      Left            =   1860
      TabIndex        =   13
      Top             =   810
      Width           =   1155
   End
   Begin VB.TextBox txtProdSerNo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Height          =   450
      Left            =   4950
      MaxLength       =   4
      TabIndex        =   19
      Top             =   2310
      Width           =   1035
   End
   Begin VB.TextBox txtGasNo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Height          =   450
      Left            =   1110
      MaxLength       =   1
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   2310
      Width           =   795
   End
   Begin VB.TextBox TxtModelName 
      Alignment       =   1  'Right Justify
      Height          =   450
      Left            =   5970
      RightToLeft     =   -1  'True
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   2310
      Width           =   3885
   End
   Begin VB.TextBox txtDescription 
      Alignment       =   1  'Right Justify
      Height          =   450
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   1530
      Width           =   5775
   End
   Begin VB.TextBox txtVoltBefor 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   5580
      TabIndex        =   11
      Top             =   780
      Width           =   1020
   End
   Begin VB.TextBox txtVoltAfter 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   3900
      TabIndex        =   12
      Top             =   780
      Width           =   1020
   End
   Begin VB.TextBox txtNotes 
      Alignment       =   1  'Right Justify
      Height          =   450
      Left            =   5850
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   1530
      Width           =   5775
   End
   Begin MSMask.MaskEdBox txtRepTimeBegin 
      Height          =   450
      Left            =   1050
      TabIndex        =   7
      Top             =   300
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   794
      _Version        =   393216
      MaxLength       =   5
      Format          =   "hh:mm"
      Mask            =   "99:99"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtRepTimeEnd 
      Height          =   450
      Left            =   90
      TabIndex        =   8
      Top             =   300
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   794
      _Version        =   393216
      MaxLength       =   5
      Format          =   "hh:mm"
      Mask            =   "99:99"
      PromptChar      =   "_"
   End
   Begin MSDataListLib.DataCombo DcboReparationStatus 
      Height          =   360
      Left            =   7320
      TabIndex        =   10
      Top             =   810
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   635
      _Version        =   393216
      ListField       =   ""
      BoundColumn     =   ""
      Text            =   ""
      RightToLeft     =   -1  'True
      Object.DataMember      =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataCombo DcboPayMethod 
      Height          =   360
      Left            =   60
      TabIndex        =   14
      Top             =   810
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   635
      _Version        =   393216
      ListField       =   ""
      BoundColumn     =   "no"
      Text            =   ""
      RightToLeft     =   -1  'True
      Object.DataMember      =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataCombo DcboCliRecever 
      Height          =   360
      Left            =   9600
      TabIndex        =   9
      Top             =   810
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   635
      _Version        =   393216
      ListField       =   ""
      BoundColumn     =   ""
      Text            =   ""
      RightToLeft     =   -1  'True
      Object.DataMember      =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSMask.MaskEdBox TxtRegisterDate 
      Height          =   450
      Left            =   2040
      TabIndex        =   6
      Top             =   300
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   794
      _Version        =   393216
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "99/99/9999"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtPurchaseDate 
      Height          =   450
      Left            =   30
      TabIndex        =   23
      Top             =   2310
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   794
      _Version        =   393216
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd/mm/yyyy"
      Mask            =   "99/99/9999"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtProductionDate 
      Height          =   450
      Left            =   3930
      TabIndex        =   20
      Top             =   2310
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   794
      _Version        =   393216
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd/mm/yyyy"
      Mask            =   "99/99/9999"
      PromptChar      =   "_"
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1650
      Top             =   1260
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
            Picture         =   "FrmReparation1.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReparation1.frx":26D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReparation1.frx":5533
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReparation1.frx":7CD8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReparation1.frx":A5D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReparation1.frx":CAF9
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReparation1.frx":F2B1
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReparation1.frx":11CC5
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReparation1.frx":14617
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReparation1.frx":1738B
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReparation1.frx":19BDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReparation1.frx":1CA81
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReparation1.frx":1F7DB
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReparation1.frx":2217B
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReparation1.frx":250DD
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReparation1.frx":27B05
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReparation1.frx":2A4BF
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReparation1.frx":2CDF0
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReparation1.frx":2F6F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReparation1.frx":3215E
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReparation1.frx":3510F
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReparation1.frx":37A38
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReparation1.frx":3A36D
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReparation1.frx":3CAAD
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReparation1.frx":3F41C
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReparation1.frx":41CA4
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReparation1.frx":43EB3
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReparation1.frx":46810
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReparation1.frx":4903A
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReparation1.frx":4BAEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReparation1.frx":4E52A
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReparation1.frx":51040
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReparation1.frx":53FDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReparation1.frx":56DA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReparation1.frx":59A20
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReparation1.frx":5C8B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReparation1.frx":5F4F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReparation1.frx":6209F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin Threed.SSFrame SSFrame4 
      DragMode        =   1  'Automatic
      Height          =   495
      Left            =   2610
      TabIndex        =   62
      Top             =   7980
      Width           =   9105
      _ExtentX        =   16060
      _ExtentY        =   873
      _Version        =   131074
      Begin Threed.SSCommand CmdEdit 
         Height          =   435
         Left            =   6480
         TabIndex        =   80
         Top             =   30
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   767
         _Version        =   131074
         ForeColor       =   8388608
         PictureAnimationDelay=   66
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   " ⁄œÌ·"
      End
      Begin Threed.SSCommand CmdSearch 
         Height          =   435
         Left            =   1320
         TabIndex        =   69
         Top             =   30
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   767
         _Version        =   131074
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "»ÕÀ"
      End
      Begin Threed.SSCommand CmdAdd 
         Height          =   435
         Left            =   7770
         TabIndex        =   0
         Top             =   30
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   767
         _Version        =   131074
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "ÃœÌœ"
      End
      Begin Threed.SSCommand CmdDelete 
         Height          =   435
         Left            =   5190
         TabIndex        =   65
         Top             =   30
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   767
         _Version        =   131074
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Õ–›"
      End
      Begin Threed.SSCommand CmdCancel 
         Height          =   435
         Left            =   2610
         TabIndex        =   64
         Top             =   30
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   767
         _Version        =   131074
         ForeColor       =   8388608
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   " —«Ã⁄"
      End
      Begin Threed.SSCommand cmdSave 
         Height          =   435
         Left            =   3900
         TabIndex        =   27
         Top             =   30
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   767
         _Version        =   131074
         ForeColor       =   8388608
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Õ›Ÿ"
      End
      Begin Threed.SSCommand CmdExit 
         Height          =   435
         Left            =   30
         TabIndex        =   63
         Top             =   30
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   767
         _Version        =   131074
         ForeColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Œ—ÊÃ"
      End
   End
   Begin Threed.SSFrame NavigatorFrame 
      DragMode        =   1  'Automatic
      Height          =   375
      Left            =   60
      TabIndex        =   70
      Top             =   8010
      Width           =   2475
      _ExtentX        =   4366
      _ExtentY        =   661
      _Version        =   131074
      Begin VB.CommandButton CmdLast 
         Height          =   285
         Left            =   2190
         Picture         =   "FrmReparation1.frx":64A4E
         Style           =   1  'Graphical
         TabIndex        =   74
         TabStop         =   0   'False
         ToolTipText     =   "Last"
         Top             =   60
         Width           =   255
      End
      Begin VB.CommandButton CmdNext 
         Height          =   285
         Left            =   1920
         Picture         =   "FrmReparation1.frx":64F80
         Style           =   1  'Graphical
         TabIndex        =   73
         TabStop         =   0   'False
         ToolTipText     =   "Next"
         Top             =   60
         Width           =   255
      End
      Begin VB.CommandButton CmdPrevious 
         Height          =   285
         Left            =   330
         Picture         =   "FrmReparation1.frx":6507A
         Style           =   1  'Graphical
         TabIndex        =   72
         TabStop         =   0   'False
         ToolTipText     =   "Previous"
         Top             =   60
         Width           =   255
      End
      Begin VB.CommandButton CmdFirst 
         Height          =   285
         Left            =   60
         Picture         =   "FrmReparation1.frx":65174
         Style           =   1  'Graphical
         TabIndex        =   71
         TabStop         =   0   'False
         ToolTipText     =   "First"
         Top             =   60
         Width           =   255
      End
      Begin VB.Label LNavigator 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   600
         TabIndex        =   75
         Top             =   60
         Width           =   1305
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid FlexCurrentDamge 
      Height          =   915
      Left            =   90
      TabIndex        =   76
      Top             =   3420
      Width           =   11655
      _cx             =   20558
      _cy             =   1614
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
   Begin VSFlex8Ctl.VSFlexGrid FlexReparatonItems 
      Height          =   2925
      Left            =   60
      TabIndex        =   77
      Top             =   5040
      Width           =   11655
      _cx             =   20558
      _cy             =   5159
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   1
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
      ExplorerBar     =   1
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
   Begin VB.TextBox TxtCurrentDamage 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   60
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Top             =   3030
      Width           =   11625
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "—ﬁ„ Ì«—ﬂÊœ «·„Õ—ﬂ"
      DragMode        =   1  'Automatic
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   0
      Left            =   3510
      TabIndex        =   87
      Top             =   4380
      Width           =   1260
   End
   Begin VB.Label Ldiff 
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
      Height          =   270
      Left            =   4800
      RightToLeft     =   -1  'True
      TabIndex        =   86
      Top             =   8550
      Width           =   1005
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "ﬁÌ„Â «·«ÃÊ—"
      DragMode        =   1  'Automatic
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   5850
      TabIndex        =   85
      Top             =   8550
      Width           =   900
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "«·ﬁ«∆„ »«·⁄„·"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   1470
      RightToLeft     =   -1  'True
      TabIndex        =   84
      Top             =   8550
      Width           =   1035
   End
   Begin VB.Label LEmployeeName 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H000000C0&
      Height          =   270
      Left            =   30
      RightToLeft     =   -1  'True
      TabIndex        =   83
      Top             =   8550
      Width           =   1395
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Õ«·Â «·«’·«Õ"
      DragMode        =   1  'Automatic
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   3750
      TabIndex        =   82
      Top             =   8550
      Width           =   975
   End
   Begin VB.Label lReparationState 
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
      Height          =   270
      Left            =   2550
      RightToLeft     =   -1  'True
      TabIndex        =   81
      Top             =   8550
      Width           =   1125
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "«·„‰ Ã"
      DragMode        =   1  'Automatic
      Height          =   195
      Index           =   2
      Left            =   11175
      TabIndex        =   79
      Top             =   2070
      Width           =   405
   End
   Begin VB.Label LteamName 
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
      ForeColor       =   &H000000C0&
      Height          =   420
      Left            =   4350
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   330
      Width           =   1545
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "«·ÊÕœÂ «·„‰›–Â"
      ForeColor       =   &H80000002&
      Height          =   195
      Index           =   1
      Left            =   4935
      RightToLeft     =   -1  'True
      TabIndex        =   78
      Top             =   60
      Width           =   960
   End
   Begin VB.Label LRepDate 
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
      ForeColor       =   &H000000C0&
      Height          =   420
      Left            =   3090
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   330
      Width           =   1215
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   " «—ÌŒ «·≈œŒ«·"
      ForeColor       =   &H80000002&
      Height          =   195
      Index           =   3
      Left            =   2160
      RightToLeft     =   -1  'True
      TabIndex        =   68
      Top             =   60
      Width           =   900
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "«·‘ﬂÊÏ"
      Height          =   195
      Left            =   11040
      RightToLeft     =   -1  'True
      TabIndex        =   67
      Top             =   60
      Width           =   570
   End
   Begin VB.Label LClient 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "«·“»Ê‰"
      Height          =   195
      Left            =   9180
      RightToLeft     =   -1  'True
      TabIndex        =   66
      Top             =   60
      Width           =   420
   End
   Begin VB.Label LClientName 
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
      ForeColor       =   &H000000C0&
      Height          =   420
      Left            =   5940
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   330
      Width           =   3675
   End
   Begin VB.Label LModStockNo 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H000000C0&
      Height          =   420
      Left            =   1950
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   2340
      Width           =   1935
   End
   Begin VB.Label LDealerPrice 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   60
      TabIndex        =   60
      Top             =   4680
      Width           =   825
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "”⁄— «· «Ã—"
      DragMode        =   1  'Automatic
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   1
      Left            =   150
      TabIndex        =   59
      Top             =   4380
      Width           =   735
   End
   Begin VB.Label LBalance 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   1860
      TabIndex        =   58
      Top             =   4680
      Width           =   945
   End
   Begin VB.Label LStkName 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      DragMode        =   1  'Automatic
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   4830
      TabIndex        =   57
      Top             =   4680
      Width           =   2475
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "«·ﬂ„Ì…"
      DragMode        =   1  'Automatic
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   11
      Left            =   1320
      TabIndex        =   56
      Top             =   4380
      Width           =   405
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "«·‘—Õ"
      DragMode        =   1  'Automatic
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   10
      Left            =   6870
      TabIndex        =   55
      Top             =   4380
      Width           =   420
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "«·—’Ìœ"
      DragMode        =   1  'Automatic
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   9
      Left            =   2400
      TabIndex        =   54
      Top             =   4380
      Width           =   435
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "«·—ﬁ„ «·„Œ“‰Ì"
      DragMode        =   1  'Automatic
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   8
      Left            =   10710
      TabIndex        =   53
      Top             =   4380
      Width           =   930
   End
   Begin VB.Label LRepNbr 
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
      ForeColor       =   &H000000C0&
      Height          =   420
      Left            =   9660
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   330
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "—ﬁ„ «·«’·«Õ"
      Height          =   195
      Left            =   9690
      RightToLeft     =   -1  'True
      TabIndex        =   52
      Top             =   60
      Width           =   930
   End
   Begin VB.Label lblSelectedEmployees 
      AutoSize        =   -1  'True
      Caption         =   "√⁄ÿ«· «·≈’·«Õ «·Õ«·Ì"
      DragMode        =   1  'Automatic
      Height          =   195
      Left            =   10170
      RightToLeft     =   -1  'True
      TabIndex        =   51
      Top             =   2760
      Width           =   1485
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "«·‘—«¡"
      DragMode        =   1  'Automatic
      Height          =   195
      Index           =   19
      Left            =   570
      TabIndex        =   50
      Top             =   2040
      Width           =   480
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   2  'Center
      Caption         =   "«·€«“"
      DragMode        =   1  'Automatic
      Height          =   195
      Index           =   18
      Left            =   1110
      TabIndex        =   49
      Top             =   2040
      Width           =   795
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   2  'Center
      Caption         =   "«·„Œ“‰Ì"
      DragMode        =   1  'Automatic
      Height          =   195
      Index           =   17
      Left            =   1935
      TabIndex        =   48
      Top             =   2040
      Width           =   1935
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   2  'Center
      Caption         =   " /«·≈‰ «Ã"
      DragMode        =   1  'Automatic
      Height          =   195
      Index           =   16
      Left            =   3885
      TabIndex        =   47
      Top             =   2040
      Width           =   1005
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   2  'Center
      Caption         =   " ”·”·"
      DragMode        =   1  'Automatic
      Height          =   195
      Index           =   14
      Left            =   4950
      TabIndex        =   46
      Top             =   2040
      Width           =   1035
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "„ÊœÌ·"
      DragMode        =   1  'Automatic
      Height          =   195
      Index           =   13
      Left            =   9420
      TabIndex        =   45
      Top             =   2040
      Width           =   450
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Õ«·… «·≈’·«Õ"
      Height          =   195
      Index           =   5
      Left            =   8730
      TabIndex        =   44
      Top             =   840
      Width           =   825
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   " «—ÌŒ «·≈’·«Õ"
      ForeColor       =   &H80000002&
      Height          =   195
      Index           =   6
      Left            =   3420
      TabIndex        =   43
      Top             =   60
      Width           =   915
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "”«⁄… «·»œ¡"
      ForeColor       =   &H80000002&
      Height          =   195
      Index           =   7
      Left            =   1290
      TabIndex        =   42
      Top             =   60
      Width           =   735
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "”«⁄… «·«‰ Â«¡"
      ForeColor       =   &H80000002&
      Height          =   195
      Index           =   8
      Left            =   150
      TabIndex        =   41
      Top             =   60
      Width           =   915
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "ﬁÌ„… «·≈’·«Õ"
      DragMode        =   1  'Automatic
      Height          =   195
      Index           =   9
      Left            =   3060
      TabIndex        =   40
      Top             =   840
      Width           =   840
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "«·≈’·«Õ« "
      ForeColor       =   &H80000002&
      Height          =   195
      Index           =   15
      Left            =   10950
      TabIndex        =   39
      Top             =   1260
      Width           =   660
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "«·„·«ÕŸ« "
      Height          =   195
      Index           =   22
      Left            =   5130
      TabIndex        =   38
      Top             =   1260
      Width           =   660
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "›Ê·  ﬁ»·"
      Height          =   195
      Index           =   23
      Left            =   6630
      TabIndex        =   37
      Top             =   810
      Width           =   645
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "›Ê·  »⁄œ"
      Height          =   195
      Index           =   24
      Left            =   4950
      TabIndex        =   36
      Top             =   840
      Width           =   615
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "«·„” ﬁ»·"
      Height          =   195
      Index           =   0
      Left            =   11010
      TabIndex        =   35
      Top             =   870
      Width           =   600
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "«·œ›⁄"
      DragMode        =   1  'Automatic
      ForeColor       =   &H80000002&
      Height          =   195
      Index           =   4
      Left            =   1440
      TabIndex        =   34
      Top             =   810
      Width           =   360
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "«·√ﬁ·«„"
      DragMode        =   1  'Automatic
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   15
      Left            =   11160
      TabIndex        =   33
      Top             =   8550
      Width           =   480
   End
   Begin VB.Label LCount 
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
      ForeColor       =   &H000000C0&
      Height          =   270
      Left            =   10440
      RightToLeft     =   -1  'True
      TabIndex        =   32
      Top             =   8550
      Width           =   615
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "ﬁÌ„… «·„Ê«œ"
      DragMode        =   1  'Automatic
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   9540
      TabIndex        =   31
      Top             =   8550
      Width           =   885
   End
   Begin VB.Label LTotalItems 
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
      Height          =   270
      Left            =   8550
      TabIndex        =   30
      Top             =   8550
      Width           =   975
   End
   Begin VB.Label LTotal 
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
      Height          =   270
      Left            =   6780
      RightToLeft     =   -1  'True
      TabIndex        =   29
      Top             =   8550
      Width           =   1065
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "«·≈Ã„«·Ì"
      DragMode        =   1  'Automatic
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   14
      Left            =   7860
      TabIndex        =   28
      Top             =   8550
      Width           =   675
   End
End
Attribute VB_Name = "FrmReparation1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsNavigator As New ADODB.Recordset
Dim Ok As Boolean, Flag As Boolean, Pos As Integer, RecNum  As Double, ReparationState As EnumState

Dim maintDataService_ As New MaintDataService
Dim stockDataService_    As New stockDataService
Dim reparationViewModelInfo_   As ReparationViewModel
Attribute reparationViewModelInfo_.VB_VarHelpID = -1
'Dim reparationInfo_ As New Reparation



Const ColSerNo = 1
Const ColStkNo = 2
Const ColStkName = 3
Const ColEngineBarcode = 4
Const ColQty = 5
Const ColDealerPrice = 6

Const ColReparationNo = 1
Const ColReparationDescription = 2

Const ColNo = 1
Const ColName = 2
'ˆˆ
''Private Sub Form_Load()
''On Error GoTo errorhandler
''Dim doc As New MSXML2.DOMDocument
''Dim isopen  As Boolean
''isopen = doc.Load("\\mainserver\d$\SystemConfigration.xml")
''
''If isopen Then
''   Dim nodeList As MSXML2.IXMLDOMNodeList
''
''    Set nodeList = doc.selectNodes("/systems/system/servcerName")
''    If Not nodeList Is Nothing Then
''         Dim node As MSXML2.IXMLDOMNode
''         Dim name As String
''         Dim Value As String
''
''         For Each node In nodeList
''            ' Could also do node.attributes.getNamedItem("name").text
''            name = node.selectSingleNode("@name").Text
''            Value = node.selectSingleNode("@value").Text
''         Next node
''    End If
''End If
''Exit Sub
''errorhandler:
''MsgBox Err.Description
''End Sub


'
Sub FillFormating(ByVal i As Integer, FlexGrid As VSFlexGrid)
If i = 1 Then

    fs = "|>" + "«·—ﬁ„"
    fs = fs + "|>" + "«·«”„"

    With FlexGrid
        .FormatString = fs
        .Cols = 3

        SetColWidths ColNo, FlexGrid
        SetColWidths ColName, FlexGrid

    End With
ElseIf i = 2 Then
        fs = "|>" + "—ﬁ„ «·«’·«Õ"
        fs = fs + "|>" + "≈”„ «·«’·«Õ"
         With FlexGrid
            .FormatString = fs
            .Cols = 3
            SetColWidths ColReparationDescription, FlexGrid
    SetColWidths ColReparationNo, FlexGrid
            '.ColWidth(ColReparationNo) = 0
    End With
ElseIf i = 3 Then
   
    fs = "|>" + "SerNo."
    fs = fs + "|>" + "«·—ﬁ„ «·„Œ“‰Ì"
    fs = fs + "|>" + "«·‘—Õ"
    fs = fs + "|>" + "»«—ﬂÊœ «·„Õ—ﬂ"
    fs = fs + "|>" + "«·ﬂ„Ì…"
    fs = fs + "|>" + "«·”⁄—"
     With FlexGrid
        .FormatString = fs
        .Cols = 7
        SetColWidths ColStkNo, FlexGrid
        SetColWidths ColStkName, FlexGrid
        SetColWidths ColEngineBarcode, FlexGrid
        SetColWidths ColQty, FlexGrid
        SetColWidths ColDealerPrice, FlexGrid

        .ColWidth(ColSerNo) = 0
    End With
End If
End Sub



Sub ChangeCursor(sender As Control)

    With sender
       Grid.top = .top + .Height
       Grid.left = .left
       Grid.Width = .Width
    End With

End Sub

Private Sub CmdAdd_Click()
Set reparationViewModelInfo_ = New ReparationViewModel
ReparationState = NewRecord
EnableCmds False, False, False, True, True, False
EnableControls True
ClearControls
txtFindCallNo.SetFocus
txtFindCallNo.Locked = False

End Sub


'Function MaxRec() As Double
'Dim RsMax As New ADODB.Recordset
'    sqlText = "Select isnull(Max(BillNo),0)MaxBillNo From MvMaintPayments"
'    Set RsMax = de.con.Execute(sqlText)
'    MaxRec = RsMax!MaxBillNo
'End Function

Private Sub CmdCancel_Click()
On Error GoTo ErrorHandler
    EnableCmds True, True, True, False, False, True
    EnableControls False
    If IsNull(RsNavigator!CallNo) Then
        MoveToRec Val(txtFindCallNo.Text)
    Else
        MoveToRec Val(RsNavigator!CallNo)
    End If
    ReparationState = DefaultRecord
    ChangeStatus ReparationState
Exit Sub
ErrorHandler:
MsgBox Err.Description
End Sub

Private Sub CmdDelete_Click()

On Error GoTo ErrorHandler
If reparationViewModelInfo_.ReparationState = ReadOnly Then
    MsgBox "«·‘ﬂÊÏ ··ﬁ—«¡Â ›ﬁÿ, „  —ÕÌ· «·‘ﬂÊÏ ··„Õ«”»Â", vbExclamation, "«·‘ﬂÊÏ „—Õ·Â"
    Exit Sub
End If

If MsgBox("Â· √‰  „ √ﬂœ „‰ Õ–› «·«’·«Õ", vbYesNo + vbDefaultButton2, "Õ–›") = vbYes Then
    If reparationViewModelInfo_.repNo <> 0 Then
        
        reparationViewModelInfo_.ReparationState = DeleteRecord
        SaveChanges False, EnumState.DeleteRecord
        FillControlsFromSql RsNavigator
        EnableCmds True, True, True, False, False, True
        MsgBox " „ Õ–› «·«’·«Õ »‰Ã«Õ", vbInformation, "Õ–› «·«’·«Õ"
    End If
End If

Exit Sub
ErrorHandler:
MsgBox Err.Description
End Sub

Sub ChangeStatus(state As EnumState)
    reparationViewModelInfo_.ReparationState = state
    lReparationState.Caption = maintDataService_.GetReparationState(state)
End Sub
Private Sub CmdEdit_Click()
If reparationViewModelInfo_.ReparationState = ReadOnly Then
    MsgBox "«·‘ﬂÊÏ ··ﬁ—«¡Â ›ﬁÿ, „  —ÕÌ· «·‘ﬂÊÏ ··„Õ«”»Â", vbExclamation, "«·‘ﬂÊÏ „—Õ·Â"
     
ElseIf reparationViewModelInfo_.repNo = 0 Then
    MsgBox "·«Ì„ﬂ‰ «· ⁄œÌ·,«·«’·«Õ ÃœÌœ", vbExclamation, "«·«’·«Õ ÃœÌœ"

Else
    ReparationState = UpdateRecord
    ChangeStatus (ReparationState)
    EnableCmds False, False, False, True, True, False
    EnableControls True
    TxtRegisterDate.SetFocus
    SendKeys "{home}+{end}"
    txtFindCallNo.Locked = True
End If
End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub CmdFirst_Click()

    MoveNavigator 1

End Sub

Private Sub CmdLast_Click()

    MoveNavigator 4

End Sub

Private Sub CmdNext_Click()

    MoveNavigator 3


End Sub

Private Sub CmdPrevious_Click()

    MoveNavigator 2

End Sub

Sub MoveNavigator(ByVal i As Integer)
Dim RSTemp As New ADODB.Recordset
On Error GoTo ErrorHandler
If RsNavigator.RecordCount = 0 Then Exit Sub
Select Case i
    Case 1 'First
        RsNavigator.MoveFirst
        RecNum = 1
    Case 2 'Previous
        RsNavigator.MovePrevious
        If RsNavigator.BOF Then
            RsNavigator.MoveFirst
            RecNum = 1
        Else
            RecNum = RecNum - 1
        End If
    Case 3 'Next
        RsNavigator.MoveNext
        If RsNavigator.EOF Then
            RsNavigator.MoveLast
            RecNum = RsNavigator.RecordCount
        Else
            RecNum = RecNum + 1
        End If
    Case 4 'Last
        RsNavigator.MoveLast
        RecNum = RsNavigator.RecordCount
End Select
FillControlsFromSql RsNavigator
LNavigator.Caption = LTrim(RTrim(Str(RecNum))) & "/" & LTrim(RTrim(Str(RsNavigator.RecordCount)))
Exit Sub
ErrorHandler:
MsgBox "Error In Navigator"
End Sub

Function SaveChanges(updateControls As Boolean, Optional vState) As ReparationResult
    On Error GoTo ErrorHandler
    Dim result As ReparationResult

    Set result = maintDataService_.SaveChanges(reparationViewModelInfo_, vState)
      
    If result.ReparationResultStatus Then
        If IsMissing(vState) Then
            If updateControls Then
                EnableCmds True, True, True, False, False, True
                EnableControls False
                reparationViewModelInfo_.ReparationState = DefaultRecord
                FillControls
                CmdAdd.SetFocus
                MsgBox " „ Õ›Ÿ «·«’·«Õ »‰Ã«Õ", vbInformation, "Õ›Ÿ «·«’·«Õ"
            Else
                LCount.Caption = reparationViewModelInfo_.ReparationPiecesViewModel.Count
                lReparationState.Caption = maintDataService_.GetReparationState(reparationViewModelInfo_.ReparationState)
                GetTotalsPriceForItems
            End If
        End If
    End If
    Set SaveChanges = result
    Exit Function
ErrorHandler:
    Set SaveChanges = result
'    MsgBox result.ReparationResultDescription, vbExclamation, "Œÿ√ ›Ì «· Œ“Ì‰"
End Function

Private Sub CmdSave_Click()

Dim result As ReparationResult
Set result = SaveChanges(True)
If Not result.ReparationResultStatus Then
    MsgBox result.ReparationResultDescription, vbExclamation + vbMsgBoxRight, "Œÿ√ ›Ì «· Œ“Ì‰ «Ê «·Õ‹–› √Ê «·≈÷«›‹Â"
    Dim ctrl As Control
    Set ctrl = Me.GetTheLastFocusControl
    ctrl.SetFocus
    SendKeys "{home}+{end}"
End If
End Sub


Sub FillReparationModel()
On Error GoTo ErrorHandler
    With reparationViewModelInfo_
        If .ReparationState = NewRecord Then
            reparationInfo_.repNo = 0
        Else
            reparationInfo_.repNo = .repNo
        End If
        
    End With
Exit Sub
ErrorHandler:
MsgBox Err.Description
End Sub

Private Sub CmdSearch_Click()
MoveToRec SearchRec
End Sub

Private Sub DcboCliRecever_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    DcboReparationStatus.SetFocus
End If
End Sub

Private Sub DcboCliRecever_LostFocus()
If Me.DcboCliRecever.MatchedWithList Then
    reparationViewModelInfo_.CliRecever = DcboCliRecever.BoundText
Else
     reparationViewModelInfo_.CliRecever = Null
End If
End Sub

Private Sub DcboPayMethod_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtNotes.SetFocus
    SendKeys "{home}+{end}"
End If
End Sub

Private Sub DcboPayMethod_LostFocus()
If Me.DcboPayMethod.MatchedWithList Then
    reparationViewModelInfo_.Cash = DcboPayMethod.BoundText
Else
    reparationViewModelInfo_.Cash = Null
End If
End Sub

Private Sub DcboReparationStatus_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtVoltBefor.SetFocus
    SendKeys "{home}+{end}"
End If
End Sub

Private Sub DcboReparationStatus_LostFocus()
If Me.DcboReparationStatus.MatchedWithList Then
    reparationViewModelInfo_.CallStatus = DcboReparationStatus.BoundText
Else
    reparationViewModelInfo_.CallStatus = Null
End If
End Sub

Private Sub FlexCurrentDamge_KeyDown(KeyCode As Integer, Shift As Integer)
Dim FirstRow As Integer, LastRow As Integer, Vrow As Integer

With Me.FlexCurrentDamge
If KeyCode = vbKeyDelete Then
    If MsgBox("Â· √‰  „ √ﬂœ „‰ ⁄„·Ì… «·Õ–›", vbYesNo + vbDefaultButton2, "Õ–› «·”Ã·«  «·„Õœœ…") = vbYes Then
        If .Row >= .RowSel Then
            FirstRow = .Row
            LastRow = .RowSel
        Else
            FirstRow = .RowSel
            LastRow = .Row
        End If
        For i = FirstRow To LastRow Step -1
            If i = 0 Then Exit For
             reparationViewModelInfo_.ReparationWorksViewModel.Remove (.TextMatrix(i, ColReparationNo))
            .RemoveItem i
        Next
        .Col = ColReparationNo
        .SetFocus
    End If
End If
End With
End Sub

Private Sub FlexReparatonItems_KeyDown(KeyCode As Integer, Shift As Integer)
Dim FirstRow As Integer, LastRow As Integer, Vrow As Integer
With Me.FlexReparatonItems
If KeyCode = vbKeyDelete Then
    If MsgBox("Â· √‰  „ √ﬂœ „‰ ⁄„·Ì… «·Õ–›", vbYesNo + vbDefaultButton2, "Õ–› «·”Ã· «·„Õœœ") = vbYes Then
        reparationViewModelInfo_.ReparationPiecesViewModel.Remove (.TextMatrix(.Row, ColStkNo))
        .RemoveItem .Row
        SaveChanges False
    End If
End If
End With

End Sub

Function GetTheLastFocusControl() As Control
    Set GetTheLastFocusControl = Me.ActiveControl
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyA And Shift = vbCtrlMask Then
        If CmdAdd.Enabled Then
            
            CmdAdd_Click
        End If
    End If
    If KeyCode = vbKeyS And Shift = vbCtrlMask Then
        
        If cmdSave.Enabled Then
            CmdSave_Click
        End If
    End If
End Sub

Private Sub Form_Load()
init
End Sub

Sub MoveToRec(CallNo As Double)
With RsNavigator
If RsNavigator.RecordCount = 0 Then
    ClearControls
    RecNum = 0
    LNavigator.Caption = ""
    Exit Sub
End If
.MoveLast
RecNum = RsNavigator.RecordCount
    Do While Not .BOF
       If !CallNo <> CallNo Then
            .MovePrevious
            RecNum = RecNum - 1
        Else
            FillControlsFromSql RsNavigator
            LNavigator.Caption = LTrim(RTrim(Str(RecNum))) & "/" & LTrim(RTrim(Str(RsNavigator.RecordCount)))
            Exit Sub
       End If
    Loop
End With
End Sub

Function SearchRec() As Double
On Error GoTo ErrorHandler
Dim i As Double
i = InputBox("√œŒ· —ﬁ„ «·‘ﬂÊÏ", "«·»ÕÀ ⁄‰ —ﬁ„ «·‘ﬂÊÏ")
If Val(i) <> 0 Then
    SearchRec = i
Else
    SearchRec = -1
End If
Exit Function
ErrorHandler:
SearchRec = -1
End Function

Sub FillControlsFromSql(rs As Recordset)
On Error GoTo ErrorHandler
If rs.RecordCount <> 0 Then
    Set reparationViewModelInfo_ = maintDataService_.GetReparationByCallNo(rs!CallNo)
    FillControls
Else
    ClearControls
End If
Exit Sub
ErrorHandler:
MsgBox Err.Description

End Sub
Sub FillControls()
On Error GoTo ErrorHandler
Ok = False

With reparationViewModelInfo_
    txtFindCallNo.Text = .CallNo
    LRepNbr.Caption = IIf(.repNo = 0, "", .repNo)
    LteamName.Tag = .TeamNo
    LteamName.Caption = .TeamName
    LClientName.Caption = .ClientName & ""
    LRepDate.Caption = .RepDate
    If Not IsNull(.regestdate) And IsDate(.regestdate) Then
    
        TxtRegisterDate.Text = .regestdate
    Else
        .regestdate = Format(Now, "dd/mm/yyyy")
        TxtRegisterDate.Text = .regestdate
    End If
    txtRepTimeBegin.Text = IIf(IsNull(.RepTimeBegin), "__:__", Format(.RepTimeBegin, "HH:MM"))
    txtRepTimeEnd.Text = IIf(IsNull(.RepTimeEnd), "__:__", Format(.RepTimeEnd, "HH:MM"))
    DcboCliRecever.BoundText = .CliRecever
    DcboReparationStatus.BoundText = .CallStatus
    txtVoltBefor.Text = .VoltBefor
    txtVoltAfter.Text = .VoltAfter
    txtNotes.Text = .Notes
    txtDescription.Text = .Description
    txtRepPrice.Text = .RepPrice
    DcboPayMethod.BoundText = .Cash
    TxtProductName.Tag = .AdhamProductViewModel.ProdFamNo
    TxtProductName.Text = .AdhamProductViewModel.ProdFamName
    TxtModelName.Text = .AdhamProductViewModel.Symbol
    TxtModelName.Tag = .AdhamProductViewModel.ItemNo
    txtProdSerNo.Text = .AdhamProductViewModel.BarcodeSerialNo
    If Not IsNull(.AdhamProductViewModel.ProdDate) And IsDate(.AdhamProductViewModel.ProdDate) Then
        txtProductionDate.Text = Format(.AdhamProductViewModel.ProdDate, "dd/mm/yyyy")
    Else
       txtProductionDate.Text = "__/__/____"
    End If
    txtGasNo.Text = .AdhamProductViewModel.GasNo
    If Not IsNull(.AdhamProductViewModel.ProdPurchaseDate) And IsDate(.AdhamProductViewModel.ProdPurchaseDate) Then
        txtPurchaseDate.Text = Format(.AdhamProductViewModel.ProdPurchaseDate, "dd/mm/yyyy")
    Else
        txtPurchaseDate.Text = "__/__/____"
    End If

        
    LModStockNo.Caption = .AdhamProductViewModel.ItemNo
    LEmployeeName.Caption = .FullName
    LCount.Caption = .ReparationPiecesViewModel.Count
    lReparationState.Caption = maintDataService_.GetReparationState(reparationViewModelInfo_.ReparationState)
'    LTotalItems.Caption = maintDataService_.GetTotalFeesReparation(reparationViewModelInfo_)
'    LTotal.Caption = .RepPrice
'    Ldiff.Caption = Val(LTotal.Caption) - Val(LTotalItems.Caption)
    GetTotalsPriceForItems
    LEmployeeName.Caption = .FullName
    LEmployeeName.Tag = .empNo
    
    TxtCurrentDamage.Text = ""
    FillGrid FlexCurrentDamge, reparationViewModelInfo_, 1
    TxtStkName.Text = ""
    LStkName.Caption = ""
    TxtStkQty.Text = ""
    LBalance.Caption = ""
    LDealerPrice.Caption = ""
    FillGrid FlexReparatonItems, reparationViewModelInfo_, 2
    FillFormating 2, FlexCurrentDamge
    FillFormating 3, FlexReparatonItems
End With
Ok = True
Exit Sub
ErrorHandler:
MsgBox Err.Description
End Sub

Sub InitNavigator()
    Set RsNavigator = maintDataService_.GetAllReparations
End Sub

Sub FillGrid(FlexGrid As VSFlexGrid, reprationInfo As ReparationViewModel, vChoise As Integer)
FlexGrid.Rows = 1
With FlexGrid
    Select Case vChoise
        Case 1:  'ReparationWorks
            Dim reparationWorkInfo As ReparationWorkViewModel
            For Each reparationWorkInfo In reprationInfo.ReparationWorksViewModel
                .AddItem ""
                .TextMatrix(.Rows - 1, ColReparationNo) = reparationWorkInfo.RepTypeNo
                .TextMatrix(.Rows - 1, ColReparationDescription) = reparationWorkInfo.RepTypeDescription
            Next
        Case 2 'ReparationItems
            Dim reparationPieceInfo As ReparationPieceViewModel
            For Each reparationPieceInfo In reprationInfo.ReparationPiecesViewModel
                .AddItem ""
                .TextMatrix(.Rows - 1, ColStkNo) = reparationPieceInfo.stkno
                .TextMatrix(.Rows - 1, ColStkName) = reparationPieceInfo.StkName
                .TextMatrix(.Rows - 1, ColEngineBarcode) = reparationPieceInfo.EngineBarcode
                .TextMatrix(.Rows - 1, ColQty) = reparationPieceInfo.Qty
                .TextMatrix(.Rows - 1, ColDealerPrice) = reparationPieceInfo.Price
            Next
        End Select
End With
End Sub


Sub init()
    top = 0
    left = 0
    Ok = True
    Script01_UpdateAdhamProducts
    ReparationState = DefaultRecord
    FillCombos
    InitNavigator
    EnableControls False
    FlexCurrentDamge.Rows = 1
    FillFormating 2, FlexCurrentDamge
    Me.FlexReparatonItems.Rows = 1
    FillFormating 3, FlexReparatonItems
    MoveNavigator 4
    Grid.Rows = 1
    Grid.SelectionMode = flexSelectionListBox

End Sub

Sub ClearControls()
Ok = False
    
txtFindCallNo.Text = ""
LRepNbr.Caption = ""
LClientName.Caption = ""
TxtRegisterDate.Text = Format(Now, "dd/mm/yyyy")
LteamName.Caption = ""
txtRepTimeBegin.Text = "__:__"
txtRepTimeEnd.Text = "__:__"
DcboCliRecever.BoundText = -1
DcboReparationStatus.BoundText = -1
txtVoltBefor.Text = ""
txtVoltAfter.Text = ""
txtNotes.Text = ""
txtDescription.Text = ""
txtRepPrice.Text = ""
DcboPayMethod.BoundText = -1

ClearProductInfo



TxtCurrentDamage.Text = ""
TxtStkName.Text = ""
LStkName.Caption = ""
TxtEngineBarcode.Text = ""
LBalance.Caption = ""
TxtStkQty.Text = ""
LDealerPrice.Caption = ""
LCount.Caption = ""
LTotalItems.Caption = ""
LTotal.Caption = ""
Ldiff.Caption = ""
lReparationState.Caption = ""
LEmployeeName.Caption = ""
FlexCurrentDamge.Rows = 1
FlexReparatonItems.Rows = 1

Ok = True
End Sub

Sub ClearProductInfo()
    TxtProductName.Tag = ""
    TxtProductName.Text = ""
    TxtModelName.Text = ""
    txtProdSerNo.Text = ""
    txtProductionDate.Text = "__/__/____"
    LModStockNo.Caption = ""
    txtGasNo.Text = ""
    txtPurchaseDate.Text = "__/__/____"
End Sub
Sub EnableCmds(FAdd As Boolean, FEdit As Boolean, FDelete As Boolean, FSave As Boolean, FUndo As Boolean, FNavigator As Boolean)
    CmdAdd.Enabled = FAdd
    CmdEdit.Enabled = FEdit
    CmdDelete.Enabled = FDelete
    cmdSave.Enabled = FSave
    CmdCancel.Enabled = FUndo
     Me.NavigatorFrame.Enabled = FNavigator
End Sub

Sub EnableControls(FControl As Boolean)
Dim ctrl As Control
For Each ctrl In Me.Controls
    If TypeOf ctrl Is TextBox Or TypeOf ctrl Is MaskEdBox Or TypeOf ctrl Is CheckBox Or TypeOf ctrl Is VSFlexGrid Or TypeOf ctrl Is DataCombo Then
        ctrl.Enabled = FControl
    End If
Next
End Sub

Sub FillCombos()
Dim RsRecievers As New ADODB.Recordset
sqlText = "select SerGroupRC ,ltrim(rtrim(GroupRcName)) as GroupRcName from MaintGroupReceivers"
Set RsRecievers = de.con.Execute(sqlText)
Set DcboCliRecever.RowSource = RsRecievers
DcboCliRecever.listField = "GroupRcName"
DcboCliRecever.BoundColumn = "SerGroupRC"


Dim rsReparationStatus As New ADODB.Recordset
sqlText = "select no , statues from ReparationStatus"
Set rsReparationStatus = de.con.Execute(sqlText)
Set DcboReparationStatus.RowSource = rsReparationStatus
DcboReparationStatus.listField = "statues"
DcboReparationStatus.BoundColumn = "no"


Dim RsPaymentMethod As New ADODB.Recordset
sqlText = "select no , name  from paymethod"
Set RsPaymentMethod = de.con.Execute(sqlText)
Set DcboPayMethod.RowSource = RsPaymentMethod
DcboPayMethod.listField = "name"
DcboPayMethod.BoundColumn = "no"

End Sub



Private Sub Grid_RowColChange()
On Error GoTo ErrorHandler
If Flag Then
    Ok = False
    With Grid
       Select Case Pos
        Case 1
            txtNotes.Tag = .TextMatrix(.Row, ColNo)
            txtNotes.Text = .TextMatrix(.Row, ColName)
        Case 2
            txtDescription.Tag = .TextMatrix(.Row, ColNo)
            txtDescription.Text = .TextMatrix(.Row, ColName)
        Case 3
            TxtModelName.Tag = .TextMatrix(.Row, ColNo)
            TxtModelName.Text = .TextMatrix(.Row, ColName)

        Case 4
            TxtCurrentDamage.Tag = .TextMatrix(.Row, ColNo)
            TxtCurrentDamage.Text = .TextMatrix(.Row, ColName)
        Case 5
            TxtStkName.Text = .TextMatrix(.Row, ColNo)
            LStkName.Caption = .TextMatrix(.Row, ColName)
        Case 6
            TxtProductName.Tag = .TextMatrix(.Row, ColNo)
            TxtProductName.Text = .TextMatrix(.Row, ColName)
       End Select
    End With
    Ok = True
End If
Exit Sub
ErrorHandler:
MsgBox Err.Description

End Sub

Private Sub TxtCurrentDamage_Change()
Dim tableSqlText As String
tableSqlText = "(select reptypeno , reptypedescription from CoReparationType c1 where  left(ltrim(rtrim(RepTypeNo)),2)='" & Right("00" + LTrim(RTrim(Str(reparationViewModelInfo_.AdhamProductViewModel.ProdFamNo))), 2) & "') CoReparationType"
Search TxtCurrentDamage, 1, tableSqlText, "reptypedescription", "reptypeno", True
End Sub

Private Sub TxtCurrentDamage_GotFocus()
ChangeToArabic
Pos = 4
End Sub

Private Sub TxtCurrentDamage_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then MoveCursor KeyCode, Grid
End Sub
'Function SaveRec() As Boolean
'On Error GoTo errorhandler
'    SaveChanges True
'SaveRec = True
'Exit Function
'errorhandler:
'MsgBox Err.Description
'SaveRec = False
'End Function

Sub ClearReparationTypeWorks()
    Ok = False
    TxtCurrentDamage.Text = ""
    TxtCurrentDamage.SetFocus
    Ok = True
End Sub
Private Sub TxtCurrentDamage_KeyPress(KeyAscii As Integer)
On Error GoTo errrohandler
If KeyAscii = 13 Then
        If Grid.Visible Then
            Ok = False
            TxtCurrentDamage.Tag = IIf(Grid.SelectedRows = 0, Grid.TextMatrix(1, ColNo), Grid.TextMatrix(Grid.Row, ColNo))
            TxtCurrentDamage.Text = IIf(Grid.SelectedRows = 0, Grid.TextMatrix(1, ColName), Grid.TextMatrix(Grid.Row, ColName))
            Ok = True
           If Not maintDataService_.ReparationWorkTypeNoExists(reparationViewModelInfo_.ReparationWorksViewModel, TxtCurrentDamage.Tag) Then
                Dim result As ReparationResult
                reparationViewModelInfo_.ReparationWorksViewModel.Add TxtCurrentDamage.Tag, TxtCurrentDamage.Text, NewRecord, TxtCurrentDamage.Tag
                Set result = SaveChanges(False)
                If Not result.ReparationResultStatus Then
                    reparationViewModelInfo_.ReparationWorksViewModel.Remove TxtCurrentDamage.Tag
                    MsgBox result.ReparationResultDescription, vbExclamation + vbMsgBoxRight, "Œÿ√ ›Ì «· Œ“Ì‰ «Ê «·Õ‹–› √Ê «·≈÷«›‹Â"
                Else
                    LRepNbr.Caption = reparationViewModelInfo_.repNo
                    ChangeStatus UpdateRecord
                    FillGrid FlexCurrentDamge, reparationViewModelInfo_, 1
                    FillFormating 2, FlexCurrentDamge
                End If
            End If
            Grid.Visible = False
            TxtCurrentDamage.Text = ""
            TxtCurrentDamage.SetFocus

        Else
            TxtCurrentDamage.Text = ""
            TxtStkName.SetFocus
            SendKeys "{home}+{end}"
        End If
        
End If
Exit Sub
errrohandler:
MsgBox Err.Description
End Sub


Private Sub txtDescription_Change()
Search txtDescription, 1, "adhamreparation", "RepName", "repnum", True
End Sub

Private Sub txtDescription_GotFocus()
ChangeToArabic
Pos = 2
End Sub

Private Sub txtDescription_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then MoveCursor KeyCode, Grid
End Sub

Private Sub txtDescription_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        If Grid.Visible Then
            Ok = False
            txtDescription.Tag = IIf(Grid.SelectedRows = 0, Grid.TextMatrix(1, ColNo), Grid.TextMatrix(Grid.Row, ColNo))
            txtDescription.Text = IIf(Grid.SelectedRows = 0, Grid.TextMatrix(1, ColName), Grid.TextMatrix(Grid.Row, ColName))
            Ok = True
            Grid.Visible = False
            txtDescription.SetFocus
            SendKeys "{End}"
        Else
            TxtProductName.SetFocus
            SendKeys "{Home}+{End}"
        End If
End If
End Sub

Private Sub txtDescription_LostFocus()
    reparationViewModelInfo_.Description = txtDescription.Text
End Sub

Private Sub TxtEngineBarcode_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        TxtStkQty.SetFocus
        SendKeys "{End}"
End If
End Sub

Private Sub txtFindCallNo_GotFocus()
ChangeToEnglish
End Sub

Private Sub txtFindCallNo_KeyPress(KeyAscii As Integer)
On Error GoTo ErrorHandler
If KeyAscii = 13 Then
    TxtRegisterDate.SetFocus
    SendKeys "{home}+{end}"
End If
Exit Sub
ErrorHandler:
MsgBox Err.Description
End Sub

Private Sub txtFindCallNo_LostFocus()
On Error GoTo ErrorHandler
    If Val(txtFindCallNo.Text) = 0 Then Exit Sub
    If ReparationState <> UpdateRecord Then
        Set reparationViewModelInfo_ = maintDataService_.GetReparationByCallNo(txtFindCallNo.Text)
        If reparationViewModelInfo_.CallNo = 0 Then
            MsgBox "—ﬁ„ «·‘ﬂÊÏ Â–« €Ì— „ÊÃÊœ", vbExclamation + vbMsgBoxRight, "Œÿ√"
            ClearControls
            txtFindCallNo.SetFocus
            SendKeys "{home}+{end}"
            Exit Sub
        End If
    
        If reparationViewModelInfo_.Carried Then
            FillControls
            EnableCmds True, True, True, False, False, True
            EnableControls False
            CmdAdd.SetFocus
            Exit Sub
        End If
        If reparationViewModelInfo_.repNo <> 0 Then
            ChangeStatus EnumState.UpdateRecord
        End If
    End If
    FillControls
    TxtRegisterDate.SetFocus
    SendKeys "{home}+{end}"
Exit Sub
ErrorHandler:
MsgBox Err.Description
End Sub

Private Sub txtGasNo_LostFocus()
If Me.txtGasNo.Text <> "" Then
    reparationViewModelInfo_.AdhamProductViewModel.GasNo = txtGasNo.Text
End If
End Sub

Private Sub TxtModelName_Change()
Dim tableSqlText As String
tableSqlText = "(select ItemNo , Symbol from AdhamModels a1  inner join AdhamProductFamily a2 on a1.FamNo = a2.ProdFamNo  where FamNo=" & reparationViewModelInfo_.AdhamProductViewModel.ProdFamNo & ") AdhamModels"
Search TxtModelName, 1, tableSqlText, "Symbol", "ItemNo", True
End Sub

Private Sub TxtModelName_GotFocus()
ChangeToEnglish
Pos = 3
End Sub

Private Sub TxtModelName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then MoveCursor KeyCode, Grid
End Sub

Private Sub TxtModelName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        If Grid.Visible Then
            Ok = False
            TxtModelName.Tag = IIf(Grid.SelectedRows = 0, Grid.TextMatrix(1, ColNo), Grid.TextMatrix(Grid.Row, ColNo))
            TxtModelName.Text = IIf(Grid.SelectedRows = 0, Grid.TextMatrix(1, ColName), Grid.TextMatrix(Grid.Row, ColName))
            Ok = True
            Grid.Visible = False
        ElseIf Not Grid.Visible And TxtModelName.Tag = "" And TxtModelName.Tag = "0" Then
            TxtModelName.Tag = ""
            TxtModelName.Text = ""
        End If
        txtProdSerNo.SetFocus
        SendKeys "{home}+{end}"
End If
End Sub

Private Sub txtGasNo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtPurchaseDate.SetFocus
    SendKeys "{home}+{end}"
End If
End Sub


Private Sub TxtModelName_LostFocus()

If Me.TxtModelName.Tag <> "" And Val(Me.TxtModelName.Tag) <> 0 Then
    Dim modelNo As Integer
    modelNo = maintDataService_.GetModelNo(TxtModelName.Tag)
    LModStockNo.Caption = IIf(TxtModelName.Tag = 0 Or TxtModelName.Tag = "", "", TxtModelName.Tag)
    reparationViewModelInfo_.AdhamProductViewModel.ModNo = modelNo
    reparationViewModelInfo_.AdhamProductViewModel.ItemNo = TxtModelName.Tag
Else
    LModStockNo.Caption = ""
    reparationViewModelInfo_.AdhamProductViewModel.ModNo = 0
    reparationViewModelInfo_.AdhamProductViewModel.ItemNo = ""
    reparationViewModelInfo_.AdhamProductViewModel.BarcodeSerialNo = ""
    reparationViewModelInfo_.AdhamProductViewModel.GasNo = ""
    reparationViewModelInfo_.AdhamProductViewModel.ProdBaCodeNo = ""
    reparationViewModelInfo_.AdhamProductViewModel.ProdDate = Null
    reparationViewModelInfo_.AdhamProductViewModel.ProdPurchaseDate = Null
    reparationViewModelInfo_.AdhamProductViewModel.ProdFamName = ""
    reparationViewModelInfo_.AdhamProductViewModel.Symbol = ""
    
    txtProdSerNo.Text = ""
    txtProductionDate.Text = "__/__/____"
    txtPurchaseDate.Text = "__/__/____"
    txtGasNo.Text = ""
    
    
End If
End Sub

Private Sub TxtNotes_Change()
Search txtNotes, 1, "adhamreparation", "RepName", "repnum", True
End Sub
Sub Search(sender As Control, Pos As Integer, tableName As String, listField As String, dataMember As String, Optional isChangeCursor = True)
On Error GoTo ErrorHandler
Dim RsSearch As New ADODB.Recordset


If Ok Then
    sender.Tag = 0

    If sender.Text = "" Then
        Grid.Visible = False
        Exit Sub
    End If
    Flag = False
    sqlText = "Select top 30 " & dataMember & "," & listField & " From " & tableName & " Where "
    sqlText = sqlText & dataMember & " Like " & LikeExpression(sender.Text) & " Or " & listField & " Like " & LikeExpression(sender.Text)
    Set RsSearch = de.con.Execute(sqlText)
    If RsSearch.RecordCount > 0 Then
        Set Grid.DataSource = RsSearch
        'Grid.Row = 0
        FillFormating Pos, Grid
        If isChangeCursor Then ChangeCursor sender
        Grid.Visible = True
    Else
        sender.Tag = 0
        Grid.Visible = False
    End If
    Flag = True
End If


   Exit Sub
ErrorHandler:
MsgBox Err.Description
End Sub

Private Sub txtNotes_GotFocus()
ChangeToArabic
Pos = 1
End Sub

Private Sub txtNotes_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then MoveCursor KeyCode, Grid
End Sub

Private Sub txtNotes_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        If Grid.Visible Then
            Ok = False
            txtNotes.Tag = IIf(Grid.SelectedRows = 0, Grid.TextMatrix(1, ColNo), Grid.TextMatrix(Grid.Row, ColNo))
            txtNotes.Text = IIf(Grid.SelectedRows = 0, Grid.TextMatrix(1, ColName), Grid.TextMatrix(Grid.Row, ColName))
            Ok = True
            Grid.Visible = False
            txtNotes.SetFocus
            SendKeys "{End}"
        Else
            txtDescription.SetFocus
            SendKeys "{Home}+{End}"
        End If
End If
End Sub

Private Sub txtNotes_LostFocus()
    reparationViewModelInfo_.Notes = txtNotes.Text
End Sub

Private Sub txtProdSerNo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtProductionDate.SetFocus
    SendKeys "{home}+{end}"
End If
End Sub

Private Sub txtProdSerNo_LostFocus()
If Me.txtProdSerNo.Text <> "" Then
    reparationViewModelInfo_.AdhamProductViewModel.BarcodeSerialNo = txtProdSerNo.Text
End If
End Sub

Private Sub txtProductionDate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtGasNo.SetFocus
SendKeys "{home}+{end}"
End If

End Sub

Private Sub txtProductionDate_LostFocus()
If IsDate(Me.txtProductionDate.Text) Then
    reparationViewModelInfo_.AdhamProductViewModel.ProdDate = txtProductionDate.Text
    reparationViewModelInfo_.AdhamProductViewModel.ProdPurchaseDate = DateAdd("d", 30, reparationViewModelInfo_.AdhamProductViewModel.ProdDate)
   txtPurchaseDate.Text = reparationViewModelInfo_.AdhamProductViewModel.ProdPurchaseDate
Else
    reparationViewModelInfo_.AdhamProductViewModel.ProdDate = Null

End If
End Sub

Private Sub TxtProductName_Change()
Search TxtProductName, 1, "adhamproductfamily", "ProdFamNameA", "ProdFamNo", True
End Sub

Private Sub TxtProductName_GotFocus()
ChangeToArabic
Pos = 6
End Sub

Private Sub TxtProductName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then MoveCursor KeyCode, Grid
End Sub

Private Sub TxtProductName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Grid.Visible Then
        Ok = False
        TxtProductName.Tag = IIf(Grid.SelectedRows = 0, Grid.TextMatrix(1, ColNo), Grid.TextMatrix(Grid.Row, ColNo))
        TxtProductName.Text = IIf(Grid.SelectedRows = 0, Grid.TextMatrix(1, ColName), Grid.TextMatrix(Grid.Row, ColName))
        Ok = True
        Grid.Visible = False
    End If
    TxtModelName.SetFocus
    SendKeys "{home}+{end}"

End If
End Sub

Private Sub TxtProductName_LostFocus()
    reparationViewModelInfo_.AdhamProductViewModel.ProdFamNo = TxtProductName.Tag
    
End Sub

Private Sub txtPurchaseDate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    TxtCurrentDamage.SetFocus
    SendKeys "{home}+{end}"
End If
End Sub

Private Sub txtPurchaseDate_LostFocus()
If IsDate(Me.txtPurchaseDate.Text) Then
    reparationViewModelInfo_.AdhamProductViewModel.ProdPurchaseDate = txtPurchaseDate.Text
Else
    reparationViewModelInfo_.AdhamProductViewModel.ProdPurchaseDate = Null
End If
End Sub

Private Sub TxtRegisterDate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtRepTimeBegin.SetFocus
SendKeys "{home}+{end}"
End If
End Sub

Private Sub TxtRegisterDate_LostFocus()
On Error GoTo ErrorHandler
If IsDate(TxtRegisterDate.Text) Then
    reparationViewModelInfo_.regestdate = TxtRegisterDate.Text
Else
    reparationViewModelInfo_.regestdate = Format(Now, "dd/mm/yyyy")
End If
Exit Sub
ErrorHandler:
MsgBox Err.Description
End Sub

Private Sub txtRepPrice_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    DcboPayMethod.SetFocus
End If
End Sub
Sub GetTotalsPriceForItems()
On Error GoTo ErrorHandler
    Dim totalItems As Double
    Dim reparationPrice As Double
    totalItems = maintDataService_.GetTotalFeesReparation(reparationViewModelInfo_)
    reparationPrice = reparationViewModelInfo_.RepPrice
    LTotalItems.Caption = totalItems
    LTotal.Caption = reparationPrice
    If reparationViewModelInfo_.Cash <> PayMethodEnum.Warranty Then
        Ldiff.Caption = reparationPrice - totalItems
    Else
      Ldiff.Caption = 0
    End If
Exit Sub
ErrorHandler:
MsgBox Err.Description
End Sub

Private Sub txtRepPrice_LostFocus()
If Me.txtRepPrice.Text <> "" And IsNumeric(txtRepPrice.Text) Then
    reparationViewModelInfo_.RepPrice = txtRepPrice.Text
    GetTotalsPriceForItems
Else
    reparationViewModelInfo_.RepPrice = 0
End If
End Sub

Private Sub txtRepTimeBegin_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtRepTimeEnd.SetFocus
SendKeys "{home}+{end}"
End If
End Sub



Private Sub txtRepTimeBegin_LostFocus()
If IsDate(txtRepTimeBegin.Text) Then
    reparationViewModelInfo_.RepTimeBegin = txtRepTimeBegin.Text
Else
   reparationViewModelInfo_.RepTimeBegin = Null
End If
End Sub

Private Sub txtRepTimeEnd_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    DcboCliRecever.SetFocus
    SendKeys "{home}+{end}"
End If
End Sub


Private Sub txtRepTimeEnd_LostFocus()
If IsDate(txtRepTimeEnd.Text) Then
    reparationViewModelInfo_.RepTimeEnd = txtRepTimeEnd.Text
Else
    reparationViewModelInfo_.RepTimeEnd = Null
End If
End Sub

Private Sub TxtStkName_Change()
Search TxtStkName, 1, "CoStock", "StkName", "StkNo", True
End Sub

Private Sub TxtStkName_GotFocus()
ChangeToArabic
Pos = 5
End Sub

Private Sub TxtStkName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then MoveCursor KeyCode, Grid
End Sub

Private Sub TxtStkName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        If Grid.Visible Then
            Ok = False
            TxtStkName.Tag = IIf(Grid.SelectedRows = 0, Grid.TextMatrix(1, ColNo), Grid.TextMatrix(Grid.Row, ColNo))
            TxtStkName.Text = IIf(Grid.SelectedRows = 0, Grid.TextMatrix(1, ColNo), Grid.TextMatrix(Grid.Row, ColNo))
            Ok = True
            Grid.Visible = False
            TxtStkQty.SetFocus
            SendKeys "{End}"
        ElseIf TxtStkName.Text <> "" Then
            ClearProductItems
        Else
            cmdSave.SetFocus
        End If
End If
End Sub

Private Sub TxtStkName_LostFocus()
If TxtStkName.Tag <> "0" And TxtStkName.Tag <> "" Then
    LStkName.Caption = IIf(Grid.SelectedRows = 0, Grid.TextMatrix(1, ColName), Grid.TextMatrix(Grid.Row, ColName))
    LBalance.Caption = stockDataService_.GetBalanceItemsByStkNoAndTeamNo(TxtStkName.Text, reparationViewModelInfo_.TeamNo)
    LDealerPrice.Caption = stockDataService_.GetItemNumberDealerPriceWithDiscount(TxtStkName.Tag)
End If
End Sub

Private Sub TxtStkQty_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
            If Not maintDataService_.ReparationPieceExists(reparationViewModelInfo_.ReparationPiecesViewModel, Me.TxtStkName.Tag) Then
                    Dim stkNoBalanceForTeamNo As Double
                    stkNoBalanceForTeamNo = stockDataService_.GetBalanceItemsByStkNoAndTeamNo(TxtStkName.Tag, reparationViewModelInfo_.TeamNo)
                    If IsNumeric(TxtStkQty.Text) And Val(TxtStkQty.Text) > 0 And stkNoBalanceForTeamNo >= Val(TxtStkQty.Text) And Val(LDealerPrice.Caption) > 0 Then
                        Dim result As ReparationResult
                        reparationViewModelInfo_.ReparationPiecesViewModel.Add stockDataService_.GetStkId(TxtStkName.Tag), TxtStkName.Tag, TxtStkQty.Text, LDealerPrice.Caption, LStkName.Caption, 0, NewRecord, TxtEngineBarcode.Text, TxtStkName.Tag
                        Set result = SaveChanges(False)
                        If Not result.ReparationResultStatus Then
                            reparationViewModelInfo_.ReparationPiecesViewModel.Remove TxtStkName.Tag
                            MsgBox result.ReparationResultDescription, vbExclamation + vbMsgBoxRight, "Œÿ√ ›Ì «· Œ“Ì‰ «Ê «·Õ‹–› √Ê «·≈÷«›‹Â"
                        Else
                            LRepNbr.Caption = reparationViewModelInfo_.repNo
                            ChangeStatus UpdateRecord
                            FillGrid Me.FlexReparatonItems, reparationViewModelInfo_, 2
                            FillFormating 3, FlexReparatonItems
                        End If
                    Else
                        MsgBox "«·—’Ìœ ·«Ì”„Õ ,«Ê ·«ÌÊÃœ ”⁄— ··„«œÂ", vbExclamation, "·«Ì„ﬂ‰ «·≈œŒ«·"
                    End If
              Else
                MsgBox "«·—ﬁ„ «·„Õ“‰Ì „ﬂ——", vbExclamation, "„«œÂ „ﬂ——Â"
                TxtStkName.SetFocus
                SendKeys "{home}+{end}"
              End If
              ClearProductItems
End If
End Sub
Sub ClearProductItems()
    Ok = False
    TxtStkName.Text = ""
    TxtStkName.Tag = ""
    LStkName.Caption = ""
    LBalance.Caption = ""
    TxtStkQty.Text = ""
    LDealerPrice.Caption = ""
    TxtEngineBarcode.Text = ""
    TxtStkName.SetFocus
    Ok = True
End Sub

Private Sub txtVoltAfter_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtRepPrice.SetFocus
    SendKeys "{home}+{end}"
End If
End Sub

Private Sub txtVoltAfter_LostFocus()
If Me.txtVoltAfter.Text <> "" And IsNumeric(Me.txtVoltAfter.Text) Then
    reparationViewModelInfo_.VoltAfter = txtVoltAfter.Text
Else
    reparationViewModelInfo_.VoltAfter = 0
End If
End Sub

Private Sub txtVoltBefor_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtVoltAfter.SetFocus
    SendKeys "{home}+{end}"
End If
End Sub

Private Sub txtVoltBefor_LostFocus()
If Me.txtVoltBefor.Text <> "" And IsNumeric(Me.txtVoltBefor.Text) Then
    reparationViewModelInfo_.VoltBefor = txtVoltBefor.Text
Else
    reparationViewModelInfo_.VoltBefor = 0
End If
End Sub
