VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmReparation2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   8850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11760
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8850
   ScaleWidth      =   11760
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      DragMode        =   1  'Automatic
      Height          =   4125
      Left            =   6420
      TabIndex        =   30
      Top             =   930
      Visible         =   0   'False
      Width           =   2415
      _cx             =   4260
      _cy             =   7276
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
   Begin VB.TextBox txtNotes 
      Alignment       =   1  'Right Justify
      Height          =   450
      Left            =   5850
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   1800
      Width           =   5775
   End
   Begin VB.TextBox txtVoltAfter 
      Alignment       =   2  'Center
      Height          =   360
      Left            =   4230
      TabIndex        =   11
      Top             =   1140
      Width           =   660
   End
   Begin VB.TextBox txtVoltBefor 
      Alignment       =   2  'Center
      Height          =   360
      Left            =   4935
      TabIndex        =   10
      Top             =   1140
      Width           =   660
   End
   Begin VB.TextBox txtDescription 
      Alignment       =   1  'Right Justify
      Height          =   450
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   1800
      Width           =   5775
   End
   Begin VB.TextBox TxtModelName 
      Alignment       =   1  'Right Justify
      Height          =   450
      Left            =   5970
      RightToLeft     =   -1  'True
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   2580
      Width           =   5655
   End
   Begin VB.TextBox txtGasNo 
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
      Height          =   450
      Left            =   1110
      MaxLength       =   1
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   2580
      Width           =   795
   End
   Begin VB.TextBox TxtCurrentDamage 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   60
      RightToLeft     =   -1  'True
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   3300
      Width           =   11625
   End
   Begin VB.TextBox txtProdSerNo 
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
      Height          =   450
      Left            =   4950
      MaxLength       =   3
      TabIndex        =   17
      Top             =   2580
      Width           =   1035
   End
   Begin VB.TextBox txtRepPrice 
      Height          =   360
      Left            =   2910
      TabIndex        =   12
      Top             =   1140
      Width           =   1305
   End
   Begin VB.TextBox TxtStkName 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   7110
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Top             =   5700
      Width           =   4545
   End
   Begin VB.TextBox TxtStkQty 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   930
      RightToLeft     =   -1  'True
      TabIndex        =   27
      Top             =   5700
      Width           =   855
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
   Begin MSMask.MaskEdBox txtRepTimeBegin 
      Height          =   450
      Left            =   1620
      TabIndex        =   6
      Top             =   300
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   794
      _Version        =   393216
      AutoTab         =   -1  'True
      MaxLength       =   5
      Format          =   "hh:mm"
      Mask            =   "99:99"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtRepTimeEnd 
      Height          =   450
      Left            =   660
      TabIndex        =   7
      Top             =   300
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   794
      _Version        =   393216
      AutoTab         =   -1  'True
      MaxLength       =   5
      Format          =   "hh:mm"
      Mask            =   "99:99"
      PromptChar      =   "_"
   End
   Begin MSDataListLib.DataCombo DcboReparationStatus 
      Height          =   360
      Left            =   5610
      TabIndex        =   9
      Top             =   1140
      Width           =   2895
      _ExtentX        =   5106
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
      Left            =   750
      TabIndex        =   13
      Top             =   1140
      Width           =   2175
      _ExtentX        =   3836
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
      Left            =   8520
      TabIndex        =   8
      Top             =   1140
      Width           =   3105
      _ExtentX        =   5477
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
   Begin MSMask.MaskEdBox txtRepDate 
      Height          =   450
      Left            =   2610
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   300
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   794
      _Version        =   393216
      AutoTab         =   -1  'True
      MaxLength       =   8
      Format          =   "dd/mm/yy"
      Mask            =   "99/99/99"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtPurchaseDate 
      Height          =   450
      Left            =   30
      TabIndex        =   21
      Top             =   2580
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   794
      _Version        =   393216
      AutoTab         =   -1  'True
      MaxLength       =   8
      Format          =   "dd/mm/yy"
      Mask            =   "99/99/99"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtProductionDate 
      Height          =   450
      Left            =   3930
      TabIndex        =   18
      Top             =   2580
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   794
      _Version        =   393216
      AutoTab         =   -1  'True
      MaxLength       =   8
      Format          =   "dd/mm/yy"
      Mask            =   "99/99/99"
      PromptChar      =   "_"
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   240
      Top             =   870
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
            Picture         =   "FrmReparation2.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReparation2.frx":26D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReparation2.frx":5533
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReparation2.frx":7CD8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReparation2.frx":A5D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReparation2.frx":CAF9
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReparation2.frx":F2B1
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReparation2.frx":11CC5
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReparation2.frx":14617
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReparation2.frx":1738B
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReparation2.frx":19BDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReparation2.frx":1CA81
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReparation2.frx":1F7DB
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReparation2.frx":2217B
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReparation2.frx":250DD
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReparation2.frx":27B05
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReparation2.frx":2A4BF
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReparation2.frx":2CDF0
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReparation2.frx":2F6F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReparation2.frx":3215E
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReparation2.frx":3510F
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReparation2.frx":37A38
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReparation2.frx":3A36D
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReparation2.frx":3CAAD
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReparation2.frx":3F41C
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReparation2.frx":41CA4
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReparation2.frx":43EB3
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReparation2.frx":46810
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReparation2.frx":4903A
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReparation2.frx":4BAEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReparation2.frx":4E52A
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReparation2.frx":51040
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReparation2.frx":53FDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReparation2.frx":56DA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReparation2.frx":59A20
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReparation2.frx":5C8B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReparation2.frx":5F4F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReparation2.frx":6209F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin Threed.SSFrame SSFrame4 
      DragMode        =   1  'Automatic
      Height          =   495
      Left            =   30
      TabIndex        =   31
      Top             =   7890
      Width           =   11715
      _ExtentX        =   20664
      _ExtentY        =   873
      _Version        =   131074
      Begin Threed.SSCommand CmdExit 
         Height          =   435
         Left            =   60
         TabIndex        =   36
         Top             =   30
         Width           =   1335
         _ExtentX        =   2355
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
      Begin Threed.SSCommand cmdSave 
         Height          =   435
         Left            =   5190
         TabIndex        =   29
         Top             =   30
         Width           =   1335
         _ExtentX        =   2355
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
      Begin Threed.SSCommand CmdEdit 
         Height          =   435
         Left            =   8610
         TabIndex        =   35
         Top             =   30
         Width           =   1335
         _ExtentX        =   2355
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
      Begin Threed.SSCommand CmdCancel 
         Height          =   435
         Left            =   3480
         TabIndex        =   34
         Top             =   30
         Width           =   1335
         _ExtentX        =   2355
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
      Begin Threed.SSCommand CmdDelete 
         Height          =   435
         Left            =   6900
         TabIndex        =   33
         Top             =   30
         Width           =   1335
         _ExtentX        =   2355
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
      Begin Threed.SSCommand CmdAdd 
         Height          =   435
         Left            =   10320
         TabIndex        =   0
         Top             =   30
         Width           =   1335
         _ExtentX        =   2355
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
      Begin Threed.SSCommand CmdSearch 
         Height          =   435
         Left            =   1770
         TabIndex        =   32
         Top             =   30
         Width           =   1335
         _ExtentX        =   2355
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
   End
   Begin Threed.SSFrame SSFrame10 
      DragMode        =   1  'Automatic
      Height          =   375
      Left            =   9180
      TabIndex        =   37
      Top             =   8430
      Width           =   2475
      _ExtentX        =   4366
      _ExtentY        =   661
      _Version        =   131074
      Begin VB.CommandButton CmdFirst 
         Height          =   285
         Left            =   60
         Picture         =   "FrmReparation2.frx":64A4E
         Style           =   1  'Graphical
         TabIndex        =   41
         TabStop         =   0   'False
         ToolTipText     =   "First"
         Top             =   60
         Width           =   255
      End
      Begin VB.CommandButton CmdPrevious 
         Height          =   285
         Left            =   330
         Picture         =   "FrmReparation2.frx":64F80
         Style           =   1  'Graphical
         TabIndex        =   40
         TabStop         =   0   'False
         ToolTipText     =   "Previous"
         Top             =   60
         Width           =   255
      End
      Begin VB.CommandButton CmdNext 
         Height          =   285
         Left            =   1920
         Picture         =   "FrmReparation2.frx":6507A
         Style           =   1  'Graphical
         TabIndex        =   39
         TabStop         =   0   'False
         ToolTipText     =   "Next"
         Top             =   60
         Width           =   255
      End
      Begin VB.CommandButton CmdLast 
         Height          =   285
         Left            =   2190
         Picture         =   "FrmReparation2.frx":65174
         Style           =   1  'Graphical
         TabIndex        =   38
         TabStop         =   0   'False
         ToolTipText     =   "Last"
         Top             =   60
         Width           =   255
      End
      Begin VB.Label LNavigator 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   600
         TabIndex        =   42
         Top             =   60
         Width           =   1305
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid FlexCurrentDamge 
      Height          =   1665
      Left            =   90
      TabIndex        =   23
      Top             =   3690
      Width           =   11595
      _cx             =   20452
      _cy             =   2937
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
      Height          =   1755
      Left            =   60
      TabIndex        =   76
      Top             =   6060
      Width           =   11565
      _cx             =   20399
      _cy             =   3096
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
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
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
      Left            =   1650
      TabIndex        =   75
      Top             =   8430
      Width           =   825
   End
   Begin VB.Label LTotal 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
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
      ForeColor       =   &H000000FF&
      Height          =   360
      Left            =   60
      RightToLeft     =   -1  'True
      TabIndex        =   74
      Top             =   8430
      Width           =   1545
   End
   Begin VB.Label LTotalItems 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
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
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   6030
      TabIndex        =   73
      Top             =   8430
      Width           =   975
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
      Left            =   7035
      TabIndex        =   72
      Top             =   8430
      Width           =   885
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
      Height          =   360
      Left            =   8010
      RightToLeft     =   -1  'True
      TabIndex        =   71
      Top             =   8430
      Width           =   615
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
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   15
      Left            =   8655
      TabIndex        =   70
      Top             =   8430
      Width           =   480
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "«·œ›⁄"
      DragMode        =   1  'Automatic
      ForeColor       =   &H80000002&
      Height          =   195
      Index           =   4
      Left            =   2520
      TabIndex        =   69
      Top             =   870
      Width           =   360
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "«·„” ﬁ»·"
      Height          =   285
      Index           =   0
      Left            =   11010
      TabIndex        =   68
      Top             =   870
      Width           =   600
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "›Ê·  »⁄œ"
      Height          =   195
      Index           =   24
      Left            =   4290
      TabIndex        =   67
      Top             =   870
      Width           =   615
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "›Ê·  ﬁ»·"
      Height          =   195
      Index           =   23
      Left            =   4950
      TabIndex        =   66
      Top             =   870
      Width           =   645
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "«·„·«ÕŸ« "
      Height          =   195
      Index           =   22
      Left            =   4980
      TabIndex        =   65
      Top             =   1560
      Width           =   660
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "«·≈’·«Õ« "
      ForeColor       =   &H80000002&
      Height          =   195
      Index           =   15
      Left            =   10950
      TabIndex        =   64
      Top             =   1530
      Width           =   660
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "ﬁÌ„… «·≈’·«Õ"
      DragMode        =   1  'Automatic
      Height          =   195
      Index           =   9
      Left            =   3270
      TabIndex        =   63
      Top             =   870
      Width           =   840
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "”«⁄… «·«‰ Â«¡"
      ForeColor       =   &H80000002&
      Height          =   195
      Index           =   8
      Left            =   720
      TabIndex        =   62
      Top             =   30
      Width           =   915
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "”«⁄… «·»œ¡"
      ForeColor       =   &H80000002&
      Height          =   195
      Index           =   7
      Left            =   1860
      TabIndex        =   61
      Top             =   0
      Width           =   735
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   " «—ÌŒ «·≈’·«Õ"
      ForeColor       =   &H80000002&
      Height          =   195
      Index           =   6
      Left            =   2700
      TabIndex        =   60
      Top             =   30
      Width           =   915
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Õ«·… «·≈’·«Õ"
      Height          =   195
      Index           =   5
      Left            =   7650
      TabIndex        =   59
      Top             =   900
      Width           =   825
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "„ÊœÌ·"
      DragMode        =   1  'Automatic
      Height          =   195
      Index           =   13
      Left            =   11190
      TabIndex        =   58
      Top             =   2310
      Width           =   450
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   " ”·”·"
      DragMode        =   1  'Automatic
      Height          =   195
      Index           =   14
      Left            =   4950
      TabIndex        =   57
      Top             =   2310
      Width           =   1035
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   " /«·≈‰ «Ã"
      DragMode        =   1  'Automatic
      Height          =   195
      Index           =   16
      Left            =   3885
      TabIndex        =   56
      Top             =   2310
      Width           =   1005
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "«·„Œ“‰Ì"
      DragMode        =   1  'Automatic
      Height          =   195
      Index           =   17
      Left            =   1935
      TabIndex        =   55
      Top             =   2310
      Width           =   1935
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "«·€«“"
      DragMode        =   1  'Automatic
      Height          =   195
      Index           =   18
      Left            =   1110
      TabIndex        =   54
      Top             =   2310
      Width           =   795
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "«·‘—«¡"
      DragMode        =   1  'Automatic
      Height          =   195
      Index           =   19
      Left            =   570
      TabIndex        =   53
      Top             =   2310
      Width           =   480
   End
   Begin VB.Label lblSelectedEmployees 
      AutoSize        =   -1  'True
      Caption         =   "√⁄ÿ«· «·≈’·«Õ «·Õ«·Ì"
      DragMode        =   1  'Automatic
      Height          =   195
      Left            =   10170
      RightToLeft     =   -1  'True
      TabIndex        =   52
      Top             =   3030
      Width           =   1485
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "—ﬁ„ «·«’·«Õ"
      Height          =   195
      Left            =   9870
      RightToLeft     =   -1  'True
      TabIndex        =   51
      Top             =   60
      Width           =   780
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
      Left            =   9510
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   330
      Width           =   1125
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "«·—ﬁ„ «·„Œ“‰Ì"
      DragMode        =   1  'Automatic
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   8
      Left            =   10695
      TabIndex        =   50
      Top             =   5430
      Width           =   930
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "«·—’Ìœ"
      DragMode        =   1  'Automatic
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   9
      Left            =   2130
      TabIndex        =   49
      Top             =   5430
      Width           =   435
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "«·‘—Õ"
      DragMode        =   1  'Automatic
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   10
      Left            =   6660
      TabIndex        =   48
      Top             =   5430
      Width           =   420
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "«·ﬂ„Ì…"
      DragMode        =   1  'Automatic
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   11
      Left            =   1290
      TabIndex        =   47
      Top             =   5430
      Width           =   405
   End
   Begin VB.Label LStkName 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   2640
      TabIndex        =   25
      Top             =   5700
      Width           =   4455
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
      Height          =   360
      Left            =   1830
      TabIndex        =   26
      Top             =   5700
      Width           =   765
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "”⁄— «· «Ã—"
      DragMode        =   1  'Automatic
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   1
      Left            =   90
      TabIndex        =   46
      Top             =   5430
      Width           =   735
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
      Height          =   360
      Left            =   60
      TabIndex        =   28
      Top             =   5670
      Width           =   825
   End
   Begin VB.Label LModStockNo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Height          =   420
      Left            =   1950
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   2610
      Width           =   1935
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
      Left            =   4920
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   330
      Width           =   4545
   End
   Begin VB.Label LClient 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "«·“»Ê‰"
      Height          =   195
      Left            =   8880
      RightToLeft     =   -1  'True
      TabIndex        =   45
      Top             =   60
      Width           =   420
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "«·‘ﬂÊÏ"
      Height          =   195
      Left            =   11040
      RightToLeft     =   -1  'True
      TabIndex        =   44
      Top             =   60
      Width           =   570
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   " «—ÌŒ «·≈œŒ«·"
      ForeColor       =   &H80000002&
      Height          =   195
      Index           =   3
      Left            =   3870
      RightToLeft     =   -1  'True
      TabIndex        =   43
      Top             =   60
      Width           =   900
   End
   Begin VB.Label LEntryDate 
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
      Left            =   3660
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   330
      Width           =   1215
   End
End
Attribute VB_Name = "FrmReparation2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsNavigator As New ADODB.Recordset
Dim OK As Boolean, Flag As Boolean, Pos As Integer, RecNum  As Double

Dim mainDataService_ As New MaintDataService
Dim reparationInfo_ As New ReparationViewModel


Const ColSerNo = 1
Const ColStkNo = 2
Const ColStkName = 3
Const ColQty = 4
Const ColDealerPrice = 5

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

Sub MoveCursor(KeyCode As Integer, FlexGrid As VSFlexGrid)
On Error Resume Next
If Not FlexGrid.Visible Then Exit Sub
With FlexGrid
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
'
Sub FillFormating(ByVal i As Integer, FlexGrid As VSFlexGrid)
If i = 1 Then

    FS = "|>" + "«·—ﬁ„"
    FS = FS + "|>" + "«·«”„"

    With FlexGrid
        .FormatString = FS
        .Cols = 3

        SetColWidths ColNo, FlexGrid
        SetColWidths ColName, FlexGrid

    End With
ElseIf i = 2 Then
        FS = "|>" + "ReparationNo."
        FS = FS + "|>" + "≈”„ «·«’·«Õ"
         With FlexGrid
            .FormatString = FS
            .Cols = 3
            SetColWidths ColReparationDescription, FlexGrid
            SetColWidths ColReparationNo, FlexGrid
            '.ColWidth(ColReparationNo) = 0
    End With
ElseIf i = 3 Then
   
    FS = "|>" + "SerNo."
    FS = FS + "|>" + "«·—ﬁ„ «·„Œ“‰Ì"
    FS = FS + "|>" + "«·‘—Õ"
    FS = FS + "|>" + "«·ﬂ„Ì…"
    FS = FS + "|>" + "«·”⁄—"
     With FlexGrid
        .FormatString = FS
        .Cols = 6
        SetColWidths ColStkNo, FlexGrid
        SetColWidths ColStkName, FlexGrid
        SetColWidths ColQty, FlexGrid
        SetColWidths ColDealerPrice, FlexGrid

        .ColWidth(ColSerNo) = 0
    End With
End If
End Sub

Sub SetColWidths(ByVal ColNo As Integer, FlexGrid As VSFlexGrid)
    With FlexGrid
        .AutoSize (ColNo)
    End With
End Sub

Sub ChangeCursor(sender As Control)

    With sender
       Grid.top = .top + .Height
       Grid.left = .left
       Grid.Width = .Width
    End With

End Sub

Private Sub CmdAdd_Click()
TypeRec = True
EnableCmds False, False, False, True, True, False, False, False, False
EnableControls True
ClearControls
txtFindCallNo.SetFocus
'txtFindCallNo.Locked = False

End Sub

Function DeleteRec(repNo As Double) As Boolean
On Error GoTo errorhandler
de.con.BeginTrans
    sqlText = "Delete From MvMaintPayments Where BillNo=" & repNo
    de.con.Execute (sqlText)
    
    sqlText = "Delete From MvMaintPaymentsDetails Where BillNo=" & repNo
    de.con.Execute (sqlText)
de.con.CommitTrans
DeleteRec = True
Exit Function
errorhandler:
DeleteRec = False
MsgBox Err.Description
de.con.RollbackTrans
End Function

Function MaxRec() As Double
Dim RsMax As New ADODB.Recordset
    sqlText = "Select isnull(Max(BillNo),0)MaxBillNo From MvMaintPayments"
    Set RsMax = de.con.Execute(sqlText)
    MaxRec = RsMax!MaxBillNo
End Function

Private Sub CmdCancel_Click()
On Error GoTo errorhandler
    EnableCmds True, True, True, False, False, True, True, True, True
    EnableControls False
    If RsNavigator!repNo = Null Then
        MoveToRec Val(LRepNbr.Caption)
    Else
        MoveToRec Val(RsNavigator!repNo)
    End If
Exit Sub
errorhandler:
MsgBox Err.Description
End Sub

Private Sub CmdDelete_Click()

On Error GoTo errorhandler
    If MsgBox("Â· √‰  „ √ﬂœ „‰ Õ–› «·›« Ê—…", vbYesNo + vbDefaultButton2, "Õ–›") = vbYes Then
        If DeleteRec(RsNavigator!repNo) Then
            InitNavigator
            MoveToRec MaxRec
            EnableCmds True, True, True, False, False, True, True, True, True
        End If
    End If

Exit Sub
errorhandler:
MsgBox Err.Description
End Sub

Private Sub CmdEdit_Click()
   
TypeRec = False
EnableCmds False, False, False, True, True, False, False, False, False
EnableControls True
txtFindCallNo.SetFocus
SendKeys "{home}+{end}"
txtFindCallNo.Locked = True
End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub CmdFirst_Click()
With RsNavigator
    MoveNavigator 1
End With
End Sub

Private Sub CmdLast_Click()
With RsNavigator
    MoveNavigator 4
End With
End Sub

Private Sub CmdNext_Click()
With RsNavigator
    MoveNavigator 3
End With

End Sub

Private Sub CmdPrevious_Click()
With RsNavigator
    MoveNavigator 2
End With
End Sub

Sub MoveNavigator(ByVal i As Integer)
Dim RSTemp As New ADODB.Recordset
On Error GoTo errorhandler
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
errorhandler:
MsgBox "Error In Navigator"
End Sub


Private Sub CmdSearch_Click()
MoveToRec SearchRec
End Sub

Private Sub DcboCliRecever_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    DcboReparationStatus.SetFocus
End If
End Sub

Private Sub DcboPayMethod_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
End If
End Sub

Private Sub DcboReparationStatus_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtVoltBefor.SetFocus
    SendKeys "{home}+{end}"
End If
End Sub

Private Sub Form_Load()
Init
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
RecNum = 1
    Do While Not .BOF
       If !CallNo <> CallNo Then
            .MovePrevious
            RecNum = RecNum + 1
        Else
            FillControlsFromSql RsNavigator
            LNavigator.Caption = LTrim(RTrim(Str(RecNum))) & "/" & LTrim(RTrim(Str(RsNavigator.RecordCount)))
            Exit Sub
       End If
    Loop
End With
End Sub

Function SearchRec() As Double
On Error GoTo errorhandler
Dim i As Double
i = InputBox("√œŒ· —ﬁ„ «·‘ﬂÊÏ", "«·»ÕÀ ⁄‰ —ﬁ„ «·‘ﬂÊÏ")
If Val(i) <> 0 Then
    SearchRec = i
Else
    SearchRec = -1
End If
Exit Function
errorhandler:
SearchRec = -1
End Function

Sub FillControlsFromSql(Rs As Recordset)
On Error GoTo errorhandler
OK = False

Set reparationInfo_ = mainDataService_.GetReparationByCallNo(Rs!CallNo)
With reparationInfo_
    txtFindCallNo.Text = .CallNo
    LRepNbr.Caption = .repNo
    LClientName.Caption = "" ' -----missing information
    LEntryDate.Caption = .regestdate
    txtRepDate.Text = Format(.RepDate, "dd/mm/yyyy")
    txtRepTimeBegin.Text = Format(.RepTimeBegin, "HH:MM")
    txtRepTimeEnd.Text = Format(.RepTimeEnd, "HH:MM")
    DcboCliRecever.BoundText = .CliRecever
'    DcboReparationStatus.BoundText =
    txtVoltBefor.Text = .VoltBefor
    txtVoltAfter.Text = .VoltAfter
    txtNotes.Text = .Notes
    txtDescription.Text = .Description
    txtRepPrice.Text = .RepPrice
    DcboPayMethod.BoundText = .Cash
    TxtModelName.Text = .AdhamProductViewModel.Symbol
    TxtModelName.Tag = .AdhamProductViewModel.ModNo
    txtProdSerNo.Text = .AdhamProductViewModel.BarcodeSerialNo
    txtProductionDate.Text = Format(.AdhamProductViewModel.ProdDate, "dd/mm/yyyy")
    txtGasNo.Text = .AdhamProductViewModel.GasNo
    txtPurchaseDate.Text = Format(.AdhamProductViewModel.ProdPurchaseDate, "dd/mm/yyyy")
    LModStockNo.Caption = .AdhamProductViewModel.ItemNo
    LCount.Caption = .ReparationPiecesViewModel.Count
    LTotalItems.Caption = mainDataService_.GetTotalFeesReparation(reparationInfo_)
    LTotal.Caption = .RepPrice
    FillGrid FlexCurrentDamge, reparationInfo_, 1
    FillGrid FlexReparatonItems, reparationInfo_, 2
    FillFormating 2, FlexCurrentDamge
    FillFormating 3, FlexReparatonItems
End With
OK = True
Exit Sub
errorhandler:
MsgBox Err.Description

End Sub


Sub InitNavigator()
    Set RsNavigator = mainDataService_.GetAllReparations
End Sub

Sub FillGrid(FlexGrid As VSFlexGrid, reprationInfo As ReparationViewModel, vChoise As Integer)
FlexGrid.Rows = 1
With FlexGrid
    Select Case vChoise
        Case 1:  'ReparationWorks
            For i = 1 To reprationInfo.ReparationWorksViewModel.Count
                .AddItem ""
                .TextMatrix(.Rows - 1, ColReparationNo) = reprationInfo.ReparationWorksViewModel.Item(i).RepTypeNo
                .TextMatrix(.Rows - 1, ColReparationDescription) = reprationInfo.ReparationWorksViewModel.Item(i).RepTypeDescription
            Next i
        Case 2 'ReparationItems
            For i = 1 To reprationInfo.ReparationPiecesViewModel.Count
                .AddItem ""
                .TextMatrix(.Rows - 1, ColStkNo) = reprationInfo.ReparationPiecesViewModel.Item(i).stkno
                .TextMatrix(.Rows - 1, ColStkName) = reprationInfo.ReparationPiecesViewModel.Item(i).StkName
                .TextMatrix(.Rows - 1, ColQty) = reprationInfo.ReparationPiecesViewModel.Item(i).Qty
                .TextMatrix(.Rows - 1, ColDealerPrice) = reprationInfo.ReparationPiecesViewModel.Item(i).Price
            Next i
        End Select
End With
End Sub


Sub Init()
    top = 0
    left = 0
    OK = True
    FillCombos
    InitNavigator
    EnableControls False
    MoveNavigator 4
    Grid.Rows = 1
    Grid.SelectionMode = flexSelectionListBox

End Sub

Sub ClearControls()
OK = False
    
txtFindCallNo.Text = ""
LRepNbr.Caption = ""
LClientName.Caption = ""
LEntryDate.Caption = ""
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
TxtModelName.Text = ""
txtProdSerNo.Text = ""
txtProductionDate.Text = "__/__/__"
LModStockNo.Caption = ""
txtGasNo.Text = ""
txtPurchaseDate.Text = "__/__/__"
TxtCurrentDamage.Text = ""
TxtStkName.Text = ""
LStkName.Caption = ""
LBalance.Caption = ""
TxtStkQty.Text = ""
LDealerPrice.Caption = ""
LCount.Caption = ""
LTotalItems.Caption = ""
LTotal.Caption = ""
FlexCurrentDamge.Rows = 1
FlexReparatonItems.Rows = 1

OK = True
End Sub

Sub EnableCmds(FAdd As Boolean, FEdit As Boolean, FDelete As Boolean, FSave As Boolean, FUndo As Boolean, FFirst As Boolean, FNext As Boolean, FPrevious As Boolean, FLast As Boolean)
    CmdAdd.Enabled = FAdd
    CmdEdit.Enabled = FEdit
    CmdDelete.Enabled = FDelete
    cmdSave.Enabled = FSave
    CmdCancel.Enabled = FUndo
    CmdFirst.Enabled = FFirst
    CmdLast.Enabled = FLast
    CmdNext.Enabled = FNext
    CmdPrevious.Enabled = FPrevious
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

Private Sub TxtCurrentDamage_Change()
Search TxtCurrentDamage, 1, "CoReparationType", "reptypedescription", "reptypeno", True
End Sub

Private Sub TxtCurrentDamage_GotFocus()
Pos = 4
End Sub

Private Sub TxtCurrentDamage_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then MoveCursor KeyCode, Grid
End Sub
Function SaveRec() As Boolean
On Error GoTo errorhandler


SaveRec = True
Exit Function
errorhandler:
MsgBox Err.Description
SaveRec = False
End Function

Private Sub TxtCurrentDamage_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        If Grid.Visible Then
            OK = False
            TxtCurrentDamage.Tag = IIf(Grid.SelectedRows = 0, Grid.TextMatrix(1, ColNo), Grid.TextMatrix(Grid.Row, ColNo))
            TxtCurrentDamage.Text = IIf(Grid.SelectedRows = 0, Grid.TextMatrix(1, ColName), Grid.TextMatrix(Grid.Row, ColName))
            OK = True
            Grid.Visible = False
            If SaveRec() Then
                TxtCurrentDamage.SetFocus
                SendKeys "{home}+{end}"
            End If
        Else
            TxtStkName.SetFocus
            SendKeys "{home}+{end}"
        End If
End If
End Sub

Private Sub txtDescription_Change()
Search txtDescription, 1, "adhamreparation", "RepName", "repnum", True
End Sub

Private Sub txtDescription_GotFocus()
Pos = 2
End Sub

Private Sub txtDescription_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then MoveCursor KeyCode, Grid
End Sub

Private Sub txtDescription_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        If Grid.Visible Then
            OK = False
            txtDescription.Tag = IIf(Grid.SelectedRows = 0, Grid.TextMatrix(1, ColNo), Grid.TextMatrix(Grid.Row, ColNo))
            txtDescription.Text = IIf(Grid.SelectedRows = 0, Grid.TextMatrix(1, ColName), Grid.TextMatrix(Grid.Row, ColName))
            OK = True
            Grid.Visible = False
            txtDescription.SetFocus
            SendKeys "{End}"
        Else
            txtRepPrice.SetFocus
            SendKeys "{Home}+{End}"
        End If
End If
End Sub

Private Sub txtFindCallNo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtRepDate.SetFocus
SendKeys "{home}+{end}"
End If
End Sub

Private Sub TxtModelName_Change()
Search TxtModelName, 1, "adhammodels", "Symbol", "ModNo", True
End Sub

Private Sub TxtModelName_GotFocus()
Pos = 3
End Sub

Private Sub TxtModelName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then MoveCursor KeyCode, Grid
End Sub

Private Sub TxtModelName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        If Grid.Visible Then
            OK = False
            TxtModelName.Tag = IIf(Grid.SelectedRows = 0, Grid.TextMatrix(1, ColNo), Grid.TextMatrix(Grid.Row, ColNo))
            TxtModelName.Text = IIf(Grid.SelectedRows = 0, Grid.TextMatrix(1, ColName), Grid.TextMatrix(Grid.Row, ColName))
            OK = True
            Grid.Visible = False
        Else
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

Private Sub txtNotes_Change()
Search txtNotes, 1, "adhamreparation", "RepName", "repnum", True
End Sub
Sub Search(sender As Control, Pos As Integer, tableName As String, listField As String, dataMember As String, Optional isChangeCursor = True)
On Error GoTo errorhandler
Dim RsSearch As New ADODB.Recordset


If OK Then
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
errorhandler:
MsgBox Err.Description
End Sub

Private Sub txtNotes_GotFocus()
Pos = 1
End Sub

Private Sub txtNotes_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then MoveCursor KeyCode, Grid
End Sub

Private Sub txtNotes_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        If Grid.Visible Then
            OK = False
            txtNotes.Tag = IIf(Grid.SelectedRows = 0, Grid.TextMatrix(1, ColNo), Grid.TextMatrix(Grid.Row, ColNo))
            txtNotes.Text = IIf(Grid.SelectedRows = 0, Grid.TextMatrix(1, ColName), Grid.TextMatrix(Grid.Row, ColName))
            OK = True
            Grid.Visible = False
            txtNotes.SetFocus
            SendKeys "{End}"
        Else
            txtDescription.SetFocus
            SendKeys "{Home}+{End}"
        End If
End If
End Sub

Private Sub txtProdSerNo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtProductionDate.SetFocus
    SendKeys "{home}+{end}"
End If
End Sub

Private Sub txtProductionDate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtGasNo.SetFocus
SendKeys "{home}+{end}"
End If

End Sub

Private Sub txtPurchaseDate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    TxtCurrentDamage.SetFocus
    SendKeys "{home}+{end}"
End If
End Sub

Private Sub txtRepDate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtRepTimeBegin.SetFocus
SendKeys "{home}+{end}"
End If
End Sub

Private Sub txtRepPrice_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    DcboPayMethod.SetFocus
End If
End Sub

Private Sub txtRepTimeBegin_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtRepTimeEnd.SetFocus
SendKeys "{home}+{end}"
End If
End Sub

Private Sub txtRepTimeEnd_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    DcboCliRecever.SetFocus
    SendKeys "{home}+{end}"
End If
End Sub

Private Sub TxtStkQty_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
End If
End Sub

Private Sub txtVoltAfter_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtNotes.SetFocus
    SendKeys "{home}+{end}"
End If
End Sub

Private Sub txtVoltBefor_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtVoltAfter.SetFocus
    SendKeys "{home}+{end}"
End If
End Sub

