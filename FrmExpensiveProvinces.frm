VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmExpensiveProvinces 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ßÔÝ ãÕÇÑíÝ æ ÅíÑÇÏÇÊ ÎÏãÉ ÇáãÓÊåáß ÇáãÍÇÝÙÇÊ"
   ClientHeight    =   8670
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14745
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   8670
   ScaleWidth      =   14745
   Begin VB.CheckBox Chk 
      Caption         =   "ÊÍÏíÏ Çáßá"
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
      Height          =   300
      Left            =   8970
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   1230
      Value           =   1  'Checked
      Width           =   2775
   End
   Begin VB.CommandButton CmdPreviousMonth 
      Caption         =   "P"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   60
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Top             =   720
      Width           =   2775
   End
   Begin VB.CommandButton CmdCurrentMonth 
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2970
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   720
      Width           =   2775
   End
   Begin VB.CommandButton CmdFromTo 
      Caption         =   "0-T"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   720
      Width           =   2775
   End
   Begin VB.CommandButton CmdYesterday 
      Caption         =   "Y"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8985
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   720
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "T"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11970
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   720
      Width           =   2775
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   495
      Left            =   90
      TabIndex        =   7
      Top             =   8100
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   873
      _Version        =   131074
      CaptionStyle    =   1
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "ÇáÝÑÞ"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Index           =   8
         Left            =   1695
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   60
         Width           =   375
      End
      Begin VB.Label LDiff 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   405
         Left            =   60
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   30
         Width           =   1545
      End
      Begin VB.Label LoutCome 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   405
         Left            =   2640
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   30
         Width           =   1695
      End
      Begin VB.Label LIncome 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   405
         Left            =   5550
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   30
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "ÇáÅíÑÇÏÇÊ"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Index           =   5
         Left            =   7365
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   60
         Width           =   600
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "ÇáãÕÇÑíÝ"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Index           =   4
         Left            =   4395
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   60
         Width           =   705
      End
   End
   Begin Crystal.CrystalReport cr1 
      Left            =   4920
      Top             =   2730
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VSFlex8Ctl.VSFlexGrid FlexType 
      Height          =   6465
      Left            =   8160
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1560
      Width           =   6555
      _cx             =   11562
      _cy             =   11404
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   700
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
      BackColorAlternate=   12648447
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
   Begin VSFlex8Ctl.VSFlexGrid FlexGridOutCome 
      Height          =   4185
      Left            =   60
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1560
      Width           =   8055
      _cx             =   14208
      _cy             =   7382
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   700
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
      BackColorAlternate=   12648447
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
      FormatString    =   $"FrmExpensiveProvinces.frx":0000
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
      Left            =   2640
      Top             =   1200
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
            Picture         =   "FrmExpensiveProvinces.frx":00E7
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmExpensiveProvinces.frx":27BD
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmExpensiveProvinces.frx":564F
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmExpensiveProvinces.frx":84AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmExpensiveProvinces.frx":AC51
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmExpensiveProvinces.frx":D54D
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmExpensiveProvinces.frx":FA72
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmExpensiveProvinces.frx":1222A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmExpensiveProvinces.frx":14C3E
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmExpensiveProvinces.frx":17590
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmExpensiveProvinces.frx":1A304
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmExpensiveProvinces.frx":1CB53
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmExpensiveProvinces.frx":1F9FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmExpensiveProvinces.frx":22754
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmExpensiveProvinces.frx":250F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmExpensiveProvinces.frx":28056
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmExpensiveProvinces.frx":2AA7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmExpensiveProvinces.frx":2D438
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmExpensiveProvinces.frx":2FD69
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmExpensiveProvinces.frx":3266F
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmExpensiveProvinces.frx":350D7
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmExpensiveProvinces.frx":38088
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmExpensiveProvinces.frx":3A9B1
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmExpensiveProvinces.frx":3D2E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmExpensiveProvinces.frx":3FA26
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmExpensiveProvinces.frx":42395
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmExpensiveProvinces.frx":44C1D
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmExpensiveProvinces.frx":46E2C
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmExpensiveProvinces.frx":49789
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmExpensiveProvinces.frx":4BFB3
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmExpensiveProvinces.frx":4EA67
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmExpensiveProvinces.frx":514A3
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmExpensiveProvinces.frx":53FB9
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmExpensiveProvinces.frx":56F53
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmExpensiveProvinces.frx":59D1F
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmExpensiveProvinces.frx":5C999
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmExpensiveProvinces.frx":5F82B
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmExpensiveProvinces.frx":62469
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmExpensiveProvinces.frx":65018
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
      Width           =   14745
      _ExtentX        =   26009
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
      Begin Threed.SSFrame SSFrame3 
         Height          =   585
         Left            =   3120
         TabIndex        =   25
         Top             =   0
         Width           =   4605
         _ExtentX        =   8123
         _ExtentY        =   1032
         _Version        =   131074
         Begin MSDataListLib.DataCombo ComboHall 
            Height          =   420
            Left            =   120
            TabIndex        =   26
            Top             =   90
            Width           =   3585
            _ExtentX        =   6324
            _ExtentY        =   741
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "ÇáãÍÇÝÙÉ"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   3735
            RightToLeft     =   -1  'True
            TabIndex        =   27
            Top             =   150
            Width           =   780
         End
      End
      Begin Threed.SSFrame SSFrame2 
         Height          =   585
         Left            =   7740
         TabIndex        =   17
         Top             =   0
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   1032
         _Version        =   131074
         Begin MSComCtl2.DTPicker TxtFromDate 
            Height          =   495
            Left            =   3570
            TabIndex        =   18
            Top             =   30
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   873
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   79626241
            CurrentDate     =   40664
         End
         Begin MSComCtl2.DTPicker txttilldate 
            Height          =   495
            Left            =   90
            TabIndex        =   19
            Top             =   30
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   873
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   79626241
            CurrentDate     =   40664
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "ãä ÊÇÑíÎ"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Index           =   0
            Left            =   6105
            RightToLeft     =   -1  'True
            TabIndex        =   21
            Top             =   60
            Width           =   795
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Åáì ÊÇÑíÎ"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   2670
            RightToLeft     =   -1  'True
            TabIndex        =   20
            Top             =   60
            Width           =   855
         End
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid FlexGridInCome 
      Height          =   1965
      Left            =   60
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   6060
      Width           =   8055
      _cx             =   14208
      _cy             =   3466
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   700
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
      BackColorAlternate=   12648447
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
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "ÇáÅíÑÇÏÇÊ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   3
      Left            =   7260
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   5790
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "ÇáãÕÇÑíÝ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   2
      Left            =   7185
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   1230
      Width           =   945
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "ÇäæÇÚ ÇáÅíÑÇÏÇÊ æ ÇáãÕÇÑíÝ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   1
      Left            =   12060
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   1230
      Width           =   2640
   End
End
Attribute VB_Name = "FrmExpensiveProvinces"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const ColNo = 1
Const ColName = 2
Const ColChk = 3



Const Colid = 1
Const ColAmount = 2
Const ColTYpe = 3
Const ColTYpeDescription = 4
Const colCode = 5
Const ColCodeDEscription = 6


Sub FillFormating(ByVal i As Integer, FlexGrid As VSFlexGrid)
If i = 1 Then
    Fs = "|>" + "Id"
    Fs = Fs + "|>" + "ÇáãÈáÛ"
    Fs = Fs + "|>" + "Type"
    Fs = Fs + "|>" + "äæÚ ÇáãÈáÛ"
    Fs = Fs + "|>" + "Code"
    Fs = Fs + "|>" + "ÇáÔÑÍ"

    With FlexGrid
        .FormatString = Fs
        .Cols = 7
        .ColWidth(Colid) = 0
        .ColWidth(ColTYpe) = 0
        .ColWidth(colCode) = 0
        SetColWidths ColAmount, FlexGrid
        .ColDataType(ColAmount) = flexDTDecimal
        SetColWidths ColTYpeDescription, FlexGrid
        SetColWidths ColCodeDEscription, FlexGrid
    End With
ElseIf i = 2 Then
    Fs = "|>" + "Code"
    Fs = Fs + "|>" + "äæÚ ÇáÅíÑÇÏ Ãæ ÇáãÕÑæÝ"
    Fs = Fs + "|>" + "    "
    With FlexGrid
        .FormatString = Fs
        .Cols = 4
        .ColWidth(ColNo) = 0
        SetColWidths ColName, FlexGrid
        SetColWidths ColChk, FlexGrid
        .ColDataType(ColChk) = flexDTBoolean
    End With
End If
End Sub

Sub SetColWidths(ByVal ColNo As Integer, FlexGrid As VSFlexGrid)
    With FlexGrid
        .AutoSize (ColNo)
    End With
End Sub

Sub init()

Top = 0
Left = 0

txtFromDate.Value = Date
TxtTillDate.Value = Format(Date, "dd/mm/yyyy")

Dim Rs As New ADODB.Recordset

Sqltext = "select HallId , HallName  from CoMaintenanceHalls Where IsVisible=1"
Set Rs = de.con.Execute(Sqltext)
Set ComboHall.RowSource = Rs
ComboHall.ListField = "HallName"
ComboHall.BoundColumn = "HallId"
ComboHall.BoundText = 1


Sqltext = "Select Code , CodeDescription ,-1 chk  From HafezDeveloper.dbo.CoExpensiveType"
Set Rs = de.con.Execute(Sqltext)
Set FlexType.DataSource = Rs
FillFormating 2, FlexType
FlexType.Editable = flexEDKbdMouse

Sqltext = "Select Id, Amount, TYpeId, TYpeDescription, code , CodeDEscription  from Hafez2012.dbo.t_ExpensiveProvincesQry where id=-1"
Set Rs = de.con.Execute(Sqltext)
Set FlexGridInCome.DataSource = Rs
FillFormating 1, FlexGridInCome
Set FlexGridOutCome.DataSource = Rs
FillFormating 1, FlexGridOutCome


End Sub
Sub ChkValues(Vindex As Integer)
With FlexType
    For i = 1 To .Rows - 1
        .TextMatrix(i, ColChk) = Vindex
    Next
End With

End Sub
Private Sub Chk_Click()
If Chk.Value Then
    Chk.Caption = "ÊÍÏíÏ Çáßá"
Else
    Chk.Caption = "ÅáÛÇÁ Çáßá"
End If
ChkValues Chk.Value
End Sub

Private Sub CmdCurrentMonth_Click()
On Error GoTo ERRORHANDLER
Screen.MousePointer = vbHourglass
txtFromDate.Value = "01/" + Right("00" + LTrim(RTrim(Str(Month(Now)))), 2) + "/" + LTrim(RTrim(Str(Year(Now))))
TxtTillDate.Value = Now
SearchData
Screen.MousePointer = vbDefault
Exit Sub
ERRORHANDLER:
Screen.MousePointer = vbDefault
MsgBox Err.Description

End Sub

Private Sub CmdPreviousMonth_Click()
On Error GoTo ERRORHANDLER
Screen.MousePointer = vbHourglass
txtFromDate.Value = "01/" + Right("00" + LTrim(RTrim(Str(Month(Now) - 1))), 2) + "/" + LTrim(RTrim(Str(Year(Now))))
TxtTillDate.Value = DateAdd("d", -1, "01/" + Right("00" + LTrim(RTrim(Str(Month(Now)))), 2) + "/" + LTrim(RTrim(Str(Year(Now)))))
SearchData
Screen.MousePointer = vbDefault
Exit Sub
ERRORHANDLER:
Screen.MousePointer = vbDefault
MsgBox Err.Description

End Sub

Private Sub CmdYesterday_Click()
On Error GoTo ERRORHANDLER
Screen.MousePointer = vbHourglass
txtFromDate.Value = DateAdd("d", -1, Now)
TxtTillDate.Value = DateAdd("d", -1, Now)
SearchData
Screen.MousePointer = vbDefault
Exit Sub
ERRORHANDLER:
Screen.MousePointer = vbDefault
MsgBox Err.Description
End Sub

Private Sub Command1_Click()
On Error GoTo ERRORHANDLER
Screen.MousePointer = vbHourglass
txtFromDate.Value = Now
TxtTillDate.Value = Now
SearchData
Screen.MousePointer = vbDefault
Exit Sub
ERRORHANDLER:
Screen.MousePointer = vbDefault
MsgBox Err.Description
End Sub

Private Sub CmdFromTo_Click()
On Error GoTo ERRORHANDLER
Screen.MousePointer = vbHourglass
txtFromDate.Value = Format("01/01/" + LTrim(RTrim(Str(Year(Now)))), "dd/mm/yyyy")
TxtTillDate.Value = Now
SearchData
Screen.MousePointer = vbDefault
Exit Sub
ERRORHANDLER:
Screen.MousePointer = vbDefault
MsgBox Err.Description
End Sub

Private Sub Command2_Click()

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{tab}"
    SendKeys "{home}+{end}"
End If
End Sub

Private Sub Form_Load()
init
End Sub
Sub printRep()
With cr1
    .Connect = ConnectName("")
    .ReportFileName = App.Path + "\Reports\RepExpensive.rpt"
    .Formulas(0) = "fromdate='" & txtFromDate.Value & "'"
    .Formulas(1) = "tilldate='" & TxtTillDate.Value & "'"
    .DiscardSavedData = True
    .WindowState = crptMaximized
    .Action = 1
End With
End Sub
Function ExecuteProcedure() As Boolean
On Error GoTo ERRORHANDLER
Sqltext = "Exec HafezDeveloper.dbo.sp_ExpensiveMaintenanceProvinces '" + ConvertControlDate(txtFromDate.Value) + "','" + ConvertControlDate(TxtTillDate.Value) + "','" + GetCode & "'," & Val(ComboHall.BoundText) & " with recompile"
de.con.Execute (Sqltext)
ExecuteProcedure = True
Exit Function
ERRORHANDLER:
ExecuteProcedure = False
MsgBox Err.Description

End Function

Function GetAmount(Vindex As Integer, HallId As Integer) As Double
Dim Rs As New Recordset
On Error GoTo ERRORHANDLER

Select Case Vindex
    Case 0
        Sqltext = "Select Sum(amount) as amount From Hafez2012.dbo.t_ExpensiveProvincesQry Where TYpeid=1 and HallId=" & HallId
    Case 1
        Sqltext = "Select Sum(amount) as amount From Hafez2012.dbo.t_ExpensiveProvincesQry Where TYpeid=2 And HallId = " & HallId
    Case 2
        Sqltext = "Select sum(case when typeId=1 then amount else -amount end) as amount From Hafez2012.dbo.t_ExpensiveProvincesQry where  HallId=" & HallId
End Select
   Set Rs = de.con.Execute(Sqltext)
   If Rs.RecordCount > 0 Then
    GetAmount = Rs!Amount
   Else
    GetAmount = 0
   End If
Exit Function
ERRORHANDLER:
GetAmount = 0
End Function
Function GetCode() As String
On Error GoTo ERRORHANDLER
Dim Str As String
Str = ""
With FlexType
    For i = 1 To .Rows - 1
        If .TextMatrix(i, ColChk) Then
            Str = Str + "," + .TextMatrix(i, ColNo)
        End If
    Next
End With
If Str <> "" Then
    Str = Mid(Str, 2)
Else
    Str = ""
End If
GetCode = Str
Exit Function
ERRORHANDLER:
GetCode = ""
End Function
Sub FillGrid(ByRef Rs As ADODB.Recordset, FlexGrid As VSFlexGrid)
With FlexGrid
.Rows = 1
If Rs.EOF Then Exit Sub
Rs.MoveFirst

For i = 1 To Rs.RecordCount
    .AddItem ""
    .TextMatrix(i, Colid) = Rs!Id
    .TextMatrix(i, ColAmount) = Format(Rs!Amount, "###,###,###")
    .TextMatrix(i, ColTYpe) = Rs!TypeId
    .TextMatrix(i, ColTYpeDescription) = Rs!TYpeDescription
    .TextMatrix(i, colCode) = Rs!code
    .TextMatrix(i, ColCodeDEscription) = Rs!codedescription
    Rs.MoveNext
Next
End With
End Sub
Sub SearchData()
Dim Rs As New ADODB.Recordset
If ExecuteProcedure Then
            Sqltext = "select id , amount , typeid , typedescription , code , codedescription  from Hafez2012.dbo.t_ExpensiveProvincesQry  Where Amount <> 0 and TypeId=1"
            Set Rs = de.con.Execute(Sqltext)
            FillGrid Rs, FlexGridInCome
'            Set FlexGridInCome.DataSource = rs
            FillFormating 1, FlexGridInCome
            LIncome.Caption = Format(GetAmount(0, Val(ComboHall.BoundText)), "###,###,###")
        
            Sqltext = "select id , amount , typeid , typedescription , code , codedescription  from Hafez2012.dbo.t_ExpensiveProvincesQry  Where Amount <> 0 and TypeId=2"
            Set Rs = de.con.Execute(Sqltext)
            FillGrid Rs, FlexGridOutCome
'            Set FlexGridOutCome.DataSource = rs
            FillFormating 1, FlexGridOutCome
            LoutCome.Caption = Format(GetAmount(1, Val(ComboHall.BoundText)), "###,###,###")
            LDiff.Caption = Format(GetAmount(2, Val(ComboHall.BoundText)), "###,###,###")
    End If
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo ERRORHANDLER
Dim Rs As New ADODB.Recordset
Screen.MousePointer = vbHourglass
Select Case Button.Index
    Case 1
        'printRep
    Case 3
    
        SearchData
    Case 5
        Unload Me
End Select
Screen.MousePointer = vbDefault
Exit Sub
ERRORHANDLER:
Screen.MousePointer = vbDefault
End Sub
