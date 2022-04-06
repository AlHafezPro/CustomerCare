VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FrmWshopItems 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ÅÔÚÇÑ ÊÓáíã ãæÇÏ ÃæáíÉ"
   ClientHeight    =   5475
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11250
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   5475
   ScaleWidth      =   11250
   Begin Threed.SSFrame SSFrame1 
      Height          =   465
      Left            =   30
      TabIndex        =   19
      Top             =   720
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   820
      _Version        =   131074
      Begin VB.TextBox TxtTeamName 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   60
         Width           =   3555
      End
      Begin MSMask.MaskEdBox TxtDate 
         Height          =   345
         Left            =   6300
         TabIndex        =   0
         Top             =   60
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   609
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "99/99/9999"
         PromptChar      =   "_"
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "ÑÞã ÇáÅÔÚÇÑ"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   10350
         TabIndex        =   23
         Top             =   60
         Width           =   780
      End
      Begin VB.Label LDocNum 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   8850
         TabIndex        =   22
         Top             =   60
         Width           =   1395
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "ÇáÊÇÑíÎ"
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   1
         Left            =   7530
         TabIndex        =   21
         Top             =   60
         Width           =   465
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "ÇáæÑÔÉ"
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   2
         Left            =   3720
         TabIndex        =   20
         Top             =   60
         Width           =   465
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   2085
      Left            =   6120
      TabIndex        =   15
      Top             =   3000
      Visible         =   0   'False
      Width           =   2385
      _cx             =   4207
      _cy             =   3678
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
   Begin VB.TextBox TxtQty 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   2730
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   1470
      Width           =   855
   End
   Begin VB.TextBox TxtitemName 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   8070
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   1470
      Width           =   3135
   End
   Begin MSDataListLib.DataCombo ComboOperationType 
      Height          =   360
      Left            =   1170
      TabIndex        =   4
      Top             =   1470
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   635
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
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
   Begin VSFlex8Ctl.VSFlexGrid flexGrid 
      Height          =   3555
      Left            =   30
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   1860
      Width           =   11175
      _cx             =   19711
      _cy             =   6271
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4290
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
            Picture         =   "FrmWshopItems.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmWshopItems.frx":26D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmWshopItems.frx":5533
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmWshopItems.frx":7CD8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmWshopItems.frx":A5D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmWshopItems.frx":CAF9
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmWshopItems.frx":F2B1
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmWshopItems.frx":11CC5
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmWshopItems.frx":14617
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmWshopItems.frx":1738B
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmWshopItems.frx":19BDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmWshopItems.frx":1CA81
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmWshopItems.frx":1F7DB
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmWshopItems.frx":2217B
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmWshopItems.frx":250DD
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmWshopItems.frx":27B05
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmWshopItems.frx":2A4BF
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmWshopItems.frx":2CDF0
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmWshopItems.frx":2F6F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmWshopItems.frx":3215E
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmWshopItems.frx":3510F
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmWshopItems.frx":37A38
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmWshopItems.frx":3A36D
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmWshopItems.frx":3CAAD
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmWshopItems.frx":3F41C
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmWshopItems.frx":41CA4
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmWshopItems.frx":43EB3
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmWshopItems.frx":46810
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmWshopItems.frx":4903A
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmWshopItems.frx":4BAEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmWshopItems.frx":4E52A
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmWshopItems.frx":51040
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmWshopItems.frx":53FDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmWshopItems.frx":56DA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmWshopItems.frx":59A20
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmWshopItems.frx":5C8B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmWshopItems.frx":5F4F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmWshopItems.frx":6209F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   690
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   11250
      _ExtentX        =   19844
      _ExtentY        =   1217
      ButtonWidth     =   1191
      ButtonHeight    =   1164
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   33
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            ImageIndex      =   38
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   37
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   15
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   2
         EndProperty
      EndProperty
      Begin VB.TextBox TxtCaption 
         Alignment       =   2  'Center
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
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
         Height          =   300
         Left            =   10140
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Text            =   "ÚÏÏ ÇáãæÇÏ"
         Top             =   30
         Width           =   1095
      End
      Begin VB.TextBox TxtNum 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   9000
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   30
         Width           =   1125
      End
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "äæÚ ÇáÚãáíÉ"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   6
      Left            =   1935
      TabIndex        =   13
      Top             =   1230
      Width           =   735
   End
   Begin VB.Label LWshopBalance 
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
      Height          =   315
      Left            =   60
      TabIndex        =   12
      Top             =   1470
      Width           =   1065
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "ÑÕíÏ ÇáæÑÔÉ"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   3
      Left            =   270
      TabIndex        =   11
      Top             =   1230
      Width           =   855
   End
   Begin VB.Label LStoreBalance 
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
      Height          =   315
      Left            =   3630
      TabIndex        =   10
      Top             =   1500
      Width           =   1665
   End
   Begin VB.Label LItemName 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   5310
      TabIndex        =   9
      Top             =   1500
      Width           =   2745
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "ÇáßãíÉ"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   11
      Left            =   3150
      TabIndex        =   8
      Top             =   1230
      Width           =   405
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "ÇáÔÑÍ"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   10
      Left            =   7650
      TabIndex        =   7
      Top             =   1230
      Width           =   420
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "ÇáÑÕíÏ ÇáãÓÊæÏÚí"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   9
      Left            =   3585
      TabIndex        =   6
      Top             =   1230
      Width           =   1710
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "ÇáÑÞã ÇáãÎÒäí"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   8
      Left            =   10245
      TabIndex        =   5
      Top             =   1230
      Width           =   930
   End
End
Attribute VB_Name = "FrmWshopItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Pos As Integer
Dim Ok As Boolean, Flag As Boolean


Const ColNo = 1
Const ColName = 2
Dim MvWshopRec As MvWshopType

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

Sub FillFormating(ByVal i As Integer)
If i = 1 Then
    Fs = "|>" + ""
    Fs = Fs + "|>" + ""
   With Grid
        .FormatString = Fs
        .Cols = 3
        SetColWidths ColNo, Grid
        SetColWidths ColName, Grid
    End With
End If
End Sub

Sub SetColWidths(ByVal ColNo As Integer, FlexGrid As VSFlexGrid)
    With FlexGrid
        .AutoSize (ColNo)
    End With
End Sub
Sub ChangeCursor(ByVal X As Integer)
If X = 1 Then
    With TxtTeamName
       Grid.Top = SSFrame1.Top + .Top + .Height
       Grid.Left = SSFrame1.Left + .Left
       Grid.Width = .Width
    End With
ElseIf X = 2 Then
    With TxtModelName
       Grid.Top = SSFrame1.Top + .Top + .Height
       Grid.Left = SSFrame1.Left + .Left
       Grid.Width = .Width
End With
ElseIf X = 3 Then
    With TxtitemName
       Grid.Top = SSFrame2.Top + .Top + .Height
       Grid.Left = SSFrame2.Left + .Left
       Grid.Width = .Width
End With
ElseIf X = 4 Then
    With TxtRecipient
       Grid.Top = SSTab1.Top + SSFrame2.Top + .Top + .Height
       Grid.Left = SSFrame2.Left + .Left
       Grid.Width = .Width
End With
ElseIf X = 5 Then
    With TxtFamNo
       Grid.Top = SSFrameModel.Top + SSTab1.Top + .Top + .Height
       Grid.Left = SSFrameModel.Left + .Left
       Grid.Width = .Width
End With
ElseIf X = 6 Then
    With TxtModelName
       Grid.Top = .Top + .Height
       Grid.Left = .Left
       Grid.Width = .Width
End With
ElseIf X = 7 Then
    With TxtUnExecutedReason
       Grid.Top = SSTab1.Top + .Top + .Height
       Grid.Left = .Left
       Grid.Width = .Width
End With
ElseIf X = 8 Then
    With TxtError
       Grid.Top = .Top + .Height
       Grid.Left = .Left
       Grid.Width = .Width
End With
ElseIf X = 9 Then
    With TxtitemName
       Grid.Top = .Top + .Height
       Grid.Left = .Left
       Grid.Width = .Width
End With
ElseIf X = 10 Then
    With txtCallNo
       Grid.Top = SSTab1.Top + .Top + .Height
       Grid.Left = SSTab1.Left + .Left
       Grid.Width = .Width
End With
End If
End Sub

Private Sub Text1_GotFocus()
Pos = 1
End Sub

Private Sub TxtDate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    TxtTeamName.SetFocus
    SendKeys "{home}+{end}"
End If
End Sub

Private Sub TxtitemName_GotFocus()
Pos = 2
End Sub

Private Sub TxtTeamName_Change()
On Error GoTo ERRORHANDLER
Dim RsSearch As New ADODB.Recordset
If TxtModelName.Text = "" Then
    TxtModelName.Tag = 0
    Grid.Visible = False
    Exit Sub
End If
If Ok Then
    Flag = False
    Sqltext = "select ModNo , Symbol , Name    from adhammodels Where Symbol    Like" & LikeExpression(TxtModelName.Text) & " or Name    like " & LikeExpression(TxtModelName.Text)
    Set RsSearch = de.con.Execute(Sqltext)
    If RsSearch.RecordCount > 0 Then
        Set Grid.DataSource = RsSearch
        Grid.Row = 0
        FillFormating 1
        ChangeCursor 2
        Grid.Visible = True
    Else
        TxtModelName.Tag = 0
        Grid.Visible = False
    End If
    Flag = True
End If
Exit Sub
ERRORHANDLER:
MsgBox Err.Description
End Sub


Private Sub TxtTeamName_GotFocus()
ChangeToArabic
Pos = 1
End Sub

Private Sub TxtTeamName_KeyDown(KeyCode As Integer, Shift As Integer)
    Flag = True
    If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then MoveCursor KeyCode, Grid
    Flag = True
End Sub

Private Sub TxtTeamName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Grid.Visible Then
        Ok = False
        TxtModelName.Tag = Grid.TextMatrix(IIf(Grid.Row = 0, 1, Grid.Row), ColNo)
        TxtModelName.Text = Grid.TextMatrix(IIf(Grid.Row = 0, 1, Grid.Row), Col3)
        LSymbol.Caption = Grid.TextMatrix(IIf(Grid.Row = 0, 1, Grid.Row), ColName)
        Ok = True
    ElseIf Grid.Visible = False And TxtModelName.Text <> "" And Val(TxtModelName.Tag) <> 0 Then
        ComboFees.SetFocus
        Exit Sub
    Else
        Ok = False
        TxtModelName.Tag = 0
        TxtModelName.Text = ""
        LSymbol.Caption = ""
        Ok = True
    End If
    Grid.Visible = False
    ComboFees.SetFocus
End If
End Sub

