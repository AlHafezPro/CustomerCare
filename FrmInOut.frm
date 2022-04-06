VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmInOut 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«·≈œŒ«· Ê «·≈Œ—«Ã"
   ClientHeight    =   6600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11745
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   6600
   ScaleWidth      =   11745
   Begin Crystal.CrystalReport cr1 
      Left            =   5130
      Top             =   2940
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin Threed.SSFrame Frame1 
      Height          =   465
      Left            =   0
      TabIndex        =   12
      Top             =   690
      Width           =   11745
      _ExtentX        =   20717
      _ExtentY        =   820
      _Version        =   131074
      Begin VB.TextBox TxtDocNum 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   60
         Width           =   1875
      End
      Begin MSMask.MaskEdBox TxtDate 
         Height          =   375
         Left            =   7530
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   60
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "99/99/9999"
         PromptChar      =   "_"
      End
      Begin MSDataListLib.DataCombo ComboStr 
         Height          =   360
         Left            =   2370
         TabIndex        =   1
         Top             =   120
         Width           =   4365
         _ExtentX        =   7699
         _ExtentY        =   635
         _Version        =   393216
         MatchEntry      =   -1  'True
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
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "«·»Ì«‰"
         Height          =   195
         Left            =   1995
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   60
         Width           =   375
      End
      Begin VB.Label LByanId 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   9750
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   90
         Width           =   1005
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   " «—ÌŒ «·Õ—ﬂ…"
         Height          =   195
         Left            =   8820
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   120
         Width           =   870
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "«·„” Êœ⁄"
         Height          =   195
         Left            =   6870
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   90
         Width           =   630
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "—ﬁ„ «·Õ—ﬂ…"
         Height          =   195
         Left            =   10950
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   150
         Width           =   735
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   4905
      Left            =   0
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1650
      Visible         =   0   'False
      Width           =   7815
      _cx             =   13785
      _cy             =   8652
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
   Begin VB.TextBox TxtQty 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1290
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   1170
      Width           =   1575
   End
   Begin VB.CheckBox Chk 
      Alignment       =   1  'Right Justify
      Caption         =   "„œŒ·« "
      Height          =   435
      Left            =   30
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1170
      Width           =   1275
   End
   Begin VB.TextBox TxtStkNo 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   9000
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   1200
      Width           =   1725
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
            Picture         =   "FrmInOut.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInOut.frx":26D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInOut.frx":5533
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInOut.frx":7CD8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInOut.frx":A5D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInOut.frx":CAF9
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInOut.frx":F2B1
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInOut.frx":11CC5
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInOut.frx":14617
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInOut.frx":1738B
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInOut.frx":19BDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInOut.frx":1CA81
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInOut.frx":1F7DB
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInOut.frx":2217B
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInOut.frx":250DD
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInOut.frx":27B05
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInOut.frx":2A4BF
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInOut.frx":2CDF0
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInOut.frx":2F6F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInOut.frx":3215E
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInOut.frx":3510F
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInOut.frx":37A38
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInOut.frx":3A36D
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInOut.frx":3CAAD
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInOut.frx":3F41C
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInOut.frx":41CA4
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInOut.frx":43EB3
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInOut.frx":46810
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInOut.frx":4903A
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInOut.frx":4BAEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInOut.frx":4E52A
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInOut.frx":51040
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInOut.frx":53FDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInOut.frx":56DA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInOut.frx":59A20
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInOut.frx":5C8B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInOut.frx":5F4F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInOut.frx":6209F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   690
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   11745
      _ExtentX        =   20717
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
            Object.Visible         =   0   'False
            ImageIndex      =   15
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   2
         EndProperty
      EndProperty
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
         Left            =   9450
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   60
         Width           =   1125
      End
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
         Left            =   10590
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Text            =   "⁄œœ «·„Ê«œ"
         Top             =   30
         Width           =   1095
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid FlexGrid 
      Height          =   4965
      Left            =   0
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1650
      Width           =   11715
      _cx             =   20664
      _cy             =   8758
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
      Rows            =   1
      Cols            =   4
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
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "—ﬁ„ „Œ“‰Ì"
      Height          =   345
      Left            =   10830
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   1230
      Width           =   825
   End
   Begin VB.Label Lbalance 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Left            =   3330
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   1200
      Width           =   1155
   End
   Begin VB.Label LStkName 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Left            =   4500
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   1200
      Width           =   4455
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "«·ﬂ„Ì…"
      Height          =   375
      Index           =   0
      Left            =   2910
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   1200
      Width           =   405
   End
End
Attribute VB_Name = "FrmInOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const ColId = 1
Const ColByanId = 2
Const ColStkId = 3
Const ColStkNo = 4
Const ColStkName = 5
Const ColStrid = 6
Const ColStrNo = 7
Const ColStrName = 8
Const ColStkQty = 9
Const ColChk = 10
Const ColDocNum = 11
Const colDescription = 12
Const ColMovDate = 13


Const ColStkId_1 = 1
Const ColStkno_1 = 2
Const ColStkname_1 = 3
Const ColCliPrice = 4
Const ColDealPrice = 5

Const colModNo = 1
Const ColModSymbol = 2
Const ColModName = 3


Dim Flag As Boolean
Dim Strid As Integer, ByanId As Integer, stkId As Integer, MovDate As String, Doctype As Integer, Qty As Double, QtyType As Integer
Dim Stmov As MvStockType
Dim PrevQty As Double, PrevChkType As Integer, TypeRec As Boolean
Sub FillControlsFromSql(rs As Recordset)
Ok = False

With rs
    TxtDate.Text = !Date
    TxtEmpNo.Tag = !empNo
    TxtEmpNo.Text = !FullName
'    txtBarcodeNBR.Text = !BarcodeNBr & ""
 '   TxtBillNBR.Text = !BillNBR & ""
    TxtReceiverName.Tag = !ReceiverId
    TxtReceiverName.Text = !ReceiverName
    txtJob.Tag = IIf(IsNull(!JobId), 0, !JobId)
    txtJob.Text = IIf(IsNull(!JobName), "", !JobName) & ""
    TxtShortCut.Text = IIf(IsNull(!Abbreviation), "", !Abbreviation) & ""
    TxtModelName.Tag = IIf(IsNull(!ModNo), 0, !ModNo)
    TxtModelName.Text = !Symbol
    TxtAmount.Text = !Amount
    LTopDiscount.Caption = IIf(IsNull(!TopDiscount), 0, !TopDiscount)
    'TxtTotalAmount.Text = !Amount
    txtNotes.Text = !Notes
    'FillFeesMop !FeesId
'    Lrest.Caption = GetRest(Val(TxtTotalAmount.Text))
End With
Ok = True
End Sub

Function GetCount(vByanId As Double) As Double
Dim RsFind As New ADODB.Recordset
sqlText = "Select isnull(Count(*),0)CountRec From Stmov  Where ByanId=" & vByanId
Set RsFind = de.con.Execute(sqlText)
GetCount = RsFind!CountRec
End Function
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
Function GetNum(Doctype As Integer) As Double
    Dim RsGetNum As New ADODB.Recordset
    sqlText = "Select isnull(max(DocNum),0)MaxDocNum From Stmov Where DocType=" & Doctype
    Set RsGetNum = de.con.Execute(sqlText)
    GetNum = RsGetNum!MaxDocNum + 1
End Function

Function ChkStkid(stkId As Double, Strid As Integer) As Boolean
Dim RsBChk As New ADODB.Recordset
    sqlText = "Select count(*)CountRec From Stkinf Where Stkid=" & stkId & " And StrId=" & Strid
    Set RsChk = de.con.Execute(sqlText)
    If RsChk!CountRec > 0 Then
        ChkStkid = True
    Else
        ChkStkid = False
    End If
End Function
'Sub PrintData()
'On Error GoTo ErrorHandler
'With cr1
'    .Connect = ConnectName
'    .ReportFileName = App.Path + "\reports\ByanRep.rpt"
''    .SQLQuery = "Select s1.Id , s1.ByanId , s1.Stkid , s1.StkidType , s1.StkNo , s1.StkName , s1.StrId , s1.StrNo , s1.StrName , s1.MovDate , s1.DocType , s1.Qty , s1.QtyType  , s2.FnlQnt From StmovQry s1 inner join stkinf s2 on s1.stkid=s2.stkid and s1.strid=s2.strid  Where ByanId = " & LByanId & "  Order by Case When QtyType=0 then Qty else 0 end ,Case When QtyType=1 then Qty else 0 end"
'        .SQLQuery = "Select StmovQry.Id , StmovQry.ByanId , StmovQry.Stkid , StmovQry.StkidType , StmovQry.StkNo , StmovQry.StkName , StmovQry.StrId , StmovQry.StrNo , StmovQry.StrName , StmovQry.MovDate , StmovQry.DocType , StmovQry.Qty , StmovQry.QtyType , StmovQry.FnlQnt   From StmovQry Where StmovQry.ByanId = " & LByanId & "  Order by StmovQry.Stkid , Case When StmovQry.QtyType=0 then StmovQry.Qty else 0 end ,Case When StmovQry.QtyType=1 then StmovQry.Qty else 0 end"
'    .DiscardSavedData = True
'    .WindowState = crptMaximized
'    .Action = 1
'End With
'Exit Sub
'ErrorHandler:
'MsgBox Err.Description
'End Sub
Sub FillGrid(ByanId As Double)
Dim rs As New ADODB.Recordset
sqlText = "Select  Id , ByanId , StkId , StkNo , StkName , StrId , StrNo , StrName , Qty , QtyType , DocNum , Case When QtyType =0 then '„œŒ·« ' Else '„Œ—Ã« ' End Description , MovDate From stmovqry Where ByanId=" & ByanId & " Order By Id"
Set rs = de.con.Execute(sqlText)
Set FlexGrid.DataSource = rs
FlexGrid.ColDataType(ColChk) = flexDTBoolean
FillFormatStringStk
LByanId.Caption = GByanId
TxtDocNum.Text = IIf(IsNull(rs!DocNum), 0, rs!DocNum)
ComboStr.BoundText = rs!Strid
TxtDate.Text = Format(rs!MovDate, "dd/mm/yyyy")
End Sub
Function GetBalance(stkId As Double, Strid As Double) As Double
Dim RsBalance As New ADODB.Recordset
    sqlText = "Select FnlQnt From Stkinf Where Stkid=" & stkId & " And StrId=" & Strid
    Set RsBalance = de.con.Execute(sqlText)
    If RsBalance.RecordCount > 0 Then GetBalance = IIf(IsNull(RsBalance!fnlqnt), 0, RsBalance!fnlqnt)
End Function

Sub DeleteRec(Vrow As Integer)
On Error GoTo ErrorHandler
Dim CurrBalance  As Double
de.con.BeginTrans
With FlexGrid
    sqlText = "Delete From Stmov Where Id=" & .TextMatrix(Vrow, ColId) & " And ByanId = " & .TextMatrix(Vrow, ColByanId)
    de.con.Execute (sqlText)
    CurrBalance = GetBalance(.TextMatrix(Vrow, ColStkId), .TextMatrix(Vrow, ColStrid))
    If CurrBalance >= 0 Then
        de.con.CommitTrans
        .RemoveItem Vrow
    Else
        de.con.RollbackTrans
        MsgBox "«·—’Ìœ ·«Ì”„Õ", vbExclamation, " ‰»ÌÂ"
    End If
End With
Exit Sub
ErrorHandler:
de.con.RollbackTrans
MsgBox (Err.Description)
End Sub
Sub ClearData()
    TxtStkNo.Text = ""
    TxtStkNo.Tag = ""
    LStkName.Caption = ""
End Sub
Function NewRec() As Double
Dim RsMax As New ADODB.Recordset
sqlText = "Select isnull(Max(ByanId),0) as MaxByanId From Stmov"
Set RsMax = de.con.Execute(sqlText)
If RsMax!maxByanId = 0 Then
    NewRec = 1
Else
    NewRec = RsMax!maxByanId + 1
End If
Frame1.Enabled = True
ComboStr.SetFocus
'TxtDate.SelLength = Len(TxtDate.Text)
'TxtDate.SetFocus

End Function
Function ChkQty(stkId As Double, Strid As Integer, Qty As Double, QtyType As Integer) As Boolean
Dim RsChk As New ADODB.Recordset
If QtyType = 0 Then
    ChkQty = True
Else
    sqlText = "Select Sum(Case When QtyType=0 then Qty else -Qty end) Qty From Stmov Where StkId= " & stkId & " And StrId =" & Strid
    Set RsChk = de.con.Execute(sqlText)
    If RsChk.RecordCount > 0 Then
        If RsChk!Qty >= Qty Then
            ChkQty = True
        Else
            ChkQty = False
        End If
    Else
        ChkQty = False
    End If
End If
End Function
Function StrNo(Id As Integer) As String
Dim RsStrNo As New ADODB.Recordset
sqlText = "Select Id , StrNo , StrName From NameStr where Id=" & Id
Set RsStrNo = de.con.Execute(sqlText)
If RsStrNo.RecordCount > 0 Then
    StrNo = RsStrNo!StrName
End If
End Function
Function StrName(Id As Integer) As String
Dim RsStrName As New ADODB.Recordset
sqlText = "Select Id , StrNo , StrName From NameStr where Id=" & Id
Set RsStrName = de.con.Execute(sqlText)
If RsStrName.RecordCount > 0 Then
    StrName = RsStrName!StrName
End If
End Function
Function maxId() As Double
Dim RsMaxId As New ADODB.Recordset
    sqlText = "Select Max(Id)MaxId From Stmov"
    Set RsMaxId = de.con.Execute(sqlText)
    maxId = RsMaxId!maxId
End Function
Sub AddToGrid()
Dim Vrow As Integer
With FlexGrid
    .AddItem ""
     Vrow = .Rows - 1
    .TextMatrix(Vrow, ColId) = maxId
    .TextMatrix(Vrow, ColByanId) = Stmov.ByanId
    .TextMatrix(Vrow, ColStkId) = Stmov.stkId
    .TextMatrix(Vrow, ColStkNo) = TxtStkNo.Text
    .TextMatrix(Vrow, ColStkName) = LStkName.Caption
    .TextMatrix(Vrow, ColStrid) = Stmov.Strid
    .TextMatrix(Vrow, ColStrNo) = StrNo(Stmov.Strid)
    .TextMatrix(Vrow, ColStrName) = StrName(Stmov.Strid)
    .TextMatrix(Vrow, ColStkQty) = Stmov.Qty
    .TextMatrix(Vrow, ColChk) = Stmov.QtyType
    .TextMatrix(Vrow, colDescription) = IIf(.TextMatrix(Vrow, ColChk) = 0, "„œŒ·« ", "„Œ—Ã« ")
    .TextMatrix(Vrow, ColMovDate) = Stmov.MovDate
    PrevChkType = Stmov.QtyType
    .Col = ColId
    .Sort = flexSortGenericDescending
    .Row = 1
    FlexGrid_RowColChange
End With
End Sub
Function ChkVariables() As Boolean
Dim Ok As Boolean
    Ok = True
    If TxtDocNum.Text = "" Then
        Ok = False
    End If
    If Not IsDate(TxtDate.Text) Then
        Ok = False
    End If
    If ComboStr.BoundText = "" Then
        Ok = False
    End If
    If TxtStkNo.Tag = "" Then
        Ok = False
    End If
    If TxtQty.Text = "" Then
        Ok = False
    End If
    ChkVariables = Ok
End Function
Sub FillRec()
With Stmov
    .ByanId = IIf(LByanId.Caption = "", NewRec, LByanId.Caption)
    LByanId.Caption = .ByanId
    .MovDate = Format(TxtDate.Text, "mm/dd/yyyy")
    .Qty = TxtQty.Text
    .QtyType = Chk.Value
    .stkId = TxtStkNo.Tag
    .Strid = ComboStr.BoundText
    .DocNum = TxtDocNum.Text
    .Doctype = 1
End With
End Sub

'Function MaxRec() As Integer
'Sqltext = "Select Max(Id) MAXID From Stmov"
'End Function

Function InsertRec() As Boolean
On Error GoTo ErrorHandler
If ChkVariables Then
FillRec
    If ChkQty(Stmov.stkId, Stmov.Strid, Stmov.Qty, Stmov.QtyType) Then
        With Stmov
            sqlText = "Insert into Stmov(ByanId , StkId  , StrId , Movdate , DocType , DocNum ,  Qty , QtyType,EmpNo)Values("
            sqlText = sqlText & .ByanId & "," & .stkId & "," & .Strid & ",'" & .MovDate & "'," & .Doctype & "," & .DocNum & "," & .Qty & "," & .QtyType & "," & empNo & ")"
            de.con.Execute (sqlText)
        End With
     Else
        InsertRec = False
        Exit Function
    End If
    InsertRec = True
Else
    InsertRec = False
End If
Exit Function
ErrorHandler:
InsertRec = False
MsgBox Err.Description
End Function
Sub init()

Dim RsSTr As New ADODB.Recordset
sqlText = "Select Id , StrNo , StrName From NameStr  Order by StrNo"
Set RsSTr = de.con.Execute(sqlText)
If RsSTr.RecordCount > 0 Then
    Set ComboStr.RowSource = RsSTr
    ComboStr.listField = "StrName"
    ComboStr.BoundColumn = "Id"
    ComboStr.BoundText = RsSTr!Id
End If
FillFormatStringStk
With FlexGrid
    .Editable = flexEDKbdMouse
    .ColDataType(ColChk) = flexDTBoolean
End With
Me.top = 0
Me.left = 0
    Chk.BackColor = vbBlue
    Chk.ForeColor = vbWhite
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
            Lbalance.Caption = 0
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

Sub MovList(List As VSFlexGrid)
    List.top = ActiveControl.top + ActiveControl.Height
    List.left = ActiveControl.left
    List.Width = ActiveControl.Width
End Sub

Sub FillList(sqlText As String, Field1 As String, Field2 As String, List As VSFlexGrid, ByVal Switch As Boolean)
On Error GoTo ErrorHandler
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
'        List.Text = ""
        List.Rows = 1
        TxtStkNo.Tag = 0
        Lbalance.Caption = ""
        LStkName.Caption = ""
        MsgBox "«·„«œ… €Ì— „ÊÃÊœ… ÷„‰ ﬁ«∆„… «·√—ﬁ«„ «·„Œ“‰Ì…", vbExclamation, " ‰»ÌÂ"
        List.Visible = False
        TxtStkNo.SelStart = 0
        TxtStkNo.SelLength = Len(TxtStkNo.Text)
        TxtStkNo.SetFocus
    End If
Exit Sub
ErrorHandler:
MsgBox Err.Description
End Sub
Sub SetColWidthsVS(ByVal ColNo As Integer, FlexGrid As VSFlexGrid)
    Dim i, J, s, w
    With FlexGrid
            s = 0
            For i = 0 To .Rows - 1
                w = TextWidth(.TextMatrix(i, ColNo))
                If w > s Then s = w
            Next i
            .ColWidth(ColNo) = s + 300
    End With
End Sub
Sub FillFormatVSFlex(FlexGrid As VSFlexGrid, Switch As Boolean)
If Switch Then
    fs = "|ModNo"
    fs = fs + "|<" + "—„“ «·„ÊœÌ·"
    fs = fs + "|<" + "≈”„ «·„ÊœÌ·"
    With FlexGrid
        .Visible = False
        .FormatString = fs
            .ColWidth(colModNo) = 0
            SetColWidthsVS ColModSymbol, FlexGrid
            SetColWidthsVS ColModName, FlexGrid
            .Visible = True
    End With
Else
    fs = "|ID"
    fs = fs + "|>" + "«·—ﬁ„ «·„Œ“‰Ì"
    fs = fs + "|>" + "«·≈”„"
    fs = fs + "|>" + "”⁄— «·“»Ê‰"
    fs = fs + "|>" + "”⁄— «· «Ã—"
    With FlexGrid
        .Visible = False
        .FormatString = fs
            .ColWidth(ColStkId_1) = 0
            SetColWidthsVS ColStkno_1, FlexGrid
            SetColWidthsVS ColStkname_1, FlexGrid
            SetColWidthsVS ColCliPrice, FlexGrid
            SetColWidthsVS ColDealPrice, FlexGrid
            .Visible = True
    End With
End If
End Sub

Sub FillFormatStringStk()
    fs = "|>" + "Id"
    fs = fs + "|>" + "—ﬁ„ «·»Ì«‰"
    fs = fs + "|>" + "StkId"
    fs = fs + "|>" + "«·—ﬁ„ «·„Œ“‰Ì"
    fs = fs + "|>" + "«·≈”„"
    fs = fs + "|>" + "ÚStrId"
    fs = fs + "|>" + "—ﬁ„ «·„” Êœ⁄"
    fs = fs + "|>" + "«·≈”„"
    fs = fs + "|>" + "«·ﬂ„Ì…"
    fs = fs + "|>" + "‰Ê⁄"
    fs = fs + "|>" + "DocNum"
    fs = fs + "|>" + "«·Õ—ﬂ…"
    fs = fs + "|>" + "MoVdate"
    With FlexGrid
        .FormatString = fs
        .ColWidth(ColId) = 0
        .ColWidth(ColStkId) = 0
        .ColWidth(ColStrid) = 0
        .ColWidth(ColStrNo) = 0
        .ColWidth(ColStrName) = 0
        .ColWidth(ColMovDate) = 0
        .ColWidth(ColDocNum) = 0
        SetColWidths ColByanId, FlexGrid
        SetColWidths ColStkNo, FlexGrid
        SetColWidths ColStkName, FlexGrid
        SetColWidths ColStkQty, FlexGrid
        .ColWidth(colDescription) = 1500
        .ColWidth(ColChk) = 400
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

Private Sub Chk_Click()
If Chk.Value Then
    Chk.Caption = "„Œ—Ã« "
    Chk.BackColor = vbRed
    Chk.ForeColor = vbBlack
Else
    Chk.Caption = "„œŒ·« "
    Chk.BackColor = vbBlue
    Chk.ForeColor = vbWhite
End If
End Sub
Sub ReturnOldData(ByVal Vrow As Integer, ByVal Vcol As Integer)
With FlexGrid
    .TextMatrix(Vrow, ColStkQty) = PrevQty
    If Vcol = ColChk Then
        .TextMatrix(Vrow, ColChk) = IIf(.TextMatrix(Vrow, ColChk) = 0, 1, 0)
        .TextMatrix(Vrow, colDescription) = IIf(.TextMatrix(Vrow, ColChk) = 0, "„œŒ·« ", "„Œ—Ã« ")
    End If
End With
End Sub


'Private Sub ChkMod_Stk_Click()
'If ChkMod_Stk.Value Then
'    ChkMod_Stk.Caption = "„ÊœÌ·"
'Else
'    ChkMod_Stk.Caption = "—ﬁ„ „Œ“‰Ì"
'End If
'ClearData
'End Sub

Private Sub ComboStr_Change()
On Error GoTo ErrorHandler
If TypeRec Then
    TxtDocNum.Text = GetNum(1)
End If
    Exit Sub
ErrorHandler:
    MsgBox Err.Description
End Sub

Private Sub FlexGrid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Dim CurrQty As Double, CompareAmount As Double
Dim CurrBalance  As Double
    With FlexGrid
    de.con.BeginTrans
    Vrow = Row
    CurrQty = .TextMatrix(Vrow, ColStkQty)
    sqlText = "Update Stmov Set Qty = " & CurrQty & ",QtyType=" & Abs(.TextMatrix(Vrow, ColChk)) & ",EmpNo=" & empNo & " Where id=" & .TextMatrix(Vrow, ColId)
    de.con.Execute (sqlText)
    CurrBalance = GetBalance(.TextMatrix(Vrow, ColStkId), .TextMatrix(Vrow, ColStrid))
    If CurrBalance >= 0 Then
        de.con.CommitTrans
        .TextMatrix(Vrow, colDescription) = IIf(.TextMatrix(Vrow, ColChk) = 0, "„œŒ·« ", "„Œ—Ã« ")
    Else
       de.con.RollbackTrans
       ReturnOldData Vrow, Col
       MsgBox "«·—’Ìœ ·«Ì”„Õ", vbExclamation, " ‰»ÌÂ"
    End If
    Lbalance.Caption = GetBalance(.TextMatrix(Vrow, ColStkId), .TextMatrix(Vrow, ColStrid))
    End With
End Sub

Private Sub FlexGrid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, cancel As Boolean)
'If TypeRec Then
    If Col <> ColChk And Col <> ColStkQty Then
        cancel = True
    Else
        PrevQty = FlexGrid.TextMatrix(Row, ColStkQty)
    End If
'Else
'    Cancel = True
'End If
End Sub



Private Sub FlexGrid_KeyDown(KeyCode As Integer, Shift As Integer)
'If TypeRec Then
    If KeyCode = vbKeyDelete Then
        If MsgBox("Â· √‰  „ √ﬂœ „‰ ⁄„·Ì… «·Õ–›", vbYesNo + vbDefaultButton2, "Õ–› «·”Ã·«  «·„Õœœ…") = vbYes Then
            DeleteRec FlexGrid.Row
            Lbalance.Caption = GetBalance(TxtStkNo.Tag, ComboStr.BoundText)
            TxtNum.Text = GetCount(LByanId.Caption)

        End If
    End If
'End If
End Sub

Private Sub FlexGrid_RowColChange()
'On Error Resume Next
With FlexGrid
If .Rows = 1 Then Exit Sub
    Flag = False
    TxtStkNo.Text = .TextMatrix(.Row, ColStkNo)
    Flag = True
    TxtStkNo.Tag = .TextMatrix(.Row, ColStkId)
    TxtQty.Text = .TextMatrix(.Row, ColStkQty)
    Chk.Value = IIf(.TextMatrix(.Row, ColChk) = 0, 0, 1)
    LStkName.Caption = .TextMatrix(.Row, ColStkName)
    Lbalance.Caption = GetBalance(.TextMatrix(.Row, ColStkId), .TextMatrix(.Row, ColStrid))
End With
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift And vbAltMask Then
    If KeyCode = vbKeyN Then
        If LByanId.Caption <> "" Then
            If MsgBox("Â·  —Ìœ ≈‰‘«¡ Õ—ﬂ… ÃœÌœ…", vbYesNo + vbQuestion + vbDefaultButton2, "Õ—ﬂ… ÃœÌœ…") = vbYes Then
                LByanId.Caption = ""
                TxtDate.Text = GetDate(Date)
                FlexGrid.Rows = 1
                FlexGrid.Cols = 14
                TypeRec = True
                TxtDocNum.Text = GetNum(1)
                Frame1.Enabled = True
                ComboStr.SetFocus
            End If
        Else
                LByanId.Caption = ""
                TxtDate.Text = GetDate(Date)
                FlexGrid.Rows = 1
                FlexGrid.Cols = 14
                TypeRec = True
                TxtDocNum.Text = GetNum(1)
                Frame1.Enabled = True
                ComboStr.SetFocus
        End If
    End If
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If ActiveControl.name <> "TxtStkNo" And ActiveControl.name <> "TxtQty" And ActiveControl.name <> "FlexGrid" Then
        SendKeys "{Tab}"
        SendKeys "{Home}+{End}"
    End If
End If
End Sub

Private Sub Form_Load()
    init
End Sub
Function GetDate(DateStr As String) As String
If IsDate(DateStr) Then
    GetDate = Right("00" + Trim(Str(Day(DateStr))), 2) + "/" + Right("00" + Trim(Str(Month(DateStr))), 2) + "/" + Right("0000" + Trim(Str(Year(DateStr))), 4)
Else
    GetDate = "01/01/" + Right("0000" + Trim(Str(Year(Date))), 4)
End If
End Function
Function SearchRec(Optional i) As String
On Error GoTo ErrorHandler
Dim sqlText As String

If IsMissing(i) Then
    sqlText = "select 0 Chk , ByanId , s1.StrId , StrNo , DocNum , Convert(Varchar(10),MovDate,103)MovDate , isnull(Correspondence,'') Correspondence, isnull(Correspondence,'') Correspondence, CountryNo , CountryName  , Count(*) , TypeName  from StmovQry s1 Where ByanId <> 0"
Else
    sqlText = "Select StmovQry.Id , StmovQry.ByanId , StmovQry.Stkid  , StmovQry.StkNo , StmovQry.StkName , StmovQry.StrId , StmovQry.StrNo , StmovQry.StrName , StmovQry.MovDate , StmovQry.DocType , StmovQry.Qty , StmovQry.QtyType , StmovQry.FnlQnt ,  StmovQry.Correspondence, CountryNo   From StmovQry Where  StmovQry.ByanId =" & Val(LByanId.Caption)
End If
If IsMissing(i) Then
    sqlText = sqlText & " Group By ByanId ,s1.StrId , StrNo , DocNum , MovDate , isnull(Correspondence,'')  , CountryNo , CountryName , TypeName  Order by MovDate"
Else
    sqlText = sqlText & " Order by Byanid , StrNo , ltrim(rtrim(StkNo))"
End If
SearchRec = sqlText
Exit Function
ErrorHandler:
MsgBox Err.Description
End Function

Sub PrintData()
On Error GoTo ErrorHandler
With cr1
    .Connect = ConnectName("")
    .ReportFileName = App.Path + "\reports\ByanRep.rpt"
'    .SQLQuery = "Select StmovQry.Id , StmovQry.ByanId , StmovQry.Stkid , StmovQry.StkidType , StmovQry.StkNo , StmovQry.StkName , StmovQry.StrId , StmovQry.StrNo , StmovQry.StrName , StmovQry.MovDate , StmovQry.DocType , StmovQry.Qty , StmovQry.QtyType , StmovQry.FnlQnt   From StmovQry Where strid in (" & Strids & ") and  StmovQry.ByanId in  (" & ByansSelect & ")  Order by StmovQry.Byanid , StmovQry.StrNo , ltrim(rtrim(StkNo))"
    .SQLQuery = SearchRec(1)
    .DiscardSavedData = True
    .WindowState = crptMaximized
    .Action = 1
End With
Exit Sub
ErrorHandler:
MsgBox Err.Description
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
    Case 1
        If LByanId.Caption <> "" Then
            If MsgBox("Â·  —Ìœ ≈‰‘«¡ Õ—ﬂ… ÃœÌœ…", vbYesNo + vbQuestion + vbDefaultButton2, "Õ—ﬂ… ÃœÌœ…") = vbYes Then
                LByanId.Caption = ""
                TxtDate.Text = GetDate(Date)
                FlexGrid.Rows = 1
                FlexGrid.Cols = 14
                TypeRec = True
                TxtDocNum.Text = GetNum(1)
                Frame1.Enabled = True
                ComboStr.SetFocus
            End If
        Else
                LByanId.Caption = ""
                TxtDate.Text = GetDate(Date)
                FlexGrid.Rows = 1
                FlexGrid.Cols = 14
                TypeRec = True
                TxtDocNum.Text = GetNum(1)
                Frame1.Enabled = True
                ComboStr.SetFocus
        End If
    Case 5
        PrintData
    Case 7
        SelectedStr = ComboStr.BoundText
        GByanType = 1 ' »Ì«‰ Õ—ﬂ…
        FrmPrintByans.Show 1
        If GByanId <> 0 Then
            FillGrid GByanId
        End If
    Case 9
        If MsgBox("Â·  —Ìœ «·Œ—ÊÃ", vbYesNo + vbQuestion + vbDefaultButton2, "Œ—ÊÃ") = vbYes Then
            Unload Me
        End If
End Select
End Sub

Private Sub TxtQty_KeyPress(KeyAscii As Integer)
On Error GoTo ErrorHandler
If KeyAscii = 13 Then
'    If TypeRec Then 'New
        If Val(TxtQty.Text) <= 0 Then
            MsgBox "√œŒ· «·ﬂ„Ì…", vbExclamation, " ‰»ÌÂ"
            TxtStkNo.SelStart = 0
            TxtStkNo.SelLength = Len(TxtStkNo.Text)
            TxtStkNo.SetFocus
            Exit Sub
        End If
        If InsertRec Then
            AddToGrid
            FillFormatStringStk
            Frame1.Enabled = False
            Lbalance.Caption = GetBalance(TxtStkNo.Tag, ComboStr.BoundText)
            TxtNum.Text = GetCount(LByanId.Caption)
            TxtStkNo.SetFocus
            SendKeys "{home}+{end}"
        Else
            TxtStkNo.SelStart = 0
            TxtStkNo.SelLength = Len(TxtStkNo.Text)
            TxtStkNo.SetFocus
            MsgBox "«·—’Ìœ ·«Ì”„Õ" & Chr(13) & "√Ê" & Chr(13) & "Œÿ√ ›Ì «·„œŒ·« ", vbExclamation, " ‰»ÌÂ"
        End If
        TxtQty.Text = ""
'    Else
'        MsgBox "⁄„·Ì… ≈” ⁄—«÷", vbExclamation, " ‰»ÌÂ"
'    End If
End If
Exit Sub
ErrorHandler:
MsgBox Err.Description
End Sub

Private Sub TxtStkNo_Change()
On Error GoTo ErrorHandler
If Flag Then
    Dim sqlText As String
    If Trim(TxtStkNo.Text) = "" Then
        TxtStkNo.Tag = ""
        Grid.Visible = False
        Exit Sub
    End If

        sqlText = "Select top 15 Id , ltrim(rtrim(StkNo))StkNo , ltrim(rtrim(StkName))StkName  ,CliPrice , DealPrice From CoStock  where StkNo like " & LikeExpression(TxtStkNo.Text) & " or stkname like  " & LikeExpression(TxtStkNo.Text) & " Order By len(ltrim(rtrim(StkNo))) , ltrim(rtrim(StkNo))"
    FillList sqlText, "Id", "StkNo", Grid, 0
End If
Exit Sub
ErrorHandler:
MsgBox Err.Description
End Sub

Private Sub TxtStkNo_KeyDown(KeyCode As Integer, Shift As Integer)
    Flag = False
    If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then MoveCursor KeyCode
    Flag = True
End Sub

Private Sub txtStkNo_KeyPress(KeyAscii As Integer)
On Error GoTo ErrorHandler
    If KeyAscii = 13 Then
        If DataOk(TxtStkNo.Text) Then  'Chk StkItem If Found in CoStock
            FillActiveControl Grid
            Lbalance.Caption = GetBalance(IIf(TxtStkNo.Tag = "", 0, TxtStkNo.Tag), ComboStr.BoundText)
        Else
            Grid.Visible = False
            SendKeys "{home}+{end}"
            MsgBox "«·„«œ… €Ì— „ÊÃÊœ… ÷„‰ ﬁ«∆„… «·√—ﬁ«„ «·„Œ“‰Ì…", vbExclamation, " ‰»ÌÂ"
            Exit Sub
        End If
        If CDbl(TxtStkNo.Tag) <> 0 Then
            If Not ChkStkid(TxtStkNo.Tag, ComboStr.BoundText) Then
                If MsgBox("«·„«œ… €Ì— „⁄—›… ›Ì «·„” Êœ⁄" & Chr(13) & "Â·  Êœ «·≈” „—«—", vbInformation + vbYesNo + vbDefaultButton2, " ‰»ÌÂ") = vbNo Then
                    SendKeys "{home}+{end}"
                    Exit Sub
                End If
            End If
        Else
            MsgBox "«·„«œ… €Ì— „ÊÃÊœ… ÷„‰ ﬁ«∆„… «·√—ﬁ«„ «·„Œ“‰Ì…", vbExclamation, " ‰»ÌÂ"
            Exit Sub
        End If
        TxtQty.SetFocus
        SendKeys "{home}+{end}"
    End If
Exit Sub
ErrorHandler:
MsgBox Err.Description
End Sub
