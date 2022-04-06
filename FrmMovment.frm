VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmMovment 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "‰ﬁ· „‰ „” Êœ⁄ ≈·Ï „” Êœ⁄"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   6825
   ScaleWidth      =   11910
   Begin VSFlex8Ctl.VSFlexGrid grid 
      Height          =   3165
      Left            =   3540
      TabIndex        =   18
      Top             =   1680
      Visible         =   0   'False
      Width           =   7305
      _cx             =   12885
      _cy             =   5583
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
   Begin Crystal.CrystalReport cr1 
      Left            =   5220
      Top             =   3030
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   690
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   11910
      _ExtentX        =   21008
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
         Left            =   8430
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Text            =   "⁄œœ «·„Ê«œ"
         Top             =   120
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
         Height          =   375
         Left            =   7140
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   90
         Width           =   1215
      End
      Begin VB.TextBox LByanId 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9570
         RightToLeft     =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   90
         Width           =   1215
      End
      Begin VB.TextBox LByanNo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10650
         RightToLeft     =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         Text            =   "»Ì«‰ «·Õ—ﬂ…"
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.TextBox TxtQty 
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
      ForeColor       =   &H00800000&
      Height          =   405
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   1170
      Width           =   1215
   End
   Begin Threed.SSFrame Frame1 
      Height          =   465
      Left            =   0
      TabIndex        =   8
      Top             =   690
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   820
      _Version        =   131074
      Begin VB.TextBox TxtDocNum 
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
         Height          =   375
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   30
         Width           =   1215
      End
      Begin MSDataListLib.DataCombo ComboStrTarget 
         Height          =   360
         Left            =   7170
         TabIndex        =   0
         Top             =   90
         Width           =   3765
         _ExtentX        =   6641
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
      Begin MSDataListLib.DataCombo ComboStrDesctination 
         Height          =   360
         Left            =   2370
         TabIndex        =   1
         Top             =   90
         Width           =   3765
         _ExtentX        =   6641
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
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "—ﬁ„ «·≈‘⁄«— "
         Height          =   285
         Index           =   2
         Left            =   1320
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   90
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   " „‰ «·„” Êœ⁄"
         Height          =   285
         Index           =   0
         Left            =   10890
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   " ≈·Ï «·„” Êœ⁄"
         Height          =   285
         Index           =   1
         Left            =   6180
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.TextBox TxtStkNo 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   8940
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   1170
      Width           =   1935
   End
   Begin VSFlex8Ctl.VSFlexGrid GridTarget 
      Height          =   5175
      Left            =   30
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1590
      Width           =   11865
      _cx             =   20929
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
      AllowUserFreezing=   2
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3450
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
            Picture         =   "FrmMovment.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMovment.frx":26D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMovment.frx":5533
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMovment.frx":7CD8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMovment.frx":A5D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMovment.frx":CAF9
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMovment.frx":F2B1
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMovment.frx":11CC5
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMovment.frx":14617
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMovment.frx":1738B
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMovment.frx":19BDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMovment.frx":1CA81
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMovment.frx":1F7DB
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMovment.frx":2217B
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMovment.frx":250DD
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMovment.frx":27B05
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMovment.frx":2A4BF
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMovment.frx":2CDF0
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMovment.frx":2F6F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMovment.frx":3215E
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMovment.frx":3510F
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMovment.frx":37A38
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMovment.frx":3A36D
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMovment.frx":3CAAD
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMovment.frx":3F41C
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMovment.frx":41CA4
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMovment.frx":43EB3
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMovment.frx":46810
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMovment.frx":4903A
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMovment.frx":4BAEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMovment.frx":4E52A
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMovment.frx":51040
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMovment.frx":53FDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMovment.frx":56DA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMovment.frx":59A20
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMovment.frx":5C8B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMovment.frx":5F4F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMovment.frx":6209F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "—ﬁ„ „Œ“‰Ì"
      Height          =   375
      Left            =   10980
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   1230
      Width           =   945
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "«·ﬂ„Ì… «·„‰ﬁÊ·…"
      ForeColor       =   &H00404040&
      Height          =   375
      Index           =   3
      Left            =   1260
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label LBalance 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   405
      Left            =   2250
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   1170
      Width           =   1305
   End
   Begin VB.Label LStkName 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
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
      Height          =   405
      Left            =   3570
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   1170
      Width           =   5355
   End
End
Attribute VB_Name = "FrmMovment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const ColIdTarget = 1
Const ColIdDestination = 2
Const ColByanId = 3
Const ColDocNum = 4
Const ColStkId = 5
Const ColStkNo = 6
Const ColStkName = 7
Const ColStrTargetId = 8
Const ColStrTargetName = 9
Const ColStrDestinationId = 10
Const ColStrDestinationName = 11
Const ColBalanceDestination = 12

Const ColStkId_1 = 1
Const ColStkno_1 = 2
Const ColStkname_1 = 3
Const ColCliPrice = 4
Const ColDealPrice = 5




Const ColId = 1
Const ColNo = 2
Const ColName = 3


Dim Flag As Boolean
Dim PrevBalance As Double
Dim TransferData As MovBetweenTowStoreType
Dim TypeRec As Boolean, PrevQty As Double
Dim Ok As Boolean
Dim Pos As Integer, RecNum   As Integer

Function GetCount(FlexGrid As VSFlexGrid) As Double
    GetCount = FlexGrid.Rows - 1
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

Sub DeleteRec(Vrow As Integer)
On Error GoTo ErrorHandler
Dim CurrBalance  As Double
de.con.BeginTrans
With GridTarget
    sqlText = "Delete From Stmov Where Id in (" & .TextMatrix(Vrow, ColIdTarget) & "," & .TextMatrix(Vrow, ColIdDestination) & ")" & " And ByanId = " & .TextMatrix(Vrow, ColByanId)
    de.con.Execute (sqlText)
    CurrBalance = GetBalance(.TextMatrix(Vrow, ColStkId), .TextMatrix(Vrow, ColStrTargetId))
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
Sub ReturnOldData(ByVal Vrow As Integer)
With GridTarget
    .TextMatrix(Vrow, ColBalanceDestination) = PrevQty
End With
End Sub

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

Function GetBalance(stkId As Double, Strid As Double) As Double
Dim RsBalance As New ADODB.Recordset
    sqlText = "Select FnlQnt From Stkinf Where Stkid=" & stkId & " And StrId=" & Strid
    Set RsBalance = de.con.Execute(sqlText)
    If RsBalance.RecordCount > 0 Then GetBalance = IIf(IsNull(RsBalance!fnlqnt), 0, RsBalance!fnlqnt)
End Function

Function ChkVariables() As Boolean
Dim Ok As Boolean
    Ok = True
    If TxtDocNum.Text = "" Then
        Ok = False
    End If
    If ComboStrTarget.BoundText = "" Then
        Ok = False
    End If
    If ComboStrDesctination.BoundText = "" Then
        Ok = False
    End If
    If TxtStkNo.Tag = "" Then
        Ok = False
    End If
    If TxtQty.Text = "" Or TxtQty.Text <= 0 Then
        Ok = False
    End If
    ChkVariables = Ok
End Function

Sub FillRec(QtyType As Integer)
With TransferData
    .ByanId = IIf(LByanId.Text = "", NewRec, LByanId.Text)
    LByanId.Text = .ByanId
    If QtyType = 0 Then 'IN
        .QtyTypeDestination = QtyType
        .QtyDestination = TxtQty.Text
        .StrDestination = ComboStrDesctination.BoundText
    Else 'Out =1
        .QtyTypeTarget = QtyType
        .QtyTarget = TxtQty.Text
        .StrTarget = ComboStrTarget.BoundText
    End If
    .MovDate = Format(Date, "mm/dd/yyyy")
    .stkId = TxtStkNo.Tag
    .DocNum = TxtDocNum.Text
    .Doctype = 3
End With
End Sub

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
Function IsDublicateStkId(ByanId As Double, stkId As Double) As Boolean
On Error GoTo ErrorHandler
Dim rs As New ADODB.Recordset
sqlText = "select count(*) CountRec from stmov where ByanId=" & ByanId & " and StkId=" & stkId & " and DocType=3 and QtyType=0"
Set rs = de.con.Execute(sqlText)
If rs!CountRec >= 1 Then
    IsDublicateStkId = True
Else
    IsDublicateStkId = False
End If
Exit Function
ErrorHandler:
IsDublicateStkId = False
MsgBox Err.Description
End Function
Function InsertRec() As Boolean
On Error GoTo ErrorHandler
de.con.BeginTrans
If LByanId.Text <> "" Then
    If IsDublicateStkId(LByanId.Text, TxtStkNo.Tag) Then
        InsertRec = False
        de.con.RollbackTrans
        Exit Function
    End If
End If
If ChkVariables Then
    FillRec 1
    If ChkQty(TransferData.stkId, TransferData.StrTarget, TransferData.QtyTarget, TransferData.QtyTypeTarget) Then
        With TransferData
            sqlText = "Insert into Stmov(ByanId , StkId  , StrId , Movdate , DocType , DocNum ,  Qty , QtyType,EmpNo)Values("
            sqlText = sqlText & .ByanId & "," & .stkId & "," & .StrTarget & ",'" & .MovDate & "'," & .Doctype & "," & .DocNum & "," & .QtyTarget & "," & .QtyTypeTarget & "," & empNo & ")"
            de.con.Execute (sqlText)
            .IdTarget = maxId
        End With
     Else
        InsertRec = False
        de.con.RollbackTrans
        Exit Function
    End If
    FillRec 0
    If ChkQty(TransferData.stkId, TransferData.StrDestination, TransferData.QtyDestination, TransferData.QtyTypeDestination) Then
        With TransferData
            sqlText = "Insert Into Stmov(ByanId , StkId  , StrId , Movdate , DocType , DocNum ,  Qty , QtyType,EmpNo)Values("
            sqlText = sqlText & .ByanId & "," & .stkId & "," & .StrDestination & ",'" & .MovDate & "'," & .Doctype & "," & .DocNum & "," & .QtyDestination & "," & .QtyTypeDestination & "," & empNo & ")"
            de.con.Execute (sqlText)
            .IdDestination = maxId
        End With
     Else
        InsertRec = False
        de.con.RollbackTrans
        Exit Function
    End If
    InsertRec = True
    de.con.CommitTrans
Else
    InsertRec = False
    de.con.RollbackTrans
End If
Exit Function
ErrorHandler:
    InsertRec = False
    de.con.RollbackTrans
    MsgBox Err.Description
End Function

Function maxId() As Double
Dim RsMaxId As New ADODB.Recordset
    sqlText = "Select Max(Id)MaxId From Stmov"
    Set RsMaxId = de.con.Execute(sqlText)
    maxId = RsMaxId!maxId
End Function

Sub AddToGrid()
Dim Vrow As Integer
With GridTarget
    .AddItem ""
     Vrow = .Rows - 1
    .TextMatrix(Vrow, ColIdTarget) = TransferData.IdTarget
    .TextMatrix(Vrow, ColIdDestination) = TransferData.IdDestination
    .TextMatrix(Vrow, ColByanId) = TransferData.ByanId
    .TextMatrix(Vrow, ColDocNum) = TransferData.DocNum
    .TextMatrix(Vrow, ColStkId) = TransferData.stkId
    .TextMatrix(Vrow, ColStkNo) = TxtStkNo.Text
    .TextMatrix(Vrow, ColStkName) = LStkName.Caption
    .TextMatrix(Vrow, ColStrTargetId) = TransferData.StrTarget
    .TextMatrix(Vrow, ColStrTargetName) = StrName(TransferData.StrTarget)
    .TextMatrix(Vrow, ColStrDestinationId) = TransferData.StrDestination
    .TextMatrix(Vrow, ColStrDestinationName) = StrName(TransferData.StrDestination)
    .TextMatrix(Vrow, ColBalanceDestination) = TransferData.QtyDestination
    
    .Col = ColIdTarget
    .Sort = flexSortGenericDescending
End With
End Sub

Sub NewMovmentRec()
        LByanId.Text = ""
        TxtDocNum.Text = ""
        GridTarget.Rows = 1
        GridTarget.Cols = 14
        FillFormatString GridTarget
        TypeRec = True
        Frame1.Enabled = True
        ComboStrTarget.SetFocus
        TxtDocNum.Text = GetMaxDocumner(3)
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

'Function InsertRec(ByVal VRow As Integer, ByVal TypeRec As Boolean) As Boolean
'On Error GoTo ErrorHandler
'FillRec VRow, TypeRec
'With Stmov
'    Sqltext = "Insert into Stmov(ByanId , StkId , StkIdType , StrId , Movdate , DocType , Qty , QtyType,EmpNo)Values("
'    Sqltext = Sqltext & .ByanId & "," & .StkId & "," & .StkIdType & "," & .StrId & ",'" & .MovDate & "'," & .DocType & "," & .Qty & "," & .QtyType & "," & EmpNo & ")"
'    de.con.Execute (Sqltext)
'End With
'InsertRec = True
'Exit Function
'ErrorHandler:
'InsertRec = False
'MsgBox Err.Description
'End Function

Function NewRec() As Double
Dim RsMax As New ADODB.Recordset
sqlText = "Select isnull(Max(ByanId),0) as MaxByanId From Stmov"
Set RsMax = de.con.Execute(sqlText)
If RsMax!maxByanId = 0 Then
    NewRec = 1
Else
    NewRec = RsMax!maxByanId + 1
End If
End Function

'Sub FillRec(VRow As Integer, TypeRec As Boolean)
'With GridTarget
'    Stmov.StkId = .TextMatrix(VRow, ColStkId)
'    Stmov.StkIdType = .TextMatrix(VRow, ColStkIdType)
'    Stmov.MovDate = Date
'    Stmov.DocType = 3 ' ‰ﬁ·
'    If TypeRec Then ' Out
'        Stmov.Qty = .TextMatrix(VRow, ColBalanceDestination)
'        Stmov.QtyType = 1
'        Stmov.StrId = .TextMatrix(VRow, ColStrTargetId)
'    Else ' In
'        Stmov.Qty = .TextMatrix(VRow, ColBalanceDestination)
'        Stmov.QtyType = 0
'        Stmov.StrId = .TextMatrix(VRow, ColStrDestinationId)
'    End If
'End With
'End Sub

'Sub SaveRec()
'On Error GoTo ErrorHandler
'With GridTarget
'    de.con.BeginTrans
'    Stmov.ByanId = NewRec
'    For i = 1 To .Rows - 1
'        If Not InsertRec(i, True) Then
'            de.con.RollbackTrans
'            Exit Sub
'        End If
'        If Not InsertRec(i, False) Then
'            de.con.RollbackTrans
'            Exit Sub
'        End If
'    Next
'    LByanId.Text = Stmov.ByanId
'    de.con.CommitTrans
'End With
'Exit Sub
'ErrorHandler:
'de.con.RollbackTrans
'MsgBox Err.Description
'End Sub
Function ChkStkNo(stkId As Double) As Boolean
Dim Ok As Boolean
Dim RsChk As New ADODB.Recordset
Ok = False
With GridTarget
    For i = 1 To .Rows - 1
        .Row = i
        If .TextMatrix(.Row, ColStkId) = stkId Then
            Ok = True
            ChkStkNo = Ok
            Exit Function
        End If
    Next
End With
ChkStkNo = Ok
End Function
'Sub InsertIntoGrid(StkId As Double, StkNo As String, StkName As String, StrTargetId As Integer, StrTargetName As String, Balance As Double, StrDesctinationId As Integer, StrDestinationName As String)
'With GridTarget
'If Not ChkStkNo(StkId) Then
'    .AddItem ""
'    .TextMatrix(.Rows - 1, ColStkId) = StkId
'    .TextMatrix(.Rows - 1, ColStkIdType) = ChkMod_Stk.Value
'    .TextMatrix(.Rows - 1, ColStkNo) = StkNo
'    .TextMatrix(.Rows - 1, ColStkName) = StkName
'
'    .TextMatrix(.Rows - 1, ColStrTargetId) = StrTargetId
'    .TextMatrix(.Rows - 1, ColStrTargetName) = StrTargetName
'    .TextMatrix(.Rows - 1, ColBalanceTarget) = Balance
'
'    .TextMatrix(.Rows - 1, ColStrDestinationId) = StrDesctinationId
'    .TextMatrix(.Rows - 1, ColStrDestinationName) = StrDestinationName
'    .TextMatrix(.Rows - 1, ColBalanceDestination) = 0
'    FillFormatString GridTarget
'End If
'End With
'End Sub
Sub FillFormatString(FlexGrid As VSFlexGrid)
    fs = "|>" + "IDTarget"
    fs = fs + "|>" + "IdDestination"
    fs = fs + "|>" + "ByanId"
    fs = fs + "|>" + "DocNum"
    fs = fs + "|>" + "StkId"
    fs = fs + "|>" + "«·—ﬁ„"
    fs = fs + "|>" + "«·≈”„"
    fs = fs + "|>" + "StrId"
    fs = fs + "|>" + "«·√”«”Ì"
    fs = fs + "|>" + "StrId"
    fs = fs + "|>" + "«·Âœ›"
    fs = fs + "|>" + "ﬂ„Ì… «·‰ﬁ·"
    
    With FlexGrid
        .Cols = 13
        .FormatString = fs
        .ColWidth(ColIdTarget) = 0
        .ColWidth(ColIdDestination) = 0
        .ColWidth(ColByanId) = 0
        .ColWidth(ColDocNum) = 0
        .ColWidth(ColStkId) = 0
        .ColWidth(ColStrTargetId) = 0
        .ColWidth(ColStrDestinationId) = 0
        .ColWidth(ColStrTargetName) = 0
        .ColWidth(ColStrDestinationName) = 0
        SetColWidths ColStkNo, FlexGrid
        SetColWidths ColStkName, FlexGrid
        SetColWidths ColBalanceDestination, FlexGrid
   End With
End Sub
Sub SetColWidths(ColNo As Integer, FlexGrid As VSFlexGrid)
    With FlexGrid
        .AutoSize ColNo
    End With
End Sub

Sub FillList(sqlText As String, Field1 As String, Field2 As String, List As VSFlexGrid)
    Set rs = de.con.Execute(sqlText)
    If rs.RecordCount > 0 Then
        Set List.DataSource = rs
        FillFormatVSFlex List
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

Sub FillFormatVSFlex(FlexGrid As VSFlexGrid)

    fs = "|ID"
    fs = fs + "|>" + "«·—ﬁ„ «·„Œ“‰Ì"
    fs = fs + "|>" + "«·≈”„"
    fs = fs + "|>" + "”⁄— «·“»Ê‰"
    fs = fs + "|>" + "”⁄— «· «Ã—"
    With FlexGrid
        .Visible = False
        .FormatString = fs
            .ColWidth(ColStkId_1) = 0
            SetColWidths ColStkno_1, FlexGrid
            SetColWidths ColStkname_1, FlexGrid
            SetColWidths ColCliPrice, FlexGrid
            SetColWidths ColDealPrice, FlexGrid
            .Visible = True
    End With

End Sub

Sub FillActiveControl(List As VSFlexGrid)
    With List
        If ActiveControl.Text <> "" Then
            If Not ActiveControl.DataChanged Then Exit Sub
            Flag = False
            ActiveControl.Text = IIf(.Visible = False, "", .TextMatrix(.Row, ColStkno_1))
            LStkName.Caption = IIf(.Visible = False, "", .TextMatrix(.Row, ColStkname_1))
            LStkName.Tag = IIf(.Visible = False, "", .TextMatrix(.Row, ColStkId_1))
            Lbalance.Caption = IIf(.Visible = False, "", .TextMatrix(.Row, ColStkFnlQnt))
            Flag = True
            ActiveControl.Tag = IIf(.Visible = False, "", .TextMatrix(.Row, ColStkId_1))
        Else
            ActiveControl.Text = ""
            ActiveControl.Tag = ""
            LStkName.Caption = ""
            LStkName.Tag = ""
            Lbalance.Caption = ""
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

Sub init()
    Flag = False
    Ok = True
    Dim RsSTr As New ADODB.Recordset
    sqlText = "Select Id , StrNo , StrName From NameStr Order By StrNo"
    Set RsSTr = de.con.Execute(sqlText)
    
    If RsSTr.RecordCount > 0 Then
        Set ComboStrTarget.RowSource = RsSTr
        ComboStrTarget.listField = "StrName"
        ComboStrTarget.BoundColumn = "Id"
        RsSTr.MoveFirst
        ComboStrTarget.BoundText = RsSTr!Id
    End If
    'Target
    GridTarget.Rows = 1
    FillFormatString GridTarget
    GridTarget.Editable = flexEDKbdMouse
    
    'Destination
    If RsSTr.RecordCount > 0 Then
        Set ComboStrDesctination.RowSource = RsSTr
        ComboStrDesctination.listField = "StrName"
        ComboStrDesctination.BoundColumn = "Id"
        ComboStrDesctination.BoundText = RsSTr!Id
    End If
    top = 0
    left = 0
    TypeRec = False
    Flag = True
End Sub

'Private Sub ChkMod_Stk_Click()
'
'If ChkMod_Stk.Value Then
'    ChkMod_Stk.Caption = "„ÊœÌ·"
'Else
'    ChkMod_Stk.Caption = "—ﬁ„ „Œ“‰Ì"
'End If
'
'End Sub

Private Sub ChkMod_Stk_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{Tab}"
    SendKeys "{Home}+{End}"
End If

End Sub

Private Sub ComboStrDesctination_Change()
    TxtDocNum.Text = GetMaxDocumner(3)
End Sub

Private Sub ComboStrDesctination_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{Tab}"
    SendKeys "{Home}+{End}"
End If

End Sub

Private Sub ComboStrDesctination_Validate(cancel As Boolean)
If ComboStrDesctination.BoundText = ComboStrTarget.BoundText Or ComboStrDesctination.BoundText = "" Then
    cancel = True
    MsgBox "√Œ· «·„” Êœ⁄ «·„ﬁ«»·", vbInformation, " ‰»ÌÂ"
    ComboStrDesctination.SetFocus
End If
End Sub

Function GetMaxDocumner(Doctype As Integer) As Integer
On Error GoTo ErrorHandler
    Dim RsGetNum As New ADODB.Recordset
    sqlText = "Select isnull(max(DocNum),0)MaxDocNum From Stmov Where DocType=" & Doctype
    Set RsGetNum = de.con.Execute(sqlText)
    GetMaxDocumner = RsGetNum!MaxDocNum + 1
Exit Function
ErrorHandler:
MsgBox Err.Description
GetMaxDocumner = -1
End Function
Private Sub ComboStrTarget_Change()
TxtDocNum.Text = GetMaxDocumner(3)
End Sub

'Function Fn_GetDocNum(FStrId As Integer, TStrId As Integer) As Integer
'
'End Function

Private Sub ComboStrTarget_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{Tab}"
    SendKeys "{Home}+{End}"
End If
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift And vbAltMask Then
    If KeyCode = vbKeyN Then
        If LByanId.Text <> "" Then
            If MsgBox("Â·  —Ìœ ≈‰‘«¡ Õ—ﬂ… ÃœÌœ…", vbYesNo + vbQuestion + vbDefaultButton2, "Õ—ﬂ… ÃœÌœ…") = vbYes Then
                NewMovmentRec
            End If
        Else
            NewMovmentRec
        End If
    End If
End If
End Sub

Private Sub Form_Load()
init
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

Private Sub GridTarget_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrorHandler
Dim CurrQty As Double, CurrBalanceTarget  As Double, CurrBalanceDestination As Double
With GridTarget
    de.con.BeginTrans
    CurrQty = .TextMatrix(Row, ColBalanceDestination)
    sqlText = "Update Stmov Set Qty = " & CurrQty & ",QtyType=1,EmpNo=" & empNo & " Where Id=" & .TextMatrix(Row, ColIdTarget)
    de.con.Execute (sqlText)
    
    sqlText = "Update Stmov Set Qty = " & CurrQty & ",QtyType=0,EmpNo=" & empNo & " Where Id=" & .TextMatrix(Row, ColIdDestination)
    de.con.Execute (sqlText)
    CurrBalanceTarget = GetBalance(.TextMatrix(Row, ColStkId), .TextMatrix(Row, ColStrTargetId))
    CurrBalanceDestination = GetBalance(.TextMatrix(Row, ColStkId), .TextMatrix(Row, ColStrDestinationId))
    
    If CurrBalanceTarget >= 0 And CurrBalanceDestination >= 0 Then
        de.con.CommitTrans
    Else
        de.con.RollbackTrans
        ReturnOldData Row
        MsgBox "«·—’Ìœ ·«Ì”„Õ", vbExclamation, " ‰»ÌÂ"
    End If
    Lbalance.Caption = GetBalance(.TextMatrix(Row, ColStkId), .TextMatrix(Row, ColStrTargetId))
End With
Exit Sub
ErrorHandler:
MsgBox Err.Description
de.con.RollbackTrans
End Sub

Private Sub GridTarget_BeforeEdit(ByVal Row As Long, ByVal Col As Long, cancel As Boolean)
If TypeRec Then
    If Col <> ColBalanceDestination Then
        cancel = True
    Else
        PrevQty = GridTarget.TextMatrix(Row, ColBalanceDestination)
    End If
Else
    cancel = True
End If
End Sub

Private Sub GridTarget_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrorHandler
If TypeRec Then
    If KeyCode = vbKeyDelete Then
        If MsgBox("Â· √‰  „ √ﬂœ „‰ ⁄„·Ì… «·Õ–›", vbYesNo + vbDefaultButton2, "Õ–› «·”Ã·«  «·„Õœœ…") = vbYes Then
            DeleteRec GridTarget.Row
            If GridTarget.Row = 0 Then
                Lbalance.Caption = GetBalance(TxtStkNo.Tag, ComboStrTarget.BoundText)
            Else
                Lbalance.Caption = GetBalance(GridTarget.TextMatrix(GridTarget.Row, ColStkId), ComboStrTarget.BoundText)
            End If
            TxtNum.Text = GetCount(GridTarget)
        End If
    End If
End If
Exit Sub
ErrorHandler:
MsgBox Err.Description
End Sub

Private Sub GridTarget_RowColChange()
'With GridTarget
'If Flag Then
'    Flag = False
'    If .Row = 0 Then Exit Sub
'    TxtStkNo.Text = .TextMatrix(.Row, ColStkNo)
'    Flag = True
'    TxtStkNo.Tag = .TextMatrix(.Row, ColStkId)
'    TxtQty.Text = .TextMatrix(.Row, ColStkQty)
'    LStkName.Caption = .TextMatrix(.Row, ColStkName)
'    Lbalance.Caption = GetBalance(.TextMatrix(.Row, ColStkId), .TextMatrix(.Row, ColStrTargetId))
'End If
'End With
End Sub
Function SearchRec(Optional i) As String
On Error GoTo ErrorHandler
Dim sqlText As String

If IsMissing(i) Then
    sqlText = "select 0 Chk , ByanId , s1.StrId , StrNo , DocNum , Convert(Varchar(10),MovDate,103)MovDate , isnull(Correspondence,'') Correspondence, isnull(Correspondence,'') Correspondence, CountryNo , CountryName  , Count(*) , TypeName  from StmovQry s1 Where ByanId <> 0"
Else
    sqlText = "Select StmovQry.Id , StmovQry.ByanId , StmovQry.Stkid  , StmovQry.StkNo , StmovQry.StkName , StmovQry.StrId , StmovQry.StrNo , StmovQry.StrName , StmovQry.MovDate , StmovQry.DocType , StmovQry.Qty , StmovQry.QtyType , StmovQry.FnlQnt ,  StmovQry.Correspondence, CountryNo , [In] , [Out]   From StmovQry Where  StmovQry.ByanId =" & Val(LByanId.Text) & " and StmovQry.strid=" & ComboStrDesctination.BoundText
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
Sub FillGrid(ByanId As Double)
Dim rs As New ADODB.Recordset
sqlText = "select  idtarget , iddestination, t1.byanid , t1.DocNum , t1.StkId   , t1.StkNo , t1.StkName , t1.StrId strtarget , t1.strname  strnametarget , t2.StrId striddestination, t2.strname strnamedestination , t1.qty   from ("
sqlText = sqlText & "select s1.Id  idtarget ,  s1.byanid , docnum ,  s1.stkid , StkNo , StkName , StrId , strname , qty   from stmov s1 inner join costock c1 on s1.StkId = c1.Id inner join namestr n1 on s1.StrId = n1.Id"
sqlText = sqlText & "    Where ByanId = " & ByanId & " And DocType = 3 And QtyType = 1"
sqlText = sqlText & ")t1  full outer join"
sqlText = sqlText & "("
sqlText = sqlText & "select s1.Id  iddestination ,  s1.byanid , docnum ,  s1.stkid  , StkNo , StkName , StrId , strname , qty   from stmov s1 inner join costock c1 on s1.StkId = c1.Id inner join namestr n1 on s1.StrId = n1.Id"
sqlText = sqlText & "    Where ByanId = " & ByanId & "  And DocType = 3 And QtyType = 0"
sqlText = sqlText & ")t2 on  t1.byanid = t2.ByanId and t1.StkId = t2.StkId"

Set rs = de.con.Execute(sqlText)
Set GridTarget.DataSource = rs
'GridTarget.ColDataType(ColChk) = flexDTBoolean
FillFormatString GridTarget
LByanId.Text = GByanId
ComboStrTarget.BoundText = rs!StrTarget
ComboStrDesctination.BoundText = rs!StridDestination
TxtDocNum.Text = IIf(IsNull(rs!DocNum), 0, rs!DocNum)

'TxtDate.Text = Format(Rs!MovDate, "dd/mm/yyyy")
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
        If LByanId.Text <> "" Then
            If MsgBox("Â·  —Ìœ ≈‰‘«¡ Õ—ﬂ… ÃœÌœ…", vbYesNo + vbQuestion + vbDefaultButton2, "Õ—ﬂ… ÃœÌœ…") = vbYes Then
                 NewMovmentRec
            End If
        Else
            NewMovmentRec
        End If
        Case 5
            PrintData
         Case 7
        GByanType = 3 ' ‰ﬁ·
        FrmPrintByans.Show 1
        If GByanId <> 0 Then
            TypeRec = True
            FillGrid GByanId
        End If
        Case 9
            If MsgBox("Â·  —Ìœ «·Œ—ÊÃ", vbYesNo + vbQuestion + vbDefaultButton2, "Œ—ÊÃ") = vbYes Then
                Unload Me
            End If
    End Select
End Sub

Private Sub TxtDocNum_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{Tab}"
    SendKeys "{Home}+{End}"
End If
End Sub

Private Sub TxtDocNum_Validate(cancel As Boolean)
If TypeRec Then
    If TxtDocNum.Text = "" Then
        cancel = True
        MsgBox "√œŒ· —ﬁ„ «·≈‘⁄«—", vbInformation, " ‰»ÌÂ"
        TxtDocNum.SetFocus
    End If
End If
End Sub

Private Sub TxtQty_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeRec Then 'New
        If InsertRec Then
            AddToGrid
            FillFormatString GridTarget
            Frame1.Enabled = False
            Lbalance.Caption = GetBalance(TxtStkNo.Tag, ComboStrTarget.BoundText)
            TxtNum.Text = GetCount(GridTarget)
            TxtStkNo.SetFocus
            SendKeys "{home}+{end}"
        Else
            TxtStkNo.SelStart = 0
            TxtStkNo.SelLength = Len(TxtStkNo.Text)
            TxtStkNo.SetFocus
            MsgBox "«·—’Ìœ ·«Ì”„Õ" & Chr(13) & "√Ê" & Chr(13) & "Œÿ√ ›Ì «·„œŒ·« " & Chr(13) & "√Ê" & Chr(13) & "«·„«œ… „ﬂ——…", vbExclamation, " ‰»ÌÂ"
        End If
        TxtQty.Text = ""
    Else
        MsgBox "⁄„·Ì… ≈” ⁄—«÷", vbExclamation, " ‰»ÌÂ"
    End If
End If
End Sub
Sub ChangeCursor(ByVal X As Integer)
If X = 1 Then
    With TxtClientName
       Grid.top = SSFrame1.top + .top + .Height
       Grid.left = SSFrame1.left + .left
       Grid.Width = .Width
    End With
ElseIf X = 2 Then
    With TxtModelName
       Grid.top = SSFrame1.top + .top + .Height
       Grid.left = SSFrame1.left + .left
       Grid.Width = .Width
End With
ElseIf X = 3 Then
    With TxtStkNo
       Grid.top = .top + .Height
       Grid.left = .left
       Grid.Width = .Width + LStkName.Width
End With
ElseIf X = 4 Then
    With TxtRecipient
       Grid.top = sstab1.top + SSFrame2.top + .top + .Height
       Grid.left = SSFrame2.left + .left
       Grid.Width = .Width
End With
ElseIf X = 5 Then
    With TxtFamNo
       Grid.top = SSFrameModel.top + sstab1.top + .top + .Height
       Grid.left = SSFrameModel.left + .left
       Grid.Width = .Width
End With
ElseIf X = 6 Then
    With TxtModelName
       Grid.top = .top + .Height
       Grid.left = .left
       Grid.Width = .Width
End With
ElseIf X = 7 Then
    With TxtUnExecutedReason
       Grid.top = sstab1.top + .top + .Height
       Grid.left = .left
       Grid.Width = .Width
End With
ElseIf X = 8 Then
    With TxtError
       Grid.top = .top + .Height
       Grid.left = .left
       Grid.Width = .Width
End With
ElseIf X = 9 Then
    With TxtitemName
       Grid.top = .top + .Height
       Grid.left = .left
       Grid.Width = .Width
End With
ElseIf X = 10 Then
    With txtCallNo
       Grid.top = sstab1.top + .top + .Height
       Grid.left = sstab1.left + .left
       Grid.Width = .Width
End With
End If
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
    sqlText = "Select top 15 Id , ltrim(rtrim(StkNo))StkNo , ltrim(rtrim(StkName))StkName  ,CliPrice , DealPrice from CoStock Where StkName Like" & LikeExpression(TxtStkNo.Text) & " or StkNo like '" & TxtStkNo.Text & "%'"
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

Private Sub TxtStkNo_GotFocus()
Pos = 3
End Sub

Private Sub TxtStkNo_KeyDown(KeyCode As Integer, Shift As Integer)
    Flag = True
    If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then MoveCursor KeyCode
    Flag = True
End Sub

Sub GoToIntoStkId(stkId As Double)
On Error GoTo ErrorHandler
With GridTarget
For i = 1 To .Rows - 1
    If .TextMatrix(i, ColStkId) = stkId Then
        .Row = i
        Exit Sub
    End If
Next
End With
Exit Sub
ErrorHandler:
MsgBox Err.Description
End Sub
Function IsExists(stkId As Double, ByanId As Double) As Boolean
On Error GoTo ErrorHandler

Dim rs As New ADODB.Recordset
sqlText = "select stkid from stmov where stkid=" & stkId & "and byanid =" & ByanId
Set rs = de.con.Execute(sqlText)
If rs.RecordCount >= 1 Then
    IsExists = True
Else
    IsExists = False
End If
Exit Function
ErrorHandler:
IsExists = False
MsgBox Err.Description
End Function
Private Sub txtStkNo_KeyPress(KeyAscii As Integer)

On Error GoTo ErrorHandler
If KeyAscii = 13 Then
        Grid.Visible = False
        If DataOk(TxtStkNo.Text) Then
            TxtStkNo.Tag = GetStkId(TxtStkNo.Text, ComboStrTarget.BoundText)
'            If IsExists(TxtStkNo.Tag, Val(LByanId.Text)) Then
'                GoToIntoStkId TxtStkNo.Tag
'            End If
            If TxtStkNo.Tag <> 0 Then
                LStkName.Caption = GetStkName(TxtStkNo.Tag)
                Lbalance.Caption = GetBalance(TxtStkNo.Tag, ComboStrTarget.BoundText)
            Else
                LStkName.Caption = ""
                Lbalance.Caption = ""
                MsgBox "«·„«œ… €Ì— „⁄—›… ›Ì «·„” Êœ⁄", vbInformation, " ‰»ÌÂ"
                SendKeys "{home}+{end}"
                Exit Sub
            End If
        Else
            Ok = False
            TxtStkNo.Text = Grid.TextMatrix(IIf(Grid.Row = 0, 1, Grid.Row), ColNo)
            Ok = True
            TxtStkNo.Tag = GetStkId(TxtStkNo.Text, ComboStrTarget.BoundText)
            If TxtStkNo.Tag <> 0 Then
                LStkName.Caption = GetStkName(TxtStkNo.Tag)
                Lbalance.Caption = GetBalance(TxtStkNo.Tag, ComboStrTarget.BoundText)
            Else
                LStkName.Caption = ""
                Lbalance.Caption = ""
                MsgBox "«·„«œ… €Ì— „⁄—›… ›Ì «·„” Êœ⁄", vbInformation, " ‰»ÌÂ"
                SendKeys "{home}+{end}"
                Exit Sub
            End If
        End If
        TxtQty.Text = 0
        TxtQty.SetFocus
        SendKeys "{home}+{end}"
    End If
Exit Sub
ErrorHandler:
Grid.Visible = False
MsgBox Err.Description
End Sub
Function GetStkId(stkno As String, Strid As Integer) As Double
    Dim rs As New ADODB.Recordset
    sqlText = "Select c1.Id From CoStock c1 inner join Stkinf s1 on c1.Id = s1.StkId  And StrId=" & Strid & " Where c1.StkNo ='" & stkno & "'"
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
