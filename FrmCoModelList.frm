VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form FrmCoModelList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ÊÑãíÒ ÞÇÆãÉ ÇáãæÏíáÇÊ"
   ClientHeight    =   5010
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9690
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   5010
   ScaleWidth      =   9690
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   2085
      Left            =   1500
      TabIndex        =   7
      Top             =   2010
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
   Begin VB.TextBox TxtModelName 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   2550
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   630
      Width           =   5955
   End
   Begin VB.TextBox TxtitemName 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   2550
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   1050
      Width           =   5955
   End
   Begin Crystal.CrystalReport cr1 
      Left            =   540
      Top             =   3870
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VSFlex8Ctl.VSFlexGrid FlexGrid 
      Height          =   3495
      Left            =   30
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1440
      Width           =   9585
      _cx             =   16907
      _cy             =   6165
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
      Rows            =   17
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5130
      Top             =   270
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   74
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCoModelList.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCoModelList.frx":05ED
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCoModelList.frx":09E5
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCoModelList.frx":0F06
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCoModelList.frx":1375
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCoModelList.frx":1761
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCoModelList.frx":1BF8
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCoModelList.frx":208F
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCoModelList.frx":267C
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCoModelList.frx":2AA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCoModelList.frx":2ED0
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCoModelList.frx":3335
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCoModelList.frx":379A
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCoModelList.frx":3BCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCoModelList.frx":3FFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCoModelList.frx":445E
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCoModelList.frx":48C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCoModelList.frx":4CCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCoModelList.frx":50D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCoModelList.frx":5551
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCoModelList.frx":59D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCoModelList.frx":5E00
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCoModelList.frx":6230
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCoModelList.frx":681D
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCoModelList.frx":6C31
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCoModelList.frx":70E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCoModelList.frx":75BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCoModelList.frx":79E7
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCoModelList.frx":7E78
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCoModelList.frx":82CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCoModelList.frx":86C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCoModelList.frx":8ABB
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCoModelList.frx":8FAF
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCoModelList.frx":94E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCoModelList.frx":99BD
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCoModelList.frx":9DD5
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCoModelList.frx":A1E9
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCoModelList.frx":A700
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCoModelList.frx":AB9D
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCoModelList.frx":B06E
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCoModelList.frx":B532
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCoModelList.frx":BA49
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCoModelList.frx":BEFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCoModelList.frx":C352
            Key             =   ""
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCoModelList.frx":C7D5
            Key             =   ""
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCoModelList.frx":CC26
            Key             =   ""
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCoModelList.frx":D065
            Key             =   ""
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCoModelList.frx":D516
            Key             =   ""
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCoModelList.frx":D8DD
            Key             =   ""
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCoModelList.frx":DDA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCoModelList.frx":E267
            Key             =   ""
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCoModelList.frx":E712
            Key             =   ""
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCoModelList.frx":EBBD
            Key             =   ""
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCoModelList.frx":F071
            Key             =   ""
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCoModelList.frx":F525
            Key             =   ""
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCoModelList.frx":F90A
            Key             =   ""
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCoModelList.frx":FD2D
            Key             =   ""
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCoModelList.frx":10150
            Key             =   ""
         EndProperty
         BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCoModelList.frx":10531
            Key             =   ""
         EndProperty
         BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCoModelList.frx":10912
            Key             =   ""
         EndProperty
         BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCoModelList.frx":10D42
            Key             =   ""
         EndProperty
         BeginProperty ListImage62 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCoModelList.frx":11172
            Key             =   ""
         EndProperty
         BeginProperty ListImage63 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCoModelList.frx":114FD
            Key             =   ""
         EndProperty
         BeginProperty ListImage64 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCoModelList.frx":1193B
            Key             =   ""
         EndProperty
         BeginProperty ListImage65 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCoModelList.frx":11CC2
            Key             =   ""
         EndProperty
         BeginProperty ListImage66 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCoModelList.frx":120C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage67 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCoModelList.frx":124D5
            Key             =   ""
         EndProperty
         BeginProperty ListImage68 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCoModelList.frx":12852
            Key             =   ""
         EndProperty
         BeginProperty ListImage69 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCoModelList.frx":12BC2
            Key             =   ""
         EndProperty
         BeginProperty ListImage70 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCoModelList.frx":13038
            Key             =   ""
         EndProperty
         BeginProperty ListImage71 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCoModelList.frx":13442
            Key             =   ""
         EndProperty
         BeginProperty ListImage72 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCoModelList.frx":137CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage73 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCoModelList.frx":13BF8
            Key             =   ""
         EndProperty
         BeginProperty ListImage74 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCoModelList.frx":14022
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   9690
      _ExtentX        =   17092
      _ExtentY        =   1058
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "ÇáÈÍÜÜÜÜÜÜË"
            ImageIndex      =   28
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "ÎÑæÌ ãä ÇáÈÑäÇãÌ"
            ImageIndex      =   74
         EndProperty
      EndProperty
   End
   Begin VB.Label LSymbol 
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
      Left            =   60
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   660
      Width           =   2445
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "ÇáãæÏíá"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   7
      Left            =   9090
      TabIndex        =   5
      Top             =   660
      Width           =   495
   End
   Begin VB.Label LItemName 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   60
      TabIndex        =   4
      Top             =   1080
      Width           =   2445
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "ÇáÑÞã ÇáãÎÒäí"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   8
      Left            =   8640
      TabIndex        =   3
      Top             =   1110
      Width           =   930
   End
End
Attribute VB_Name = "FrmCoModelList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Ok As Boolean, Flag As Boolean
Dim Pos As Integer

Const ColNo = 1
Const ColName = 2
Const col3 = 3

Const ColSerial = 1
Const ColModNo = 2
Const Colsymbol = 3
Const ColModelName = 4
Const ColStkNo = 5
Const ColStkName = 6

Dim ModelListItemRec  As ModelListItem

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
    Fs = Fs + "|>" + ""
    With Grid
        .FormatString = Fs
        .Cols = 4
        If Pos = 2 Then
            SetColWidths ColNo, Grid
        Else
            .ColWidth(ColNo) = 0
        End If
        If Pos = 1 Then
            SetColWidths col3, Grid
        Else
            .ColWidth(col3) = 0
        End If
        SetColWidths ColName, Grid
    End With
ElseIf i = 2 Then
    Fs = "|>" + "Serial"
    Fs = Fs + "|>" + "ModNo"
    Fs = Fs + "|>" + "ÑãÒ ÇáãæÏíá"
    Fs = Fs + "|>" + "ÔÑÍ ÇáãæÏíá"
    Fs = Fs + "|>" + "ÑÞã ÇáãÇÏÉ"
    Fs = Fs + "|>" + "ÅÓã ÇáãÇÏÉ"
    With FlexGrid
        .FormatString = Fs
        .Cols = 7
        .ColWidth(ColSerial) = 0
        .ColWidth(ColModNo) = 0
        SetColWidths Colsymbol, FlexGrid
        SetColWidths ColModelName, FlexGrid
        SetColWidths ColStkNo, FlexGrid
        SetColWidths ColStkName, FlexGrid
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
    With TxtModelName
       Grid.Top = .Top + .Height
       Grid.Left = .Left
       Grid.Width = .Width
End With
ElseIf X = 2 Then
    With TxtitemName
       Grid.Top = .Top + .Height
       Grid.Left = .Left
       Grid.Width = .Width
    End With
End If
End Sub

Sub init()
Top = 0
Left = 0
Ok = True
FlexGrid.Rows = 1
FillFormating 2
End Sub

Private Sub FlexGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim FirstRow As Integer, LastRow As Integer, Vrow As Integer
With FlexGrid
If KeyCode = vbKeyDelete Then
    If MsgBox("åá ÃäÊ ãÊÃßÏ ãä ÚãáíÉ ÇáÍÐÝ", vbYesNo + vbDefaultButton2, "ÍÐÝ ÇáÓÌáÇÊ ÇáãÍÏÏÉ") = vbYes Then
        If .Row >= .RowSel Then
            FirstRow = .Row
            LastRow = .RowSel
        Else
            FirstRow = .RowSel
            LastRow = .Row
        End If
        For i = FirstRow To LastRow Step -1
            Vrow = i
            If DeleteRow(FlexGrid, Vrow, ColSerial, "CoMaintModelItems", "Id") Then
                .RemoveItem Vrow
            End If
        Next
    End If
End If
End With
End Sub

Private Sub Form_Load()
init
End Sub

Sub SearchData(ModNo As Integer, StkNo As String)
Dim Rs As New ADODB.Recordset
Sqltext = "select id , Modno , symbol , name , Stkno , stkname from CoMaintModelItemsQry where Id<> -1 "
If ModNo <> 0 Then
    Sqltext = Sqltext & " and ModNo=" & ModNo
End If
If StkNo <> "" Then
    Sqltext = Sqltext & " and StkNo='" & StkNo & "'"
End If
Sqltext = Sqltext & " Order By Id Desc"
Set Rs = de.con.Execute(Sqltext)
Set FlexGrid.DataSource = Rs
FillFormating 2
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
        Case 1
            SearchData Val(TxtModelName.Tag), TxtitemName.Tag
        Case 3
            Unload Me
End Select
End Sub

Private Sub TxtitemName_Change()
On Error GoTo ERRORHANDLER
Dim RsSearch As New ADODB.Recordset
If TxtitemName.Text = "" Then
    TxtitemName.Tag = "0"
    Grid.Visible = False
    Exit Sub
End If

If Ok Then
    Flag = False
    Sqltext = "Select Top 100 StkNo , StkName  from CoStock Where StkName Like" & LikeExpression(TxtitemName.Text) & " or StkNo like '" & TxtitemName.Text & "%'"
    Set RsSearch = de.con.Execute(Sqltext)
    If RsSearch.RecordCount > 0 Then
        Set Grid.DataSource = RsSearch
        Grid.Row = 0
        FillFormating 1
        ChangeCursor 2
        Grid.Visible = True
    Else
        TxtitemName.Tag = ""
        Grid.Visible = False
    End If
    Flag = True
End If
Exit Sub
ERRORHANDLER:
MsgBox Err.Description
End Sub

Private Sub TxtitemName_GotFocus()
Pos = 2
End Sub


Private Sub TxtitemName_KeyDown(KeyCode As Integer, Shift As Integer)
    Flag = True
    If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then MoveCursor KeyCode, Grid
    Flag = True

End Sub
Function FillVariables() As Boolean
If TxtModelName.Tag <> 0 And TxtitemName.Tag <> "" Then
    FillVariables = True
Else
    FillVariables = False
End If

End Function
Function fillstructure() As Boolean
On Error GoTo ERRORHANDLER
If FillVariables Then
    With ModelListItemRec
        .ModNo = TxtModelName.Tag
        .StkNo = TxtitemName.Tag
    End With
    fillstructure = True
End If
Exit Function
ERRORHANDLER:

fillstructure = False

End Function
Sub insertintoGrid(Id As Double)
Dim Vrow As Integer
Dim Rs As New ADODB.Recordset
    Sqltext = "select id, Modno, symbol, name, Stkno, StkName From CoMaintModelItemsQry Where Id=" & Id
    Set Rs = de.con.Execute(Sqltext)
    With FlexGrid
        .AddItem ""
        Vrow = .Rows - 1
        .TextMatrix(Vrow, Colid) = Rs!Id
        .TextMatrix(Vrow, ColModNo) = Rs!ModNo
        .TextMatrix(Vrow, Colsymbol) = Rs!Symbol & ""
        .TextMatrix(Vrow, ColName) = Rs!Name & ""
        .TextMatrix(Vrow, ColStkNo) = Rs!StkNo & ""
        .TextMatrix(Vrow, ColStkName) = Rs!StkName & ""
    End With
End Sub
Function GetMaxId() As Double
On Error GoTo ERRORHANDLER
Dim RsMax As New Recordset
Sqltext = "Select Max(Id) as MaxId From CoMaintModelItems"
Set RsMax = de.con.Execute(Sqltext)
GetMaxId = RsMax!MaxId
Exit Function
ERRORHANDLER:
GetMaxId = 0
End Function
Sub SaveRec()
On Error GoTo ERRORHANDLER
If fillstructure Then
    With ModelListItemRec
        Sqltext = "Insert Into CoMaintModelItems(ModNo ,  StkNo)Values(" & .ModNo & ",'" & .StkNo & "')"
        de.con.Execute (Sqltext)
        SearchData .ModNo, ""
    End With
End If
Exit Sub
ERRORHANDLER:
MsgBox Err.Description
    
End Sub
Private Sub TxtitemName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Grid.Visible Then
        Ok = False
        TxtitemName.Tag = Grid.TextMatrix(IIf(Grid.Row = 0, 1, Grid.Row), ColNo)
        TxtitemName.Text = Grid.TextMatrix(IIf(Grid.Row = 0, 1, Grid.Row), ColNo)
        LItemName.Caption = Grid.TextMatrix(IIf(Grid.Row = 0, 1, Grid.Row), ColName)
        Ok = True
    ElseIf Grid.Visible = False And TxtitemName.Text <> "" And TxtitemName.Tag <> "" Then
        TxtitemName.SetFocus
        TxtitemName.SelStart = 0
        TxtitemName.SelLength = Len(TxtitemName.Text)
        Exit Sub
    Else
        Ok = False
        TxtitemName.Tag = ""
        TxtitemName.Text = ""
        LItemName.Caption = ""
        Ok = True
    End If
    SaveRec
    TxtitemName.SetFocus
    TxtitemName.SelStart = 0
    TxtitemName.SelLength = Len(TxtitemName.Text)
    Grid.Visible = False
End If
End Sub


Private Sub TxtModelName_Change()
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
        ChangeCursor 1
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


Private Sub TxtModelName_GotFocus()
ChangeToArabic
Pos = 1
End Sub

Private Sub TxtModelName_KeyDown(KeyCode As Integer, Shift As Integer)
    Flag = True
    If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then MoveCursor KeyCode, Grid
    Flag = True
End Sub

Private Sub TxtModelName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Grid.Visible Then
        Ok = False
        TxtModelName.Tag = Grid.TextMatrix(IIf(Grid.Row = 0, 1, Grid.Row), ColNo)
        TxtModelName.Text = Grid.TextMatrix(IIf(Grid.Row = 0, 1, Grid.Row), col3)
        LSymbol.Caption = Grid.TextMatrix(IIf(Grid.Row = 0, 1, Grid.Row), ColName)
        Ok = True
    ElseIf Grid.Visible = False And TxtModelName.Text <> "" And Val(TxtModelName.Tag) <> 0 Then
        TxtitemName.SetFocus
        SendKeys "{home}+{end}"
        Exit Sub
    Else
        Ok = False
        TxtModelName.Tag = 0
        TxtModelName.Text = ""
        LSymbol.Caption = ""
        Ok = True
    End If
    Grid.Visible = False
    TxtitemName.SetFocus
    SendKeys "{home}+{end}"
End If
End Sub

