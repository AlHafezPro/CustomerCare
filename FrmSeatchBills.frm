VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmSeatchBills 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "»ÕÀ"
   ClientHeight    =   7245
   ClientLeft      =   45
   ClientTop       =   450
   ClientWidth     =   11700
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   7245
   ScaleWidth      =   11700
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   2085
      Left            =   4290
      TabIndex        =   46
      Top             =   3690
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
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   405
      Left            =   120
      TabIndex        =   74
      Top             =   5910
      Visible         =   0   'False
      Width           =   11505
      _ExtentX        =   20294
      _ExtentY        =   714
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin Threed.SSFrame SSFrame2 
      Height          =   465
      Left            =   90
      TabIndex        =   63
      Top             =   6750
      Width           =   11565
      _ExtentX        =   20399
      _ExtentY        =   820
      _Version        =   131074
      Begin VB.Label LSumAccount 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   9090
         RightToLeft     =   -1  'True
         TabIndex        =   80
         Top             =   60
         Width           =   1665
      End
      Begin VB.Label LSumTotal 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   5190
         RightToLeft     =   -1  'True
         TabIndex        =   79
         Top             =   30
         Width           =   1665
      End
      Begin VB.Label LCountDetails 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   73
         Top             =   60
         Width           =   675
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "«·⁄œœ"
         ForeColor       =   &H000000FF&
         Height          =   405
         Index           =   11
         Left            =   690
         TabIndex        =   72
         Top             =   30
         Width           =   510
      End
      Begin VB.Label LCountTotal 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   3960
         RightToLeft     =   -1  'True
         TabIndex        =   71
         Top             =   60
         Width           =   675
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "«·⁄œœ"
         ForeColor       =   &H000000FF&
         Height          =   405
         Index           =   10
         Left            =   4680
         TabIndex        =   70
         Top             =   30
         Width           =   510
      End
      Begin VB.Label LCountAccount 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   7860
         RightToLeft     =   -1  'True
         TabIndex        =   69
         Top             =   60
         Width           =   675
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "«·⁄œœ"
         ForeColor       =   &H000000FF&
         Height          =   405
         Index           =   9
         Left            =   8550
         TabIndex        =   68
         Top             =   30
         Width           =   510
      End
      Begin VB.Label LSumDetails 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   1230
         RightToLeft     =   -1  'True
         TabIndex        =   67
         Top             =   60
         Width           =   1665
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "«·„Ã„Ê⁄ «· ›’Ì·Ì"
         ForeColor       =   &H000000FF&
         Height          =   405
         Index           =   8
         Left            =   2790
         TabIndex        =   66
         Top             =   30
         Width           =   780
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "«·„Ã„Ê⁄ «·≈Ã„«·Ì"
         ForeColor       =   &H000000FF&
         Height          =   405
         Index           =   5
         Left            =   6810
         TabIndex        =   65
         Top             =   30
         Width           =   735
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "«·„Ã„Ê⁄ «·„Õ«”»Ì"
         ForeColor       =   &H000000FF&
         Height          =   405
         Index           =   2
         Left            =   10650
         TabIndex        =   64
         Top             =   30
         Width           =   840
      End
   End
   Begin TabDlg.SSTab sstab1 
      Height          =   3465
      Left            =   60
      TabIndex        =   59
      Top             =   3270
      Width           =   11595
      _ExtentX        =   20452
      _ExtentY        =   6112
      _Version        =   393216
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   " ›«’Ì·"
      TabPicture(0)   =   "FrmSeatchBills.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "flexGridDetails"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "≈Ã„«·Ì"
      TabPicture(1)   =   "FrmSeatchBills.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "FlexSummary"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "«·„Õ«”»…"
      TabPicture(2)   =   "FrmSeatchBills.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "FlexAccount"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "«·€Ì— „—Õ·"
      TabPicture(3)   =   "FrmSeatchBills.frx":0054
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "FlexNotTransfered"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      Begin VSFlex8Ctl.VSFlexGrid flexGridDetails 
         Height          =   3045
         Left            =   -74940
         TabIndex        =   60
         Top             =   360
         Width           =   11475
         _cx             =   20241
         _cy             =   5371
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
      Begin VSFlex8Ctl.VSFlexGrid FlexSummary 
         Height          =   3015
         Left            =   -74940
         TabIndex        =   61
         Top             =   390
         Width           =   11475
         _cx             =   20241
         _cy             =   5318
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
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
      Begin VSFlex8Ctl.VSFlexGrid FlexAccount 
         Height          =   3045
         Left            =   -74940
         TabIndex        =   62
         Top             =   360
         Width           =   11475
         _cx             =   20241
         _cy             =   5371
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
      Begin VSFlex8Ctl.VSFlexGrid FlexNotTransfered 
         Height          =   3045
         Left            =   60
         TabIndex        =   78
         Top             =   360
         Width           =   11475
         _cx             =   20241
         _cy             =   5371
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
   Begin Crystal.CrystalReport cr1 
      Left            =   5280
      Top             =   2130
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.ComboBox ComboFixDate 
      Height          =   315
      Index           =   0
      ItemData        =   "FrmSeatchBills.frx":0070
      Left            =   5790
      List            =   "FrmSeatchBills.frx":0086
      RightToLeft     =   -1  'True
      TabIndex        =   54
      TabStop         =   0   'False
      Text            =   "="
      Top             =   2520
      Width           =   600
   End
   Begin VB.ComboBox ComboFixDate 
      Height          =   315
      Index           =   1
      ItemData        =   "FrmSeatchBills.frx":009F
      Left            =   5790
      List            =   "FrmSeatchBills.frx":00B5
      RightToLeft     =   -1  'True
      TabIndex        =   53
      TabStop         =   0   'False
      Text            =   "="
      Top             =   2910
      Width           =   600
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   975
      Left            =   30
      TabIndex        =   49
      Top             =   720
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1720
      _Version        =   131074
      Begin Threed.SSOption SSOption1 
         Height          =   345
         Index           =   2
         Left            =   30
         TabIndex        =   52
         Top             =   600
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
         _Version        =   131074
         ForeColor       =   192
         Caption         =   "ﬂ·ÌÂ„«"
         Value           =   -1
      End
      Begin Threed.SSOption SSOption1 
         Height          =   345
         Index           =   1
         Left            =   30
         TabIndex        =   51
         Top             =   30
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
         _Version        =   131074
         Caption         =   "„À» …"
      End
      Begin Threed.SSOption SSOption1 
         Height          =   345
         Index           =   0
         Left            =   30
         TabIndex        =   50
         Top             =   330
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
         _Version        =   131074
         Caption         =   "„ƒﬁ …"
      End
   End
   Begin VB.ComboBox ComboFeesOperation 
      Height          =   315
      Index           =   1
      ItemData        =   "FrmSeatchBills.frx":00CE
      Left            =   1560
      List            =   "FrmSeatchBills.frx":00E4
      RightToLeft     =   -1  'True
      TabIndex        =   20
      TabStop         =   0   'False
      Text            =   "="
      Top             =   2910
      Width           =   600
   End
   Begin VB.ComboBox ComboFeesOperation 
      Height          =   315
      Index           =   0
      ItemData        =   "FrmSeatchBills.frx":00FD
      Left            =   1560
      List            =   "FrmSeatchBills.frx":0113
      RightToLeft     =   -1  'True
      TabIndex        =   19
      TabStop         =   0   'False
      Text            =   "="
      Top             =   2520
      Width           =   600
   End
   Begin VB.TextBox TxtitemName 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   1860
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   1710
      Width           =   2505
   End
   Begin VB.TextBox TxtFeesAmount 
      Alignment       =   1  'Right Justify
      Height          =   345
      Index           =   1
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   2880
      Width           =   1065
   End
   Begin VB.TextBox TxtFeesAmount 
      Alignment       =   1  'Right Justify
      Height          =   345
      Index           =   0
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   2490
      Width           =   1065
   End
   Begin VB.TextBox TxtAmount 
      Alignment       =   1  'Right Justify
      Height          =   345
      Index           =   1
      Left            =   2190
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   2910
      Width           =   1065
   End
   Begin VB.TextBox TxtAmount 
      Alignment       =   1  'Right Justify
      Height          =   345
      Index           =   0
      Left            =   2190
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   2520
      Width           =   1065
   End
   Begin VB.TextBox TxtBillNo 
      Alignment       =   1  'Right Justify
      Height          =   345
      Index           =   1
      Left            =   8730
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   2910
      Width           =   885
   End
   Begin VB.TextBox TxtBillNo 
      Alignment       =   1  'Right Justify
      Height          =   345
      Index           =   0
      Left            =   8730
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   2520
      Width           =   885
   End
   Begin VB.ComboBox ComboAmount 
      Height          =   315
      Index           =   1
      ItemData        =   "FrmSeatchBills.frx":012C
      Left            =   3600
      List            =   "FrmSeatchBills.frx":0142
      RightToLeft     =   -1  'True
      TabIndex        =   16
      TabStop         =   0   'False
      Text            =   "="
      Top             =   2910
      Width           =   600
   End
   Begin VB.ComboBox ComboAmount 
      Height          =   315
      Index           =   0
      ItemData        =   "FrmSeatchBills.frx":015B
      Left            =   3600
      List            =   "FrmSeatchBills.frx":0171
      RightToLeft     =   -1  'True
      TabIndex        =   15
      TabStop         =   0   'False
      Text            =   "="
      Top             =   2520
      Width           =   600
   End
   Begin VB.ComboBox ComboDate 
      Height          =   315
      Index           =   1
      ItemData        =   "FrmSeatchBills.frx":018A
      Left            =   8160
      List            =   "FrmSeatchBills.frx":01A0
      RightToLeft     =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Text            =   "="
      Top             =   2910
      Width           =   600
   End
   Begin VB.ComboBox ComboDate 
      Height          =   315
      Index           =   0
      ItemData        =   "FrmSeatchBills.frx":01B9
      Left            =   8160
      List            =   "FrmSeatchBills.frx":01CF
      RightToLeft     =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Text            =   "="
      Top             =   2520
      Width           =   600
   End
   Begin VB.ComboBox ComboBillNo 
      Height          =   315
      Index           =   1
      ItemData        =   "FrmSeatchBills.frx":01E8
      Left            =   10500
      List            =   "FrmSeatchBills.frx":01FE
      RightToLeft     =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Text            =   "="
      Top             =   2910
      Width           =   600
   End
   Begin VB.ComboBox ComboBillNo 
      Height          =   315
      Index           =   0
      ItemData        =   "FrmSeatchBills.frx":0217
      Left            =   10500
      List            =   "FrmSeatchBills.frx":022D
      RightToLeft     =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Text            =   "="
      Top             =   2520
      Width           =   600
   End
   Begin VB.TextBox TxtModelName 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   6960
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   1710
      Width           =   4695
   End
   Begin VB.TextBox TxtClientName 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   4410
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   1020
      Width           =   2505
   End
   Begin MSDataListLib.DataCombo ComboType 
      Height          =   360
      Left            =   8610
      TabIndex        =   1
      Top             =   1020
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   635
      _Version        =   393216
      Style           =   2
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
   Begin MSDataListLib.DataCombo ComboOperationType 
      Height          =   360
      Left            =   10110
      TabIndex        =   0
      Top             =   1020
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   635
      _Version        =   393216
      Style           =   2
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
   Begin MSDataListLib.DataCombo ComboPayment 
      Height          =   360
      Left            =   6930
      TabIndex        =   2
      Top             =   1020
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   635
      _Version        =   393216
      Style           =   2
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
   Begin MSMask.MaskEdBox TxtDate 
      Height          =   345
      Index           =   0
      Left            =   6600
      TabIndex        =   13
      Top             =   2520
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   609
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "99/99/9999"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox TxtDate 
      Height          =   345
      Index           =   1
      Left            =   6600
      TabIndex        =   14
      Top             =   2910
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   609
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "99/99/9999"
      PromptChar      =   "_"
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3180
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
            Picture         =   "FrmSeatchBills.frx":0246
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSeatchBills.frx":291C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSeatchBills.frx":5779
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSeatchBills.frx":7F1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSeatchBills.frx":A81A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSeatchBills.frx":CD3F
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSeatchBills.frx":F4F7
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSeatchBills.frx":11F0B
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSeatchBills.frx":1485D
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSeatchBills.frx":175D1
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSeatchBills.frx":19E20
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSeatchBills.frx":1CCC7
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSeatchBills.frx":1FA21
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSeatchBills.frx":223C1
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSeatchBills.frx":25323
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSeatchBills.frx":27D4B
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSeatchBills.frx":2A705
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSeatchBills.frx":2D036
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSeatchBills.frx":2F93C
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSeatchBills.frx":323A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSeatchBills.frx":35355
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSeatchBills.frx":37C7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSeatchBills.frx":3A5B3
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSeatchBills.frx":3CCF3
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSeatchBills.frx":3F662
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSeatchBills.frx":41EEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSeatchBills.frx":440F9
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSeatchBills.frx":46A56
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSeatchBills.frx":49280
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSeatchBills.frx":4BD34
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSeatchBills.frx":4E770
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSeatchBills.frx":51286
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSeatchBills.frx":54220
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSeatchBills.frx":56FEC
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSeatchBills.frx":59C66
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSeatchBills.frx":5CAF8
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSeatchBills.frx":5F736
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSeatchBills.frx":622E5
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   690
      Left            =   0
      TabIndex        =   41
      Top             =   0
      Width           =   11700
      _ExtentX        =   20638
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
            ImageIndex      =   15
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   37
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Tag             =   "1"
                  Text            =   " ﬁ—Ì— ≈Ã„«·Ì"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Tag             =   "2"
                  Text            =   " ﬁ—Ì—  ›’Ì·Ì"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   2
         EndProperty
      EndProperty
      Begin Threed.SSFrame SSFrame3 
         Height          =   555
         Left            =   6300
         TabIndex        =   75
         Top             =   30
         Width           =   5325
         _ExtentX        =   9393
         _ExtentY        =   979
         _Version        =   131074
         Begin MSDataListLib.DataCombo ComboDestination 
            Height          =   360
            Left            =   60
            TabIndex        =   76
            Top             =   90
            Width           =   4215
            _ExtentX        =   7435
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
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "«·ÃÂ‹‹‹‹…"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   4
            Left            =   4680
            RightToLeft     =   -1  'True
            TabIndex        =   77
            Top             =   90
            Width           =   540
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   525
         Left            =   9300
         RightToLeft     =   -1  'True
         TabIndex        =   42
         Top             =   90
         Width           =   2325
      End
   End
   Begin MSDataListLib.DataCombo ComboFees 
      Height          =   360
      Left            =   6930
      TabIndex        =   6
      Top             =   2100
      Width           =   4185
      _ExtentX        =   7382
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
   Begin MSMask.MaskEdBox TxtFixBillDate 
      Height          =   345
      Index           =   0
      Left            =   4230
      TabIndex        =   55
      Top             =   2520
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   609
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "99/99/9999"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox TxtFixBillDate 
      Height          =   345
      Index           =   1
      Left            =   4230
      TabIndex        =   56
      Top             =   2910
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   609
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "99/99/9999"
      PromptChar      =   "_"
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   " «—ÌŒ"
      Height          =   195
      Index           =   9
      Left            =   5400
      RightToLeft     =   -1  'True
      TabIndex        =   58
      Top             =   2580
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   " «—ÌŒ"
      Height          =   195
      Index           =   8
      Left            =   5400
      RightToLeft     =   -1  'True
      TabIndex        =   57
      Top             =   2940
      Width           =   375
   End
   Begin VB.Label lPhoneNbr 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   2490
      TabIndex        =   48
      Top             =   1050
      Width           =   1905
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "—ﬁ„ «·Â« ›"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   2
      Left            =   3600
      RightToLeft     =   -1  'True
      TabIndex        =   47
      Top             =   750
      Width           =   780
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "≈·Ï"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   1
      Left            =   11370
      TabIndex        =   45
      Top             =   2970
      Width           =   255
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "„‰"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   0
      Left            =   11400
      TabIndex        =   44
      Top             =   2580
      Width           =   180
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "«·√ÃÊ—"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   13
      Left            =   11220
      TabIndex        =   43
      Top             =   2130
      Width           =   405
   End
   Begin VB.Label LItemName 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   90
      TabIndex        =   40
      Top             =   1740
      Width           =   1725
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
      Left            =   4410
      RightToLeft     =   -1  'True
      TabIndex        =   39
      Top             =   1740
      Width           =   2505
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "«·„ÊœÌ·"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   7
      Left            =   11100
      TabIndex        =   38
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "‰Ê⁄ «·“»Ê‰"
      Height          =   195
      Index           =   3
      Left            =   1410
      RightToLeft     =   -1  'True
      TabIndex        =   37
      Top             =   750
      Width           =   1050
   End
   Begin VB.Label LClientType 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1410
      TabIndex        =   36
      Top             =   1050
      Width           =   1035
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "«·—ﬁ„ «·„Œ“‰Ì"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   1
      Left            =   3420
      RightToLeft     =   -1  'True
      TabIndex        =   35
      Top             =   1410
      Width           =   930
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "«·√ÃÊ—"
      Height          =   195
      Index           =   7
      Left            =   1170
      RightToLeft     =   -1  'True
      TabIndex        =   34
      Top             =   2970
      Width           =   405
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "«·√ÃÊ—"
      Height          =   195
      Index           =   6
      Left            =   1170
      RightToLeft     =   -1  'True
      TabIndex        =   33
      Top             =   2550
      Width           =   405
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "«·“»Ê‰"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   0
      Left            =   6120
      RightToLeft     =   -1  'True
      TabIndex        =   32
      Top             =   750
      Width           =   780
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "‰Ê⁄ «·’Ì«‰…"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   4
      Left            =   9300
      TabIndex        =   31
      Top             =   750
      Width           =   780
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "‰Ê⁄ «·⁄„·Ì…"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   6
      Left            =   10860
      TabIndex        =   30
      Top             =   750
      Width           =   735
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "ÿ—Ìﬁ… «·œ›⁄"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   3
      Left            =   7710
      TabIndex        =   29
      Top             =   750
      Width           =   795
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "„»·€"
      Height          =   195
      Index           =   5
      Left            =   3270
      RightToLeft     =   -1  'True
      TabIndex        =   28
      Top             =   2940
      Width           =   285
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "„»·€"
      Height          =   195
      Index           =   4
      Left            =   3270
      RightToLeft     =   -1  'True
      TabIndex        =   27
      Top             =   2580
      Width           =   285
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   " «—ÌŒ"
      Height          =   195
      Index           =   3
      Left            =   7770
      RightToLeft     =   -1  'True
      TabIndex        =   26
      Top             =   2940
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   " «—ÌŒ"
      Height          =   195
      Index           =   2
      Left            =   7770
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Top             =   2580
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "—ﬁ„ «·›« Ê—…"
      Height          =   195
      Index           =   1
      Left            =   9660
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Top             =   2940
      Width           =   825
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "—ﬁ„ «·›« Ê—…"
      Height          =   195
      Index           =   0
      Left            =   9660
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Top             =   2550
      Width           =   825
   End
   Begin VB.Menu mnufile 
      Caption         =   "file"
      Visible         =   0   'False
      Begin VB.Menu mnu0 
         Caption         =   " —ÕÌ· ≈·Ï «·„Õ«”»…"
         HelpContextID   =   6
         Index           =   0
      End
   End
   Begin VB.Menu mnufile1 
      Caption         =   "file1"
      Visible         =   0   'False
      Begin VB.Menu mnu1 
         Caption         =   "≈·€«¡ «· À»Ì» "
         HelpContextID   =   6
         Index           =   0
      End
   End
   Begin VB.Menu mnufile2 
      Caption         =   "file2"
      Visible         =   0   'False
      Begin VB.Menu mnu2 
         Caption         =   " —ÕÌ· ≈·Ï «·„” Êœ⁄« "
         HelpContextID   =   6
         Index           =   0
      End
   End
   Begin VB.Menu mnufile3 
      Caption         =   "file3"
      Visible         =   0   'False
      Begin VB.Menu mnu3 
         Caption         =   " ’œÌ— ≈‘⁄«—«  «·„Ê«œ «·√Ê·Ì…"
         HelpContextID   =   6
         Index           =   0
      End
   End
   Begin VB.Menu mnufile4 
      Caption         =   "file4"
      Visible         =   0   'False
      Begin VB.Menu mnu4 
         Caption         =   "ÿ»«⁄… «·›« Ê— «·√’·Ì…"
         HelpContextID   =   6
         Index           =   0
      End
   End
   Begin VB.Menu mnufile5 
      Caption         =   "file5"
      Visible         =   0   'False
      Begin VB.Menu mnu5 
         Caption         =   " —ÕÌ· ≈·Ï „·› ‰’Ì"
         Index           =   0
      End
   End
End
Attribute VB_Name = "FrmSeatchBills"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Pos As Integer, RecNum   As Integer
Dim Ok As Boolean, Flag As Boolean

Const ColNo = 1
Const ColName = 2
Const col3 = 3
Const Col4 = 4

Const ColId = 1
Const ColCallNo = 2
Const ColBillNo = 3
Const ColSerByYear = 4
Const Colbilldate = 5
Const ColFixbilldate = 6
Const ColOperationTypeId = 7
Const ColOperationTypeName = 8
Const ColStat = 9
Const Colstatdescription = 10
Const ColPaymentTypeId = 11
Const ColPaymentTypeName = 12
Const ColClientName = 13
Const ColClientPhoneNBr = 14
Const ColModelName = 15
Const colSymbol = 16
Const ColModelQty = 17
Const ColStkNo = 18
Const ColStkAccNo = 19
Const ColStkName = 20
Const ColQty = 21
Const ColPriceTypeName = 22
Const ColPrice = 23
Const ColTotPrice = 24

Const Colisfixed = 25
Const ColIsTransfered = 26
Const ColisTransferedToStock = 27

Const ColIsError = 28


Const ColTotalBillNo = 1
Const ColTotalBillTotal = 2

Const ColAccountNum = 1
Const ColAccountDate = 2
Const ColAccountstkno = 3
Const ColAccountAccNo = 4
Const ColAccountStkname = 5
Const ColAccountDeb = 6
Const ColAccountDebAccName = 7
Const ColAccountCre = 8
Const ColAccountCreAccName = 9
Const ColAccountDescription = 10
Const ColAccountOperationType = 11
Const ColAccountOperationTypeName = 12
Const ColAccountpaymenttypeid = 13
Const ColAccountpaymentTypeName = 14
Const ColAccountStat = 15
Const ColAccountstatdescription = 16
Const ColAccountTotPrice = 17
Const ColAccountCountRec = 18
Const ColAccountType = 19




Const ColNotTransferedBillno = 1
Const ColNotTransferedTotPrice = 2
Const ColNotTransferedFixbilldate = 3
Const ColNotTransferedOperationType = 4
Const ColNotTransferedOperationTypeName = 5
Const ColNotTransferedStat = 6
Const ColNotTransferedstatdescription = 7
Const ColNotTransferedpaymenttypeid = 8
Const ColNotTransferedpaymentTypeName = 9


Dim SearchRec As SearchRecType
Dim Vindex  As Integer
'Sub UpdateDebCre(maxi As Double)
'    Sqltext = GetDatabaseName(TxtTransferDate.Text) & ".dbo.UpdateForAmountfDEb_Cre(" & maxi & ")"
'    Dacc.con.Execute (Sqltext)
'    Sqltext = "Select Amount , AmountfDeb  , AmountfCre From " & GetDatabaseName(TxtTransferDate.Text) & ".dbo.AccRegTemp Where AccRegNoTemp=" & maxi
'    Set Rs = Dacc.con.Execute(Sqltext)
'    AmountfDeb = Rs!Amount / Rs!AmountfDeb
'    AmountfCre = Rs!Amount / Rs!AmountfCre
'End Sub
'Function Gettag(empNo As Integer, TagId As Integer) As Boolean
'On Error GoTo errorhandler
'Dim rs As New ADODB.Recordset
'    sqlText = "Select * from comaintpermission Where empno = " & empNo & " and TagId=" & TagId
'    Set rs = de.con.Execute(sqlText)
'    If rs.RecordCount > 0 Then
'        Gettag = True
'    Else
'        Gettag = False
'    End If
'Exit Function
'errorhandler:
'Gettag = False
'End Function


Sub TransMaintAccRegTemp()
On Error GoTo ErrorHandler
Dim FirstRow As Integer, LastRow As Integer, Vrow As Integer

ProgressBar1.Visible = True
Screen.MousePointer = vbHourglass
de.con.BeginTrans
With FlexAccount
    ProgressBar1.Min = 1
    ProgressBar1.Max = .Rows
    For i = 1 To .Rows - 1
        Vrow = i
        sqlText = "Exec Sp_TransferRecToAccRegTemp '" & .TextMatrix(Vrow, ColAccountDeb) & "','" & .TextMatrix(Vrow, ColAccountCre) & "'," & .TextMatrix(Vrow, ColAccountTotPrice) & "," & .TextMatrix(Vrow, ColAccountCountRec) & "," & .TextMatrix(Vrow, ColAccountType) & ",'" & ConvertControlDate(.TextMatrix(Vrow, ColAccountDate)) & "','" & .TextMatrix(Vrow, ColAccountDescription) & "'," & empNo
        de.con.Execute (sqlText)
        ProgressBar1.Value = i
    Next
End With

SearchData 1
de.con.CommitTrans

ProgressBar1.Visible = False
Screen.MousePointer = vbDefault
Exit Sub
ErrorHandler:
de.con.RollbackTrans
Screen.MousePointer = vbDefault
MsgBox Err.Description
End Sub

Function AllowPrintBills(ByVal Level As Integer, ByVal empNo As Integer) As Boolean
On Error GoTo ErrorHandler
Dim rs As New ADODB.Recordset
    sqlText = "select EmpNo , AllowPrintBills from MaintUsers Where EmpNo=" & empNo
    Set rs = de.con.Execute(sqlText)
    If rs!AllowPrintBills And (2 ^ Level) Then
        AllowPrintBills = True
    Else
        AllowPrintBills = False
    End If
Exit Function
ErrorHandler:
MsgBox Err.Description

End Function
Function AllowExportItems(ByVal Level As Integer, ByVal empNo As Integer) As Boolean
On Error GoTo ErrorHandler
Dim rs As New ADODB.Recordset
    sqlText = "select EmpNo , AllowExportItems from MaintUsers Where EmpNo=" & empNo
    Set rs = de.con.Execute(sqlText)
    If rs!AllowExportItems And (2 ^ Level) Then
        AllowExportItems = True
    Else
        AllowExportItems = False
    End If
Exit Function
ErrorHandler:
MsgBox Err.Description
End Function


Function AllowTransferToStock(ByVal Level As Integer, ByVal empNo As Integer) As Boolean
On Error GoTo ErrorHandler
Dim rs As New ADODB.Recordset
    sqlText = "select EmpNo , AllowTransferStock from MaintUsers Where EmpNo=" & empNo
    Set rs = de.con.Execute(sqlText)
    If rs!AllowTransferStock And (2 ^ Level) Then
        AllowTransferToStock = True
    Else
        AllowTransferToStock = False
    End If
Exit Function
ErrorHandler:
MsgBox Err.Description
End Function

Function AllowTransferToAccount(ByVal Level As Integer, ByVal empNo As Integer) As Boolean
Dim rs As New ADODB.Recordset
    sqlText = "select EmpNo , AllowTransfer from MaintUsers Where EmpNo=" & empNo
    Set rs = de.con.Execute(sqlText)
    If rs!AllowTransfer And (2 ^ Level) Then
        AllowTransferToAccount = True
    Else
        AllowTransferToAccount = False
    End If
End Function
Function AllowCancelAccount(ByVal Level As Integer, ByVal empNo As Integer) As Boolean
Dim rs As New ADODB.Recordset
    sqlText = "select EmpNo , CancelAccount from MaintUsers Where EmpNo=" & empNo
    Set rs = de.con.Execute(sqlText)
    If rs!CancelAccount And (2 ^ Level) Then
        AllowCancelAccount = True
    Else
        AllowCancelAccount = False
    End If
End Function



Function GetPhoneNbr(clientId As Double, Class As Integer) As String
On Error GoTo ErrorHandler
Dim RsPhone As New ADODB.Recordset
sqlText = "Select isnull(ClientPhonenbr,'') ClientPhonenbr From ClientQry Where ClientId=" & clientId & " and Class =" & Class
Set RsPhone = de.con.Execute(sqlText)
If RsPhone.RecordCount > 0 Then
    GetPhoneNbr = RsPhone!ClientPhoneNBr
Else
    GetPhoneNbr = ""
End If
Exit Function
ErrorHandler:
GetPhoneNbr = ""
End Function

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

Sub FillFormating(ByVal i As Integer)
If i = 1 Then
    fs = "|>" + ""
    fs = fs + "|>" + ""
    fs = fs + "|>" + ""
    fs = fs + "|>" + ""
    With Grid
        .FormatString = fs
        .Cols = 5
        If Pos = 3 Then
            SetColWidths ColNo, Grid
        Else
            .ColWidth(ColNo) = 0
        End If
        If Pos = 1 Or Pos = 2 Then
            SetColWidths col3, Grid
        Else
            .ColWidth(col3) = 0
        End If
        SetColWidths ColName, Grid
        .ColWidth(Col4) = 0
    End With
ElseIf i = 2 Then
    fs = "|>" + "Id"
    fs = fs + "|>" + "—ﬁ„ «·‘ﬂÊÏ"
    fs = fs + "|>" + "—ﬁ„ «·›« Ê—…"
    fs = fs + "|>" + "—ﬁ„ «·›« Ê—…"
'    Fs = Fs + "|>" + "≈Ã„«·Ì «·›« Ê—…"
    fs = fs + "|>" + " «—ÌŒ «·›« Ê—…"
    fs = fs + "|>" + " «—ÌŒ «·›« Ê—… «·„À» …"
    fs = fs + "|>" + "OperationTYpeId"
    fs = fs + "|>" + "«·⁄„·Ì…"
    fs = fs + "|>" + "—ﬁ„ ‰Ê⁄ «·’Ì«‰…"
    fs = fs + "|>" + "‰Ê⁄ «·’Ì«‰…"
    
    
    fs = fs + "|>" + "—ﬁ„ ÿ—Ìﬁ… «·œ›⁄"
    fs = fs + "|>" + "ÿ—Ìﬁ… «·œ›⁄"
    
    
     
    fs = fs + "|>" + "≈”„ «·“»Ê‰"
    fs = fs + "|>" + "—ﬁ„ «·Â« ›"
    fs = fs + "|>" + "«·„ÊœÌ·"
    fs = fs + "|>" + "—„“ «·„ÊœÌ·"
    
    
    fs = fs + "|>" + "«·⁄œœ «·„ÿ·Ê»"
    fs = fs + "|>" + "«·—ﬁ„ «·„Œ“‰Ì"
    fs = fs + "|>" + "—ﬁ„ «·Õ”«»"
    fs = fs + "|>" + "«·‘—Õ"
    
    fs = fs + "|>" + "«·ﬂ„Ì…"
    fs = fs + "|>" + "‰Ê⁄ «·”⁄—"
    fs = fs + "|>" + "«·”⁄—"
    
    
    fs = fs + "|>" + "«·≈Ã„«·Ì"
    
    
    
    fs = fs + "|>" + "„À» "
    fs = fs + "|>" + "„—Õ·"
    fs = fs + "|>" + "„—Õ· ≈·Ï «·„” Êœ⁄« "
    fs = fs + "|>" + "IsError"
    
    With flexGridDetails
        .FormatString = fs
        .Cols = 29
        SetColWidths ColId, flexGridDetails
        SetColWidths ColCallNo, flexGridDetails
        .ColWidth(ColSerByYear) = 0
        SetColWidths ColBillNo, flexGridDetails
        'SetColWidths ColSerByYear, flexGridDetails
        SetColWidths ColAllTotalPrice, flexGridDetails
        SetColWidths Colbilldate, flexGridDetails
        SetColWidths ColFixbilldate, flexGridDetails
        
        
        .ColWidth(ColOperationTypeId) = 0
        SetColWidths ColOperationTypeName, flexGridDetails
        
        
        .ColWidth(ColStat) = 0
        SetColWidths Colstatdescription, flexGridDetails
        
        
        .ColWidth(ColPaymentTypeId) = 0
        SetColWidths ColPaymentTypeName, flexGridDetails


        SetColWidths ColClientName, flexGridDetails
        SetColWidths ColClientPhoneNBr, flexGridDetails
        SetColWidths ColModelName, flexGridDetails
        SetColWidths colSymbol, flexGridDetails
        SetColWidths ColModelQty, flexGridDetails
        SetColWidths ColStkNo, flexGridDetails
        SetColWidths ColStkAccNo, flexGridDetails
        SetColWidths ColStkName, flexGridDetails
        SetColWidths ColQty, flexGridDetails
        SetColWidths ColPriceTypeName, flexGridDetails
        SetColWidths ColPrice, flexGridDetails
        SetColWidths ColTotPrice, flexGridDetails
        
        
        SetColWidths Colisfixed, flexGridDetails
        SetColWidths ColIsTransfered, flexGridDetails
        .ColDataType(Colisfixed) = flexDTBoolean
        .ColDataType(ColIsTransfered) = flexDTBoolean
        .ColDataType(ColisTransferedToStock) = flexDTBoolean
        .ColWidth(ColIsError) = 0
        '.ColWidth(ColisTransferedToStock) = 300
        
End With
ElseIf i = 3 Then
    fs = "|>" + "—ﬁ„ «·›« Ê—…"
    fs = fs + "|>" + "≈Ã„«·Ì «·›« Ê—…"
    With FlexSummary
        .FormatString = fs
        .Cols = 3
        SetColWidths ColTotalBillNo, FlexSummary
        SetColWidths ColTotalBillTotal, FlexSummary
    End With
ElseIf i = 4 Then
    fs = "|>" + " ”·”·"
    fs = fs + "|>" + " ‹‹«—ÌŒ «·Õ—ﬂ…"
    fs = fs + "|>" + "«·—ﬁ„ «·„Œ“‰Ì"
    fs = fs + "|>" + "—ﬁ„ Õ”«» «·„Œ“‰Ì"
    fs = fs + "|>" + "‘—Õ «·Õ”«»"
    fs = fs + "|>" + "«·„œÌ‹‹‹‰"
    fs = fs + "|>" + "‘—Õ «·Õ”«» «·„œÌ‰"
    fs = fs + "|>" + "«·œ«∆‹‹‰"
    fs = fs + "|>" + "‘—Õ «·Õ”«» «·œ«∆‰"
    fs = fs + "|>" + "«·‘‹‹‹‹‹‹—Õ"
    fs = fs + "|>" + "‰Ê⁄ «·⁄„·Ì…"
    fs = fs + "|>" + "‘—Õ ‰Ê⁄ «·⁄„·Ì…"
    fs = fs + "|>" + "—ﬁ„ ÿ—Ìﬁ… «·œ›⁄"
    fs = fs + "|>" + "ÿ—Ìﬁ… «·œ›⁄"
    fs = fs + "|>" + "—ﬁ„ ‰Ê⁄ «·’Ì«‰…"
    fs = fs + "|>" + "‰Ê⁄ «·’Ì«‰…"
    fs = fs + "|>" + "«·„»·€ «·≈Ã„«·Ì"
    fs = fs + "|>" + "«·ﬂ„Ì…"
    fs = fs + "|>" + "‰Ê⁄ «·⁄„·Ì…"
    
    With FlexAccount
        .FormatString = fs
        .Cols = 20
        SetColWidths ColAccountNum, FlexAccount
        SetColWidths ColAccountDate, FlexAccount
        SetColWidths ColAccountstkno, FlexAccount
        SetColWidths ColAccountAccNo, FlexAccount
        SetColWidths ColAccountStkname, FlexAccount
        SetColWidths ColAccountDeb, FlexAccount
        SetColWidths ColAccountDebAccName, FlexAccount
        
        SetColWidths ColAccountCre, FlexAccount
        SetColWidths ColAccountCreAccName, FlexAccount
        SetColWidths ColAccountDescription, FlexAccount
        
        .ColWidth(ColAccountOperationType) = 0
        SetColWidths ColAccountOperationTypeName, FlexAccount
        .ColWidth(ColAccountpaymenttypeid) = 0
        SetColWidths ColAccountpaymentTypeName, FlexAccount
        .ColWidth(ColAccountStat) = 0
        SetColWidths ColAccountstatdescription, FlexAccount
        SetColWidths ColAccountTotPrice, FlexAccount
        SetColWidths ColAccountCountRec, FlexAccount
        SetColWidths ColAccountType, FlexAccount
End With
ElseIf i = 5 Then
    fs = "|>" + "—ﬁ„ «·›« Ê—…"
    fs = fs + "|>" + "≈Ã„«·Ì «·›« Ê—…"
    fs = fs + "|>" + " «—ÌŒ «·›« Ê—…"
    fs = fs + "|>" + "‰Ê⁄ «·⁄„·Ì…"
    fs = fs + "|>" + "‘—Õ ‰Ê⁄ «·⁄„·Ì…"
    fs = fs + "|>" + "—ﬁ„ ÿ—Ìﬁ… «·œ›⁄"
    fs = fs + "|>" + "ÿ—Ìﬁ… «·œ›⁄"
    fs = fs + "|>" + "—ﬁ„ ‰Ê⁄ «·’Ì«‰…"
    fs = fs + "|>" + "‰Ê⁄ «·’Ì«‰…"
    
    With FlexNotTransfered
        .FormatString = fs
        .Cols = 10
        SetColWidths ColNotTransferedBillno, FlexNotTransfered
        SetColWidths ColNotTransferedTotPrice, FlexNotTransfered
        SetColWidths ColNotTransferedFixbilldate, FlexNotTransfered
        .ColWidth(ColNotTransferedOperationType) = 0
        SetColWidths ColNotTransferedOperationTypeName, FlexNotTransfered
        .ColWidth(ColNotTransferedStat) = 0
        SetColWidths ColNotTransferedstatdescription, FlexNotTransfered
        .ColWidth(ColNotTransferedpaymenttypeid) = 0
        SetColWidths ColNotTransferedpaymentTypeName, FlexNotTransfered
    End With
End If
End Sub

Sub SetColWidths(ByVal ColNo As Integer, flexGrid As VSFlexGrid)
    With flexGrid
        .AutoSize (ColNo)
    End With
End Sub
Sub ChangeCursor(ByVal X As Integer)
If X = 1 Then
    With TxtClientName
       Grid.top = .top + .Height
       Grid.left = .left
       Grid.Width = .Width
    End With
ElseIf X = 2 Then
    With TxtModelName
       Grid.top = .top + .Height
       Grid.left = .left
       Grid.Width = .Width
End With
ElseIf X = 3 Then
    With TxtitemName
       Grid.top = .top + .Height
       Grid.left = .left
       Grid.Width = .Width
End With
ElseIf X = 4 Then
    With TxtRecipient
       Grid.top = .top + .Height
       Grid.left = .left
       Grid.Width = .Width
End With
ElseIf X = 5 Then
    With TxtFamNo
       Grid.top = .top + .Height
       Grid.left = .left
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
       Grid.top = .top + .Height
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
       Grid.top = .top + .Height
       Grid.left = .left
       Grid.Width = .Width
End With
End If
End Sub

Sub FillCombos()
    
    Dim RsOperationType As New ADODB.Recordset
    If OperationEmpStr = "" Or PaymentEmpStr = "" Or MaintTYpeEmpStr = "" Then Exit Sub
    sqlText = "select OpNo , OpName  from operkind Where OpNo in (" & OperationEmpStr & ")"
    Set RsOperationType = de.con.Execute(sqlText)
    Set ComboOperationType.RowSource = RsOperationType
    ComboOperationType.listField = "OpName"
    ComboOperationType.BoundColumn = "OpNo"
    ComboOperationType.BoundText = 1
    
    
    Dim rsPayment As New ADODB.Recordset
    sqlText = "Select No , Name  From PayMethod Where No in (" & PaymentEmpStr & ")"
    Set rsPayment = de.con.Execute(sqlText)
    Set ComboPayment.RowSource = rsPayment
    ComboPayment.listField = "Name"
    ComboPayment.BoundColumn = "No"
    ComboPayment.BoundText = 0
    
    Dim rsType As New ADODB.Recordset
    sqlText = "select no , stat from dbo.maintypestat where No in (" & MaintTYpeEmpStr & ")"
    Set rsType = de.con.Execute(sqlText)
    Set ComboType.RowSource = rsType
    ComboType.listField = "Stat"
    ComboType.BoundColumn = "No"
    ComboType.BoundText = 1
    
    Dim rsFees As New ADODB.Recordset
    sqlText = "Select FeesId , FeesName  From  CoMaintFees Where isnull(CliPriceafterdiscount,0) <> 0 or isnull(DealPriceafterdiscount,0) <> 0 or isnull(DistPriceafterdiscount,0) <> 0  "
    Set rsFees = de.con.Execute(sqlText)
    Set ComboFees.RowSource = rsFees
    ComboFees.listField = "FeesName"
    ComboFees.BoundColumn = "FeesId"

    Dim RsDestination As New ADODB.Recordset
    sqlText = "Select Id , Destination From CoDestination"
    Set RsDestination = de.con.Execute(sqlText)
    Set ComboDestination.RowSource = RsDestination
    ComboDestination.listField = "Destination"
    ComboDestination.BoundColumn = "Id"

'    Dim rsCurrency As New ADODB.Recordset
'    Sqltext = "select PriceNo , PriceTYpe , col   from dbo.PriceTypes where PriceNo in (1,2,3)"
'    Set rsCurrency = de.con.Execute(Sqltext)
'    Set ComboCurrencyType.RowSource = rsCurrency
'    ComboCurrencyType.ListField = "PriceTYpe"
'    ComboCurrencyType.BoundColumn = "PriceNo"
    
'        Set ComboFeesPriceType.RowSource = rsCurrency
'    ComboFeesPriceType.ListField = "PriceTYpe"
'    ComboFeesPriceType.BoundColumn = "PriceNo"





End Sub
Function GetModelName(ModNo As Integer) As String
On Error GoTo ErrorHandler
Dim RsModelName As New ADODB.Recordset

sqlText = "Select Name    from adhammodels  Where ModNo=" & ModNo
Set RsModelName = de.con.Execute(sqlText)
If RsModelName.RecordCount > 0 Then
    GetModelName = RsModelName!name
Else
    GetModelName = ""
End If
Exit Function
ErrorHandler:
GetModelName = ""
End Function

Function GetSymbol(ModNo As Integer) As String
On Error GoTo ErrorHandler
Dim RsSymbol As New ADODB.Recordset

sqlText = "Select Symbol    from adhammodels   Where ModNo=" & ModNo
Set RsSymbol = de.con.Execute(sqlText)
If RsSymbol.RecordCount > 0 Then
    GetSymbol = RsSymbol!Symbol
Else
    GetSymbol = ""
End If
Exit Function
ErrorHandler:
GetSymbol = ""
End Function
Sub init()
    
    top = 0
    left = 0
    sstab1.Tab = 0
    Vindex = 2
    FillCombos
    Ok = True
    flexGridDetails.Rows = 1
    FillFormating 2
    
    FlexSummary.Rows = 1
    FillFormating 3
    
    FlexAccount.Rows = 1
    FillFormating 4
    
    FlexNotTransfered.Rows = 1
    FillFormating 5
    
    
    ComboBillNo(0).Text = ">="
    ComboBillNo(1).Text = "<="
    ComboDate(0).Text = ">="
    ComboDate(1).Text = "<="
    
    ComboFixDate(0).Text = ">="
    ComboFixDate(1).Text = "<="
    
    ComboAmount(0).Text = ">="
    ComboAmount(1).Text = "<="
    
    ComboFeesOperation(0).Text = ">="
    ComboFeesOperation(1).Text = "<="
    TxtDate(0).Text = Format(Now, "dd/mm/yyyy")
    TxtDate(1).Text = Format(Now, "dd/mm/yyyy")
    
    TxtFixBillDate(0).Text = Format(Now, "dd/mm/yyyy")
    TxtFixBillDate(1).Text = Format(Now, "dd/mm/yyyy")

End Sub

Function ShowPermission(ByVal Level As Integer, ByVal empNo As Integer) As Boolean
Dim rs As New ADODB.Recordset
    sqlText = "select EmpNo , AllowTransfer from MaintUsers Where EmpNo=" & empNo
    Set rs = de.con.Execute(sqlText)
    If rs!AllowTransfer And (2 ^ Level) Then
        ShowPermission = True
    Else
        ShowPermission = False
    End If
End Function

Private Sub ComboFees_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    TxtBillNo(0).SetFocus
    SendKeys "{home}+{end}"
End If
End Sub

Private Sub ComboOperationType_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    ComboType.SetFocus
End If
End Sub

Private Sub ComboPayment_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    TxtClientName.SetFocus
    SendKeys "{HOME}+{END}"
End If
End Sub

Private Sub ComboType_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    ComboPayment.SetFocus
End If
End Sub

Function HavePer(ByVal Level As Double, ByVal empNo As Integer) As Boolean
On Error GoTo ErrorHandler
    sqlText = "Select Permission From MaintUsers Where EmpNo=" & empNo
    Set rs = de.con.Execute(sqlText)
    sqlText = "Select dbo.BitwizeAnd(" & 2 ^ Level & "," & rs!Permission & ") as Result"
    Set rs = de.con.Execute(sqlText)
    HavePer = rs!result
Exit Function
ErrorHandler:
HavePer = False
End Function

Private Sub FlexAccount_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button And vbRightButton Then
        If Gettag(empNo, 24) Then
            PopupMenu mnufile
        End If
    End If
End Sub

Private Sub flexGridDetails_DblClick()
On Error GoTo ErrorHandler
If Gettag(empNo, 8) Then
    With flexGridDetails
        IDBill = .TextMatrix(.Row, ColBillNo)
        LoadForm = True
        FrmBills.Show
    End With
End If
Exit Sub
ErrorHandler:
MsgBox Err.Description
End Sub

Private Sub flexGridDetails_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button And vbRightButton Then
        If Gettag(empNo, 25) Then
            PopupMenu mnufile1
        End If
'        If Gettag(empNo, 23) Then
'            PopupMenu mnufile2
'        End If
'        If Gettag(6, empNo) Then
'            PopupMenu mnufile3
'        End If
'        If Gettag(empNo, 26) Then
'            PopupMenu mnufile4
'        End If
        PopupMenu mnufile5
    End If
End Sub

Private Sub Form_Load()
    init
End Sub

Sub ColorRow(Color As Long)
With flexGridDetails

For i = 1 To .Rows - 1
    If .TextMatrix(i, ColIsError) Then
            Row = i
             For J = 1 To .Cols - 1
                .Col = J
                .Row = i
                .CellBackColor = Color
            Next
    End If
Next
End With
End Sub

Sub SearchData(IsTransfered As Integer)
On Error GoTo ErrorHandler
Dim RsSum  As New ADODB.Recordset
    With SearchRec
        .DestinationId = Val(ComboDestination.BoundText)
        .OperationId = Val(ComboOperationType.BoundText)
        .TYpeId = Val(ComboType.BoundText)
        If ComboPayment.BoundText = "" Then
            .PaymentId = -1
        Else
            .PaymentId = ComboPayment.BoundText
        End If
        
        .clientId = Val(TxtClientName.Tag)
        .ClientType = Val(LClientType.Tag)
        .ModNo = Val(TxtModelName.Tag)
        .stkno = TxtitemName.Tag
        .FeesId = Val(ComboFees.BoundText)
        .OperationFromBillNo = ComboBillNo(0).Text
        .FromBillNo = Val(TxtBillNo(0).Text)
        .OperationTillBillno = ComboBillNo(1).Text
        .TillBillNo = Val(TxtBillNo(1).Text)
        
        .OperationFromDate = ComboDate(0).Text
        .FromDate = TxtDate(0).Text
        .OperationTillDate = ComboDate(1).Text
        .TillDate = TxtDate(1).Text
        
        
        
        
        .OperationFixFromDate = ComboFixDate(0).Text
        .FixFromDate = TxtFixBillDate(0).Text
        .OperationFixTillDate = ComboFixDate(1).Text
        .FixTillDate = TxtFixBillDate(1).Text
        
        
        
        .OperationFromAmount = ComboAmount(0).Text
        .FromAmount = Val(TxtAmount(0).Text)
        .OperationTillAmount = ComboAmount(1).Text
        .TillAmount = Val(TxtAmount(1).Text)
        
        .OperationFromFees = ComboFeesOperation(0).Text
        .FromFees = Val(TxtFeesAmount(0).Text)
        .OperationTillFees = ComboFeesOperation(1).Text
        .TillFees = Val(TxtFeesAmount(1).Text)
        .Voption = Vindex
        
        sstab1.Visible = False
        Screen.MousePointer = vbHourglass
        sqlText = "Exec sp_MaintData " & .OperationId & "," & .TYpeId & "," & .PaymentId & "," & .clientId & "," & .ClientType & "," & .ModNo & ",'" & .stkno & "'," & .FeesId & ",'" & .OperationFromBillNo & "','" & .OperationTillBillno & "'," & .FromBillNo & "," & .TillBillNo & ",'" & .OperationFromDate & "','" & .OperationTillDate & "','" & ConvertControlDate(.FromDate) & "','" & ConvertControlDate(.TillDate) & "','" & .OperationFixFromDate & "','" & .OperationFixTillDate & "','" & ConvertControlDate(.FixFromDate) & "','" & ConvertControlDate(.FixTillDate) & "','" & .OperationFromAmount & "','" & .OperationTillAmount & "'," & .FromAmount & "," & .TillAmount & ",'" & .OperationFromFees & "','" & .OperationTillFees & "'," & .FromFees & "," & .TillFees & ",'" & .OperationFromTotal & "','" & .OperationTillTotal & "'," & .FromTotal & "," & .TillTotal & "," & .Voption & "," & IsTransfered & "," & empNo & "," & .DestinationId
        de.con.Execute (sqlText)
        
        sqlText = "Select Id, CallNo , Billno, SerByYear , Billdate, FixBillDate ,  OperationType , OperationTypeName, Stat, Statdescription,      PaymentTYpeId , PaymentTypeName ,ClientName , ClientPhoneNBr, Name, Symbol, ModelQty, StkNo, AccNo , stkname, Qty, PriceTYpeName, Price, TotPrice, IsFixed, IsTransfered , IsTransferedToStock , IsError From t_MvMaintPayment"
        Set rs = de.con.Execute(sqlText)
        Set flexGridDetails.DataSource = rs
        FillFormating 2
        
        ColorRow &HC0E0FF
        
        sqlText = "Select Sum(TotPrice) as TotPrice , Count(*) CountRec From t_MvMaintPayment where isnull(IsError,0)=0"
        Set RsSum = de.con.Execute(sqlText)
        LSumDetails.Caption = Format(IIf(IsNull(RsSum!TotPrice), 0, RsSum!TotPrice), "###,###,###.00")
        LCountDetails.Caption = IIf(IsNull(RsSum!CountRec), 0, RsSum!CountRec)
        
        
        If .TYpeId = 6 Then
            sqlText = "Select t1.CallNo , m1.AllTotalPrice  From mvmainttotalpriceqry m1 inner join (Select distinct CallNo From t_MvMaintPayment) t1 on m1.billno = t1.callno Where Class=1 Order By t1.callno"
        Else
            sqlText = "Select m1.BillNo , m1.AllTotalPrice  From mvmainttotalpriceqry m1 inner join (Select distinct BillNo From t_MvMaintPayment) t1 on m1.billno = t1.billno Where Class=0 Order By m1.billno"
        End If
        Set rs = de.con.Execute(sqlText)
        Set FlexSummary.DataSource = rs
        FillFormating 3
        
        If .TYpeId = 6 Then
            sqlText = "Select Sum(AllTotalPrice) as AllTotalPrice , Count(*) CountRec From MvMainttotalpriceqry m1 inner join (Select distinct CallNo From t_MvMaintPayment) t1 on m1.billno = t1.callNo Where Class=1"
        Else
            sqlText = "Select Sum(AllTotalPrice) as AllTotalPrice , Count(*) CountRec From MvMainttotalpriceqry m1 inner join (Select distinct BillNo From t_MvMaintPayment) t1 on m1.billno = t1.billno Where Class=0"
        End If
        Set RsSum = de.con.Execute(sqlText)
        LSumTotal.Caption = Format(IIf(IsNull(RsSum!AllTotalPrice), 0, RsSum!AllTotalPrice), "###,###,###.00")
        LCountTotal.Caption = IIf(IsNull(RsSum!CountRec), 0, RsSum!CountRec)

        
        sqlText = "Select [rowcount] , FixBillDate , Stkno, AccNo, Stkname, Deb, DebAccName, Cre, CreAccName, StkName + ' ' + ' ⁄œœ ' +  Convert(varchar(10),CountRec) as Description ,     OperationType, OperationTypeName, paymenttypeid, paymentTypeName, Stat, statdescription, TotPrice , CountRec , Type From MvMaintAccountQry Order By StkNo"
        Set rs = de.con.Execute(sqlText)
        Set FlexAccount.DataSource = rs
        FillFormating 4

        sqlText = "Select case when isnull(callNo,0)=0 then billno else callNo end billno, Sum(totprice)TotPrice  , Fixbilldate , OperationType , OperationTypeName , paymenttypeid , paymentTypeName ,Stat , statdescription from MvMaintPaymentsQry_ar MvMaintPaymentsQry Where IsFixed=1 And IsTransfered=0 And Fixbilldate <  '" & ConvertControlDate(.FixFromDate) & "' group by case when isnull(callNo,0)=0 then billno else callNo end  , Fixbilldate , OperationType , OperationTypeName , Stat , statdescription , paymenttypeid , paymentTypeName"
        Set rs = de.con.Execute(sqlText)
        Set FlexNotTransfered.DataSource = rs
        FillFormating 5

        sqlText = "Select Sum(TotPrice)TotPrice From MvMaintAccountQry"
        Set RsSum = de.con.Execute(sqlText)
        LSumAccount.Caption = Format(IIf(IsNull(RsSum!TotPrice), 0, RsSum!TotPrice), "###,###,###.00")
        LCountAccount.Caption = FlexAccount.Rows - 1
        sstab1.Visible = True
        Screen.MousePointer = vbDefault
    End With
Exit Sub
ErrorHandler:
sstab1.Visible = True
Screen.MousePointer = vbDefault
MsgBox Err.Description
End Sub
Sub PrintData(Vindex As Integer, Optional BillNo As Variant)
On Error GoTo ErrorHandler
With Cr1
    .Connect = ConnectName("")
    Select Case Vindex
        Case 0
            .SQLQuery = "SELECT   OperationType, OperationTypeName, MaintTYpe, statdescription, paymenttypeid, paymentTypeName, StkNo, AccNo, StkName, Qty, TotalPrice FROM    t_MvMaintPaymentQry  ORDER BY    OperationType ASC,    MaintTYpe ASC,    paymenttypeid ASC,    AccNo ASC"
            .ReportFileName = App.Path + "\Reports\RepDocument.rpt"
            .Formulas(0) = "fromDate='" & TxtFixBillDate(0).Text & "'"
            .Formulas(1) = "TillDate='" & TxtFixBillDate(1).Text & "'"
        Case 1
            .SQLQuery = "SELECT     Clientaccno, ClientName, classname, OperationType, OperationTypeName, stat, statdescription, paymenttypeid, paymentTypeName, totprice FROM     dbo.MvMaintCreditQry   ORDER BY    OperationType ASC,    stat ASC,    paymenttypeid ASC,    totprice DESC"
            .ReportFileName = App.Path + "\Reports\RepCreditReport.rpt"
            .Formulas(0) = "fromDate='" & TxtFixBillDate(0).Text & "'"
            .Formulas(1) = "TillDate='" & TxtFixBillDate(1).Text & "'"
    
       Case 2
            .SQLQuery = "SELECT     SerByYear , billno, billdate, OperationTypeName, statdescription, ClientName, stkno, stkname, qty, price, paymentTypeName, TotPrice, Row FROM     dbo.t_MvMaintPayment ORDER BY    billno ASC,    ClientName ASC,    Row ASC,    stkno ASC"
            .ReportFileName = App.Path + "\Reports\RepCreditDetailsReport.rpt"
            .Formulas(0) = "fromDate='" & TxtFixBillDate(0).Text & "'"
            .Formulas(1) = "TillDate='" & TxtFixBillDate(1).Text & "'"
'       Case 3
'            .Formulas(0) = ""
'            .Formulas(1) = ""
'
'            If IsMissing(BillNo) Then
'                MsgBox "·„ Ì „  ÕœÌœ —ﬁ„ «·›« Ê—…", vbExclamation
'                Exit Sub
'            Else
'                .SQLQuery = "SELECT    billno, billdate , FixBillDate , MaintTYpe, ClientName, ClientPhoneNBr, Name, Symbol, stkno, stkname, qty, price, TotPrice   FROM    MvMaintPaymentsQry_ar  MvMaintPaymentsQry   Where stat<>6 and  billno = " & BillNo & "  Order By Row , StkNo"
'                .ReportFileName = App.Path + "\Reports\RepBill1.rpt"
'            End If
    End Select
    .DiscardSavedData = True
    .WindowState = crptMaximized
    .Action = 1
End With
Exit Sub
ErrorHandler:
MsgBox Err.Description
End Sub

Private Sub mnu0_Click(Index As Integer)
Select Case Index
    Case 0
        TransMaintAccRegTemp
End Select

End Sub

Sub CancelAccount()
On Error GoTo ErrorHandler
Dim FirstRow As Integer, LastRow As Integer
With flexGridDetails
            If .Row >= .RowSel Then
            FirstRow = .Row
            LastRow = .RowSel
        Else
            FirstRow = .RowSel
            LastRow = .Row
        End If
        For i = FirstRow To LastRow Step -1
            Vrow = i
            If .TextMatrix(i, ColStat) = 6 Then
                sqlText = "update r1 set carried=0 from Reparation r1  where r1.RepNo=" & .TextMatrix(i, ColBillNo)
                de.con.Execute (sqlText)
'                sqlText = "update r2 set carried=0 from Reparation r1 inner join ReparationPieces r2 on r1.RepNo = r2.repno   where r1.RepNo=" & .TextMatrix(i, ColBillNo)
'                de.con.Execute (sqlText)
                .TextMatrix(i, ColIsTransfered) = 0
            Else
                sqlText = " Update m1 set IsTransfered=0 from MvMaintPayments m1 Where BillNo=" & .TextMatrix(i, ColBillNo)
                de.con.Execute (sqlText)
                .TextMatrix(i, ColIsTransfered) = 0
            End If
        Next
End With
Exit Sub
ErrorHandler:
MsgBox Err.Description
End Sub
Private Sub mnu1_Click(Index As Integer)
Select Case Index
    Case 0
        CancelAccount
End Select
End Sub
'Function TransferToMvStock() As Boolean
''‰—ÕÌ· «·„Ê«œ «·√Ê·Ì… Ê «·„ÀÌ …
'On Error GoTo ErrorHandler
'Dim sqlText As String
'ProgressBar1.Visible = True
'With flexGridDetails
'    ProgressBar1.Min = 1
'    ProgressBar1.Max = .Rows
'    ProgressBar1.Value = 1
'    For i = 1 To .Rows - 1
'        If .TextMatrix(i, ColStkNo) <> "" And Abs(.TextMatrix(i, Colisfixed)) = 1 And .TextMatrix(i, ColStat) = 1 Then
'            sqlText = "Exec SP_TransferMaintToStock " & .TextMatrix(i, ColId) & ",'" & Trim(.TextMatrix(i, ColStkNo)) & "'," & .TextMatrix(i, ColQty) & "," & .TextMatrix(i, ColOperationTypeId) & "," & .TextMatrix(i, ColBillNo) & "," & empNo
'            de.con.Execute (sqlText)
'        End If
'        ProgressBar1.Value = ProgressBar1.Value + 1
'    Next
'End With
'ProgressBar1.Visible = False
'TransferToMvStock = True
'Exit Function
'ErrorHandler:
'TransferToMvStock = False
'MsgBox Err.Description
'End Function
'Private Sub mnu2_Click(Index As Integer)
'If TransferToMvStock Then
'    MsgBox " „  —ÕÌ· «·„Ê«œ", vbInformation, " —ÕÌ· «·„Ê«œ ≈·Ï «·„” Êœ⁄« "
'End If
'End Sub

'Sub ExportItemsToTextFile()
'On Error Resume Next
'        sqlText = "select billno NotificationNo, stkno PieceStockNo , Fixbilldate MvtDate, qty , OperationType OpKind  from t_MvMaintPayment where ISNULL(stkno,'') <> ''"
'        Set rs = de.con.Execute(sqlText)
'        Kill "E:\MainData\≈‘⁄«—« .txt"
'        rs.Save "E:\MainData\≈‘⁄«—« .txt", adPersistADTG
'        rs.Close
'        MsgBox " „ Õ›Ÿ ≈‘⁄«—«  «·„Ê«œ «·„—”·…" & " " & "E:\MainData\≈‘⁄«—« .txt"
'End Sub
'
'Private Sub mnu3_Click(Index As Integer)
'Select Case Index
'    Case 0 ' ’œÌ— ≈‘⁄«—«  «·„Ê«œ «·«Ê·Ì…
'        ExportItemsToTextFile
'End Select
'End Sub

'Private Sub mnu4_Click(Index As Integer)
'Select Case Index
'    Case 0 ' ’œÌ— ≈‘⁄«—«  «·„Ê«œ «·«Ê·Ì…
'        PrintData 3, Val(flexGridDetails.TextMatrix(flexGridDetails.Row, ColBillNo))
'End Select
'End Sub

Private Sub mnu5_Click(Index As Integer)
If TransferToTextFile Then
    MsgBox " „  Õ“Ì‰ «·„·› «·‰’Ì ⁄·Ï «·„”«—" & Chr(13) & "E:\MainData\_HallsItems.txt", vbInformation, " Œ“Ì‰ «·„·› «·‰’Ì"
Else
    MsgBox "Œÿ√ ›Ì  —ÕÌ· «·„·›", vbExclamation, "Œÿ√ ›Ì  —ÕÌ· «·„·›"
End If

End Sub
Function TransferToTextFile() As Boolean
On Error GoTo ErrorHandler
Dim sqlText As String
Dim rs As New ADODB.Recordset

        sqlText = "Select  Billno , billdate   ,stkid , StkNo , stkname, Qty From t_MvMaintPayment"
        Set rs = de.con.Execute(sqlText)
        
        If Dir$("e:\mainData\_HallItems.txt") <> "" Then
            Kill "E:\mainData\_HallItems.txt"
        End If
        rs.Save "E:\mainData\_HallItems.txt", adPersistADTG
        rs.Close
TransferToTextFile = True
Exit Function
ErrorHandler:
TransferToTextFile = False
End Function

Private Sub SSOption1_Click(Index As Integer, Value As Integer)
Vindex = Index
End Sub



Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
    Case 1
        SearchData 0
    Case 3
        PrintData 0
        
    Case 5
        Unload Me
End Select
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Select Case ButtonMenu.Tag
    Case 1
        PrintData 1
   Case 2
        PrintData 2
End Select
End Sub

Private Sub TxtAmount_Change(Index As Integer)
If Index = 0 Then
    TxtAmount(1).Text = TxtAmount(0).Text
End If

End Sub

Private Sub TxtAmount_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    If Index = 0 Then
        TxtAmount(1).SetFocus
        SendKeys "{home}+{end}"
    Else
        TxtFeesAmount(0).SetFocus
        SendKeys "{home}+{end}"
    End If
End If

End Sub

Private Sub TxtBillNo_Change(Index As Integer)
If Index = 0 Then
    TxtBillNo(1).Text = TxtBillNo(0).Text
End If
End Sub

Private Sub TxtBillNo_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    If Index = 0 Then
        TxtBillNo(1).SetFocus
        SendKeys "{home}+{end}"
    Else
        TxtDate(0).SetFocus
        SendKeys "{home}+{end}"
    End If
End If
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
    sqlText = "Select Top 10 [ClientId] , [ClientName]  , [ClientTypeName] , Class  From ClientQry Where ClientName like" & LikeExpression(TxtClientName.Text) & " or ClientPhoneNBr like " & LikeExpression(TxtClientName.Text)
    Set RsSearch = de.con.Execute(sqlText)
    If RsSearch.RecordCount > 0 Then
        Set Grid.DataSource = RsSearch
        Grid.Row = 0
        FillFormating 1
        ChangeCursor 1
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
Pos = 1
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
        LClientType.Caption = Grid.TextMatrix(IIf(Grid.Row = 0, 1, Grid.Row), col3)
        LClientType.Tag = Grid.TextMatrix(IIf(Grid.Row = 0, 1, Grid.Row), Col4)
        lPhoneNbr.Caption = GetPhoneNbr(TxtClientName.Tag, LClientType.Tag)
        Ok = True
    ElseIf Grid.Visible = False And TxtClientName.Text <> "" And Val(TxtClientName.Tag) <> 0 Then
        TxtModelName.SetFocus
        TxtModelName.SelStart = 0
        TxtModelName.SelLength = Len(TxtModelName.Text)
        Exit Sub
    ElseIf Grid.Visible = False And TxtClientName.Text <> "" And Val(TxtClientName.Tag) = 0 Then
        Ok = False
        TxtClientName.Tag = 0
        LClientType.Tag = 0
        LClientType.Caption = ""
        Ok = True
    Else
        Ok = False
        TxtClientName.Tag = 0
        TxtClientName.Text = ""
        LClientType.Caption = ""
        LClientType.Tag = 0
        lPhoneNbr.Caption = ""
        Ok = True
    End If
    Grid.Visible = False
    TxtModelName.SetFocus
    TxtModelName.SelStart = 0
    TxtModelName.SelLength = Len(TxtModelName.Text)
End If
End Sub

Private Sub TxtDate_Change(Index As Integer)
If Index = 0 Then
    TxtDate(1).Text = TxtDate(0).Text
End If

End Sub

Private Sub TxtDate_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    If Index = 0 Then
        TxtDate(1).SetFocus
        SendKeys "{home}+{end}"
    Else
        TxtFixBillDate(0).SetFocus
        SendKeys "{home}+{end}"
    End If
End If

End Sub

Private Sub TxtFeesAmount_Change(Index As Integer)
If Index = 0 Then
    TxtFeesAmount(1).Text = TxtFeesAmount(0).Text
End If
End Sub

Private Sub TxtFeesAmount_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    If Index = 0 Then
        TxtFeesAmount(1).SetFocus
        SendKeys "{home}+{end}"
    Else
        ComboOperationType.SetFocus
        SendKeys "{home}+{end}"
    End If
End If

End Sub

Private Sub TxtFixBillDate_Change(Index As Integer)
If Index = 0 Then
    TxtFixBillDate(1).Text = TxtFixBillDate(0).Text
End If

End Sub

Private Sub TxtFixBillDate_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    If Index = 0 Then
        TxtFixBillDate(1).SetFocus
        SendKeys "{home}+{end}"
    Else
        TxtAmount(0).SetFocus
        SendKeys "{home}+{end}"
    End If
End If

End Sub

Private Sub TxtItemName_Change()
On Error GoTo ErrorHandler
Dim RsSearch As New ADODB.Recordset
If TxtitemName.Text = "" Then
    TxtitemName.Tag = ""
    Grid.Visible = False
    Exit Sub
End If
If Ok Then
    Flag = False
    sqlText = "select Top 10 StkNo , StkName  , FnlQnt from CoStock Where StkName Like" & LikeExpression(TxtitemName.Text) & " or StkNo like '" & TxtitemName.Text & "%'"
    Set RsSearch = de.con.Execute(sqlText)
    If RsSearch.RecordCount > 0 Then
        Set Grid.DataSource = RsSearch
        Grid.Row = 0
        FillFormating 1
        ChangeCursor 3
        Grid.Visible = True
    Else
        TxtitemName.Tag = ""
        Grid.Visible = False
    End If
    Flag = True
End If
Exit Sub
ErrorHandler:
MsgBox Err.Description
End Sub

Private Sub TxtItemName_GotFocus()
Pos = 3
End Sub


Private Sub TxtitemName_KeyDown(KeyCode As Integer, Shift As Integer)
    Flag = True
    If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then MoveCursor KeyCode, Grid
    Flag = True

End Sub

Private Sub TxtitemName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Grid.Visible Then
        Ok = False
        TxtitemName.Tag = Grid.TextMatrix(IIf(Grid.Row = 0, 1, Grid.Row), ColNo)
        TxtitemName.Text = Grid.TextMatrix(IIf(Grid.Row = 0, 1, Grid.Row), ColNo)
        LItemName.Caption = Grid.TextMatrix(IIf(Grid.Row = 0, 1, Grid.Row), ColName)
        'LBalance.Caption = Grid.TextMatrix(IIf(Grid.Row = 0, 1, Grid.Row), Col3)
        'LPrice.Caption = GetPrice(Val(ComboCurrencyType.BoundText), TxtitemName.Tag)
        Ok = True
    ElseIf Grid.Visible = False And TxtitemName.Text <> "" And TxtitemName.Tag <> "" Then
        TxtQty.SetFocus
        TxtQty.SelStart = 0
        TxtQty.SelLength = Len(TxtQty.Text)
        Exit Sub
    Else
        Ok = False
        TxtitemName.Tag = ""
        TxtitemName.Text = ""
        LItemName.Caption = ""
        'LBalance.Caption = ""
        Ok = True
    End If
    Grid.Visible = False
    ComboFees.SetFocus
End If

End Sub

Private Sub TxtModelName_Change()
On Error GoTo ErrorHandler
Dim RsSearch As New ADODB.Recordset
If TxtModelName.Text = "" Then
    TxtModelName.Tag = 0
    Grid.Visible = False
    Exit Sub
End If
If Ok Then
    Flag = False
    sqlText = "select ModNo , Symbol , Name    from adhammodels Where Symbol    Like" & LikeExpression(TxtModelName.Text) & " or Name    like " & LikeExpression(TxtModelName.Text)
    Set RsSearch = de.con.Execute(sqlText)
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
ErrorHandler:
MsgBox Err.Description
End Sub


Private Sub TxtModelName_GotFocus()
ChangeToArabic
Pos = 2
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


    
