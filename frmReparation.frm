VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmReparation 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8130
   ClientLeft      =   45
   ClientTop       =   450
   ClientWidth     =   11805
   BeginProperty Font 
      Name            =   "Arabic Transparent"
      Size            =   12
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   8130
   ScaleWidth      =   11805
   Begin Crystal.CrystalReport Cr 
      Left            =   30
      Top             =   150
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Timer TimerFindClient 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   11430
      Top             =   210
   End
   Begin VB.TextBox txtFindCallNo 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   405
      Left            =   8940
      RightToLeft     =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   150
      Width           =   1785
   End
   Begin TabDlg.SSTab SSTabRep 
      Height          =   7395
      Left            =   90
      TabIndex        =   35
      Top             =   750
      Width           =   11685
      _ExtentX        =   20611
      _ExtentY        =   13044
      _Version        =   393216
      TabOrientation  =   3
      Tab             =   2
      TabHeight       =   520
      ForeColor       =   -2147483646
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arabic Transparent"
         Size            =   14.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "»Ì«‰«  «·≈’·«Õ"
      TabPicture(0)   =   "frmReparation.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "SSFrameDip"
      Tab(0).Control(1)=   "SSFrameRepInfo"
      Tab(0).Control(2)=   "SSFrameModel"
      Tab(0).Control(3)=   "cmdDiscount"
      Tab(0).Control(4)=   "CmdSave"
      Tab(0).Control(5)=   "CmdDelete"
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "«·√⁄ÿ«·"
      TabPicture(1)   =   "frmReparation.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "SSFrameRep"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "«·„Ê«œ"
      TabPicture(2)   =   "frmReparation.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label6(8)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label6(9)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label6(10)"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Label6(11)"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "LItemName"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "LBalance"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "LPrice"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "Label6(2)"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "Label6(0)"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "Label2"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "Flexgrid"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "FrmePieces"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).Control(12)=   "TxtitemName"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).Control(13)=   "TxtQty"
      Tab(2).Control(13).Enabled=   0   'False
      Tab(2).ControlCount=   14
      Begin VB.TextBox TxtQty 
         Alignment       =   1  'Right Justify
         Height          =   405
         Left            =   3150
         RightToLeft     =   -1  'True
         TabIndex        =   60
         Top             =   570
         Width           =   855
      End
      Begin VB.TextBox TxtitemName 
         Alignment       =   1  'Right Justify
         Height          =   405
         Left            =   7920
         RightToLeft     =   -1  'True
         TabIndex        =   59
         Top             =   570
         Width           =   3345
      End
      Begin Threed.SSFrame FrmePieces 
         Height          =   555
         Left            =   120
         TabIndex        =   79
         Top             =   6720
         Width           =   11145
         _ExtentX        =   19659
         _ExtentY        =   979
         _Version        =   131074
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
            Height          =   345
            Left            =   8640
            RightToLeft     =   -1  'True
            TabIndex        =   89
            Top             =   120
            Width           =   1545
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "«·≈Ã„«·Ì"
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
            Height          =   315
            Index           =   14
            Left            =   10260
            TabIndex        =   88
            Top             =   120
            Width           =   825
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid Flexgrid 
         Height          =   5625
         Left            =   120
         TabIndex        =   61
         Top             =   1050
         Width           =   11145
         _cx             =   19659
         _cy             =   9922
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arabic Transparent"
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
      Begin Threed.SSFrame SSFrameDip 
         Height          =   1905
         Left            =   -74790
         TabIndex        =   74
         Top             =   4200
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   3360
         _Version        =   131074
         Begin VB.TextBox txtFindClientName 
            Alignment       =   1  'Right Justify
            Height          =   405
            Left            =   150
            RightToLeft     =   -1  'True
            TabIndex        =   28
            Top             =   720
            Width           =   2235
         End
         Begin MSDataListLib.DataCombo DcboClientDipositName 
            Height          =   405
            Left            =   150
            TabIndex        =   29
            Top             =   1230
            Width           =   2235
            _ExtentX        =   3942
            _ExtentY        =   714
            _Version        =   393216
            BackColor       =   -2147483643
            ListField       =   ""
            BoundColumn     =   ""
            Text            =   ""
            RightToLeft     =   -1  'True
            Object.DataMember      =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arabic Transparent"
               Size            =   12
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSDataListLib.DataCombo DcboClientKind 
            Bindings        =   "frmReparation.frx":0054
            Height          =   405
            Left            =   150
            TabIndex        =   27
            Top             =   210
            Width           =   2235
            _ExtentX        =   3942
            _ExtentY        =   714
            _Version        =   393216
            BackColor       =   -2147483643
            ListField       =   "ClKindName"
            BoundColumn     =   "ClKindNo"
            Text            =   ""
            RightToLeft     =   -1  'True
            Object.DataMember      =   "ClientKind"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arabic Transparent"
               Size            =   12
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label45 
            AutoSize        =   -1  'True
            Caption         =   "»ÕÀ «”„-Â« ›"
            Height          =   285
            Left            =   2490
            RightToLeft     =   -1  'True
            TabIndex        =   77
            Top             =   720
            Width           =   1200
         End
         Begin VB.Label Label44 
            AutoSize        =   -1  'True
            Caption         =   "«”„ «·“»Ê‰"
            Height          =   285
            Left            =   2490
            RightToLeft     =   -1  'True
            TabIndex        =   76
            Top             =   1230
            Width           =   810
         End
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            Caption         =   "‰Ê⁄ «·“»Ê‰"
            Height          =   285
            Left            =   2490
            RightToLeft     =   -1  'True
            TabIndex        =   75
            Top             =   210
            Width           =   870
         End
      End
      Begin Threed.SSFrame SSFrameRep 
         Height          =   6645
         Left            =   -74760
         TabIndex        =   52
         Top             =   270
         Width           =   10905
         _ExtentX        =   19235
         _ExtentY        =   11721
         _Version        =   131074
         Begin VB.ListBox lstCoReparation 
            Height          =   5190
            Left            =   5910
            RightToLeft     =   -1  'True
            Sorted          =   -1  'True
            TabIndex        =   34
            TabStop         =   0   'False
            Top             =   1200
            Width           =   4740
         End
         Begin VB.ListBox lstSelectedReparation 
            Height          =   5190
            Left            =   270
            RightToLeft     =   -1  'True
            Sorted          =   -1  'True
            TabIndex        =   53
            TabStop         =   0   'False
            Top             =   1200
            Width           =   4740
         End
         Begin VB.TextBox txtSearch 
            Alignment       =   1  'Right Justify
            Height          =   405
            Left            =   5910
            RightToLeft     =   -1  'True
            TabIndex        =   33
            Top             =   720
            Width           =   4740
         End
         Begin MSDataListLib.DataCombo DcboRepClass 
            Bindings        =   "frmReparation.frx":006A
            Height          =   405
            Left            =   5910
            TabIndex        =   32
            TabStop         =   0   'False
            Top             =   240
            Width           =   4740
            _ExtentX        =   8361
            _ExtentY        =   714
            _Version        =   393216
            BackColor       =   -2147483643
            ListField       =   "RepClassName"
            BoundColumn     =   "RepClassNo"
            Text            =   ""
            RightToLeft     =   -1  'True
            Object.DataMember      =   "CoRepClassification"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arabic Transparent"
               Size            =   12
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label lblSelectedEmployees 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "√⁄ÿ«· «·≈’·«Õ «·Õ«·Ì"
            Height          =   315
            Left            =   1710
            RightToLeft     =   -1  'True
            TabIndex        =   56
            Top             =   780
            Width           =   1785
         End
         Begin Threed.SSCommand CmdSelect 
            Height          =   405
            Left            =   5190
            TabIndex        =   55
            TabStop         =   0   'False
            Top             =   3030
            Width           =   555
            _ExtentX        =   979
            _ExtentY        =   714
            _Version        =   131074
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arabic Transparent"
               Size            =   12
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "<<"
         End
         Begin Threed.SSCommand CmdUnSelect 
            Height          =   405
            Left            =   5190
            TabIndex        =   54
            TabStop         =   0   'False
            Top             =   3450
            Width           =   555
            _ExtentX        =   979
            _ExtentY        =   714
            _Version        =   131074
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arabic Transparent"
               Size            =   12
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   ">>"
         End
      End
      Begin Threed.SSFrame SSFrameRepInfo 
         Height          =   3735
         Left            =   -74790
         TabIndex        =   38
         Top             =   270
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   6588
         _Version        =   131074
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arabic Transparent"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.TextBox txtNotes 
            Alignment       =   1  'Right Justify
            Height          =   900
            Left            =   240
            MultiLine       =   -1  'True
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   17
            Top             =   1695
            Width           =   4335
         End
         Begin MSMask.MaskEdBox txtRepTimeBegin 
            Height          =   405
            Left            =   8790
            TabIndex        =   7
            Top             =   1695
            Width           =   705
            _ExtentX        =   1244
            _ExtentY        =   714
            _Version        =   393216
            AutoTab         =   -1  'True
            MaxLength       =   5
            Format          =   "hh:mm"
            Mask            =   "99:99"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtRepTimeEnd 
            Height          =   405
            Left            =   8790
            TabIndex        =   8
            Top             =   2190
            Width           =   705
            _ExtentX        =   1244
            _ExtentY        =   714
            _Version        =   393216
            AutoTab         =   -1  'True
            MaxLength       =   5
            Format          =   "hh:mm"
            Mask            =   "99:99"
            PromptChar      =   "_"
         End
         Begin VB.TextBox txtTeamNo 
            Alignment       =   2  'Center
            Height          =   405
            Left            =   4950
            RightToLeft     =   -1  'True
            TabIndex        =   11
            Top             =   180
            Width           =   735
         End
         Begin VB.TextBox txtVoltAfter 
            Alignment       =   2  'Center
            Height          =   405
            Left            =   2700
            TabIndex        =   15
            Top             =   1185
            Width           =   1020
         End
         Begin VB.TextBox txtVoltBefor 
            Alignment       =   2  'Center
            Height          =   405
            Left            =   4665
            TabIndex        =   14
            Top             =   1185
            Width           =   1020
         End
         Begin VB.TextBox txtDescription 
            Alignment       =   1  'Right Justify
            Height          =   900
            Left            =   240
            MultiLine       =   -1  'True
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   19
            Top             =   2700
            Width           =   4335
         End
         Begin VB.TextBox txtRepPrice 
            Alignment       =   2  'Center
            Height          =   405
            Left            =   7380
            TabIndex        =   9
            Top             =   2700
            Width           =   2115
         End
         Begin VB.TextBox txtRepNo 
            Alignment       =   2  'Center
            Height          =   405
            Left            =   6630
            Locked          =   -1  'True
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   180
            Width           =   1755
         End
         Begin VB.TextBox txtNotesFind 
            Alignment       =   2  'Center
            Height          =   405
            Left            =   4665
            RightToLeft     =   -1  'True
            TabIndex        =   16
            Top             =   1695
            Width           =   1020
         End
         Begin VB.TextBox txtDescriptionFind 
            Alignment       =   2  'Center
            Height          =   405
            Left            =   4665
            RightToLeft     =   -1  'True
            TabIndex        =   18
            Top             =   2700
            Width           =   1020
         End
         Begin MSDataListLib.DataCombo DcboReparationStatus 
            Bindings        =   "frmReparation.frx":0080
            Height          =   405
            Left            =   2700
            TabIndex        =   13
            Top             =   690
            Width           =   2985
            _ExtentX        =   5265
            _ExtentY        =   714
            _Version        =   393216
            ListField       =   "statues"
            BoundColumn     =   "no"
            Text            =   ""
            RightToLeft     =   -1  'True
            Object.DataMember      =   "ReparationStatus"
         End
         Begin MSDataListLib.DataCombo DcboPayMethod 
            Bindings        =   "frmReparation.frx":0096
            Height          =   405
            Left            =   7380
            TabIndex        =   10
            Top             =   3195
            Width           =   2115
            _ExtentX        =   3731
            _ExtentY        =   714
            _Version        =   393216
            ListField       =   "name"
            BoundColumn     =   "no"
            Text            =   ""
            RightToLeft     =   -1  'True
            Object.DataMember      =   "paymethod"
         End
         Begin MSDataListLib.DataCombo DcboCliRecever 
            Bindings        =   "frmReparation.frx":00AC
            Height          =   405
            Left            =   7380
            TabIndex        =   6
            Top             =   1185
            Width           =   2115
            _ExtentX        =   3731
            _ExtentY        =   714
            _Version        =   393216
            ListField       =   "GroupRcName"
            BoundColumn     =   "SerGroupRC"
            Text            =   ""
            RightToLeft     =   -1  'True
            Object.DataMember      =   "MaintGroupReceivers"
         End
         Begin MSMask.MaskEdBox txtRegestDate 
            Height          =   405
            Left            =   8490
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   180
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   714
            _Version        =   393216
            AutoTab         =   -1  'True
            MaxLength       =   8
            Format          =   "dd/mm/yy"
            Mask            =   "99/99/99"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtRepDate 
            Height          =   405
            Left            =   8490
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   690
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   714
            _Version        =   393216
            AutoTab         =   -1  'True
            MaxLength       =   8
            Format          =   "dd/mm/yy"
            Mask            =   "99/99/99"
            PromptChar      =   "_"
         End
         Begin MSDataListLib.DataCombo DcboTeam 
            Height          =   405
            Left            =   2700
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   180
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   714
            _Version        =   393216
            ListField       =   "TeamName"
            BoundColumn     =   "TeamNo"
            Text            =   ""
            RightToLeft     =   -1  'True
            Object.DataMember      =   ""
         End
         Begin VB.Label lblFieldLabel 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "«·œ›⁄"
            ForeColor       =   &H80000002&
            Height          =   285
            Index           =   4
            Left            =   9645
            TabIndex        =   51
            Top             =   3195
            Width           =   330
         End
         Begin VB.Label lblFieldLabel 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   " «—ÌŒ «·≈œŒ«·"
            ForeColor       =   &H80000002&
            Height          =   285
            Index           =   3
            Left            =   9615
            RightToLeft     =   -1  'True
            TabIndex        =   50
            Top             =   180
            Width           =   1050
         End
         Begin VB.Label lblFieldLabel 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "«·ÊÕœ…"
            ForeColor       =   &H80000002&
            Height          =   285
            Index           =   2
            Left            =   5835
            TabIndex        =   49
            Top             =   180
            Width           =   465
         End
         Begin VB.Label lblFieldLabel 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "«·„” ﬁ»·"
            Height          =   285
            Index           =   0
            Left            =   9615
            TabIndex        =   48
            Top             =   1185
            Width           =   600
         End
         Begin VB.Label lblFieldLabel 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "›Ê·  »⁄œ"
            Height          =   285
            Index           =   24
            Left            =   3825
            TabIndex        =   47
            Top             =   1245
            Width           =   675
         End
         Begin VB.Label lblFieldLabel 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "›Ê·  ﬁ»·"
            Height          =   285
            Index           =   23
            Left            =   5835
            TabIndex        =   46
            Top             =   1185
            Width           =   690
         End
         Begin VB.Label lblFieldLabel 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "«·„·«ÕŸ« "
            Height          =   285
            Index           =   22
            Left            =   5775
            TabIndex        =   45
            Top             =   2700
            Width           =   810
         End
         Begin VB.Label lblFieldLabel 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "«·≈’·«Õ« "
            ForeColor       =   &H80000002&
            Height          =   285
            Index           =   15
            Left            =   5775
            TabIndex        =   44
            Top             =   1695
            Width           =   885
         End
         Begin VB.Label lblFieldLabel 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "ﬁÌ„… «·≈’·«Õ"
            Height          =   285
            Index           =   9
            Left            =   9615
            TabIndex        =   43
            Top             =   2700
            Width           =   1050
         End
         Begin VB.Label lblFieldLabel 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "”«⁄… «·«‰ Â«¡"
            ForeColor       =   &H80000002&
            Height          =   285
            Index           =   8
            Left            =   9615
            TabIndex        =   42
            Top             =   2190
            Width           =   1020
         End
         Begin VB.Label lblFieldLabel 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "”«⁄… «·»œ¡"
            ForeColor       =   &H80000002&
            Height          =   285
            Index           =   7
            Left            =   9615
            TabIndex        =   41
            Top             =   1695
            Width           =   795
         End
         Begin VB.Label lblFieldLabel 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   " «—ÌŒ «·≈’·«Õ"
            ForeColor       =   &H80000002&
            Height          =   285
            Index           =   6
            Left            =   9615
            TabIndex        =   40
            Top             =   690
            Width           =   1170
         End
         Begin VB.Label lblFieldLabel 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Õ«·… «·≈’·«Õ"
            ForeColor       =   &H80000002&
            Height          =   285
            Index           =   5
            Left            =   5835
            TabIndex        =   39
            Top             =   690
            Width           =   1065
         End
      End
      Begin Threed.SSFrame SSFrameModel 
         Height          =   2925
         Left            =   -70770
         TabIndex        =   57
         Top             =   4200
         Width           =   6885
         _ExtentX        =   12144
         _ExtentY        =   5159
         _Version        =   131074
         Begin VB.TextBox txtFindModels 
            Alignment       =   2  'Center
            Height          =   405
            Left            =   2790
            RightToLeft     =   -1  'True
            TabIndex        =   21
            Top             =   210
            Width           =   855
         End
         Begin VB.TextBox txtFindModelsName 
            Alignment       =   2  'Center
            Height          =   405
            Left            =   240
            RightToLeft     =   -1  'True
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   210
            Width           =   1905
         End
         Begin VB.TextBox txtProdSerNo 
            Alignment       =   2  'Center
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
            Height          =   420
            Left            =   4710
            MaxLength       =   3
            TabIndex        =   25
            Top             =   1770
            Width           =   1035
         End
         Begin VB.TextBox txtModStockNo 
            Alignment       =   2  'Center
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
            Height          =   420
            Left            =   1125
            MaxLength       =   8
            TabIndex        =   62
            TabStop         =   0   'False
            Top             =   1770
            Width           =   2385
         End
         Begin VB.TextBox txtGazNo 
            Alignment       =   2  'Center
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
            Height          =   420
            Left            =   240
            MaxLength       =   1
            TabIndex        =   58
            TabStop         =   0   'False
            Top             =   1770
            Width           =   795
         End
         Begin MSDataListLib.DataCombo DcboProduct 
            Height          =   405
            Left            =   4050
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   210
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   714
            _Version        =   393216
            ForeColor       =   -2147483646
            ListField       =   ""
            BoundColumn     =   "ProdFamNo"
            Text            =   ""
            RightToLeft     =   -1  'True
            Object.DataMember      =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arabic Transparent"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSDataListLib.DataCombo DcboModel 
            Height          =   405
            Left            =   240
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   720
            Width           =   5505
            _ExtentX        =   9710
            _ExtentY        =   714
            _Version        =   393216
            ListField       =   "FullName"
            BoundColumn     =   "ModNo"
            Text            =   ""
            RightToLeft     =   -1  'True
            Object.DataMember      =   ""
         End
         Begin MSMask.MaskEdBox txtPurchaseDate 
            Height          =   405
            Left            =   4710
            TabIndex        =   26
            Top             =   2310
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   714
            _Version        =   393216
            AutoTab         =   -1  'True
            MaxLength       =   8
            Format          =   "dd/mm/yy"
            Mask            =   "99/99/99"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtProductionDate 
            Height          =   405
            Left            =   3615
            TabIndex        =   24
            Top             =   1770
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   714
            _Version        =   393216
            AutoTab         =   -1  'True
            MaxLength       =   8
            Format          =   "dd/mm/yy"
            Mask            =   "99/99/99"
            PromptChar      =   "_"
         End
         Begin VB.Label lblFieldLabel 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "«·„‰ Ã"
            Height          =   285
            Index           =   10
            Left            =   5865
            TabIndex        =   73
            Top             =   210
            Width           =   420
         End
         Begin VB.Label lblFieldLabel 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "«·„ÊœÌ·"
            Height          =   285
            Index           =   11
            Left            =   5865
            TabIndex        =   72
            Top             =   720
            Width           =   540
         End
         Begin VB.Label lblFieldLabel 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "—ﬁ„"
            Height          =   285
            Index           =   12
            Left            =   3735
            TabIndex        =   71
            Top             =   210
            Width           =   255
         End
         Begin VB.Label lblFieldLabel 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "„ÊœÌ·"
            Height          =   285
            Index           =   13
            Left            =   2280
            TabIndex        =   70
            Top             =   210
            Width           =   450
         End
         Begin VB.Label lblFieldLabel 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Caption         =   " ”·”·"
            Height          =   420
            Index           =   14
            Left            =   4710
            TabIndex        =   69
            Top             =   1230
            Width           =   1035
         End
         Begin VB.Label lblFieldLabel 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Caption         =   " /«·≈‰ «Ã"
            Height          =   405
            Index           =   16
            Left            =   3615
            TabIndex        =   68
            Top             =   1230
            Width           =   1005
         End
         Begin VB.Label lblFieldLabel 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Caption         =   "«·„Œ“‰Ì"
            Height          =   420
            Index           =   17
            Left            =   1125
            TabIndex        =   67
            Top             =   1230
            Width           =   2385
         End
         Begin VB.Label lblFieldLabel 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Caption         =   "«·€«“"
            Height          =   420
            Index           =   18
            Left            =   240
            TabIndex        =   66
            Top             =   1230
            Width           =   795
         End
         Begin VB.Label lblFieldLabel 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "«·‘—«¡"
            Height          =   285
            Index           =   19
            Left            =   5865
            TabIndex        =   65
            Top             =   2310
            Width           =   480
         End
         Begin VB.Label lblFieldLabel 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "«·»«—ﬂÊœ"
            Height          =   285
            Index           =   20
            Left            =   5865
            TabIndex        =   64
            Top             =   1770
            Width           =   600
         End
         Begin VB.Label lblBarCodeNo 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Caption         =   "123456789"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   420
            Left            =   240
            TabIndex        =   63
            Top             =   2310
            Width           =   4365
         End
      End
      Begin VB.Label Label2 
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
         Height          =   375
         Left            =   1620
         TabIndex        =   91
         Top             =   570
         Width           =   1485
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "”⁄— «· «Ã—"
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   0
         Left            =   2100
         TabIndex        =   90
         Top             =   180
         Width           =   855
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "«·„»·€"
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   2
         Left            =   1065
         TabIndex        =   87
         Top             =   180
         Width           =   390
      End
      Begin VB.Label LPrice 
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
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   120
         TabIndex        =   86
         Top             =   570
         Width           =   1455
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
         Height          =   375
         Left            =   4050
         TabIndex        =   85
         Top             =   570
         Width           =   765
      End
      Begin VB.Label LItemName 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   4890
         TabIndex        =   84
         Top             =   570
         Width           =   3015
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "«·ﬂ„Ì…"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   11
         Left            =   3510
         TabIndex        =   83
         Top             =   180
         Width           =   405
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "«·‘—Õ"
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   10
         Left            =   7140
         TabIndex        =   82
         Top             =   180
         Width           =   750
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "«·—’Ìœ"
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   9
         Left            =   3840
         TabIndex        =   81
         Top             =   180
         Width           =   945
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "«·—ﬁ„ «·„Œ“‰Ì"
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   8
         Left            =   9345
         TabIndex        =   80
         Top             =   180
         Width           =   1890
      End
      Begin Threed.SSCommand cmdDiscount 
         Height          =   885
         Left            =   -74790
         TabIndex        =   78
         TabStop         =   0   'False
         ToolTipText     =   "«·Õ”„"
         Top             =   6240
         Visible         =   0   'False
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   1561
         _Version        =   131074
         ButtonStyle     =   1
      End
      Begin Threed.SSCommand CmdSave 
         Height          =   795
         Left            =   -71970
         TabIndex        =   30
         ToolTipText     =   "„Ê«›ﬁ"
         Top             =   6240
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   1402
         _Version        =   131074
         PictureFrames   =   1
         Picture         =   "frmReparation.frx":00C2
         ButtonStyle     =   1
      End
      Begin Threed.SSCommand CmdDelete 
         Height          =   795
         Left            =   -72765
         TabIndex        =   31
         ToolTipText     =   "Õ–›"
         Top             =   6240
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   1402
         _Version        =   131074
         PictureFrames   =   1
         Picture         =   "frmReparation.frx":0E23
         ButtonStyle     =   1
      End
   End
   Begin MSDataListLib.DataCombo DcboClientName 
      Height          =   405
      Left            =   390
      TabIndex        =   2
      TabStop         =   0   'False
      Tag             =   "0"
      Top             =   150
      Width           =   7875
      _ExtentX        =   13891
      _ExtentY        =   714
      _Version        =   393216
      ForeColor       =   255
      ListField       =   "ClientName"
      BoundColumn     =   "CallNo"
      Text            =   ""
      RightToLeft     =   -1  'True
      Object.DataMember      =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arabic Transparent"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblStatusPiece 
      ForeColor       =   &H000000C0&
      Height          =   105
      Left            =   2070
      RightToLeft     =   -1  'True
      TabIndex        =   37
      Top             =   150
      Width           =   75
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "«·‘ﬂÊÏ"
      Height          =   285
      Left            =   10815
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   150
      Width           =   555
   End
   Begin VB.Label LClient 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "«·“»Ê‰"
      Height          =   285
      Left            =   8370
      RightToLeft     =   -1  'True
      TabIndex        =   36
      Top             =   150
      Width           =   495
   End
   Begin VB.Menu MnuHelp 
      Caption         =   "ﬁ«∆„… „”«⁄œ…"
      Visible         =   0   'False
      Begin VB.Menu MnuClientLastProd 
         Caption         =   "¬Œ— „‰ Ã „‰ “Ì«—… ”«»ﬁ…"
      End
      Begin VB.Menu MnuChangeCallNo 
         Caption         =   " »œÌ· —ﬁ„ «·‘ﬂÊÏ"
      End
      Begin VB.Menu mnuLine 
         Caption         =   "-"
      End
      Begin VB.Menu mntStatistics 
         Caption         =   "≈Õ’«∆Ì« "
         Begin VB.Menu mnuRepititionAll 
            Caption         =   "‰”» «· ﬂ—«—"
         End
         Begin VB.Menu mntRepitition 
            Caption         =   "‰”» «· ﬂ—«— -   ﬁ—Ì— «·„œÌ— «·⁄«„"
         End
      End
   End
End
Attribute VB_Name = "frmReparation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsModel As New ADODB.Recordset '«·„ÊœÌ· «·Õ«·Ì
Dim DisableModSearch As Boolean
Dim ProdSerNoForRep As String
Dim ProdSerNo As String
Dim NewRepOp As Boolean
Dim NewProdOp As Boolean
Dim RsRep As New ADODB.Recordset
Dim RsPrices As New ADODB.Recordset
Dim DblTotalPiecesOnly As Double
Dim DblOldTotal As Double
Dim RsAllPieceInfo As New ADODB.Recordset
Dim RsModels As New ADODB.Recordset
Dim RsCallCheck As New ADODB.Recordset
Dim RsPieces As New ADODB.Recordset
Dim RsPiecesInfo As New ADODB.Recordset
Dim ComNo As String

Private Sub CmdAddPiece_Click()
    If IsRepCarried(txtRepNo.Text) Then
        MsgBox "«·ﬁ”Ì„… „—Õ·…", vbExclamation + vbMsgBoxRight, "Œÿ√"
        Exit Sub
    End If
    AddPiece
End Sub
Public Sub LockCtls(t As Boolean)
    '≈Ã—«∆Ì… ﬁ›· ⁄‰«’— «· Õﬂ„ Ê «·√“—«—
    '«· «»⁄… ·„Ã„Ê⁄… «·”Ã·«  «·√»
    'Rs
    txtFindCallNo.Enabled = t
    txtQty.Locked = t
    'txtPrice.Locked = T
    'txtFindPieceNo.Locked = t
    
    DcboPieces.Locked = t
    DcboPriceType.Locked = t
    CmdAddPiece.Enabled = t
    CmdCancelPiece.Enabled = Not t
    CmdEditPiece.Enabled = t
    CmdSavePiece.Enabled = Not t
    CmdDeletePiece.Enabled = t
    'CmdExit.Enabled = T
    DgrdPieces.Enabled = t
End Sub

Private Sub CmdCancelPiece_Click()
On Error GoTo Err_handler
    If Not CmdCancelPiece.Enabled Then Exit Sub
    CancelUpdateRec
    RsPieces.CancelUpdate
    LockCtls True
    Exit Sub
Err_handler:
    MsgBox Err.Description
End Sub

Private Sub CmdDelete_Click()
    If txtRepNo.Text & "" = "" Then Exit Sub
    If NewRepOp Then Exit Sub
    If CmdCancelPiece.Enabled Then CmdCancelPiece_Click
    If RsRep!Carried = 1 And RsPieces.RecordCount > 0 Then
        MsgBox "«·ﬁ”Ì„… „—Õ·…", vbExclamation + vbMsgBoxRight, "Œÿ√"
        Exit Sub
    End If
    Dim t As Integer
    t = MsgBox("Â· √‰  „ √ﬂœ „‰ Õ–› «·ﬁ”Ì„… «·Õ«·Ì… ø", vbYesNo + vbDefaultButton2 + vbQuestion + vbMsgBoxRight, " ‰»ÌÂ")
    If t = vbNo Then Exit Sub
    Dim sqltext As String
    sqltext = " delete from AdhamProducts where ProdNo = "
    sqltext = sqltext & "(select top 1 ProdNo from Reparation where RepNo = " & Trim(txtRepNo.Text) & " And ID_ComNo = " & ComNo & ")"
    sqltext = sqltext & Chr(13) & " delete from ReparationPieces where RepNo = " & Trim(txtRepNo.Text) & " And ID_ComNo = " & ComNo
    sqltext = sqltext & Chr(13) & " delete from ReparationWorks  where RepNo = " & Trim(txtRepNo.Text) & " And ID_ComNo = " & ComNo
    sqltext = sqltext & Chr(13) & " delete from Reparation       where RepNo = " & Trim(txtRepNo.Text) & " And ID_ComNo = " & ComNo
    sqltext = sqltext & Chr(13) & " update MaintCall set CallStatus = 0 where CallNo = " & Trim(txtFindCallNo.Text) & " And ID_ComNo = " & ComNo
    'MsgBox SqlText
    DeMaint.CnnMaint.Execute sqltext
    txtFindCallNo.SetFocus
    
End Sub

Private Sub CmdDeletePiece_Click()
On Error GoTo Err_handler
    If Not CmdDeletePiece.Enabled Then Exit Sub
    If RsPieces!Carried = 1 Then
        MsgBox "„«œ… „—Õ·…", vbExclamation + vbMsgBoxRight, "Œÿ√"
        Exit Sub
    End If
    Dim t As Integer
    If RsPieces.RecordCount = 0 Then Exit Sub
    t = MsgBox("Â· √‰  „ √ﬂœ ﬂ„ «·Õ–›", vbYesNo + vbQuestion + vbDefaultButton1 + vbMsgBoxRight, " ‰»ÌÂ")
    If t = vbYes Then
        RsPieces.Delete
        If Not RsPieces.RecordCount = 0 Then
            RsPieces.MoveNext
        End If
'        FormatGrdPieces
        DoEvents
        CmdEditPiece_Click
        CmdSavePiece_Click
    End If
    Exit Sub
Err_handler:
    MsgBox Err.Description
End Sub

Private Sub cmdDiscount_Click()
    If Not NewRepOp And Val(txtRepPrice.Text) > 0 Then
        Load frmdiscount
        frmdiscount.callNo = txtFindCallNo.Text
        frmdiscount.MntOrdSerNo = ""
        frmdiscount.TxtClientName = DcboClientName.Text
        frmdiscount.maxDiscount = txtRepPrice.Text
        Set frmdiscount.Cnn = DeMaint.CnnMaint
        frmdiscount.loadData
        frmdiscount.Show 1
    End If
End Sub

Private Sub CmdEditPiece_Click()
On Error GoTo Err_handler
    If Not CmdEditPiece.Enabled Then Exit Sub
    If RsPieces!Carried = 1 Then
        MsgBox "„«œ… „—Õ·…", vbExclamation + vbMsgBoxRight, "Œÿ√"
        Exit Sub
    End If
    If RsPieces.RecordCount > 0 Then LockCtls False
    DblOldTotal = Val(txtPrice.Text) * Val(txtQty.Text)
    Exit Sub
Err_handler:
    MsgBox Err.Description
End Sub

Private Sub CmdSavePiece_Click()
On Error GoTo Err_handler
    If Not CmdSavePiece.Enabled Then Exit Sub
    If Val(txtQty.Text) > Val(txtTeamPieceBalance.Text) Then 'And Format(txtRepDate.Text, "dd/mm/yyyy") > Format("21/01/2008", "dd/mm/yyyy") Then
        MsgBox "—’Ìœ «·„«œ… ·œÏ «·Ê—‘… €Ì— ﬂ«›Ì", vbExclamation + vbMsgBoxRight
        Exit Sub
    End If
    If Not DcboPieces.MatchedWithList Then
        MsgBox "Ì—ÃÏ ≈Œ Ì«— «·„«œ… »«·‘ﬂ· «·’ÕÌÕ", vbExclamation + vbMsgBoxRight, "Œÿ√"
        Exit Sub
    End If
    If Not IsNumeric(txtQty.Text) Or (txtQty.Text & "" = "") Then
        MsgBox "Ì—ÃÏ ≈œŒ«· «·ﬂ„Ì… »«·‘ﬂ· «·’ÕÌÕ", vbExclamation + vbMsgBoxRight, "Œÿ√"
        Exit Sub
    End If
    If Not IsNumeric(txtPrice.Text) Or (txtPrice.Text & "" = "") Then
        MsgBox "Ì—ÃÏ ≈œŒ«· «·ﬁÌ„… »«·‘ﬂ· «·’ÕÌÕ", vbExclamation + vbMsgBoxRight, "Œÿ√"
        Exit Sub
    End If
    
    If Val(txtAmount.Text) > 0 Then
    Call GetTotal
        DblTotalPiecesOnly = (DblTotalPiecesOnly + (Val(txtPrice.Text) * Val(RsPieces!Qty))) - DblOldTotal
        If DblTotalPiecesOnly > Val(txtAmount.Text) Then
            MsgBox "·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ≈Ã„«·Ì «·„Ê«œ «·„œŒ·… √ﬂ»— „‰ ﬁÌ„… «·ﬁ”Ì„…", vbExclamation + vbMsgBoxRight, "Œÿ√"
            Exit Sub
        End If
    End If
    
    RsPieces.Move 0
    
    RsPieces.Filter = "PieceNo = '999999'"
    If RsPieces.RecordCount > 0 Then
        RsPieces!Price = Val(txtAmount.Text) - DblTotalPiecesOnly
        RsPieces.Update
    End If
    RsPieces.Filter = "PieceNo <> '9999999999999999999999'"
    RsPieces.Move 0
    DblTotal = 0
    DblTotalPiecesOnly = 0
    DblOldTotal = 0
    GetPieces
    LockCtls True
    CmdAddPiece.SetFocus
    Exit Sub
Err_handler:
    MsgBox Err.Description
End Sub

Private Sub CmdSave_Click()
    Dim sqltext As String
    sqltext = ""
    If CmdCancelPiece.Enabled Then CmdCancelPiece_Click
    If Not NewProdOp Then
        Dim Qold As Integer
        Qold = MsgBox("≈‰ Â–Â «·ﬁ”Ì„… „ÊÃÊœ… ”«»ﬁ« Â·  —Ìœ Õ›Ÿ «· ⁄œÌ·«  «· Ì √Ã—Ì Â« ⁄·ÌÂ« ø", vbQuestion + vbYesNo + vbDefaultButton2 + vbMsgBoxRight, " ‰»ÌÂ")
        If Qold = vbNo Then
            DcboCliRecever.SetFocus
            Exit Sub
        End If
    End If
    txtProdSerNo.Text = IIf(txtProdSerNo.Text & "" = "", "", Right("00" + txtProdSerNo.Text, 3))
    If Not CheckRepEnteries Then Exit Sub
    If DcboModel.MatchedWithList Then
        If CheckModelEnteries Then
            sqltext = sqltext & " " & Chr(13) & GetSaveModelStr
        Else
            Dim t As Integer
            t = MsgBox("·‰ Ì „  Œ“Ì‰ «·„‰ Ã Â·  —Ìœ «·≈ „«„ ⁄·Ï √Ì Õ«·", vbYesNo + vbQuestion + vbMsgBoxRight + vbDefaultButton1, " ‰»ÌÂ")
            If Not (t = vbYes) Then Exit Sub
        End If
    Else
        If Not NewProdOp Then
            sqltext = sqltext & " " & Chr(13) & GetDeleteModelStr
        Else
            Dim t1 As Integer
            t1 = MsgBox("·‰ Ì „  Œ“Ì‰ «·„‰ Ã Â·  —Ìœ «·≈ „«„ ⁄·Ï √Ì Õ«·", vbYesNo + vbQuestion + vbMsgBoxRight + vbDefaultButton1, " ‰»ÌÂ")
            If Not (t1 = vbYes) Then Exit Sub
        End If
    End If
    
    sqltext = sqltext & " " & Chr(13) & GetSaveRepStr
    sqltext = sqltext & " " & Chr(13) & GetWagesDefaults
    sqltext = sqltext & " " & Chr(13) & GetUpdateCallStr
    DeMaint.CnnMaint.Execute sqltext
    If RsCallCheck!ModNo <> DcboProduct.BoundText Then
        GetProdCoRep
        If Not NewRepOp Then
            GetRepWorks
        End If
    End If
    GetPieces
    NewRepOp = False
    MsgBox " „ Õ›Ÿ «·ﬁ”Ì„…", vbInformation + vbMsgBoxRight, ""
    SSTabRep.Tab = 1
    txtSearch.SetFocus
End Sub

Private Sub CmdSearch_Click()
    FrmSearch.Show 1
End Sub

Private Sub CmdSelect_Click()
    AddRepToWorks
End Sub

Private Sub CmdUnSelect_Click()
    DeleteWork
End Sub

Private Sub DcboClientDipositName_KeyPress(KeyAscii As Integer)
    Form_KeyPress KeyAscii
End Sub

Private Sub DcboClientKind_KeyPress(KeyAscii As Integer)
    Form_KeyPress KeyAscii
End Sub

Private Sub DcboClientName_Click(Area As Integer)
    If Area <> 2 Then Exit Sub
    If DcboClientName.MatchedWithList Then
        txtFindCallNo.Text = DcboClientName.BoundText
        txtFindCallNo_GotFocus
        DoEvents
        txtFindCallNo_KeyPress 13
    End If
End Sub

Private Sub DcboClientName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        DcboClientName_Click 2
    End If
    If KeyAscii > 48 Then
        TimerFindClient.Enabled = True
        DcboClientName.Tag = 1
    End If
End Sub

Private Sub DcboCliRecever_KeyPress(KeyAscii As Integer)
    Form_KeyPress KeyAscii
End Sub

Private Sub DcboModel_Click(Area As Integer)
'    If Area = 2 Then
        If DcboModel.MatchedWithList Then
            If RsModels.State = adStateOpen And RsModels.RecordCount > 0 Then
                Dim ModNo As String
                ModNo = DcboModel.BoundText
                'ClearModelCtrls , , , True
                RsModels.MoveFirst
                RsModels.Find "ModNo = " & ModNo
                If Not (RsModels.EOF Or RsModels.BOF) Then
                    DisableModSearch = True
                    txtModStockNo.Text = RsModels!ItemNo
                    txtFindModels.Text = RsModels!ModNo
                    DisableModSearch = False
                End If
            End If
        End If
'    End If
End Sub

Private Sub DcboModel_KeyPress(KeyAscii As Integer)
    Form_KeyPress KeyAscii
End Sub

Private Sub DcboPayMethod_Change()
    If DcboPayMethod.BoundText = "1" Then
        SSFrameDip.Visible = True
    Else
        SSFrameDip.Visible = False
        txtFindClientName.Text = ""
        DcboClientKind.BoundText = ""
        DcboClientDipositName.BoundText = ""
    End If
End Sub

Private Sub DcboPayMethod_KeyPress(KeyAscii As Integer)
    Form_KeyPress KeyAscii
End Sub

Private Sub DcboPieces_Click(Area As Integer)
    If DcboPieces.MatchedWithList Then
        Select Case DcboPayMethod.BoundText
            Case 0, 1
            DcboPriceType.BoundText = 3
            Case 2
            DcboPriceType.BoundText = 1
        End Select
        GetPrice
        txtTeamPieceBalance.Text = GetTeamPieceBalance
    End If
End Sub

Private Sub DcboPieces_KeyPress(KeyAscii As Integer)
    Form_KeyPress KeyAscii
End Sub

Private Sub DcboPriceType_Change()
    If DcboPriceType.DataChanged Then GetPrice
End Sub

Private Sub DcboPriceType_Click(Area As Integer)
    GetPrice
End Sub

Private Sub DcboPriceType_KeyPress(KeyAscii As Integer)
    Form_KeyPress KeyAscii
End Sub

Private Sub DcboProduct_Change()
    ClearModelCtrls
    Select Case DcboProduct.BoundText
        Case 1, 17, 16, 11
            txtGazNo.Text = "1"
        Case Else
            txtGazNo.Text = "3"
    End Select
    If DcboProduct.MatchedWithList Then
        GetProdFamModels DcboProduct.BoundText
        GetProdCoRep
    End If
End Sub

Private Sub DcboProduct_KeyPress(KeyAscii As Integer)
    Form_KeyPress KeyAscii
End Sub

Private Sub DcboReparationStatus_KeyPress(KeyAscii As Integer)
    Form_KeyPress KeyAscii
End Sub

Private Sub DcboRepClass_Click(Area As Integer)
    Dim RsClassRep As New ADODB.Recordset, sqltext As String
    If Not DcboRepClass.MatchedWithList Then Exit Sub
    sqltext = "select RepTypeNo , RepTypeDescription from CoReparationType where "
    sqltext = sqltext & " left(ltrim(rtrim(RepTypeNo)),2) = '"
    sqltext = sqltext & IIf(RsCallCheck!ModNo <> 11, Right("0" + Trim(Str(RsCallCheck!ModNo)), 2), "01") & "'"
    sqltext = sqltext & " And RepClassNo = " & Trim(DcboRepClass.BoundText)
    If RsClassRep.State <> adStateClosed Then RsClassRep.Close
    RsClassRep.Open sqltext, DeMaint.CnnMaint, adOpenForwardOnly, adLockReadOnly, adCmdText
    If RsClassRep.RecordCount > 0 Then
        FillList lstCoReparation, RsClassRep
        lstCoReparation.Selected(0) = True
    Else
        lstCoReparation.Clear
    End If
End Sub

Private Sub DcboRepClass_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        AddRepToWorks
    Else
        Form_KeyPress KeyAscii
    End If
End Sub

Private Sub DcboTeam_KeyPress(KeyAscii As Integer)
    Form_KeyPress KeyAscii
End Sub

Private Sub DgrdPieces_Click()
On Error GoTo Err_handler
    If RsPieces.RecordCount = 0 Then Exit Sub
    RsPieces.AbsolutePosition = DgrdPieces.Row
    Exit Sub
Err_handler:
    MsgBox Err.Description
End Sub

Private Sub DgrdPieces_GotFocus()
On Error GoTo Err_handler
    If RsPieces.RecordCount = 0 Then Exit Sub
    RsPieces.AbsolutePosition = DgrdPieces.Row
    Set DcboPieces.RowSource = RsAllPieceInfo
    Exit Sub
Err_handler:
    MsgBox Err.Description
End Sub

Private Sub DgrdPieces_RowColChange()
On Error GoTo Err_handler
    If RsPieces.RecordCount = 0 Then Exit Sub
    RsPieces.AbsolutePosition = DgrdPieces.Row
    Exit Sub
Err_handler:
    MsgBox Err.Description
End Sub

Private Sub Form_Activate()

    FindClient
    ComNo = "1"
    NewRepOp = True
    LockCtls True
    txtFindCallNo.SetFocus
    MnuHelp.Visible = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Not (Me.ActiveControl Is lstCoReparation Or Me.ActiveControl Is txtSearch) Then
        If KeyCode = 13 Then SendKeys "{tab}"
    End If
End Sub
'
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 23 And Not NewRepOp Then
        If CmdCancelPiece.Enabled Then CmdCancelPiece_Click
        SSTabRep.Tab = 1
        txtSearch.SetFocus
    End If
    If KeyAscii = 19 And Not NewRepOp Then
        If CmdCancelPiece.Enabled Then CmdCancelPiece_Click
        SSTabRep.Tab = 2
        CmdAddPiece_Click
    End If
    If KeyAscii = 1 Then
        If CmdCancelPiece.Enabled Then CmdCancelPiece_Click
        SSTabRep.Tab = 0
        txtFindCallNo.SetFocus
    End If
    If KeyAscii = 17 Then
        CmdSave_Click
    End If
End Sub

Private Sub Form_Load()
On Error GoTo Err_handler
top = 0
left = 0
'RepInfo Handling
    If DeMaint.rsMaintGroupReceivers.State <> adStateClosed Then DeMaint.rsMaintGroupReceivers.Close
    DeMaint.MaintGroupReceivers
    
    If DeMaint.rspaymethod.State <> adStateClosed Then DeMaint.rspaymethod.Close
    DeMaint.paymethod

    'If DeMaint.rsMaintTeam.State <> adStateClosed Then DeMaint.rsMaintTeam.Close
    'DeMaint.MaintTeam
    
    If DeMaint.rsReparationStatus.State <> adStateClosed Then DeMaint.rsReparationStatus.Close
    DeMaint.ReparationStatus
    
    If DeMaint.rsAdhamProductFamily.State <> adStateClosed Then DeMaint.rsAdhamProductFamily.Close
    DeMaint.AdhamProductFamily
        
    If DeMaint.rsCoRepClassification.State <> adStateClosed Then DeMaint.rsCoRepClassification.Close
    DeMaint.CoRepClassification
'Pieces Handling
    If DeMaint.CnnMaint.State <> adStateClosed Then DeMaint.CnnMaint.Close
    DeMaint.CnnMaint.Open
    RsAllPieceInfo.Open "select PieceStockNo , PieceName from Pieces", DeMaint.CnnMaint
    Set DcboPieces.RowSource = RsAllPieceInfo
    
    Dim RsProduct As New ADODB.Recordset
    Set RsProduct = de.con.Execute("select prodfamno , ProdFamNameA from adhamproductfamily ")
    Set DcboProduct.RowSource = RsProduct
    DcboProduct.listField = "ProdFamNameA"
    DcboProduct.BoundColumn = "prodfamno"
    
    
    If DeMaint.rsPriceTypes.State <> adStateClosed Then DeMaint.rsPriceTypes.Close
    DeMaint.PriceTypes
    Exit Sub
Err_handler:
    MsgBox Err.Description
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Me.PopupMenu MnuHelp
    End If
End Sub

Private Sub lstCoReparation_DblClick()
    AddRepToWorks
End Sub

Private Sub lstCoReparation_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        AddRepToWorks
        txtSearch.SetFocus
    End If
End Sub

Private Sub lstSelectedReparation_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Then
        DeleteWork
    End If
End Sub

Private Sub mntRepitition_Click()
    With Cr
        .Connect = ConnectReport
        .ReportFileName = App.Path + "\mntCallRepititionLandscape.rpt"
        .DiscardSavedData = True
        .WindowState = crptMaximized
        .Action = 1
    End With
End Sub

Private Sub MnuChangeCallNo_Click()
On Error GoTo Err_handler
    If NewRepOp Then
        MsgBox "Ì—ÃÏ  Œ“Ì‰ «·≈’·«Õ √Ê·«", vbExclamation + vbMsgBoxRight, "Œÿ√"
        Exit Sub
    End If
    If txtFindCallNo.Text & "" = "" Or Not IsNumeric(txtFindCallNo.Text) Then
        MsgBox "Ì—ÃÏ ≈œŒ«· —ﬁ„ «·‘ﬂÊÏ »«·‘ﬂ· «·’ÕÌÕ", vbExclamation + vbMsgBoxRight, "Œÿ√"
        Exit Sub
    End If
    Dim NewCallNo As String
    NewCallNo = InputBox("Ì—ÃÏ ≈œŒ«· «·—ﬁ„ «·ÃœÌœ ··‘ﬂÊÏ " & txtFindCallNo.Text, " »œÌ· —ﬁ„ «·‘ﬂÊÏ")
    If NewCallNo & "" = "" Or Not IsNumeric(NewCallNo) Or _
        Trim(NewCallNo) = Trim(txtFindCallNo.Text) Then Exit Sub
    Dim sqltext As String
    Dim RsTest As New ADODB.Recordset
    sqltext = " select 1 from MaintCall where CallNo = " & NewCallNo
    Set RsTest = DeMaint.CnnMaint.Execute(sqltext)
    If RsTest.RecordCount = 0 Then
        MsgBox "—ﬁ„ «·‘ﬂÊÏ €Ì— „⁄—›", vbExclamation + vbMsgBoxRight, "Œÿ√"
        RsTest.Close
        Exit Sub
    End If
    sqltext = " select 1 from Reparation where CallNo = " & NewCallNo
    Set RsTest = DeMaint.CnnMaint.Execute(sqltext)
    If RsTest.RecordCount > 0 Then
        MsgBox "—ﬁ„ «·‘ﬂÊÏ Â–« ÌÕ ÊÌ ⁄·Ï ≈’·«Õ „”»ﬁ«", vbExclamation + vbMsgBoxRight, "Œÿ√"
        RsTest.Close
        Exit Sub
    End If
    Dim t As Integer
    t = 0
    t = MsgBox("Â· √‰  „ √ﬂœ „‰  »œÌ· —ﬁ„ «·‘ﬂÊÏ ≈·Ï «·ﬁÌ„… " & NewCallNo, vbMsgBoxRight + vbYesNo + vbDefaultButton1 + vbQuestion, " ‰»ÌÂ")
    If t = vbYes Then
        sqltext = " update Reparation Set CallNo "
        sqltext = sqltext & " = " & NewCallNo & " where CallNo = " & txtFindCallNo.Text
        DeMaint.CnnMaint.Execute sqltext
    End If
    Exit Sub
Err_handler:
    MsgBox Err.Description
End Sub

Private Sub MnuClientLastProd_Click()
    If DcboProduct.MatchedWithList Then
        If NewRepOp Then
            If RsCallCheck.State = adStateOpen And RsCallCheck.RecordCount = 1 Then
                GetClientFinalModInfo DcboProduct.BoundText, RsCallCheck!CliNo
            End If
        End If
    End If
End Sub

Private Sub mnuRepititionAll_Click()
    With Cr
        .Connect = ConnectReport
        .ReportFileName = App.Path + "\mntCallRepitition.rpt"
        .DiscardSavedData = True
        .WindowState = crptMaximized
        .Action = 1
    End With
End Sub

Private Sub SSFrameModel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Me.PopupMenu MnuHelp
    End If
End Sub

Private Sub SSFrameRepInfo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Me.PopupMenu MnuHelp
    End If
End Sub

Private Sub SSTabRep_Click(PreviousTab As Integer)
    SSFrameRep.Enabled = Not NewRepOp
    SSFramePiecesTab.Enabled = Not NewRepOp
End Sub

Private Sub SSTabRep_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Me.PopupMenu MnuHelp
    End If
End Sub

Private Sub TimerFindClient_Timer()
    If DcboClientName.Tag = 1 Then
        Screen.MousePointer = vbHourglass
        TimerFindClient.Enabled = False
        FindClient
        DcboClientName.Tag = 0
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub txtDescription_GotFocus()
    SendKeys "{end}"
End Sub

Private Sub txtDescription_LostFocus()
    txtDescription.Text = Replace(txtDescription.Text, Chr(13), "")
End Sub

Private Sub txtDescriptionFind_Change()
    Dim Rs As New ADODB.Recordset
    If Not IsNumeric(txtDescriptionFind.Text) Then Exit Sub
    If txtDescriptionFind.Text & "" = "" Then
        txtDescription.Text = ""
        Exit Sub
    End If
    Rs.Open "select RepName from adhamreparation where repnum = " & Trim(txtDescriptionFind.Text), DeMaint.CnnMaint, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Rs.RecordCount = 0 Then Exit Sub
    txtDescription.Text = Rs!RepName
End Sub

Private Sub txtFindCallNo_GotFocus()
    ClearAll
End Sub

Private Sub txtFindCallNo_KeyPress(KeyAscii As Integer)
On Error GoTo Err_handler
    DoEvents
    DoEvents
    Dim sqltext As String
    If KeyAscii = 13 Then
        ClearAll
        DoEvents
        If txtFindCallNo.Text & "" = "" Then Exit Sub
        If Not IsNumeric(txtFindCallNo.Text) Then
            MsgBox "Ì—ÃÏ ≈œŒ«· —ﬁ„ «·‘ﬂÊÏ »«·‘ﬂ· «·’ÕÌÕ", vbExclamation + vbMsgBoxRight, "Œÿ√"
            txtFindCallNo.SetFocus
            Exit Sub
        End If
        If Not IsCallExists(txtFindCallNo.Text) Then
            MsgBox "—ﬁ„ «·‘ﬂÊÏ Â–« €Ì— „ÊÃÊœ", vbExclamation + vbMsgBoxRight, "Œÿ√"
            txtFindCallNo.SetFocus
            Exit Sub
        End If
        DcboClientName.Text = GetClientName
        DcboProduct.BoundText = IIf(IsNull(RsCallCheck!ModNo), "", RsCallCheck!ModNo)
        DcboProduct.DataChanged = False
        
        DcboReparationStatus.BoundText = IIf(IsNull(RsCallCheck!CallStatus), "", RsCallCheck!CallStatus)
        GetRepInfo
    End If
    Exit Sub
Err_handler:
    MsgBox Err.Description
End Sub

Private Sub txtFindCallNo_LostFocus()
    DoEvents
    DoEvents
    DoEvents
End Sub

Private Sub txtFindClientName_Change()
    Dim Rs As New ADODB.Recordset, sqltext As String
    On Error GoTo Err_handler
    If Rs.State <> adStateClosed Then Rs.Close
    If DcboClientKind.MatchedWithList Then
        If txtFindClientName.Text & "" = "" Then Exit Sub
        If IsNumeric(txtFindClientName.Text) Then
            sqltext = " select top 30 * from adhamview7 where "
            sqltext = sqltext & " Left(adhamphon, " & Len(txtFindClientName.Text) & ") " & " = '" & Trim(txtFindClientName.Text) & "' "
            sqltext = sqltext & " and kind = " & DcboClientKind.BoundText
        Else
            sqltext = " declare @Find as varchar(50)"
            sqltext = sqltext & " set @Find= dbo.find('" & txtFindClientName.Text & "') "
            sqltext = sqltext & " select top 30 adhamno ,adhamName from adhamview7 where adhamName Like @Find  "
            sqltext = sqltext & " and kind=" & DcboClientKind.BoundText
        End If
    Else
        MsgBox "Ì—ÃÏ  ÕœÌœ ‰Ê⁄ «·“»Ê‰", vbExclamation + vbMsgBoxRight, "Œÿ√"
        DcboClientKind.SetFocus
        Exit Sub
    End If
    Rs.Open sqltext, DeMaint.CnnMaint, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Rs.RecordCount = 0 Then Exit Sub
    With DcboClientDipositName
        Set .RowSource = Rs
        .listField = "adhamName"
        .BoundColumn = "adhamnO"
    End With
    Exit Sub
Err_handler:
    MsgBox Err.Description
End Sub

Private Sub txtFindModels_Change()
    If DisableModSearch Then Exit Sub
    If Not IsNumeric(txtFindModels.Text) Or txtFindModels.Text & "" = "" Then Exit Sub
    'ClearModelCtrls True
    If Not DcboProduct.MatchedWithList Then
        MsgBox "Ì—ÃÏ  ÕœÌœ «·„‰ Ã √Ê·«", vbExclamation + vbMsgBoxRight, "Œÿ√"
        txtFindModels.Text = ""
        Exit Sub
    End If
    GetProdFamModels DcboProduct.BoundText, txtFindModels.Text
    DcboModel_Click 2
End Sub

Private Sub txtFindModelsName_Change()
    If DisableModSearch Then Exit Sub
    If txtFindModelsName.Text & "" = "" Then Exit Sub
    'ClearModelCtrls , True
    If Not DcboProduct.MatchedWithList Then
        MsgBox "Ì—ÃÏ  ÕœÌœ «·„‰ Ã √Ê·«", vbExclamation + vbMsgBoxRight, "Œÿ√"
        txtFindModelsName.Text = ""
        Exit Sub
    End If
    GetProdFamModels DcboProduct.BoundText, , txtFindModelsName.Text
End Sub

Private Sub txtFindPieceNo_Change()
If Not txtFindPieceNo.Text & "" = "" Then
    Dim sqltext As String
    If IsNumeric(txtFindPieceNo.Text) Then
        sqltext = "Select p1.PieceStockNo , p1.PieceName from Pieces p1 Where p1.PieceStockNo like '" & txtFindPieceNo.Text & "%' Order by p1.PieceStockNo "
    Else
        sqltext = "Select PieceStockNo , PieceName from Pieces p1 Where PieceName like '%" & txtFindPieceNo.Text & "%' or PieceStockNo like '%" & txtFindPieceNo.Text & "%' Order by PieceStockNo "
    End If
    If RsPiecesInfo.State <> adStateClosed Then RsPiecesInfo.Close
    RsPiecesInfo.Open sqltext, DeMaint.CnnMaint
    Set DcboPieces.RowSource = Nothing
    If RsPiecesInfo.RecordCount > 0 Then
        Set DcboPieces.RowSource = RsPiecesInfo
        DcboPieces.BoundText = RsPiecesInfo!PieceStockNo
        Select Case DcboPayMethod.BoundText
            Case 0, 1
            DcboPriceType.BoundText = 3
            Case 2
            DcboPriceType.BoundText = 1
        End Select
        GetPrice
        DcboPriceType.DataChanged = False
    Else
        DcboPieces.BoundText = ""
        DcboPriceType.BoundText = ""
        txtPrice.Text = ""
    End If
Else
    DcboPieces.BoundText = ""
    DcboPriceType.BoundText = ""
    txtPrice.Text = ""
End If
End Sub
Sub EmptyPiecesCtrls()
    txtAmount.Text = ""
    Set txtPrice.DataSource = Nothing
    Set txtQty.DataSource = Nothing
    Set DcboPieces.DataSource = Nothing
'    FormatGrdPieces
    txtFindPieceNo.Text = ""
    txtPrice.Text = ""
    txtQty.Text = ""
    DcboPieces.Text = ""
End Sub

'Sub FormatGrdPieces()
'    DoEvents
'    With DgrdPieces
'        .Clear
'        .Rows = 2
'        .Cols = 9
'        If RsPieces.State = adStateOpen Then
'            If RsPieces.RecordCount > 0 Then Set DgrdPieces.DataSource = RsPieces
'        End If
'        .TextMatrix(0, 2) = "«·—ﬁ„ «·„Œ“‰Ì"
'        .TextMatrix(0, 3) = "«·„«œ…"
'        .TextMatrix(0, 5) = "«·ﬂ„Ì…"
'        .TextMatrix(0, 6) = "«·”⁄—"
'        flexColWidth DgrdPieces, Me
'        .ColWidth(1) = 0
'        .ColWidth(4) = 0
'        .ColWidth(7) = 0
'        .ColWidth(8) = 0
'        .ColWidth(9) = 0
'    End With
'End Sub

Sub GetTotal()
    DblTotalPiecesOnly = 0
    DblTotal = 0
    With DgrdPieces
        For i = 1 To .Rows - 1
            If Not Trim(.TextMatrix(i, 2)) = "999999" Then
                DblTotalPiecesOnly = DblTotalPiecesOnly + (Val(.TextMatrix(i, 5)) * Val(.TextMatrix(i, 6)))
            End If
        Next i
    End With
End Sub
Sub GetPrice()
On Error GoTo errhnadler
    If DcboPriceType.MatchedWithList Then
        SQL = "select cliprice,distprice,dealprice , isnull(discount,0) discount , dealpriceafterdiscount , CliPriceafterdiscount , DistPriceafterdiscount  from pieces where piecestockno='" & DcboPieces.BoundText & "'and cliprice>0"
        If RsPrices.State <> adStateClosed Then RsPrices.Close
        RsPrices.Open SQL, DeMaint.CnnMaint
        If RsPrices.RecordCount = 0 Then
            DcboPriceType.Text = ""
            txtPrice.Text = ""
            Exit Sub
        End If
            Select Case DcboPriceType.BoundText
               Case 1
                  If Not IsNull(RsPrices!DealPrice) Then
                  
                     txtPrice.Text = RsPrices!dealpriceafterdiscount
                  End If
               Case 2
                  If Not IsNull(RsPrices!DistPrice) Then
                     txtPrice.Text = RsPrices!DistPriceafterdiscount
                  End If
               Case 3
                  If Not IsNull(RsPrices!CliPrice) Then
                     txtPrice.Text = RsPrices!CliPriceafterdiscount
                  End If
               Case Else
               txtPrice.Text = ""
            End Select
    End If
    Exit Sub
errhnadler:
    MsgBox "Error on getcomp"
End Sub

Sub CancelUpdateRec()
        DcboPieces.DataChanged = False
        txtQty.DataChanged = False
        txtPrice.DataChanged = False
End Sub
Sub GetPieces()
    If RsPieces.State <> adStateClosed Then RsPieces.Close
    sqltext = "select Id_ComNo , PieceNo , dbo.GetPieceName(PieceNo) as PieceName , RepNo , "
    sqltext = sqltext & " Qty , price , Notes , Carried , AccRegNoTemp  from ReparationPieces "
    sqltext = sqltext & " Where RepNo = " & Trim(txtRepNo.Text)
    sqltext = sqltext & " And ID_ComNo = " & ComNo
    RsPieces.Open sqltext, DeMaint.CnnMaint, adOpenStatic, adLockOptimistic, adCmdText
    txtAmount.Text = IIf(txtRepPrice.Text & "" = "", "", txtRepPrice.Text)
    Set txtPrice.DataSource = RsPieces
    Set txtQty.DataSource = RsPieces
    Set DcboPieces.DataSource = RsPieces
    If RsPieces.RecordCount = 0 Then Exit Sub
'    FormatGrdPieces
End Sub

Sub GetRepInfo()
    If RsRep.State <> adStateClosed Then RsRep.Close
    RsRep.Open "select * from Reparation where CallNo = " & txtFindCallNo.Text, DeMaint.CnnMaint, adOpenStatic, adLockBatchOptimistic, adCmdText
    If RsRep.RecordCount = 1 Then
        Dim qp As Integer
        qp = MsgBox("≈‰ Â–« «·≈’·«Õ „ÊÃÊœ „”»ﬁ« Â·  —Ìœ «·„ «»⁄… ø", vbQuestion + vbMsgBoxRight + vbDefaultButton2 + vbYesNo, " ‰»ÌÂ")
        If qp = vbYes Then
            FillTeams RsRep!RepDate
            FillRepCtrls
            DoEvents
            NewRepOp = False
            GetModelInfo
            GetRepWorks
            GetPieces
            'If RsRep!Carried = 1 Then
            '    CmdSave.Enabled = False
            '    CmdDelete.Enabled = False
            '    SSFrameRep.Enabled = False
            'Else
            '    CmdSave.Enabled = True
            '    CmdDelete.Enabled = True
            '    SSFrameRep.Enabled = True
            'End If
        Else
            txtFindCallNo.SetFocus
        End If
    Else
        SetRepDefaults
        GetClientFinalModInfo DcboProduct.BoundText, RsCallCheck!CliNo
        NewRepOp = True
        NewProdOp = True
    End If
End Sub

Function IsCallExists(strCallNo As String) As Boolean
    If RsCallCheck.State <> adStateClosed Then RsCallCheck.Close
    RsCallCheck.Open "select ModNo , CallStatus , CliNo , CallDateTime from MaintCall where CallNo = " & strCallNo, DeMaint.CnnMaint, adOpenForwardOnly, adLockReadOnly, adCmdText
    If RsCallCheck.RecordCount = 1 Then
        IsCallExists = True
    Else
        IsCallExists = False
    End If
End Function
Sub EmptyRepCtrls()
    DcboClientName.Text = ""
    txtRepNo.Text = ""
    txtRegestDate.Text = "__/__/__"
    txtRepDate.Text = "__/__/__"
    DcboCliRecever.Text = ""
    txtRepTimeBegin.Text = "__:__"
    txtRepTimeEnd.Text = "__:__"
    txtRepPrice.Text = ""
    DcboPayMethod.Text = ""
    DcboTeam.Text = ""
    txtTeamNo.Text = ""
    DcboReparationStatus.Text = ""
    txtVoltBefor.Text = ""
    txtVoltAfter.Text = ""
    txtNotes.Text = ""
    txtDescription.Text = ""
    txtNotesFind.Text = ""
    txtDescriptionFind.Text = ""
    DcboProduct.BoundText = ""
    'model Handling
    ClearModelCtrls
    'rep Handling
    txtSearch.Text = ""
    lstCoReparation.Clear
    lstSelectedReparation.Clear
    SSFrameRep.Enabled = False
    'Pieces Handling
    If RsPieces.State <> adStateClosed Then RsPieces.Close
    EmptyPiecesCtrls
    SSFramePiecesTab.Enabled = False
End Sub
Sub FillRepCtrls()
    'txtRepNo.Text = IIf(IsNull(RsRep!RepNo), "", RsRep!RepNo)
    txtRepNo.Text = RsRep!RepNo
    txtRegestDate.Text = IIf(IsNull(RsRep!regestdate), "__/__/__", Format(RsRep!regestdate, "dd/mm/yy"))
    txtRepDate.Text = IIf(IsNull(RsRep!RepDate), "__/__/__", Format(RsRep!RepDate, "dd/mm/yy"))
    DcboCliRecever.BoundText = IIf(IsNull(RsRep!CliRecever), "", RsRep!CliRecever)
    txtRepTimeBegin.Text = IIf(IsNull(RsRep!RepTimeBegin), "__:__", Format(RsRep!RepTimeBegin, "hh:mm"))
    txtRepTimeEnd.Text = IIf(IsNull(RsRep!RepTimeEnd), "__:__", Format(RsRep!RepTimeEnd, "hh:mm"))
    txtRepPrice.Text = IIf(IsNull(RsRep!RepPrice), "", RsRep!RepPrice)
    DcboPayMethod.BoundText = IIf(IsNull(RsRep!Cash), "", RsRep!Cash)
    DcboTeam.BoundText = IIf(IsNull(RsRep!TeamNo), "", RsRep!TeamNo)
    txtTeamNo.Text = IIf(IsNull(RsRep!TeamNo), "", RsRep!TeamNo)
    txtVoltBefor.Text = IIf(IsNull(RsRep!VoltBefor), "", RsRep!VoltBefor)
    txtVoltAfter.Text = IIf(IsNull(RsRep!VoltAfter), "", RsRep!VoltAfter)
    txtNotes.Text = IIf(IsNull(RsRep!Notes), "", Trim(RsRep!Notes))
    txtDescription.Text = IIf(IsNull(RsRep!Description), "", Trim(RsRep!Description))
    
    txtRepPrice.Locked = (RsRep!Carried = 1)
End Sub
Sub GetModelInfo()
    Dim sqltext As String
    If IsNull(RsRep!ProdNo) Then
        NewProdOp = True
        Exit Sub
    End If
    sqltext = " select top 1 * from AdhamProducts where "
    sqltext = sqltext & " ProdNo = " & RsRep!ProdNo
    If RsModel.State <> adStateClosed Then RsModel.Close
    RsModel.Open sqltext, DeMaint.CnnMaint, adOpenForwardOnly, adLockReadOnly, adCmdText
    If RsModel.RecordCount = 1 Then
        FillCurrModelInfo RsModel!ModNo
        ProdSerNo = RsModel!ProdNo
        NewProdOp = False
    Else
        ProdSerNo = ""
        NewProdOp = True
    End If
End Sub

Sub GetProdFamModels(ProdNo As String, Optional StrModNo As String, _
    Optional StrModName As String, Optional StrStkNo As String)
    Dim sqltext As String
    Dim GetDefMod As Boolean
    sqltext = "select ModNo , Symbol , Name , left(ltrim(rtrim(Symbol))+ space(15),15) + ' ' + "
    sqltext = sqltext & " ltrim(rtrim(Name)) as FullName ,  len(ltrim(rtrim(symbol))) , ItemNo from AdhamModels"
    If Not ProdNo & "" = "" Then
        sqltext = sqltext & " where FamNo  " & IIf(ProdNo = "1" Or ProdNo = "11", "in (1,11) ", " = " & ProdNo)
    End If
    If Not StrModNo & "" = "" Then
        sqltext = sqltext & " And ModNo = " & StrModNo
        GetDefMod = True
    End If
    If Not StrModName & "" = "" Then
        sqltext = sqltext & " And left(ltrim(rtrim(Symbol))+ space(15),15) + ' ' + ltrim(rtrim(Name)) like  '%" & StrModName & "%'"
        GetDefMod = True
    End If
    If Not StrStkNo & "" = "" Then
        sqltext = sqltext & " And ltrim(rtrim(ItemNo)) like  '%" & StrStkNo & "%'"
        GetDefMod = True
    End If

    DcboModel.Text = ""
    If RsModels.State <> adStateClosed Then RsModels.Close
    RsModels.Open sqltext, DeMaint.CnnMaint, adOpenStatic, adLockOptimistic, adCmdText
    Set DcboModel.RowSource = RsModels
    
    If Not GetDefMod Then Exit Sub
    If RsModels.RecordCount > 0 Then
        RsModels.MoveFirst
        DcboModel.BoundText = RsModels!ModNo
    Else
        DcboModel.Text = ""
    End If
End Sub

Sub FillCurrModelInfo(ModNo As String)
    DcboModel.BoundText = RsModel!ModNo
    If RsModels.State = adStateOpen And RsModels.RecordCount > 0 Then
        RsModels.MoveFirst
        RsModels.Find "ModNo = " & ModNo
        If Not (RsModels.EOF Or RsModels.BOF) Then
            Dim StrProdDate As String
            DisableModSearch = True
            txtModStockNo.Text = RsModels!ItemNo
            DisableModSearch = False
            txtGazNo.Text = left(Trim(RsModel!ProdBaCodeNo), 1)
            If Len(Trim(RsModel!ProdBaCodeNo)) = 16 Or Len(Trim(RsModel!ProdBaCodeNo)) = 18 Then
                txtProdSerNo.Text = Right(Trim(RsModel!ProdBaCodeNo), 3)
                txtProductionDate.Text = Format(RsModel!ProdDate, "dd/mm/yy")
            End If
            lblBarCodeNo.Caption = Trim(RsModel!ProdBaCodeNo)
            txtPurchaseDate.Text = Format(RsModel!ProdPurchaseDate, "dd/mm/yy")
        End If
    End If
End Sub

Sub SetRepDefaults()
    txtVoltBefor.Text = "190"
    txtVoltAfter.Text = "210"
    txtRegestDate.Text = Format(Date, "dd/mm/yy")
    txtRepDate.Text = Format(RsCallCheck!CallDatetime, "dd/mm/yy")
    FillTeams txtRepDate.Text
    txtRepPrice.Text = "0"
    DcboCliRecever.BoundText = "1"
    DcboPayMethod.BoundText = "0"
    DcboReparationStatus.BoundText = "1"
    
    Select Case DcboProduct.BoundText
        Case 1, 17, 16, 11
            txtGazNo.Text = "1"
        Case Else
            txtGazNo.Text = "3"
    End Select
End Sub

Sub GetClientFinalModInfo(ProdNo As String, CliNo As String)
    Dim RsCliLastMod As New ADODB.Recordset
    Dim sqltext As String
    sqltext = " select top 1 * from MntClientRep "
    sqltext = sqltext & " where ModNo = " & ProdNo & " And CliNo  = " & CliNo
    sqltext = sqltext & " Order by ProdNo desc "
    If RsCliLastMod.State <> adStateClosed Then RsCliLastMod.Close
    RsCliLastMod.Open sqltext, DeMaint.CnnMaint, adOpenForwardOnly, adLockReadOnly, adCmdText
    If RsCliLastMod.RecordCount = 0 Then Exit Sub
    Dim RsCliLastModStkNo As New ADODB.Recordset
    If RsCliLastModStkNo.State <> adStateClosed Then RsCliLastModStkNo.Close
    sqltext = "select top 1 ItemNo from AdhamModels where ModNo = " & RsCliLastMod!ModelNo
    RsCliLastModStkNo.Open sqltext, DeMaint.CnnMaint, adOpenForwardOnly, adLockReadOnly, adCmdText
    If RsCliLastModStkNo.RecordCount = 1 Then
        DisableModSearch = True
        txtModStockNo.Text = RsCliLastModStkNo!ItemNo
        DisableModSearch = False
    End If
    Dim StrProdDate As String
    DcboModel.BoundText = RsCliLastMod!ModelNo
    txtGazNo.Text = left(Trim(RsCliLastMod!ProdBaCodeNo), 1)
    If Len(Trim(RsCliLastMod!ProdBaCodeNo)) = 16 Or Len(Trim(RsCliLastMod!ProdBaCodeNo)) = 18 Then
        txtProdSerNo.Text = Right(Trim(RsCliLastMod!ProdBaCodeNo), 3)
    End If
    txtProductionDate.Text = Format(RsCliLastMod!ProdDate, "dd/mm/yy")
    lblBarCodeNo.Caption = Trim(RsCliLastMod!ProdBaCodeNo)
    txtPurchaseDate.Text = Format(RsCliLastMod!ProdPurchaseDate, "dd/mm/yy")
End Sub

Sub ClearModelCtrls(Optional KeepModNo As Boolean, _
                    Optional KeepModName As Boolean, _
                    Optional KeepModStkNo As Boolean, _
                    Optional KeepModFullName As Boolean)
    
    If Not KeepModFullName Then DcboModel.Text = ""
    DisableModSearch = True
    If Not KeepModNo Then txtFindModels.Text = ""
    If Not KeepModName Then txtFindModelsName.Text = ""
    If Not KeepModStkNo Then txtModStockNo.Text = ""
    DisableModSearch = False
    
    txtGazNo.Text = ""
    txtProductionDate.Text = "__/__/__"
    txtProdSerNo.Text = ""
    txtPurchaseDate.Text = "__/__/__"
    lblBarCodeNo.Caption = ""
End Sub

Private Sub txtFindPieceNo_LostFocus()
    If DcboPieces.MatchedWithList Then txtTeamPieceBalance.Text = GetTeamPieceBalance
End Sub

Private Sub txtModStockNo_Change()
    If DisableModSearch Then Exit Sub
    If txtModStockNo.Text & "" = "" Then Exit Sub
    'ClearModelCtrls , , True
    If Not DcboProduct.MatchedWithList Then
        MsgBox "Ì—ÃÏ  ÕœÌœ «·„‰ Ã √Ê·«", vbExclamation + vbMsgBoxRight, "Œÿ√"
        txtModStockNo.Text = ""
        Exit Sub
    End If
    GetProdFamModels DcboProduct.BoundText, , , Trim(txtModStockNo.Text)
End Sub

Private Sub txtNotes_GotFocus()
    SendKeys "{end}"
End Sub

Private Sub txtNotes_LostFocus()
    txtNotes.Text = Replace(txtNotes.Text, Chr(13), "")
End Sub

Private Sub txtNotesFind_Change()
    Dim Rs As New ADODB.Recordset
    If Not IsNumeric(txtNotesFind.Text) Then Exit Sub
    If txtNotesFind.Text & "" = "" Then
        txtNotes.Text = ""
        Exit Sub
    End If
    Rs.Open "select RepName from adhamreparation where repnum = " & Trim(txtNotesFind.Text), DeMaint.CnnMaint, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Rs.RecordCount = 0 Then Exit Sub
    txtNotes.Text = Rs!RepName
End Sub

Private Sub txtProdSerNo_GotFocus()
    If txtProdSerNo.Text & "" = "" Then
        txtProdSerNo.Text = "001"
    End If
    SendKeys "{home}+{end}"
End Sub

Private Sub txtProductionDate_Change()
    If IsDate(txtProductionDate.Text) Then
        txtPurchaseDate.Text = Format(DateAdd("m", 1, txtProductionDate.Text), "dd/mm/yy")
    End If
End Sub
Function CheckRepEnteries() As Boolean
    If DcboPayMethod.BoundText = 2 And Val(txtRepPrice.Text) > 0 Then
        MsgBox "·« Ì„ﬂ‰ √‰  ﬂÊ‰ «·ﬁ”Ì„… ÷„‰ «·ﬂ›«·… Ê «·ﬁÌ„… √ﬂ»— „‰ «·’›—", vbExclamation, "Œÿ√"
        txtRepPrice.Text = ""
        txtRepPrice.SetFocus
        CheckRepEnteries = False
        Exit Function
    End If
    If txtFindCallNo.Text & "" = "" Then
        MsgBox "Ì—ÃÏ ≈œŒ«· —ﬁ„ «·‘ﬂÊÏ »«·‘ﬂ· «·’ÕÌÕ", vbExclamation, "Œÿ√"
        txtFindCallNo.SetFocus
        CheckRepEnteries = False
        Exit Function
    End If
    If Not IsDate(txtRegestDate.Text) Then
        MsgBox "Ì—ÃÏ ≈œŒ«·  «—ÌŒ «·≈œŒ«· »«·‘ﬂ· «·’ÕÌÕ", vbExclamation, "Œÿ√"
        txtRegestDate.SetFocus
        CheckRepEnteries = False
        Exit Function
    End If
    If Not IsDate(txtRepDate.Text) Then
        MsgBox "Ì—ÃÏ ≈œŒ«·  «—ÌŒ «·≈’·«Õ »«·‘ﬂ· «·’ÕÌÕ", vbExclamation, "Œÿ√"
        txtRepDate.SetFocus
        CheckRepEnteries = False
        Exit Function
    End If
    If Not IsDate(txtRepTimeBegin.Text) Then
        MsgBox "Ì—ÃÏ ≈œŒ«· ”«⁄… »œ¡ «·≈’·«Õ »«·‘ﬂ· «·’ÕÌÕ", vbExclamation, "Œÿ√"
        txtRepTimeBegin.SetFocus
        CheckRepEnteries = False
        Exit Function
    End If
    If Not IsDate(txtRepTimeEnd.Text) Then
        MsgBox "Ì—ÃÏ ≈œŒ«· ”«⁄… «‰ Â«¡ «·≈’·«Õ »«·‘ﬂ· «·’ÕÌÕ", vbExclamation, "Œÿ√"
        txtRepTimeEnd.SetFocus
        CheckRepEnteries = False
        Exit Function
    End If
    If Not DcboPayMethod.MatchedWithList Then
        MsgBox "Ì—ÃÏ «Œ Ì«— ÿ—Ìﬁ… «·œ›⁄ »«·‘ﬂ· «·’ÕÌÕ", vbExclamation, "Œÿ√"
        DcboPayMethod.SetFocus
        CheckRepEnteries = False
        Exit Function
    End If
    If Not DcboTeam.MatchedWithList Then
        MsgBox "Ì—ÃÏ «Œ Ì«— «·ÊÕœ… »«·‘ﬂ· «·’ÕÌÕ", vbExclamation, "Œÿ√"
        DcboTeam.SetFocus
        CheckRepEnteries = False
        Exit Function
    End If
    If Not DcboReparationStatus.MatchedWithList Then
        MsgBox "Ì—ÃÏ «Œ Ì«— Õ«·… «·≈’·«Õ »«·‘ﬂ· «·’ÕÌÕ", vbExclamation, "Œÿ√"
        DcboReparationStatus.SetFocus
        CheckRepEnteries = False
        Exit Function
    End If
    If txtNotes.Text & "" = "" Then
        MsgBox "Ì—ÃÏ ﬂ «»… «·≈’·«Õ« ", vbExclamation, "Œÿ√"
        txtNotes.SetFocus
        CheckRepEnteries = False
        Exit Function
    End If
    'If DateDiff("d", txtRepDate.Text, txtRegestDate.Text) < 0 Then
    '    MsgBox "·« Ì„ﬂ‰ √‰ ÌﬂÊ‰  «—ÌŒ «·≈œŒ«· √’€— „‰  «—ÌŒ «·≈’·«Õ", vbExclamation, "Œÿ√"
    '    txtRepDate.SetFocus
    '    CheckRepEnteries = False
    '    Exit Function
    'End If
    If DateDiff("n", txtRepTimeBegin.Text, txtRepTimeEnd.Text) < 0 Then
        MsgBox "·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ Êﬁ  «·«‰ Â«¡ √’€— „‰ Êﬁ  «·»œ¡", vbExclamation, "Œÿ√"
        txtRepTimeBegin.SetFocus
        CheckRepEnteries = False
        Exit Function
    End If
    If DcboPayMethod.BoundText = "1" And Not DcboClientKind.MatchedWithList Then
        MsgBox "·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ‰Ê⁄ «·œ›⁄ –„… Ê ·« ÌÊÃœ ‰Ê⁄ “»Ê‰", vbExclamation, "Œÿ√"
        DcboClientKind.SetFocus
        CheckRepEnteries = False
        Exit Function
    End If
    If DcboPayMethod.BoundText = "1" And Not DcboClientDipositName.MatchedWithList Then
        MsgBox "·« Ì„ﬂ‰ √‰ ÌﬂÊ‰ ‰Ê⁄ «·œ›⁄ –„… Ê ·« ÌÊÃœ «”„ “»Ê‰", vbExclamation, "Œÿ√"
        DcboClientDipositName.SetFocus
        CheckRepEnteries = False
        Exit Function
    End If
    CheckRepEnteries = True
End Function

Function CheckModelEnteries() As Boolean
    If txtModStockNo.Text & "" = "" Then
        MsgBox "Ì—ÃÏ ≈œŒ«· «·—ﬁ„ «·„Œ“‰Ì ··„ÊœÌ·", vbExclamation, "Œÿ√"
        txtModStockNo.SetFocus
        CheckModelEnteries = False
        Exit Function
    End If
    If Not IsDate(txtProductionDate.Text) Then
        MsgBox "Ì—ÃÏ ≈œŒ«·  «—ÌŒ «·≈‰ «Ã", vbExclamation, "Œÿ√"
        txtProductionDate.SetFocus
        CheckModelEnteries = False
        Exit Function
    End If
    If Not IsDate(txtPurchaseDate.Text) Then
        MsgBox "Ì—ÃÏ ≈œŒ«·  «—ÌŒ «·‘—«¡", vbExclamation, "Œÿ√"
        txtPurchaseDate.SetFocus
        CheckModelEnteries = False
        Exit Function
    End If
    If txtProdSerNo.Text & "" = "" Then
        MsgBox "Ì—ÃÏ ≈œŒ«· —ﬁ„ «·„‰ Ã «·„ ”·”·", vbExclamation, "Œÿ√"
        txtProdSerNo.SetFocus
        CheckModelEnteries = False
        Exit Function
    End If
    If DateDiff("d", txtProductionDate.Text, txtPurchaseDate.Text) < 0 Then
        MsgBox "·« Ì„ﬂ‰ √‰ ÌﬂÊ‰  «—ÌŒ «·‘—«¡ √’€— „‰  «—ÌŒ «·≈‰ «Ã", vbExclamation, "Œÿ√"
        txtPurchaseDate.SetFocus
        CheckModelEnteries = False
        Exit Function
    End If
    If DateDiff("d", txtRepDate.Text, txtPurchaseDate.Text) > 0 Then
        MsgBox "·« Ì„ﬂ‰ √‰ ÌﬂÊ‰  «—ÌŒ «·‘—«¡ √ﬂ»— „‰  «—ÌŒ «·≈’·«Õ", vbExclamation, "Œÿ√"
        txtPurchaseDate.SetFocus
        CheckModelEnteries = False
        Exit Function
    End If
    If DateDiff("y", txtProductionDate.Text, txtRepDate.Text) / 365 > 30 Then
        MsgBox "·« Ì„ﬂ‰ √‰ ÌﬂÊ‰  «—ÌŒ «·≈‰ «Ã √’€— „‰  «—ÌŒ «·≈’·«Õ »√ﬂÀ— „‰ 30 ⁄«„", vbExclamation + vbMsgBoxRight, "Œÿ√"
        txtProductionDate.SetFocus
        CheckModelEnteries = False
        Exit Function
    End If
    RsModels.Find "ItemNo = '" & Trim(txtModStockNo.Text) & "'"
    If RsModels.EOF Or RsModels.BOF Then
        MsgBox "≈‰ Â–« «·—ﬁ„ «·„Œ“‰Ì €Ì— „⁄—›", vbExclamation, "Œÿ√"
        txtModStockNo.SetFocus
        CheckModelEnteries = False
        Exit Function
    End If
    CheckModelEnteries = True
End Function
Function GetSaveModelStr() As String
    Dim sqltext As String
    Dim NewProdSerNo As String
    Dim StrBarCodeNo As String
    StrBarCodeNo = Trim(txtGazNo.Text) & Trim(txtModStockNo.Text) & _
    Right(Trim(txtProductionDate.Text), 2) & Mid(Trim(txtProductionDate.Text), 4, 2) & _
    left(Trim(txtProductionDate.Text), 2) & Trim(txtProdSerNo.Text)
    If NewProdOp Then
        Dim RsNewRep As New ADODB.Recordset
        sqltext = "select isnull(Max(ProdNo),0) as MaxProdNo from AdhamProducts where Id_ComNo = " & ComNo
        If RsNewRep.State <> adStateClosed Then RsNewRep.Close
        RsNewRep.Open sqltext, DeMaint.CnnMaint, adOpenForwardOnly, adLockReadOnly, adCmdText
        NewProdSerNo = IIf(RsNewRep.RecordCount = 0, "1", Trim(Str(RsNewRep!MaxProdNo + 1)))
        sqltext = ""
        sqltext = " Insert into AdhamProducts "
        sqltext = sqltext & " ( Id_ComNo , ProdNo , ProdBaCodeNo , ProdDate , ProdPurchaseDate , ModNo ) values ( "
        sqltext = sqltext & ComNo & " , " & NewProdSerNo & " , "
        sqltext = sqltext & IIf(StrBarCodeNo & "" = "", " null ", "'" & Trim(StrBarCodeNo) & "'") & " , "
        sqltext = sqltext & IIf(Not IsDate(txtProductionDate.Text), " null ", "'" & TransDateToSql(txtProductionDate.Text) & "'") & " , "
        sqltext = sqltext & IIf(Not IsDate(txtPurchaseDate.Text), " null ", "'" & TransDateToSql(txtPurchaseDate.Text) & "'") & " , "
        sqltext = sqltext & IIf(Not DcboModel.MatchedWithList, " null ", DcboModel.BoundText) & " ) "
        ProdSerNoForRep = NewProdSerNo
    Else
        sqltext = ""
        sqltext = " update AdhamProducts set "
        sqltext = sqltext & " ProdBaCodeNo = " & IIf(StrBarCodeNo & "" = "", " null ", "'" & Trim(StrBarCodeNo) & "'") & " , "
        sqltext = sqltext & " ProdDate = " & IIf(Not IsDate(txtProductionDate.Text), " null ", "'" & TransDateToSql(txtProductionDate.Text) & "'") & " , "
        sqltext = sqltext & " ProdPurchaseDate = " & IIf(Not IsDate(txtPurchaseDate.Text), " null ", "'" & TransDateToSql(txtPurchaseDate.Text) & "'") & " , "
        sqltext = sqltext & " ModNo = " & IIf(Not DcboModel.MatchedWithList, " null ", DcboModel.BoundText)
        sqltext = sqltext & " Where ProdNo = " & ProdSerNo & " And ID_ComNo = " & ComNo
        ProdSerNoForRep = ProdSerNo
    End If
    GetSaveModelStr = sqltext
End Function
Function GetSaveRepStr() As String
    Dim sqltext As String
    If NewRepOp Then
        Dim RsNew As New ADODB.Recordset
        If RsNew.State <> adStateClosed Then RsNew.Close
        RsNew.Open "select isnull(Max(RepNo),0) as Rep from Reparation ", DeMaint.CnnMaint, adOpenForwardOnly, adLockReadOnly, adCmdText
        If RsNew.RecordCount > 0 Then
            txtRepNo.Text = RsNew!Rep + 1
        End If
        
        sqltext = " Insert into Reparation ( ID_ComNo , RepNo , CallNo , ProdNo , regestdate , RepDate , CliRecever , "
        sqltext = sqltext & " RepTimeBegin , RepTimeEnd , RepPrice , Cash , Clikind , CliNo ,  TeamNo , VoltBefor , "
        sqltext = sqltext & " VoltAfter , Notes , Description ) Values ( "
        sqltext = sqltext & ComNo & " , "
        sqltext = sqltext & Trim(txtRepNo.Text) & " , "
        sqltext = sqltext & Trim(txtFindCallNo.Text) & " , "
        sqltext = sqltext & IIf(Not ProdSerNoForRep & "" = "", ProdSerNoForRep, " null ") & " , "
        sqltext = sqltext & IIf(IsDate(txtRegestDate.Text), "'" & TransDateToSql(txtRegestDate.Text) & "'", " null ") & " , "
        sqltext = sqltext & IIf(IsDate(txtRepDate.Text), "'" & TransDateToSql(txtRepDate.Text) & "'", " null ") & " , "
        sqltext = sqltext & IIf(DcboCliRecever.MatchedWithList, DcboCliRecever.BoundText, " null ") & " , "
        sqltext = sqltext & IIf(IsDate(txtRepTimeBegin.Text), "'" & txtRepTimeBegin.Text & "'", " null ") & " , "
        sqltext = sqltext & IIf(IsDate(txtRepTimeEnd.Text), "'" & txtRepTimeEnd.Text & "'", " null ") & " , "
        sqltext = sqltext & IIf(IsNumeric(txtRepPrice.Text), Val(txtRepPrice.Text), " 0 ") & " , "
        sqltext = sqltext & IIf(DcboPayMethod.MatchedWithList, DcboPayMethod.BoundText, " null ") & " , "
        sqltext = sqltext & IIf(DcboClientKind.MatchedWithList, DcboClientKind.BoundText, " Null ") & " , "
        sqltext = sqltext & IIf(DcboClientDipositName.MatchedWithList, DcboClientDipositName.BoundText, " Null ") & " , "
        sqltext = sqltext & IIf(DcboTeam.MatchedWithList, DcboTeam.BoundText, " null ") & " , "
        sqltext = sqltext & IIf(IsNumeric(txtVoltBefor.Text), txtVoltBefor.Text, " null ") & " , "
        sqltext = sqltext & IIf(IsNumeric(txtVoltAfter.Text), txtVoltAfter.Text, " null ") & " , "
        sqltext = sqltext & IIf(Not (txtNotes.Text & "" = ""), "'" & Trim(txtNotes.Text) & "'", " null ") & " , "
        sqltext = sqltext & IIf(Not (txtDescription.Text & "" = ""), "'" & Trim(txtDescription.Text) & "'", " null ") & " )"
    Else
        sqltext = "update Reparation set "
        sqltext = sqltext & " ProdNo = " & IIf(Not ProdSerNoForRep & "" = "", ProdSerNoForRep, " null ") & " , "
        sqltext = sqltext & " regestdate = " & IIf(IsDate(txtRegestDate.Text), "'" & TransDateToSql(txtRegestDate.Text) & "'", " null ") & " , "
        sqltext = sqltext & " RepDate = " & IIf(IsDate(txtRepDate.Text), "'" & TransDateToSql(txtRepDate.Text) & "'", " null ") & " , "
        sqltext = sqltext & " CliRecever = " & IIf(DcboCliRecever.MatchedWithList, DcboCliRecever.BoundText, " null ") & " , "
        sqltext = sqltext & " RepTimeBegin = " & IIf(IsDate(txtRepTimeBegin.Text), "'" & txtRepTimeBegin.Text & "'", " null ") & " , "
        sqltext = sqltext & " RepTimeEnd = " & IIf(IsDate(txtRepTimeEnd.Text), "'" & txtRepTimeEnd.Text & "'", " null ") & " , "
        sqltext = sqltext & " RepPrice = " & IIf(IsNumeric(txtRepPrice.Text), txtRepPrice.Text, " 0 ") & " , "
        sqltext = sqltext & " Cash = " & IIf(DcboPayMethod.MatchedWithList, DcboPayMethod.BoundText, " null ") & " , "
        sqltext = sqltext & " CliKind = " & IIf(DcboClientKind.MatchedWithList, DcboClientKind.BoundText, " Null ") & " , "
        sqltext = sqltext & " CliNo = " & IIf(DcboClientDipositName.MatchedWithList, DcboClientDipositName.BoundText, " Null ") & " , "
        sqltext = sqltext & " TeamNo = " & IIf(DcboTeam.MatchedWithList, DcboTeam.BoundText, " null ") & " , "
        sqltext = sqltext & " VoltBefor = " & IIf(IsNumeric(txtVoltBefor.Text), txtVoltBefor.Text, " null ") & " , "
        sqltext = sqltext & " VoltAfter = " & IIf(IsNumeric(txtVoltAfter.Text), txtVoltAfter.Text, " null ") & " , "
        sqltext = sqltext & " Notes = " & IIf(Not (txtNotes.Text & "" = ""), "'" & Trim(txtNotes.Text) & "'", " null ") & " , "
        sqltext = sqltext & " Description = " & IIf(Not (txtDescription.Text & "" = ""), "'" & Trim(txtDescription.Text) & "'", " null ")
        sqltext = sqltext & " where RepNo = " & Trim(txtRepNo.Text) & " And ID_ComNo = " & ComNo & Chr(13)
        
        sqltext = sqltext & " Delete from ReparationWorks where RepNo = " & txtRepNo.Text
        sqltext = sqltext & " And ID_ComNo = " & ComNo
        sqltext = sqltext & " And Convert(int,left(ltrim(rtrim(RepTypeNo)),2)) <> " & DcboProduct.BoundText
    End If
    GetSaveRepStr = sqltext
End Function
Function GetDeleteModelStr() As String
    If ProdSerNo & "" = "" Then
        GetDeleteModelStr = ""
    Else
        ProdSerNoForRep = ""
        GetDeleteModelStr = " delete from AdhamProducts where ProdNo = " & ProdSerNo & " And ID_ComNo = " & ComNo
    End If
End Function
Sub GetRepWorks()
    Dim RsWorks As New ADODB.Recordset, sqltext As String
    If RsWorks.State <> adStateClosed Then RsWorks.Close
    sqltext = "select * from CoReparationType  where RepTypeNo in ( "
    sqltext = sqltext & " select RepTypeNo from ReparationWorks where RepNo = " & Trim(txtRepNo.Text) & " ) "
    RsWorks.Open sqltext, DeMaint.CnnMaint, adOpenForwardOnly, adLockReadOnly, adCmdText
    FillList lstSelectedReparation, RsWorks
End Sub
Sub GetProdCoRep()
    Dim RsCurrentProdReps As New ADODB.Recordset, sqltext As String
    If RsCurrentProdReps.State <> adStateClosed Then RsCurrentProdReps.Close
    sqltext = "select * from CoReparationType where Left(ltrim(rtrim(RepTypeNo)),2) = '"
    sqltext = sqltext & IIf(DcboProduct.BoundText <> 11, Right("0" + Trim(Str(DcboProduct.BoundText)), 2), "01") & "'"
    sqltext = sqltext & " And RepClassNo is not null "
    RsCurrentProdReps.Open sqltext, DeMaint.CnnMaint, adOpenForwardOnly, adLockReadOnly, adCmdText
    FillList lstCoReparation, RsCurrentProdReps
End Sub

Private Sub txtProductionDate_LostFocus()
    If IsDate(txtProductionDate.Text) Then
        txtProductionDate.Text = Format(txtProductionDate.Text, "dd/mm/yy")
    End If
End Sub

Private Sub txtRepPrice_GotFocus()
    SendKeys "{home}+{end}"
End Sub

Private Sub txtRepTimeBegin_GotFocus()
    SendKeys "{home}+{end}"
End Sub

Private Sub txtRepTimeEnd_GotFocus()
    SendKeys "{home}+{end}"
End Sub

Private Sub TxtSearch_Change()
    Dim RsRepFilter As New ADODB.Recordset, sqltext As String
    If Not (RsRep.State = adStateOpen) Then Exit Sub
    If IsNumeric(txtSearch.Text) Then
        sqltext = "select RepTypeNo , RepTypeDescription from CoReparationType where "
        sqltext = sqltext & " left(ltrim(rtrim(RepTypeNo)),2) = '"
        sqltext = sqltext & IIf(DcboProduct.BoundText <> 11, Right("0" + Trim(Str(DcboProduct.BoundText)), 2), "01") & "'"
        sqltext = sqltext & " And ltrim(rtrim(RepTypeNo)) like '" & Trim(txtSearch.Text) & "%'"
        sqltext = sqltext & " And RepClassNo is not null"
    Else
        sqltext = "select RepTypeNo , RepTypeDescription from CoReparationType where "
        sqltext = sqltext & " left(ltrim(rtrim(RepTypeNo)),2) = '"
        sqltext = sqltext & IIf(DcboProduct.BoundText <> 11, Right("0" + Trim(Str(DcboProduct.BoundText)), 2), "01") & "'"
        sqltext = sqltext & " And ltrim(rtrim(RepTypeDescription)) like '" & Trim(txtSearch.Text) & "%'"
        sqltext = sqltext & " And RepClassNo is not null"
    End If
    If RsRepFilter.State <> adStateClosed Then RsRepFilter.Close
    RsRepFilter.Open sqltext, DeMaint.CnnMaint, adOpenForwardOnly, adLockReadOnly, adCmdText
    If RsRepFilter.RecordCount > 0 Then
        FillList lstCoReparation, RsRepFilter
        lstCoReparation.Selected(0) = True
    Else
        lstCoReparation.Clear
    End If
End Sub

Sub FillList(ByRef Lst As ListBox, ByRef Rs As ADODB.Recordset)
    With Lst
        .Visible = False
        .Clear
        For i = 0 To Rs.RecordCount - 1
            .AddItem Rs.Fields(1).Value & ""
            .ItemData(.NewIndex) = Rs.Fields(0).Value
            Rs.MoveNext
        Next i
        .Visible = True
    End With
End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 40 Then
        lstCoReparation.SetFocus
    End If
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        AddRepToWorks
    End If
End Sub
Sub AddRepToWorks()
    Dim sqltext As String
    If Not (lstCoReparation.ListCount > 0 And lstCoReparation.ListIndex > -1) Then Exit Sub
    sqltext = "if not exists (select 1 from ReparationWorks where RepNo = " & txtRepNo.Text
    sqltext = sqltext & " And RepTypeNo = " & lstCoReparation.ItemData(lstCoReparation.ListIndex) & " ) "
    sqltext = sqltext & "Insert into ReparationWorks ( ID_ComNo , RepNo , RepTypeNo ) Values ( "
    sqltext = sqltext & ComNo & "," & txtRepNo.Text & " , "
    sqltext = sqltext & "'" & Right("0" + Trim(lstCoReparation.ItemData(lstCoReparation.ListIndex)), 5) & "' )"
    DeMaint.CnnMaint.Execute sqltext
    lstSelectedReparation.AddItem lstCoReparation.List(lstCoReparation.ListIndex)
    lstSelectedReparation.ItemData(lstSelectedReparation.NewIndex) = lstCoReparation.ItemData(lstCoReparation.ListIndex)
    txtSearch.SetFocus
    SendKeys "{home}+{end}"
End Sub
Function GetUpdateCallStr() As String
    Dim sqltext As String
    sqltext = " Update MaintCall set "
    sqltext = sqltext & " CallStatus = " & IIf(DcboReparationStatus.MatchedWithList, DcboReparationStatus.BoundText, " null ")
    sqltext = sqltext & IIf(DcboProduct.MatchedWithList, " , ModNo = " & DcboProduct.BoundText, "")
    sqltext = sqltext & " where CallNo = " & txtFindCallNo.Text & " And ID_ComNo = " & ComNo
    GetUpdateCallStr = sqltext
End Function
Function GetClientName() As String
    Dim Rs As New ADODB.Recordset
    If Rs.State <> adStateClosed Then Rs.Close
    Rs.Open "select AdhamName from AdhamView7 where AdhamNo = " & RsCallCheck!CliNo, DeMaint.CnnMaint, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Rs.RecordCount > 0 Then
        GetClientName = Rs!AdhamName
    Else
        GetClientName = ""
    End If
End Function
Sub DeleteWork()
    On Error GoTo Err_handler
    Dim sqltext As String
    If Not (lstSelectedReparation.ListCount > 0 And lstSelectedReparation.ListIndex > -1) Then Exit Sub
    sqltext = "Delete from ReparationWorks where "
    sqltext = sqltext & " ID_ComNo = " & ComNo & " And "
    sqltext = sqltext & " RepNo = " & Trim(txtRepNo.Text) & " And "
    sqltext = sqltext & " RepTypeNo = '" & Right("0" + Trim(lstSelectedReparation.ItemData(lstSelectedReparation.ListIndex)), 5) & "' "
    DeMaint.CnnMaint.Execute sqltext
    lstSelectedReparation.RemoveItem (lstSelectedReparation.ListIndex)
    Exit Sub
Err_handler:
    MsgBox Err.Description
End Sub

Private Sub txtTeamNo_Change()
    DcboTeam.BoundText = txtTeamNo.Text
End Sub

Function GetWagesDefaults() As String
    Dim sqltext As String
    If NewRepOp Then
        If Val(DcboPayMethod.BoundText) <> 2 And Val(txtRepPrice.Text) > 0 Then
            sqltext = ""
            sqltext = " Insert into ReparationPieces ( ID_ComNo , RepNo , PieceNo ,  Qty , Price , Carried ) Values ( "
            sqltext = sqltext & ComNo & " , " & Trim(txtRepNo.Text) & " , "
            sqltext = sqltext & " '999999'  , 1 , " & Trim(txtRepPrice.Text) & " , 0 )"
        Else
        End If
    Else
    End If
    GetWagesDefaults = sqltext
End Function

Private Sub txtTeamNo_GotFocus()
    SendKeys "{home}+{end}"
End Sub

Private Sub txtVoltAfter_GotFocus()
    SendKeys "{home}+{end}"
End Sub

Private Sub txtVoltBefor_GotFocus()
    SendKeys "{home}+{end}"
End Sub
Function LikeExpression(Expr As String, Optional MultiLeter As String = "%", Optional OneLetter As String = "_") As String
    Dim X As String
    X = Replace(Expr, "«", "*«*")
    X = Replace(X, "√", "*√*")
    X = Replace(X, "≈", "*≈*")
    X = Replace(X, "¬", "*¬*")
    X = Replace(X, "Ï", "*Ï*")
    X = Replace(X, "Ì", "*Ì*")
    
    X = Replace(X, "*«*", "[«√≈¬Ï]")
    X = Replace(X, "*√*", "[«√≈¬]")
    X = Replace(X, "*≈*", "[«√≈¬]")
    X = Replace(X, "*¬*", "[«√≈¬]")
    X = Replace(X, "*Ï*", "[«ÏÌ]")
    X = Replace(X, "*Ì*", "[ÏÌ]")
    
    X = Replace(X, MultiLeter, "!" & MultiLeter)
    X = Replace(X, OneLetter, "!" & OneLetter)
    'Replace each space with "_%"
    X = Replace(X, " ", OneLetter & MultiLeter)
    X = Replace(MultiLeter & X & MultiLeter, MultiLeter & MultiLeter, MultiLeter)
    X = "'" & X & "'"
    If InStr(1, X, "!" & MultiLeter) > 0 Or InStr(1, X, "!" & oneLeter) > 0 Then
        X = X & " ESCAPE '!'"
    End If
    LikeExpression = X
End Function
Sub FindClient()
    Dim Rs As New ADODB.Recordset
    Dim sqltext As String
    DoEvents
    If Not DcboClientName.Text & "" = "" Then
        If Not DcboClientName.MatchedWithList Then
            sqltext = " select top 50 left(ltrim(rtrim(str(m.CallNo)))+ '      ',6) + ' ' + "
            sqltext = sqltext & " left(ltrim(rtrim(a.AdhamPhon))+ '          ',10) + ' ' + "
            sqltext = sqltext & " ltrim(rtrim(a.Adhamname)) as ClientName , m.CallNo from AdhamView7 a , MaintCall m "
            sqltext = sqltext & " where m.CliNo = a.AdhamNo And "
            sqltext = sqltext & IIf(Not IsNumeric(DcboClientName.Text), " a.AdhamName like " & LikeExpression(DcboClientName.Text), "AdhamPhon like '" & DcboClientName.Text & "%' ")
            Set Rs = DeMaint.CnnMaint.Execute(sqltext)
            Set DcboClientName.RowSource = Rs
        End If
    End If
End Sub
Sub AddPiece()
On Error GoTo Err_handler
    If Not CmdAddPiece.Enabled Then Exit Sub
    If DcboPayMethod.BoundText = 0 And Val(txtRepPrice.Text) = 0 Then
        MsgBox "·« Ì„ﬂ‰ ≈÷«›… „Ê«œ ·ﬁ”Ì„… „œ›Ê⁄… ‰ﬁœ« ≈–« ﬂ«‰  «·ﬁÌ„… ’›—«", vbExclamation, "Œÿ√"
        Exit Sub
    End If
    RsPieces.AddNew
    RsPieces!RepNo = txtRepNo.Text
    RsPieces!ID_ComNo = ComNo
    LockCtls False
    txtFindPieceNo.SetFocus
    SendKeys "{home}+{end}"
    Exit Sub
Err_handler:
    MsgBox Err.Description
End Sub
Sub ClearAll()
On Error GoTo Err_handler
    DoEvents
    SSTabRep.Tab = 0
    
    If RsModel.State <> adStateClosed Then RsModel.Close
    If RsModels.State <> adStateClosed Then RsModels.Close
    If RsCallCheck.State <> adStateClosed Then RsCallCheck.Close
    If RsRep.State <> adStateClosed Then RsRep.Close
    
    DblTotalPiecesOnly = 0
    DblOldTotal = 0
    ProdSerNoForRep = ""
    ProdSerNo = ""
    
    NewRepOp = True
    txtRepPrice.Locked = False
    EmptyRepCtrls
    SendKeys "{home}+{end}"
    Exit Sub
Err_handler:
    MsgBox Err.Description
End Sub

Function IsRepCarried(RepNo As String) As Boolean
    Dim Rs As New Recordset
    If Rs.State <> adStateClosed Then Rs.Close
    Rs.Open "select RepNo , isnull(Carried,0) as Carried from Reparation where RepNo = " & RepNo, DeMaint.CnnMaint, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Rs.RecordCount > 0 Then
        If Rs!Carried = 1 Then
            IsRepCarried = True
        Else
            IsRepCarried = False
        End If
    Else
        IsRepCarried = False
    End If
End Function
Sub FillTeams(dt As Date)
    If DeMaint.rsMaintTeam.State <> adStateClosed Then DeMaint.rsMaintTeam.Close
    DeMaint.MaintTeam Format(dt, "dd/mm/yyyy")
    Set DcboTeam.RowSource = DeMaint.rsMaintTeam
    DcboTeam.BoundText = GetTeamNo
End Sub

Function GetTeamNo() As Integer
    Dim Rs As New ADODB.Recordset
    Set Rs = DeMaint.CnnMaint.Execute("select isnull(TeamNo,0) as TeamNo from mntMoveSummary where CallNo = " & txtFindCallNo.Text)
    If Rs.RecordCount = 1 Then
        GetTeamNo = Rs!TeamNo
    Else
        GetTeamNo = 0
    End If
End Function

Function GetTeamPieceBalance() As String
    'GetTeamPieceBalance = "10000000"
    'Exit Function
    Dim sqltext As String, Rs As New ADODB.Recordset
    If Not DcboTeam.MatchedWithList Or Not DcboPieces.MatchedWithList Then Exit Function
    sqltext = " select TeamPieceBalance from MvtTeamPiecebalanceFinal where PieceStockNo = '" & DcboPieces.BoundText & "'"
    sqltext = sqltext & " And TeamNo = " & DcboTeam.BoundText
    If Rs.State <> adStateClosed Then Rs.Close
    Rs.Open sqltext, DeMaint.CnnMaint, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Rs.RecordCount = 1 And Not IsNull(Rs!TeamPieceBalance) Then
        GetTeamPieceBalance = Str(Rs!TeamPieceBalance)
    Else
        GetTeamPieceBalance = "0"
    End If
End Function
