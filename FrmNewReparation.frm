VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form FrmNewReparation 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ÅÕáÇÍ ÇáÔßÇæì"
   ClientHeight    =   7155
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   11235
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   7155
   ScaleWidth      =   11235
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   2895
      Left            =   2700
      TabIndex        =   75
      Top             =   1110
      Visible         =   0   'False
      Width           =   2895
      _cx             =   5106
      _cy             =   5106
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   5415
      Left            =   60
      TabIndex        =   31
      Top             =   690
      Width           =   11115
      _ExtentX        =   19606
      _ExtentY        =   9551
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "ÈíÇäÇÊ ÇáÅÕáÇÍ"
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "LTotal"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "LPrice"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(23)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "LUnitBalance"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1(7)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "LItemName"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label1(29)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label1(28)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label1(27)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label1(26)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label1(25)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label1(22)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label1(21)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label1(20)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label1(13)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label1(8)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label1(6)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label1(5)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "VSFlexGrid2"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "VSFlexGrid1"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "ComboPriceType"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "ComboErrorReason"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "ComboGuarantyStatus"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "comboOperationTYpe"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "comboGuarantyKind"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "TxtFees"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "TxtQty"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "TxtItemName"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "TxtDescription"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "TxtError"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).ControlCount=   30
      TabCaption(1)   =   "ÇáÈíÇäÇÊ ÇáÑÆíÓÉ"
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label1(19)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label1(18)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label1(17)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label1(16)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label1(15)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label1(14)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "LClient(2)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label1(0)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "SSFrameModel"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "TxtCallDate"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "SSFrame2"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "TxtAnotherReason"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "TxtUnExecutedReason"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "TxtAfterVolt"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "TxtBeforeVolt"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "TxtTechnicalNotes"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "TxtOtherNotes"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "txtCallNo"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).ControlCount=   18
      Begin VB.TextBox txtCallNo 
         Alignment       =   1  'Right Justify
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
         Left            =   9630
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   600
         Width           =   1365
      End
      Begin VB.TextBox TxtOtherNotes 
         Alignment       =   1  'Right Justify
         Height          =   405
         Left            =   60
         RightToLeft     =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   3750
         Width           =   5415
      End
      Begin VB.TextBox TxtTechnicalNotes 
         Alignment       =   1  'Right Justify
         Height          =   405
         Left            =   5550
         RightToLeft     =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   3750
         Width           =   5415
      End
      Begin VB.TextBox TxtBeforeVolt 
         Alignment       =   1  'Right Justify
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
         Left            =   9630
         RightToLeft     =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   4620
         Width           =   1335
      End
      Begin VB.TextBox TxtAfterVolt 
         Alignment       =   1  'Right Justify
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
         Left            =   8280
         RightToLeft     =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   4620
         Width           =   1335
      End
      Begin VB.TextBox TxtUnExecutedReason 
         Alignment       =   1  'Right Justify
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
         Left            =   5250
         RightToLeft     =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   4620
         Width           =   2985
      End
      Begin VB.TextBox TxtAnotherReason 
         Alignment       =   1  'Right Justify
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
         Left            =   2220
         RightToLeft     =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   4620
         Width           =   2985
      End
      Begin VB.TextBox TxtError 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   -70710
         RightToLeft     =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   780
         Width           =   2835
      End
      Begin VB.TextBox TxtDescription 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   -74970
         RightToLeft     =   -1  'True
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   780
         Width           =   2955
      End
      Begin VB.TextBox TxtItemName 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   -66450
         RightToLeft     =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   2940
         Width           =   2325
      End
      Begin VB.TextBox TxtQty 
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
         Left            =   -70380
         RightToLeft     =   -1  'True
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   2940
         Width           =   1305
      End
      Begin VB.TextBox TxtFees 
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
         Left            =   -66240
         RightToLeft     =   -1  'True
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   4860
         Width           =   1155
      End
      Begin MSDataListLib.DataCombo comboGuarantyKind 
         Height          =   315
         Left            =   -66390
         TabIndex        =   16
         Top             =   780
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo comboOperationTYpe 
         Height          =   315
         Left            =   -67830
         TabIndex        =   17
         Top             =   780
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo ComboGuarantyStatus 
         Height          =   315
         Left            =   -65220
         TabIndex        =   15
         Top             =   780
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo ComboErrorReason 
         Height          =   360
         Left            =   -71970
         TabIndex        =   19
         Top             =   780
         Width           =   1245
         _ExtentX        =   2196
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
      Begin MSDataListLib.DataCombo ComboPriceType 
         Height          =   360
         Left            =   -73320
         TabIndex        =   23
         Top             =   2940
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
      Begin Threed.SSFrame SSFrame2 
         Height          =   1065
         Left            =   60
         TabIndex        =   57
         Top             =   1050
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   1879
         _Version        =   131074
         ForeColor       =   8388608
         Caption         =   "ÇáÈíÇäÇÊ ÇáÃÓÇÓíÉ"
         Alignment       =   1
         Begin VB.TextBox TxtRecipient 
            Alignment       =   1  'Right Justify
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
            Left            =   30
            RightToLeft     =   -1  'True
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   540
            Width           =   2715
         End
         Begin VB.TextBox TxtCallEndKm 
            Alignment       =   1  'Right Justify
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
            Left            =   3930
            RightToLeft     =   -1  'True
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   540
            Width           =   1485
         End
         Begin VB.TextBox TxtEndHour 
            Alignment       =   1  'Right Justify
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
            Left            =   6640
            RightToLeft     =   -1  'True
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   540
            Width           =   1485
         End
         Begin VB.TextBox TxtBeginHour 
            Alignment       =   1  'Right Justify
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
            Left            =   9360
            RightToLeft     =   -1  'True
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   540
            Width           =   1485
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "ÇáÒÈæä ÇáãÓÊÞÈá"
            Height          =   195
            Index           =   12
            Left            =   1665
            RightToLeft     =   -1  'True
            TabIndex        =   61
            Top             =   330
            Width           =   1065
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "ÇáßíáæãÊÑÇÌ"
            Height          =   195
            Index           =   11
            Left            =   4620
            RightToLeft     =   -1  'True
            TabIndex        =   60
            Top             =   300
            Width           =   765
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "ÓÇ : ÇáÅäÊåÇÁ"
            Height          =   195
            Index           =   10
            Left            =   7260
            RightToLeft     =   -1  'True
            TabIndex        =   59
            Top             =   270
            Width           =   825
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "ÓÇ : ÇáÈÏÃ"
            Height          =   195
            Index           =   9
            Left            =   10170
            RightToLeft     =   -1  'True
            TabIndex        =   58
            Top             =   270
            Width           =   630
         End
      End
      Begin MSMask.MaskEdBox TxtCallDate 
         Height          =   405
         Left            =   8400
         TabIndex        =   2
         Top             =   600
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   714
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "99/99/9999"
         PromptChar      =   "_"
      End
      Begin Threed.SSFrame SSFrameModel 
         Height          =   1335
         Left            =   90
         TabIndex        =   62
         Top             =   2130
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   2355
         _Version        =   131074
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
            Left            =   1290
            MaxLength       =   1
            TabIndex        =   84
            TabStop         =   0   'False
            Top             =   360
            Width           =   795
         End
         Begin VB.TextBox txtModStockNo 
            BackColor       =   &H00FFFFFF&
            Height          =   420
            Left            =   2085
            MaxLength       =   8
            TabIndex        =   83
            TabStop         =   0   'False
            Top             =   360
            Width           =   2385
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
            Height          =   420
            Left            =   5520
            MaxLength       =   3
            TabIndex        =   82
            Top             =   360
            Width           =   1035
         End
         Begin VB.TextBox TxtFamNo 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   405
            Left            =   9840
            RightToLeft     =   -1  'True
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   360
            Width           =   1035
         End
         Begin VB.TextBox TxtModelName 
            Alignment       =   1  'Right Justify
            Height          =   405
            Left            =   6600
            RightToLeft     =   -1  'True
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   360
            Width           =   3135
         End
         Begin MSMask.MaskEdBox txtPurchaseDate 
            Height          =   405
            Left            =   8730
            TabIndex        =   85
            Top             =   810
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
            Height          =   435
            Left            =   4485
            TabIndex        =   86
            Top             =   360
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   767
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
            Caption         =   "ÇáÈÇÑßæÏ"
            Height          =   195
            Index           =   0
            Left            =   6615
            TabIndex        =   93
            Top             =   930
            Width           =   555
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
            Height          =   450
            Left            =   1260
            TabIndex        =   92
            Top             =   810
            Width           =   5355
         End
         Begin VB.Label lblFieldLabel 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "ÊÇÑíÎ ÇáÔÑÇÁ"
            Height          =   195
            Index           =   19
            Left            =   9930
            TabIndex        =   91
            Top             =   840
            Width           =   855
         End
         Begin VB.Label lblFieldLabel 
            Alignment       =   2  'Center
            Caption         =   "ÇáÛÇÒ"
            Height          =   240
            Index           =   18
            Left            =   1350
            TabIndex        =   90
            Top             =   60
            Width           =   765
         End
         Begin VB.Label lblFieldLabel 
            Alignment       =   2  'Center
            Caption         =   "ÇáãÎÒäí"
            Height          =   240
            Index           =   17
            Left            =   2145
            TabIndex        =   89
            Top             =   60
            Width           =   2355
         End
         Begin VB.Label lblFieldLabel 
            Alignment       =   2  'Center
            Caption         =   "Ê/ÇáÅäÊÇÌ"
            Height          =   225
            Index           =   16
            Left            =   4545
            TabIndex        =   88
            Top             =   60
            Width           =   975
         End
         Begin VB.Label lblFieldLabel 
            Alignment       =   2  'Center
            Caption         =   "ÊÓáÓá"
            Height          =   240
            Index           =   14
            Left            =   5550
            TabIndex        =   87
            Top             =   60
            Width           =   1005
         End
         Begin VB.Label lblFieldLabel 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "ÇáÚÇÆáÉ"
            Height          =   195
            Index           =   10
            Left            =   10395
            TabIndex        =   64
            Top             =   90
            Width           =   420
         End
         Begin VB.Label lblFieldLabel 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "ãæÏíá"
            Height          =   285
            Index           =   13
            Left            =   9180
            TabIndex        =   63
            Top             =   60
            Width           =   450
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid1 
         Height          =   1515
         Left            =   -74940
         TabIndex        =   80
         Top             =   1140
         Width           =   10875
         _cx             =   19182
         _cy             =   2672
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
      Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid2 
         Height          =   1515
         Left            =   -74940
         TabIndex        =   81
         Top             =   3360
         Width           =   10875
         _cx             =   19182
         _cy             =   2672
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
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "ÑÞã ÇáÞÓíãÉ"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   0
         Left            =   10140
         RightToLeft     =   -1  'True
         TabIndex        =   72
         Top             =   360
         Width           =   795
      End
      Begin VB.Label LClient 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "ÊÇÑíÎ ÇáÅÕáÇÍ"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   2
         Left            =   8550
         RightToLeft     =   -1  'True
         TabIndex        =   71
         Top             =   360
         Width           =   915
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "ãáÇÍÙÇÊ ÝäíÉ"
         Height          =   195
         Index           =   14
         Left            =   10020
         RightToLeft     =   -1  'True
         TabIndex        =   70
         Top             =   3510
         Width           =   900
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "ãáÇÍÙÇÊ ÃÎÑì"
         Height          =   195
         Index           =   15
         Left            =   4500
         RightToLeft     =   -1  'True
         TabIndex        =   69
         Top             =   3480
         Width           =   1020
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "ÝæáÊ ÞÈá "
         Height          =   195
         Index           =   16
         Left            =   10290
         RightToLeft     =   -1  'True
         TabIndex        =   68
         Top             =   4350
         Width           =   690
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "ÝæáÊ ÈÚÏ"
         Height          =   195
         Index           =   17
         Left            =   8940
         RightToLeft     =   -1  'True
         TabIndex        =   67
         Top             =   4350
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "ÇáÚäÇæíä ÇáÛíÑ ãäÝÐÉ"
         Height          =   195
         Index           =   18
         Left            =   6900
         RightToLeft     =   -1  'True
         TabIndex        =   66
         Top             =   4350
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "ÈíÇä ÇáÂÎÑ"
         Height          =   195
         Index           =   19
         Left            =   4560
         RightToLeft     =   -1  'True
         TabIndex        =   65
         Top             =   4350
         Width           =   660
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "äæÚ ÇáÚãáíÉ"
         Height          =   195
         Index           =   5
         Left            =   -67215
         RightToLeft     =   -1  'True
         TabIndex        =   56
         Top             =   450
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "ÇáÚØá æ ÅÕáÇÍå"
         Height          =   195
         Index           =   6
         Left            =   -68910
         RightToLeft     =   -1  'True
         TabIndex        =   55
         Top             =   450
         Width           =   1050
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "äæÚ ÇáßÝÇáÉ"
         Height          =   195
         Index           =   8
         Left            =   -66060
         RightToLeft     =   -1  'True
         TabIndex        =   54
         Top             =   450
         Width           =   750
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "ÍÇáÉ ÇáßÝÇáÉ"
         Height          =   195
         Index           =   13
         Left            =   -64860
         RightToLeft     =   -1  'True
         TabIndex        =   53
         Top             =   450
         Width           =   780
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "ÓÈÈ ÇáÚØá"
         Height          =   195
         Index           =   20
         Left            =   -71490
         RightToLeft     =   -1  'True
         TabIndex        =   52
         Top             =   450
         Width           =   750
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "ÇáÈíÇä"
         Height          =   195
         Index           =   21
         Left            =   -72360
         RightToLeft     =   -1  'True
         TabIndex        =   51
         Top             =   450
         Width           =   405
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "ÇáÑÞã ÇáãÎÒäí"
         Height          =   195
         Index           =   22
         Left            =   -65100
         RightToLeft     =   -1  'True
         TabIndex        =   50
         Top             =   2700
         Width           =   930
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "ÇáßãíÉ ÇáãÓÊåáßÉ"
         Height          =   195
         Index           =   25
         Left            =   -70170
         RightToLeft     =   -1  'True
         TabIndex        =   49
         Top             =   2670
         Width           =   1125
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "ÑÕíÏ ÇáæÍÏÉ ÇáÍÇáí"
         Height          =   195
         Index           =   26
         Left            =   -71730
         RightToLeft     =   -1  'True
         TabIndex        =   48
         Top             =   2670
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "ÇáÓÚÑ"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   27
         Left            =   -74550
         RightToLeft     =   -1  'True
         TabIndex        =   47
         Top             =   2700
         Width           =   1140
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "ÃÌæÑ ÇáÅÕáÇÍ"
         Height          =   195
         Index           =   28
         Left            =   -65025
         RightToLeft     =   -1  'True
         TabIndex        =   46
         Top             =   4860
         Width           =   885
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "ÇáÅÌãÇáí"
         Height          =   195
         Index           =   29
         Left            =   -73290
         RightToLeft     =   -1  'True
         TabIndex        =   45
         Top             =   4890
         Width           =   570
      End
      Begin VB.Label LItemName 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   405
         Left            =   -69030
         RightToLeft     =   -1  'True
         TabIndex        =   44
         Top             =   2940
         Width           =   2565
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "ÇáãÇÏÉ"
         Height          =   195
         Index           =   7
         Left            =   -66870
         RightToLeft     =   -1  'True
         TabIndex        =   43
         Top             =   2730
         Width           =   390
      End
      Begin VB.Label LUnitBalance 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   405
         Left            =   -71760
         RightToLeft     =   -1  'True
         TabIndex        =   42
         Top             =   2940
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "äãØ ÇáÓÚÑ"
         Height          =   195
         Index           =   23
         Left            =   -72615
         RightToLeft     =   -1  'True
         TabIndex        =   41
         Top             =   2700
         Width           =   675
      End
      Begin VB.Label LPrice 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   -74880
         RightToLeft     =   -1  'True
         TabIndex        =   40
         Top             =   2940
         Width           =   1515
      End
      Begin VB.Label LTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   -74880
         RightToLeft     =   -1  'True
         TabIndex        =   39
         Top             =   4920
         Width           =   1515
      End
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   615
      Left            =   60
      TabIndex        =   26
      Top             =   30
      Width           =   11145
      _ExtentX        =   19659
      _ExtentY        =   1085
      _Version        =   131074
      Begin VB.Label LModelName 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2940
         RightToLeft     =   -1  'True
         TabIndex        =   79
         Top             =   270
         Width           =   1905
      End
      Begin VB.Label LClient 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "ãæÏíá"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   6
         Left            =   4365
         RightToLeft     =   -1  'True
         TabIndex        =   78
         Top             =   60
         Width           =   405
      End
      Begin VB.Label LFamNo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   4860
         RightToLeft     =   -1  'True
         TabIndex        =   77
         Top             =   270
         Width           =   1275
      End
      Begin VB.Label LClient 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "ÇáÚÇÆáÉ"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   5
         Left            =   5715
         RightToLeft     =   -1  'True
         TabIndex        =   76
         Top             =   30
         Width           =   420
      End
      Begin VB.Label LEmployeeName 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   74
         Top             =   270
         Width           =   2835
      End
      Begin VB.Label LClient 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "ÅÓã ÇáãÏÎá"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   4
         Left            =   2160
         RightToLeft     =   -1  'True
         TabIndex        =   73
         Top             =   30
         Width           =   765
      End
      Begin VB.Label LCallNo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   9870
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   270
         Width           =   1185
      End
      Begin VB.Label LClient 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "ÑÞã ÇáÞÓíãÉ"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   3
         Left            =   10290
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   30
         Width           =   795
      End
      Begin VB.Label LClient 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "ÇáÒÈæä"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   9300
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   30
         Width           =   495
      End
      Begin VB.Label LClientName 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   7470
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   270
         Width           =   2385
      End
      Begin VB.Label LClient 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "ÊÇÑíÎ  ÇáÅÏÎÇá"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   1
         Left            =   6480
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   30
         Width           =   945
      End
      Begin VB.Label LEntryDate 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   6180
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   270
         Width           =   1275
      End
   End
   Begin Threed.SSFrame SSFrame4 
      Height          =   495
      Left            =   60
      TabIndex        =   34
      Top             =   6120
      Width           =   11085
      _ExtentX        =   19553
      _ExtentY        =   873
      _Version        =   131074
      Begin Threed.SSCommand CmdSearch 
         Height          =   435
         Left            =   1640
         TabIndex        =   100
         Top             =   30
         Width           =   1515
         _ExtentX        =   2672
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
         Caption         =   "ÈÍÜÜÜË"
      End
      Begin Threed.SSCommand CmdAdd 
         Height          =   435
         Left            =   9540
         TabIndex        =   1
         Top             =   30
         Width           =   1515
         _ExtentX        =   2672
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
         Caption         =   "ÌÏíÏ"
      End
      Begin Threed.SSCommand CmdDelete 
         Height          =   435
         Left            =   6380
         TabIndex        =   38
         Top             =   30
         Width           =   1515
         _ExtentX        =   2672
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
         Caption         =   "ÍÐÝ"
      End
      Begin Threed.SSCommand CmdCancel 
         Height          =   435
         Left            =   3220
         TabIndex        =   37
         Top             =   30
         Width           =   1515
         _ExtentX        =   2672
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
         Caption         =   "ÊÑÇÌÚ"
      End
      Begin Threed.SSCommand CmdEdit 
         Height          =   435
         Left            =   7950
         TabIndex        =   36
         Top             =   30
         Width           =   1515
         _ExtentX        =   2672
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
         Caption         =   "ÊÚÏíá"
      End
      Begin Threed.SSCommand CmdSave 
         Height          =   435
         Left            =   4800
         TabIndex        =   25
         Top             =   30
         Width           =   1515
         _ExtentX        =   2672
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
         Caption         =   "ÊÎÒíä"
      End
      Begin Threed.SSCommand CmdExit 
         Height          =   435
         Left            =   60
         TabIndex        =   35
         Top             =   30
         Width           =   1515
         _ExtentX        =   2672
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
         Caption         =   "ÎÑæÌ"
      End
   End
   Begin Threed.SSFrame SSFrame10 
      Height          =   375
      Left            =   60
      TabIndex        =   94
      Top             =   6660
      Width           =   2475
      _ExtentX        =   4366
      _ExtentY        =   661
      _Version        =   131074
      Begin VB.CommandButton CmdLast 
         Height          =   285
         Left            =   2190
         Picture         =   "FrmNewReparation.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   98
         TabStop         =   0   'False
         ToolTipText     =   "Last"
         Top             =   60
         Width           =   255
      End
      Begin VB.CommandButton CmdNext 
         Height          =   285
         Left            =   1920
         Picture         =   "FrmNewReparation.frx":0532
         Style           =   1  'Graphical
         TabIndex        =   97
         TabStop         =   0   'False
         ToolTipText     =   "Next"
         Top             =   60
         Width           =   255
      End
      Begin VB.CommandButton CmdPrevious 
         Height          =   285
         Left            =   330
         Picture         =   "FrmNewReparation.frx":062C
         Style           =   1  'Graphical
         TabIndex        =   96
         TabStop         =   0   'False
         ToolTipText     =   "Previous"
         Top             =   60
         Width           =   255
      End
      Begin VB.CommandButton CmdFirst 
         Height          =   285
         Left            =   60
         Picture         =   "FrmNewReparation.frx":0726
         Style           =   1  'Graphical
         TabIndex        =   95
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
         TabIndex        =   99
         Top             =   60
         Width           =   1305
      End
   End
End
Attribute VB_Name = "FrmNewReparation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim OK As Boolean, Flag As Boolean
Dim RsNavigator As New ADODB.Recordset
Dim Pos As Integer, RecNum   As Integer
Dim TypeRec As Boolean


Const ColNo = 1
Const ColName = 2



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
        If Pos = 11 Then
            SetColWidths ColNo, Grid
        Else
        .ColWidth(ColNo) = 0
        End If
        SetColWidths ColName, Grid
    End With
ElseIf i = 2 Then
    Fs = "|>" + ""
    Fs = Fs + "|>" + ""
    With Grid
        .FormatString = Fs
        .Cols = 3
        If Pos = 11 Then
            SetColWidths ColNo, Grid
        Else
        .ColWidth(ColNo) = 0
        End If
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
       Grid.top = sstab1.top + SSFrame2.top + .top + .Height
       Grid.left = SSFrame2.left + .left
       Grid.Width = .Width
    End With
ElseIf X = 2 Then
    With TxtAssistance
       Grid.top = .top + .Height
       Grid.left = .left
       Grid.Width = .Width
End With
ElseIf X = 3 Then
    With TxtCarNum
       Grid.top = .top + .Height
       Grid.left = .left
       Grid.Width = .Width
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
       Grid.top = SSFrameModel.top + sstab1.top + .top + .Height
       Grid.left = SSFrame2.left + .left
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
       Grid.top = sstab1.top + .top + .Height
       Grid.left = sstab1.left + .left
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
Dim Rs As New ADODB.Recordset

Sqltext = "select Id, GuarantyStatusName from CoGuaranty"
Set Rs = de.con.Execute(Sqltext)
Set ComboGuarantyStatus.RowSource = Rs
ComboGuarantyStatus.listField = "GuarantyStatusName"
ComboGuarantyStatus.BoundColumn = "Id"

Sqltext = "select Id , GuarantyTypeName from CoGuarantyType"
Set Rs = de.con.Execute(Sqltext)
Set comboGuarantyKind.RowSource = Rs
comboGuarantyKind.listField = "GuarantyTypeName"
comboGuarantyKind.BoundColumn = "Id"


Sqltext = "Select Id , OperationTypeName from CoOperationTYpe"
Set Rs = de.con.Execute(Sqltext)
Set ComboOperationType.RowSource = Rs
ComboOperationType.listField = "OperationTypeName"
ComboOperationType.BoundColumn = "Id"


Sqltext = "Select Errorreasonid  , errorreasonname from CoErrorsReason"
Set Rs = de.con.Execute(Sqltext)
Set ComboErrorReason.RowSource = Rs
ComboErrorReason.listField = "errorreasonname"
ComboErrorReason.BoundColumn = "Errorreasonid"


Dim rsCurrency As New ADODB.Recordset
Sqltext = "Select PriceNo , PriceTYpe , col   from dbo.PriceTypes where PriceNo in (1,3)"
Set rsCurrency = de.con.Execute(Sqltext)
Set ComboPriceType.RowSource = rsCurrency
ComboPriceType.listField = "PriceTYpe"
ComboPriceType.BoundColumn = "PriceNo"

End Sub

Sub init()
'ReadIniFile App.Path & "\init.txt", ";"
'ConnectString = "Provider=SQLOLEDB.1 " & ";Initial Catalog=" & DataBase & ";Data Source=" & ServerName
'If de.con.State <> adStateOpen Then de.con.Open ConnectString, "user1", GetPass
    top = 0
    left = 0
    FillCombos
    OK = True
   ' EnableControls False
   ' InitNavigator
   ' If FormLoad Then
   '     FormLoad = False
   '     MoveToRec IdLetter
   ' Else
   '     MoveNavigator 4    'Move Last
   ' End If
End Sub
Sub ClearControls()
OK = False
   ' ComboBox.BoundText = ""
    'TxtClass.Tag = 0
    'TxtClass.Text = ""
    
    TxtBenefit.Tag = 0
    TxtBenefit.Text = ""
    lClass.Caption = ""
    lClass.Tag = 0
    
    txtDescription.Tag = 0
    txtDescription.Text = ""
    TxtDate.Text = Format(Now, "dd/mm/yyyy")
    txtAmount.Text = ""
OK = True
End Sub

Private Sub Chk_Click()
If Chk.Value Then
    Chk.Caption = "ÇáãÚÇæä"
Else
    Chk.Caption = "ÑÆíÓ ÇáæÍÏÉ"
End If
End Sub

Private Sub CmdAdd_Click()
TypeRec = True
EnableCmds False, False, False, True, True, False, False, False, False
EnableControls True

End Sub

Private Sub CmdCancel_Click()
 EnableCmds True, True, True, False, False, True, True, True, True
    EnableControls False
End Sub

Private Sub CmdEdit_Click()
 TypeRec = False
        EnableCmds False, False, False, True, True, False, False, False, False
        EnableControls True
End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub CmdSave_Click()
EnableCmds True, True, True, False, False, True, True, True, True
    EnableControls False
    CmdAdd.SetFocus
End Sub

Private Sub Form_Load()
init

End Sub

Private Sub Grid_RowColChange()
If Flag Then
    OK = False
    With Grid
       Select Case Pos
        Case 1
            TxtTeamName.Tag = .TextMatrix(.Row, ColNo)
            TxtTeamName.Text = .TextMatrix(.Row, ColName)
        Case 2
            TxtAssistance.Tag = .TextMatrix(.Row, ColNo)
            TxtAssistance.Text = .TextMatrix(.Row, ColName)
        Case 3
            TxtCarNum.Tag = .TextMatrix(.Row, ColNo)
            TxtCarNum.Text = .TextMatrix(.Row, ColName)
        Case 4
            TxtRecipient.Tag = .TextMatrix(.Row, ColNo)
            TxtRecipient.Text = .TextMatrix(.Row, ColName)
        Case 5
            TxtFamNo.Tag = .TextMatrix(.Row, ColNo)
            TxtFamNo.Text = .TextMatrix(.Row, ColName)
        Case 6
            TxtModelName.Tag = .TextMatrix(.Row, ColNo)
            TxtModelName.Text = .TextMatrix(.Row, ColName)
        Case 7
            TxtUnExecutedReason.Tag = .TextMatrix(.Row, ColNo)
            TxtUnExecutedReason.Text = .TextMatrix(.Row, ColName)
        Case 8
            TxtError.Tag = .TextMatrix(.Row, ColNo)
            TxtError.Text = .TextMatrix(.Row, ColName)
        Case 9
            TxtitemName.Tag = .TextMatrix(.Row, ColNo)
            TxtitemName.Text = .TextMatrix(.Row, ColName)
       End Select
    End With
    OK = True
End If
End Sub

Private Sub TxtAssistance_GotFocus()
Pos = 2

End Sub

Private Sub TxtCarNum_GotFocus()
Pos = 3

End Sub



Sub FillLabels(CallNo As Double)
On Error GoTo ERRORHANDLER
Dim Rs As New ADODB.Recordset
Sqltext = "Select CallNo , adhamname  From MaintCall M1 INNER JOIN Adhamview7 A1 ON M1.CliNo = A1.AdhamNo Where Callno=" & CallNo
Set Rs = de.con.Execute(Sqltext)
If Rs.RecordCount > 0 Then
    LCallNo.Caption = Rs!CallNo
    LClientName.Caption = Rs!AdhamName
Else
    LCallNo.Caption = ""
    LClientName.Caption = ""
End If

Exit Sub
ERRORHANDLER:
MsgBox Err.Description

End Sub

Private Sub TxtAfterVolt_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    TxtUnExecutedReason.SetFocus
    SendKeys "{home}+{end}"
End If

End Sub

Private Sub TxtBeforeVolt_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    TxtAfterVolt.SetFocus
    SendKeys "{home}+{end}"
End If

End Sub

Private Sub TxtBeginHour_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    TxtEndHour.SetFocus
    SendKeys "{home}+{end}"
End If

End Sub

Private Sub TxtCallDate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    TxtBeginHour.SetFocus
    SendKeys "{home}+{end}"
End If
End Sub

Private Sub TxtCallEndKm_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    TxtRecipient.SetFocus
    SendKeys "{home}+{end}"
End If

End Sub

Private Sub txtCallNo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    FillLabels Val(txtCallNo.Text)
    TxtCallDate.SetFocus
    SendKeys "{home}+{end}"
End If
End Sub

Private Sub TxtEndHour_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    TxtCallEndKm.SetFocus
    SendKeys "{home}+{end}"
End If

End Sub

Private Sub TxtError_Change()
On Error GoTo ERRORHANDLER
Dim RsSearch As New ADODB.Recordset
If TxtError.Text = "" Then
    TxtError.Tag = 0
    Grid.Visible = False
    Exit Sub
End If
If OK Then
    Flag = flase
    Sqltext = "Select Id, ErrorName From CoErrorsByFamilyQty Where ErrorName Like" & LikeExpression(TxtError.Text) & " and OperationTYpeid=" & Val(ComboOperationType.BoundText) & " and FamNo=" & Val(TxtFamNo.Tag)
    Set RsSearch = de.con.Execute(Sqltext)
    If RsSearch.RecordCount > 0 Then
        Set Grid.DataSource = RsSearch
        Grid.Row = 0
        FillFormating 1
        ChangeCursor 8
        Grid.Visible = True
    Else
        TxtError.Tag = 0
        Grid.Visible = False
    End If
    Flag = True
End If
Exit Sub
ERRORHANDLER:
MsgBox Err.Description

End Sub

Private Sub TxtError_GotFocus()
Pos = 8

End Sub

Private Sub TxtError_KeyDown(KeyCode As Integer, Shift As Integer)
    Flag = True
    If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then MoveCursor KeyCode, Grid
    Flag = True

End Sub

Private Sub TxtError_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Grid.Visible Then
        OK = False
        TxtError.Tag = Grid.TextMatrix(IIf(Grid.Row = 0, 1, Grid.Row), ColNo)
        TxtError.Text = Grid.TextMatrix(IIf(Grid.Row = 0, 1, Grid.Row), ColName)
        OK = True
    ElseIf Grid.Visible = False And TxtError.Text <> "" And Val(TxtError.Tag) <> 0 Then
        ComboErrorReason.SetFocus
        ComboErrorReason.SelStart = 0
        ComboErrorReason.SelLength = Len(ComboErrorReason.Text)
        Exit Sub
    Else
        OK = False
        TxtError.Tag = 0
        TxtError.Text = ""
        OK = True
    End If
    Grid.Visible = False
End If
End Sub

Private Sub TxtFamNo_GotFocus()
Pos = 5

End Sub

Private Sub txtGazNo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtPurchaseDate.SetFocus
    txtPurchaseDate.SelStart = 0
    txtPurchaseDate.SelLength = Len(txtPurchaseDate.Text)
End If

End Sub

Private Sub TxtitemName_GotFocus()
Pos = 10
End Sub


Private Sub TxtModelName_GotFocus()
Pos = 6
End Sub

Private Sub txtModStockNo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtGazNo.SetFocus
    txtGazNo.SelStart = 0
    txtGazNo.SelLength = Len(txtGazNo.Text)
End If
End Sub

Private Sub TxtOtherNotes_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    TxtBeforeVolt.SetFocus
    SendKeys "{home}+{end}"
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
    txtModStockNo.SetFocus
    SendKeys "{home}+{end}"
End If
End Sub

Private Sub txtPurchaseDate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    TxtTechnicalNotes.SetFocus
    TxtTechnicalNotes.SelStart = 0
    TxtTechnicalNotes.SelLength = Len(TxtTechnicalNotes.Text)
End If

End Sub

Private Sub TxtRecipient_GotFocus()
Pos = 4
End Sub

Private Sub TxtTeamName_GotFocus()
Pos = 1
End Sub

Private Sub TxtRecipient_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    TxtFamNo.SetFocus
    SendKeys "{home}+{end}"
End If

End Sub

Private Sub TxtTechnicalNotes_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    TxtOtherNotes.SetFocus
    TxtOtherNotes.SelStart = 0
    TxtOtherNotes.SelLength = Len(TxtOtherNotes.Text)
End If

End Sub

Private Sub TxtUnExecutedReason_GotFocus()
Pos = 7
End Sub

Private Sub TxtUnExecutedReason_Change()
On Error GoTo ERRORHANDLER
Dim RsSearch As New ADODB.Recordset
If TxtUnExecutedReason.Text = "" Then
    TxtUnExecutedReason.Tag = 0
    Grid.Visible = False
    Exit Sub
End If
If OK Then
    Flag = flase
    Sqltext = "select id, UnExecutedReason from CoUnexecutedReason Where UnExecutedReason  Like" & LikeExpression(TxtRecipient.Text)
    Set RsSearch = de.con.Execute(Sqltext)
    If RsSearch.RecordCount > 0 Then
        Set Grid.DataSource = RsSearch
        Grid.Row = 0
        FillFormating 1
        ChangeCursor 7
        Grid.Visible = True
    Else
        TxtUnExecutedReason.Tag = 0
        Grid.Visible = False
    End If
    Flag = True
End If
Exit Sub
ERRORHANDLER:
MsgBox Err.Description
End Sub


Private Sub TxtUnExecutedReason_KeyDown(KeyCode As Integer, Shift As Integer)
    Flag = True
    If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then MoveCursor KeyCode, Grid
    Flag = True
End Sub

Private Sub TxtUnExecutedReason_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Grid.Visible Then
        OK = False
        TxtUnExecutedReason.Tag = Grid.TextMatrix(IIf(Grid.Row = 0, 1, Grid.Row), ColNo)
        TxtUnExecutedReason.Text = Grid.TextMatrix(IIf(Grid.Row = 0, 1, Grid.Row), ColName)
        OK = True
    ElseIf Grid.Visible = False And TxtUnExecutedReason.Text <> "" And Val(TxtUnExecutedReason.Tag) <> 0 Then
        TxtAnotherReason.SetFocus
        TxtAnotherReason.SelStart = 0
        TxtAnotherReason.SelLength = Len(TxtAnotherReason.Text)
        Exit Sub
    Else
        OK = False
        TxtUnExecutedReason.Tag = 0
        TxtUnExecutedReason.Text = ""
        OK = True
    End If
    TxtAnotherReason.SetFocus
    TxtAnotherReason.SelStart = 0
    TxtAnotherReason.SelLength = Len(TxtAnotherReason.Text)
    Grid.Visible = False
End If
End Sub

'--------------


Private Sub TxtFamNo_Change()
On Error GoTo ERRORHANDLER
Dim RsSearch As New ADODB.Recordset
If TxtFamNo.Text = "" Then
    TxtFamNo.Tag = 0
    Grid.Visible = False
    Exit Sub
End If
If OK Then
    Flag = flase
    Sqltext = "select ProdFamNo , ProdFamNameA   from AdhamProductFamily Where ProdFamNameA    Like" & LikeExpression(TxtFamNo.Text)
    Set RsSearch = de.con.Execute(Sqltext)
    If RsSearch.RecordCount > 0 Then
        Set Grid.DataSource = RsSearch
        Grid.Row = 0
        FillFormating 1
        ChangeCursor 5
        Grid.Visible = True
    Else
        TxtFamNo.Tag = 0
        Grid.Visible = False
    End If
    Flag = True
End If
Exit Sub
ERRORHANDLER:
MsgBox Err.Description
End Sub


Private Sub TxtFamNo_KeyDown(KeyCode As Integer, Shift As Integer)
    Flag = True
    If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then MoveCursor KeyCode, Grid
    Flag = True
End Sub

Private Sub TxtFamNo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    
    If Grid.Visible Then
        OK = False
        TxtFamNo.Tag = Grid.TextMatrix(IIf(Grid.Row = 0, 1, Grid.Row), ColNo)
        TxtFamNo.Text = Grid.TextMatrix(IIf(Grid.Row = 0, 1, Grid.Row), ColName)
        OK = True
    ElseIf Grid.Visible = False And TxtFamNo.Text <> "" And Val(TxtFamNo.Tag) <> 0 Then
        TxtModelName.SetFocus
        TxtModelName.SelStart = 0
        TxtModelName.SelLength = Len(TxtModelName.Text)
        Exit Sub
    Else
        OK = False
        TxtFamNo.Tag = 0
        TxtFamNo.Text = ""
        OK = True
    End If
    TxtModelName.SetFocus
    TxtModelName.SelStart = 0
    TxtModelName.SelLength = Len(TxtModelName.Text)
    Grid.Visible = False
End If
End Sub

'-----




Private Sub TxtModelName_Change()
On Error GoTo ERRORHANDLER
Dim RsSearch As New ADODB.Recordset
If TxtModelName.Text = "" Then
    TxtModelName.Tag = 0
    Grid.Visible = False
    Exit Sub
End If
If OK Then
    Flag = flase
    Sqltext = "select ModNo , Symbol   from adhammodels Where Symbol    Like" & LikeExpression(TxtModelName.Text)
    Set RsSearch = de.con.Execute(Sqltext)
    If RsSearch.RecordCount > 0 Then
        Set Grid.DataSource = RsSearch
        Grid.Row = 0
        FillFormating 1
        ChangeCursor 6
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


Private Sub TxtModelName_KeyDown(KeyCode As Integer, Shift As Integer)
    Flag = True
    If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then MoveCursor KeyCode, Grid
    Flag = True
End Sub

Private Sub TxtModelName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Grid.Visible Then
        OK = False
        TxtModelName.Tag = Grid.TextMatrix(IIf(Grid.Row = 0, 1, Grid.Row), ColNo)
        TxtModelName.Text = Grid.TextMatrix(IIf(Grid.Row = 0, 1, Grid.Row), ColName)
        OK = True
    ElseIf Grid.Visible = False And TxtFamNo.Text <> "" And Val(TxtFamNo.Tag) <> 0 Then
        txtProdSerNo.SetFocus
        txtProdSerNo.SelStart = 0
        txtProdSerNo.SelLength = Len(txtProdSerNo.Text)
        Exit Sub
    Else
        OK = False
        TxtModelName.Tag = 0
        TxtModelName.Text = ""
        OK = True
    End If
        txtProdSerNo.SetFocus
        txtProdSerNo.SelStart = 0
        txtProdSerNo.SelLength = Len(txtProdSerNo.Text)
        Grid.Visible = False
End If
End Sub
'----------------------------------------------




