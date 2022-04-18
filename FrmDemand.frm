VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form FrmDemand 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ÿ·»Ì«  «·ﬁÿ⁄"
   ClientHeight    =   6600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10455
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   6600
   ScaleWidth      =   10455
   Begin VB.TextBox txtDescription 
      Alignment       =   2  'Center
      DragMode        =   1  'Automatic
      Height          =   360
      Left            =   60
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   1140
      Width           =   3045
   End
   Begin VB.TextBox TxtStkName 
      Alignment       =   1  'Right Justify
      DragMode        =   1  'Automatic
      Height          =   360
      Left            =   7410
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   1140
      Width           =   2985
   End
   Begin VB.TextBox TxtQty 
      Alignment       =   2  'Center
      DragMode        =   1  'Automatic
      Height          =   360
      Left            =   3150
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   1140
      Width           =   855
   End
   Begin Threed.SSFrame SSFrame4 
      DragMode        =   1  'Automatic
      Height          =   495
      Left            =   0
      TabIndex        =   5
      Top             =   5610
      Width           =   10395
      _ExtentX        =   18336
      _ExtentY        =   873
      _Version        =   131074
      Begin Threed.SSCommand CmdPrint 
         Height          =   435
         Left            =   1350
         TabIndex        =   24
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
         Caption         =   "ÿ»«⁄Â"
      End
      Begin Threed.SSCommand CmdExit 
         Height          =   435
         Left            =   30
         TabIndex        =   10
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
      Begin Threed.SSCommand cmdSave 
         Height          =   435
         Left            =   5220
         TabIndex        =   4
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
      Begin Threed.SSCommand CmdCancel 
         Height          =   435
         Left            =   3930
         TabIndex        =   9
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
      Begin Threed.SSCommand CmdAdd 
         Height          =   435
         Left            =   9090
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
      Begin Threed.SSCommand CmdSearch 
         Height          =   435
         Left            =   2640
         TabIndex        =   7
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
      Begin Threed.SSCommand CmdEdit 
         Height          =   435
         Left            =   7800
         TabIndex        =   6
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
      Begin Threed.SSCommand CmdDelete 
         Height          =   435
         Left            =   6510
         TabIndex        =   8
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
   End
   Begin Threed.SSFrame SSFrame1 
      DragMode        =   1  'Automatic
      Height          =   795
      Left            =   30
      TabIndex        =   13
      Top             =   0
      Width           =   10395
      _ExtentX        =   18336
      _ExtentY        =   1402
      _Version        =   131074
      Begin VB.Label LDemandDate 
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
         ForeColor       =   &H00C00000&
         Height          =   345
         Left            =   7455
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   390
         Width           =   1245
      End
      Begin VB.Label LdemandId 
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
         ForeColor       =   &H00C00000&
         Height          =   345
         Left            =   9090
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   390
         Width           =   1245
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "—ﬁ„ «·ÿ·»ÌÂ"
         DragMode        =   1  'Automatic
         Height          =   195
         Index           =   5
         Left            =   9600
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   120
         Width           =   720
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   " «—ÌŒ «·ÿ·»ÌÂ"
         Height          =   195
         Index           =   1
         Left            =   7815
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   90
         Width           =   855
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      DragMode        =   1  'Automatic
      Height          =   2685
      Left            =   5040
      TabIndex        =   18
      Top             =   1560
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
   Begin Threed.SSFrame NavigatorFrame 
      DragMode        =   1  'Automatic
      Height          =   375
      Left            =   0
      TabIndex        =   25
      Top             =   6150
      Width           =   2475
      _ExtentX        =   4366
      _ExtentY        =   661
      _Version        =   131074
      Begin VB.CommandButton CmdFirst 
         Height          =   285
         Left            =   60
         Picture         =   "FrmDemand.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   29
         TabStop         =   0   'False
         ToolTipText     =   "First"
         Top             =   60
         Width           =   255
      End
      Begin VB.CommandButton CmdPrevious 
         Height          =   285
         Left            =   330
         Picture         =   "FrmDemand.frx":0532
         Style           =   1  'Graphical
         TabIndex        =   28
         TabStop         =   0   'False
         ToolTipText     =   "Previous"
         Top             =   60
         Width           =   255
      End
      Begin VB.CommandButton CmdNext 
         Height          =   285
         Left            =   1920
         Picture         =   "FrmDemand.frx":062C
         Style           =   1  'Graphical
         TabIndex        =   27
         TabStop         =   0   'False
         ToolTipText     =   "Next"
         Top             =   60
         Width           =   255
      End
      Begin VB.CommandButton CmdLast 
         Height          =   285
         Left            =   2190
         Picture         =   "FrmDemand.frx":0726
         Style           =   1  'Graphical
         TabIndex        =   26
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
         TabIndex        =   30
         Top             =   60
         Width           =   1305
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid FlexGrid 
      Height          =   4035
      Left            =   60
      TabIndex        =   31
      Top             =   1560
      Width           =   10335
      _cx             =   18230
      _cy             =   7117
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
   Begin VB.Label lblFieldLabel 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "«·„·«ÕŸ« "
      DragMode        =   1  'Automatic
      Height          =   195
      Index           =   22
      Left            =   2460
      TabIndex        =   23
      Top             =   870
      Width           =   660
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "«·—ﬁ„ «·„Œ“‰Ì"
      DragMode        =   1  'Automatic
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   8
      Left            =   9390
      TabIndex        =   22
      Top             =   870
      Width           =   930
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "«·‘—Õ"
      DragMode        =   1  'Automatic
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   10
      Left            =   6930
      TabIndex        =   21
      Top             =   870
      Width           =   420
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "«·ﬂ„Ì…"
      DragMode        =   1  'Automatic
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   11
      Left            =   3540
      TabIndex        =   20
      Top             =   870
      Width           =   405
   End
   Begin VB.Label LStkName 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      DragMode        =   1  'Automatic
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   4050
      TabIndex        =   19
      Top             =   1170
      Width           =   3315
   End
   Begin VB.Label LCount 
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
      ForeColor       =   &H000000C0&
      Height          =   270
      Left            =   9120
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   6240
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
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   15
      Left            =   9840
      TabIndex        =   11
      Top             =   6240
      Width           =   480
   End
End
Attribute VB_Name = "FrmDemand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsNavigator As New ADODB.Recordset
Dim Ok As Boolean, Flag As Boolean, Pos As Integer, RecNum  As Double, DemandState As EnumState

Dim maintDataService_ As New MaintDataService
Dim demandInfo_    As Demand




Const ColDemandId = 1
Const ColDemandDate = 2
Const ColStkId = 3
Const ColStkNo = 4
Const ColStkName = 5
Const ColQty = 6
Const colDescription = 7

Const ColNo = 1
Const ColName = 2

Sub FillFormating(ByVal i As Integer, FlexGrid As VSFlexGrid)
If i = 1 Then
   
    fs = "|>" + "«·—ﬁ„"
    fs = fs + "|>" + "«·≈”„"

    With FlexGrid
        .FormatString = fs
        .Cols = 3

        SetColWidths ColNo, FlexGrid
        SetColWidths ColName, FlexGrid

    End With
ElseIf i = 2 Then
    fs = "|>" + "—ﬁ„ «·ÿ·»ÌÂ"
    fs = fs + "|>" + " «—ÌŒ «·ÿ·»ÌÂ"
    fs = fs + "|>" + "StkId"
    fs = fs + "|>" + "«·—ﬁ„ «·„Œ“‰Ì"
    fs = fs + "|>" + "«·‘—Õ"
    fs = fs + "|>" + "«·ﬂ„ÌÂ"
    fs = fs + "|>" + "„·«ÕŸ« "
    
    With FlexGrid
        .FormatString = fs
        .Cols = 8
        
        .ColWidth(ColDemandId) = 0
        SetColWidths ColDemandDate, FlexGrid
        .ColWidth(ColStkId) = 0
        SetColWidths ColStkNo, FlexGrid
        SetColWidths ColStkName, FlexGrid
        SetColWidths ColQty, FlexGrid
        SetColWidths colDescription, FlexGrid
    End With
End If
End Sub

Sub InitNavigator()
    Set RsNavigator = maintDataService_.GetAllDemands
End Sub

Private Sub CmdAdd_Click()
Set demandInfo_ = New Demand
DemandState = EnumState.NewRecord
EnableCmds False, False, False, True, True, False
EnableControls True
ClearControls
TxtStkName.SetFocus
End Sub

Private Sub CmdDelete_Click()
On Error GoTo ErrorHandler
If MsgBox("Â· √‰  „ √ﬂœ „‰ Õ–› «·ÿ·»ÌÂ", vbYesNo + vbDefaultButton2, "Õ–›") = vbYes Then
    If demandInfo_.DemandNo <> 0 Then
        SaveDemandChanges DemandState
        'FillControlsFromSql RsNavigator
        EnableCmds True, True, True, False, False, True
        MsgBox " „ Õ–› «·ÿ·»ÌÂ »‰Ã«Õ", vbInformation, "Õ–› «·ÿ·»ÌÂ"
    End If
End If

Exit Sub
ErrorHandler:
MsgBox Err.Description
End Sub

Private Sub CmdEdit_Click()
    DemandState = EnumState.UpdateRecord
    EnableCmds False, False, False, True, True, False
    EnableControls True
    TxtStkName.SetFocus
    Sendkeys "{home}+{end}"
End Sub

Private Sub cmdSave_Click()
Dim result As DemandResult
Set result = SaveDemandChanges(DemandState)

If Not result.DemandResultStatus Then
    MsgBox result.DemandResultDescription, vbExclamation + vbMsgBoxRight, "Œÿ√ ›Ì «· Œ“Ì‰ «Ê «·Õ‹–› √Ê «·≈÷«›‹Â"
    Dim ctrl As Control
    Set ctrl = Me.GetTheLastFocusControl
    ctrl.SetFocus
    Sendkeys "{home}+{end}"
Else
    MsgBox " „ Õ›Ÿ «·ÿ·»ÌÂ »‰Ã«Õ", vbInformation, "Õ›Ÿ «·ÿ·»ÌÂ"
    EnableCmds True, True, True, False, False, True
    EnableControls False
    reparationViewModelInfo_.ReparationState = DefaultRecord
    FillControls
    CmdAdd.SetFocus
End If
End Sub



Function SaveDemandChanges(vState) As ReparationResult
    On Error GoTo ErrorHandler
    Dim result As DemandResult

    Set result = maintDataService_.SaveDemandChanges(demandInfo_, vState)
    Set SaveDemandChanges = result
    Exit Function
ErrorHandler:
    Set SaveDemandChanges = result
End Function

Private Sub Form_Load()
init
End Sub

Sub ClearControls()
Ok = False
    
LdemandId.Caption = ""
LDemandDate.Caption = Format(Date, "dd/mm/yyyy")
LCount.Caption = ""
TxtStkName.text = ""
LStkName.Caption = ""
TxtQty.text = ""
txtDescription.text = ""
FlexGrid.Rows = 1
FillFormating 2, FlexGrid

Ok = True
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


Sub init()
InitNavigator
EnableControls False
FlexGrid.Rows = 1
FillFormating 2, FlexGrid
End Sub
