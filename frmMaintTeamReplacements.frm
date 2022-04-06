VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMaintTeamReplacements 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ÅÏÎÇá ÇÓãÇÁ ÑÄÓÇÁ ÇáæÑÔ æ ãÑÇÝÞíåã"
   ClientHeight    =   7365
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11685
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   7365
   ScaleWidth      =   11685
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   5235
      Left            =   60
      TabIndex        =   5
      Top             =   2100
      Visible         =   0   'False
      Width           =   5805
      _cx             =   10239
      _cy             =   9234
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
   Begin VSFlex8Ctl.VSFlexGrid FlexGrid 
      Height          =   5235
      Left            =   90
      TabIndex        =   2
      Top             =   2100
      Width           =   11595
      _cx             =   20452
      _cy             =   9234
      Appearance      =   1
      BorderStyle     =   0
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
   Begin VB.TextBox TxtTeamName 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   3720
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   1230
      Width           =   6705
   End
   Begin VB.TextBox TxtSearch 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      Height          =   525
      Left            =   1200
      RightToLeft     =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   750
      Visible         =   0   'False
      Width           =   6765
   End
   Begin VB.TextBox TxtTeamLeaderEmpNo 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   3735
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   1680
      Width           =   6705
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3210
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
            Picture         =   "frmMaintTeamReplacements.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMaintTeamReplacements.frx":26D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMaintTeamReplacements.frx":5533
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMaintTeamReplacements.frx":7CD8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMaintTeamReplacements.frx":A5D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMaintTeamReplacements.frx":CAF9
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMaintTeamReplacements.frx":F2B1
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMaintTeamReplacements.frx":11CC5
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMaintTeamReplacements.frx":14617
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMaintTeamReplacements.frx":1738B
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMaintTeamReplacements.frx":19BDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMaintTeamReplacements.frx":1CA81
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMaintTeamReplacements.frx":1F7DB
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMaintTeamReplacements.frx":2217B
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMaintTeamReplacements.frx":250DD
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMaintTeamReplacements.frx":27B05
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMaintTeamReplacements.frx":2A4BF
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMaintTeamReplacements.frx":2CDF0
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMaintTeamReplacements.frx":2F6F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMaintTeamReplacements.frx":3215E
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMaintTeamReplacements.frx":3510F
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMaintTeamReplacements.frx":37A38
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMaintTeamReplacements.frx":3A36D
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMaintTeamReplacements.frx":3CAAD
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMaintTeamReplacements.frx":3F41C
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMaintTeamReplacements.frx":41CA4
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMaintTeamReplacements.frx":43EB3
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMaintTeamReplacements.frx":46810
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMaintTeamReplacements.frx":4903A
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMaintTeamReplacements.frx":4BAEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMaintTeamReplacements.frx":4E52A
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMaintTeamReplacements.frx":51040
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMaintTeamReplacements.frx":53FDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMaintTeamReplacements.frx":56DA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMaintTeamReplacements.frx":59A20
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMaintTeamReplacements.frx":5C8B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMaintTeamReplacements.frx":5F4F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMaintTeamReplacements.frx":6209F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   690
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   11685
      _ExtentX        =   20611
      _ExtentY        =   1217
      ButtonWidth     =   1191
      ButtonHeight    =   1164
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   33
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Tag             =   "1"
                  Text            =   "ÊÞÑíÑ ááÈäß"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Tag             =   "2"
                  Text            =   "ÊÞÑíÑ ãÕÇÑíÝ ÇáÊæØíä"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   2
         EndProperty
      EndProperty
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   645
         Left            =   4530
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   0
         Width           =   6825
      End
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "ÇÓã ÇáæÑÔå"
      Height          =   195
      Left            =   10950
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   1290
      Width           =   765
   End
   Begin VB.Label LTeamNo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   9360
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   750
      Width           =   1065
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "ÑÞã ÇáæÑÔå"
      Height          =   195
      Left            =   10905
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   750
      Width           =   750
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "ÑÆíÓ ÇáæÑÔå"
      Height          =   195
      Left            =   10800
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   1680
      Width           =   840
   End
End
Attribute VB_Name = "frmMaintTeamReplacements"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim OK As Boolean, Flag As Boolean, Pos As Integer
'Dim BankRec As BankEmpType

Const colTeamNo = 1
Const ColTeamName = 2
Const ColLeaderEmpNo = 3
Const ColLeaderFullName = 4



Const ColNo = 1
Const ColName = 2

Dim IsGridContainsTeams As Boolean

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

Sub FillFormating(ByVal i As Integer, FlexGrid As VSFlexGrid)
If i = 1 Then
   
    Fs = "|>" + "ÇáÑÞã"
    Fs = Fs + "|>" + "ÇáÇÓã"

    With FlexGrid
        .FormatString = Fs
        .Cols = 3

        SetColWidths ColNo, FlexGrid
        SetColWidths ColName, FlexGrid

    End With
ElseIf i = 2 Then
    Fs = "|>" + "TeamNo."
    Fs = Fs + "|>" + "ÅÓã ÇáæÑÔå"
    Fs = Fs + "|>" + "ÑÞã ÑÆíÓ ÇáæÑÔå"
    Fs = Fs + "|>" + "ÅÓã ÑÆíÓ ÇáæÑÔå"
    With FlexGrid
        .FormatString = Fs
        .Cols = 5
        SetColWidths colTeamNo, FlexGrid
        SetColWidths ColTeamName, FlexGrid
        
        .ColWidth(ColLeaderEmpNo) = 0
        SetColWidths ColLeaderFullName, FlexGrid '= 6000 ', FlexGrid
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

Sub init()

    top = 0
    left = 0
    OK = True
    sqlText = "select m1.TeamNo , m1.teamName ,  m1.EmpNo  , e1.FullName LeaderFullName   from MaintTeam m1 Left outer join EmpFullName e1 on m1.Empno = e1.EmpNo order by m1.teamno desc "
    Set rs = de.con.Execute(sqlText)
    IsGridContainsTeams = False
    Set FlexGrid.DataSource = rs
    FillFormating 2, FlexGrid
    IsGridContainsTeams = True
    Grid.Rows = 1
    Grid.SelectionMode = flexSelectionListBox
End Sub

Private Sub FlexGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim FirstRow As Integer, LastRow As Integer, Vrow As Integer
With FlexGrid
If KeyCode = vbKeyDelete Then
    If MsgBox("åá ÃäÊ ãÊÃßÏ ãä ÚãáíÉ ÇáÍÐÝ", vbYesNo + vbDefaultButton2, "ÍÐÝ ÇáæÑÔ ÇáãÍÏÏÉ") = vbYes Then
        If .Row >= .RowSel Then
            FirstRow = .Row
            LastRow = .RowSel
        Else
            FirstRow = .RowSel
            LastRow = .Row
        End If
        For i = FirstRow To LastRow Step -1
            Vrow = i
            If .Rows = 1 Then
                'UpdateRecords i
            Else
                  'RemoveFromMvStock .TextMatrix(i, Colid)
                  If DeleteRow(FlexGrid, Vrow, colTeamNo, "MaintTeam", "TeamNo") Then
                    .RemoveItem Vrow
                End If
            End If
        Next
        .Col = colTeamNo
        .SetFocus
    End If
End If
End With
End Sub

Private Sub flexGrid_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    With FlexGrid
    OK = False
    Select Case .Col
        Case ColTeamName
            ShowEditor FlexGrid, ColTeamName, colTeamNo
        Case ColLeaderFullName
            ShowEditor FlexGrid, ColLeaderFullName, ColLeaderEmpNo

    End Select

        End With
End If
End Sub

Sub ShowEditor(FlexGrid As VSFlexGrid, colText As Integer, ColId As Integer)
With FlexGrid
    OK = False
    If .Rows > 23 Then
        TxtSearch.Move .left + .CellLeft + 220, .top + .CellTop, .CellWidth, .CellHeight
    Else
        TxtSearch.Move .left + .CellLeft, .top + .CellTop, .CellWidth, .CellHeight
    End If
    TxtSearch.Tag = .TextMatrix(.Row, ColId)
     
     TxtSearch.Text = .TextMatrix(.Row, colText)
    TxtSearch.Visible = True
    TxtSearch.SelStart = 0
    TxtSearch.SelLength = Len(TxtSearch.Text)
    TxtSearch.SetFocus
    OK = True
End With
                     
End Sub

Private Sub FlexGrid_RowColChange()
If IsGridContainsTeams Then
With FlexGrid
       
                OK = False
LTeamNo.Caption = .TextMatrix(.Row, colTeamNo)
TxtTeamName.Text = .TextMatrix(.Row, ColTeamName)
TxtTeamLeaderEmpNo.Tag = .TextMatrix(.Row, ColLeaderEmpNo)
TxtTeamLeaderEmpNo.Text = .TextMatrix(.Row, ColLeaderFullName)


                OK = True


        
    End With
End If
End Sub



Private Sub Form_Load()
init
End Sub



Private Sub Grid_RowColChange()
If Grid.Row = 0 Then Exit Sub
If Flag Then
    OK = False
    With Grid
    
            ActiveControl.Tag = .TextMatrix(.Row, ColNo)
            ActiveControl.Text = .TextMatrix(.Row, ColName)
            
    End With
    OK = True
End If
End Sub
Sub NewRec()
OK = False
LTeamNo.Caption = ""
TxtTeamName.Tag = 0
TxtTeamName.Text = ""
TxtTeamLeaderEmpNo.Tag = 0
TxtTeamLeaderEmpNo.Text = ""
TxtTeamName.SetFocus
OK = True
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
    Case 1
        NewRec
    Case 3
        Unload Me
End Select
End Sub

'Private Sub TxtAssistantEmpNo_Change()
'Search TxtAssistantEmpNo, 1, "Employee", "FirstName + ' ' + LastName", "EmpNo", True
'End Sub
'
'
'Private Sub TxtAssistantEmpNo_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then MoveCursor KeyCode, Grid
'End Sub
'
'Private Sub TxtAssistantEmpNo_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'    If Grid.Visible Then
'        OK = False
'        TxtAssistantEmpNo.Tag = Grid.TextMatrix(1, ColNo)
'        TxtAssistantEmpNo.Text = Grid.TextMatrix(1, ColName)
'        OK = True
'        Grid.Visible = False
'    End If
'    If SaveRec(Val(LTeamNo.Caption), TxtTeamName.Text, Val(TxtTeamLeaderEmpNo.Tag), Val(TxtAssistantEmpNo.Tag)) Then
'        AddToGrid TxtTeamName.Text
'    End If
'    TxtTeamName.SetFocus
'    SendKeys "{Home}+{End}"
'End If
'End Sub
Function SaveRec(TeamNo As Double, TeamName As String, LeaderEmpNo As Double)
On Error GoTo errorhandler
    sqlText = "Sp_AddOrUpdateMaintTeam " & TeamNo & ",'" & TeamName & "'," & LeaderEmpNo
    de.con.Execute (sqlText)
SaveRec = True
Exit Function
errorhandler:
SaveRec = False
MsgBox Err.Description
End Function
Function GetRow(TeamNo As Integer) As Integer
On Error GoTo errorhandler

With FlexGrid
    For i = 1 To .Rows - 1
        If (.TextMatrix(i, colTeamNo) = TeamNo) Then
            GetRow = i
            Exit Function
        End If
    Next i
GetRow = -1
End With
Exit Function
errorhandler:
GetRow = -1
MsgBox Err.Description
End Function
Sub AddToGrid(TeamName As String)
On Error GoTo errorhandler
Dim Vrow As Integer

    Dim teamInformation As TeamInfo
    teamInformation = GetTeamInformationByTeamName(TeamName)
    With FlexGrid
        Vrow = GetRow(teamInformation.TeamNo)
        If Vrow = -1 Then
            .AddItem ""
            Vrow = .Rows - 1
        End If

        .TextMatrix(Vrow, colTeamNo) = teamInformation.TeamNo
       .TextMatrix(Vrow, ColTeamName) = teamInformation.TeamName
        
        .TextMatrix(Vrow, ColLeaderEmpNo) = teamInformation.LeaderEmpNo
        .TextMatrix(Vrow, ColLeaderFullName) = teamInformation.LeaderFullName & ""

        FillFormating 2, FlexGrid
        If Not .RowIsVisible(Vrow) Then
            .TopRow = Vrow
            
        End If
        .Row = Vrow
        .Col = 0
        .ColSel = .Cols - 1
    End With

Exit Sub
errorhandler:
MsgBox Err.Description
End Sub

Private Sub TxtSearch_Change()
With FlexGrid

    Select Case .Col
        Case ColTeamName
            
            Search TxtSearch, 1, "MaintTeam", "TeamName", "TeamNo", False
        Case ColLeaderFullName
            Search TxtSearch, 1, "Employee", "FirstName + ' ' + LastName", "EmpNo", False
    End Select

        

End With
End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyCancel Then
        OK = False
        Grid.Visible = False
        TxtSearch.Visible = False
        TxtSearch.Text = ""
        OK = True
        Exit Sub
    End If
    
    If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then MoveCursor KeyCode, Grid
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


Private Sub txtSearch_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Grid.Visible Then
            With FlexGrid
                Select Case .Col
                    Case ColLeaderFullName
                        .TextMatrix(.Row, ColLeaderEmpNo) = IIf(Val(TxtSearch.Tag) = 0, IIf(Grid.SelectedRows = 0, Grid.TextMatrix(1, ColNo), Grid.TextMatrix(Grid.Row, ColNo)), TxtSearch.Tag)
                        .TextMatrix(.Row, .Col) = IIf(Val(TxtSearch.Tag) = 0, IIf(Grid.SelectedRows = 0, Grid.TextMatrix(1, ColName), Grid.TextMatrix(Grid.Row, ColName)), TxtSearch.Text)
                    Case ColTeamName
                       .TextMatrix(.Row, ColTeamName) = IIf(Val(TxtSearch.Tag) = 0, IIf(Grid.SelectedRows = 0, Grid.TextMatrix(1, ColName), Grid.TextMatrix(Grid.Row, ColName)), TxtSearch.Tag)
                End Select
                SetColWidths .Col, FlexGrid
            End With
        TxtSearch.Visible = False
        Grid.Visible = False
   Else
        If Val(TxtSearch.Tag) = 0 Then
            ClearRow FlexGrid
        End If
         TxtSearch.Visible = False
    End If
    With FlexGrid
        SaveRec Val(.TextMatrix(.Row, colTeamNo)), .TextMatrix(.Row, ColTeamName), Val(.TextMatrix(.Row, ColLeaderEmpNo))
    End With
ElseIf KeyAscii = 27 Then
    OK = False
    TxtSearch.Tag = 0
    TxtSearch.Text = ""
    TxtSearch.Visible = False
    
    OK = True
End If
Grid.Rows = 1
FlexGrid.SetFocus

End Sub
Sub ClearRow(FlexGrid As VSFlexGrid)
        With FlexGrid
        
        
                Select Case .Col
                    Case ColAssistantFullName
                        .TextMatrix(.Row, ColAssistantEmpNo) = 0
                        .Text = ""
                    Case ColLeaderFullName
                    .TextMatrix(.Row, ColLeaderEmpNo) = 0
                    .Text = ""
                    Case ColTeamName
                        .TextMatrix(.Row, .Col) = TxtSearch.Text
                End Select

        End With
End Sub

Private Sub TxtTeamLeaderEmpNo_Change()
Search TxtTeamLeaderEmpNo, 1, "Employee", "FirstName + ' ' + LastName", "EmpNo", True
End Sub


Private Sub TxtTeamLeaderEmpNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then MoveCursor KeyCode, Grid
End Sub

'Function GetAssistantEmpNo(LeaderEmpNo As Integer)
'On Error GoTo errorhandler
'Dim Rs As Recordset
'sqltext = "Select AssistantEmpNo from MaintTeam Where Empno=" & LeaderEmpNo & " and Isnull(AssistantEmpNo ,0) <> 0 "
'Set Rs = de.con.Execute(sqltext)
'If Rs.RecordCount > 0 Then
'    GetAssistantEmpNo = Rs!AssistantEmpNo
'    Exit Function
'End If
'GetAssistantEmpNo = 0
'Exit Function
'errorhandler:
'GetAssistantEmpNo = 0
'MsgBox Err.Description
'End Function

'Function GetAssistantEmpFullName(AssistantEmpNo As Integer)
'On Error GoTo errorhandler
'Dim Rs As Recordset
'sqltext = "Select firstName + ' ' + LastName as AssistantFullName from Employee  Where Empno=" & AssistantEmpNo
'Set Rs = de.con.Execute(sqltext)
'If Rs.RecordCount > 0 Then
'    GetAssistantEmpFullName = Rs!AssistantFullName
'    Exit Function
'End If
'GetAssistantEmpFullName = ""
'Exit Function
'errorhandler:
'GetAssistantEmpFullName = ""
'MsgBox Err.Description
'End Function

Function GetTeamNo(LeaderEmpNo As Integer) As Integer
On Error GoTo errorhandler

Dim rs As New ADODB.Recordset
sqlText = "Select TeamNo From MaintTeam Where EmpNo = " & LeaderEmpNo
Set rs = de.con.Execute(sqlText)
If rs.RecordCount > 0 Then
    GetTeamNo = rs!TeamNo
Else
GetTeamNo = 0
End If
Exit Function
errorhandler:
GetTeamNo = 0
End Function
Private Sub TxtTeamLeaderEmpNo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Grid.Visible Then
        OK = False
        TxtTeamLeaderEmpNo.Tag = IIf(Grid.SelectedRows = 0, Grid.TextMatrix(1, ColNo), Grid.TextMatrix(Grid.Row, ColNo))
        TxtTeamLeaderEmpNo.Text = IIf(Grid.SelectedRows = 0, Grid.TextMatrix(1, ColName), Grid.TextMatrix(Grid.Row, ColName))
        OK = True
        Grid.Visible = False
    End If
    If SaveRec(Val(LTeamNo.Caption), TxtTeamName.Text, Val(TxtTeamLeaderEmpNo.Tag)) Then
        AddToGrid TxtTeamName.Text
    End If
    TxtTeamName.SetFocus
    SendKeys "{Home}+{End}"
End If
End Sub

Private Sub TxtTeamName_Change()
Search TxtTeamName, 1, "MaintTeam", "TeamName", "TeamNo", True
End Sub

Private Sub TxtTeamName_GotFocus()
Pos = 2
End Sub

Private Sub TxtTeamName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then MoveCursor KeyCode, Grid
End Sub

Friend Function GetTeamInformationByTeamName(TeamName As String) As TeamInfo
On Error GoTo errorhandler
Dim teamInformation As TeamInfo

sqlText = "Select m1.Empno , m1.TeamNo , m1.TeamName ,  e1.FirstName + ' ' + e1.LastName LeaderTeamName  From " _
    & "MaintTeam m1 left outer join Employee e1 on m1.EmpNo = e1.EmpNo " _
    & "Where TeamName = '" & TeamName & "'"

Set rs = de.con.Execute(sqlText)
If rs.RecordCount > 0 Then
    
    teamInformation.TeamNo = IIf(IsNull(rs!TeamNo), 0, rs!TeamNo)
    teamInformation.TeamName = rs!TeamName & ""
    teamInformation.LeaderEmpNo = IIf(IsNull(rs!empNo), 0, rs!empNo)
    teamInformation.LeaderFullName = rs!LeaderTeamName & ""
    GetTeamInformationByTeamName = teamInformation
End If

Exit Function
errorhandler:
MsgBox Err.Description

End Function


Friend Function GetTeamInformationByTeamNo(TeamNo As Integer) As TeamInfo
On Error GoTo errorhandler
Dim teamInformation As TeamInfo

sqlText = "Select m1.Empno , m1.TeamName ,  e1.FirstName + ' ' + e1.LastName LeaderTeamName  From " _
    & "MaintTeam m1 left outer join Employee e1 on m1.EmpNo = e1.EmpNo " _
    & "Where TeamNo = " & TeamNo

Set rs = de.con.Execute(sqlText)
If rs.RecordCount > 0 Then
    
    teamInformation.TeamNo = IIf(IsNull(rs!empNo), 0, rs!empNo)
    teamInformation.TeamName = rs!TeamName & ""
    teamInformation.LeaderEmpNo = IIf(IsNull(rs!empNo), 0, rs!empNo)
    teamInformation.LeaderFullName = rs!LeaderTeamName & ""
    GetTeamInformationByTeamNo = teamInformation
End If

Exit Function
errorhandler:
MsgBox Err.Description

End Function
Private Sub TxtTeamName_KeyPress(KeyAscii As Integer)
Dim teamInformation As TeamInfo
If KeyAscii = 13 Then
        If Grid.Visible Then
            OK = False
            TxtTeamName.Tag = IIf(Grid.SelectedRows = 0, Grid.TextMatrix(1, ColNo), Grid.TextMatrix(Grid.Row, ColNo))
            TxtTeamName.Text = IIf(Grid.SelectedRows = 0, Grid.TextMatrix(1, ColName), Grid.TextMatrix(Grid.Row, ColName))
            LTeamNo.Caption = TxtTeamName.Tag
            teamInformation = GetTeamInformationByTeamNo(IIf(Grid.SelectedRows = 0, Grid.TextMatrix(1, ColNo), Grid.TextMatrix(Grid.Row, ColNo)))
            TxtTeamLeaderEmpNo.Tag = teamInformation.LeaderEmpNo
            TxtTeamLeaderEmpNo.Text = teamInformation.LeaderFullName
            OK = True
            Grid.Visible = False
        Else
            LTeamNo.Caption = ""
            TxtTeamLeaderEmpNo.Tag = ""
            TxtTeamLeaderEmpNo.Text = ""
        End If

    TxtTeamLeaderEmpNo.SetFocus
    SendKeys "{Home}+{End}"
End If
End Sub
