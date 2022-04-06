VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form FrmCoItems 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " —„Ì“ «·√—ﬁ«„ «·„Œ“‰Ì…"
   ClientHeight    =   7110
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10215
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   7110
   ScaleWidth      =   10215
   Begin VB.TextBox TxtItemName 
      Alignment       =   1  'Right Justify
      Height          =   405
      Left            =   60
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   60
      Width           =   9255
   End
   Begin VSFlex8Ctl.VSFlexGrid FlexGrid 
      Height          =   6525
      Left            =   30
      TabIndex        =   1
      Top             =   510
      Width           =   10155
      _cx             =   17912
      _cy             =   11509
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
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "—ﬁ„ «·„«œÂ"
      Height          =   195
      Left            =   9480
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   120
      Width           =   675
   End
End
Attribute VB_Name = "FrmCoItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Ok As Boolean, Flag As Boolean, Pos As Integer


Const ColItemNbr = 1
Const ColItemName = 2
Const ColClientPrice = 3

Const ColNo = 1
Const ColName = 2
Dim mainDataService As New MaintDataService


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
   
    fs = "|>" + "«·—ﬁ„"
    fs = fs + "|>" + "«·≈”„"

    With FlexGrid
        .FormatString = fs
        .Cols = 3

        SetColWidths ColNo, FlexGrid
        SetColWidths ColName, FlexGrid

    End With
ElseIf i = 2 Then
    fs = "|>" + "—ﬁ„ «·„«œÂ"
    fs = fs + "|>" + "≈”„ «·„«œÂ"
    fs = fs + "|>" + "”⁄— «·„” Â·ﬂ"
    With FlexGrid
        .FormatString = fs
        .Cols = 4
        
        SetColWidths ColItemNbr, FlexGrid
        SetColWidths ColItemName, FlexGrid
        SetColWidths ColClientPrice, FlexGrid
        
    End With
End If
End Sub



Sub ChangeCursor(sender As Control, Optional top As Variant = 0, Optional left As Variant = 0)

    With sender
       Grid.top = top + .top + .Height
       Grid.left = left + .left
       Grid.Width = .Width
    End With

End Sub

Sub FillGrid(Optional searchFormula As String)
On Error GoTo ErrorHandler
Dim rsItems As New ADODB.Recordset
    sqlText = "Select stkno , stkname , cliprice From CoStock"
    sqlText = sqlText & " where  id <> -1"
    
    If Not IsMissing(searchFormula) Then
        sqlText = sqlText & " and StkName like " & LikeExpression(searchFormula) & " or Stkno like" & LikeExpression(searchFormula)
    End If
    sqlText = sqlText & " order by StkNo"
    Set rsItems = de.con.Execute(sqlText)
    
    Set FlexGrid.DataSource = rsItems
    FillFormating 2, FlexGrid
Exit Sub
ErrorHandler:
MsgBox Err.Description
End Sub
Sub init()
    top = 0
    left = 0
    Ok = True
    FillGrid
    FlexGrid.Editable = flexEDKbdMouse
End Sub


Private Sub Form_Load()
init
FillGrid ""
End Sub

Private Sub TxtItemName_Change()
    FillGrid TxtItemName.Text
End Sub

Private Sub TxtItemName_GotFocus()
ChangeToArabic
End Sub

