VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form FrmCoZone 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " —„Ì“ «·„‰«ÿﬁ"
   ClientHeight    =   7125
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6630
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   7125
   ScaleWidth      =   6630
   Begin VB.TextBox TxtZoneName 
      Alignment       =   1  'Right Justify
      Height          =   405
      Left            =   60
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   60
      Width           =   5445
   End
   Begin VSFlex8Ctl.VSFlexGrid FlexGrid 
      Height          =   6525
      Left            =   90
      TabIndex        =   1
      Top             =   570
      Width           =   6465
      _cx             =   11404
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
      Caption         =   "≈”„ «·„‰ÿﬁÂ"
      Height          =   195
      Left            =   5730
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   90
      Width           =   795
   End
End
Attribute VB_Name = "FrmCoZone"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Ok As Boolean, Flag As Boolean, Pos As Integer
Dim oldZoneName As String

Const ColZoneNo = 1
Const ColZoneName = 2
Const ColZoneClientsCount = 3

Const ColNo = 1
Const ColName = 2
Dim mainDataService As New MaintDataService





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

Sub FillFormating(ByVal i As Integer, flexGrid As VSFlexGrid)
If i = 1 Then
   
    fs = "|>" + "«·—ﬁ„"
    fs = fs + "|>" + "«·≈”„"

    With flexGrid
        .FormatString = fs
        .Cols = 3

        SetColWidths ColNo, flexGrid
        SetColWidths ColName, flexGrid

    End With
ElseIf i = 2 Then
    fs = "|>" + "ZoneNo"
    fs = fs + "|>" + "«·„‰ŸﬁÂ"
    fs = fs + "|>" + "⁄œœ «·“»«∆‰ ›Ì «·„‰ŸﬁÂ"
    With flexGrid
        .FormatString = fs
        .Cols = 4
        
        .ColWidth(ColZoneNo) = 0
        SetColWidths ColZoneName, flexGrid
        SetColWidths ColZoneClientsCount, flexGrid
        
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
Dim rsZone As New ADODB.Recordset
    sqlText = "select ZoneNo , ZoneName , (select Count(*) from adhamview7 a1 where c1.zoneno = a1.zone ) from CoZone c1 "
    sqlText = sqlText & " where  ZoneNo <> -1"
    
    If Not IsMissing(searchFormula) Then
        sqlText = sqlText & " and ZoneName like " & LikeExpression(searchFormula)
    End If
    sqlText = sqlText & " order by ZoneNo"
    Set rsZone = de.con.Execute(sqlText)
    
    Set flexGrid.DataSource = rsZone
    FillFormating 2, flexGrid
Exit Sub
ErrorHandler:
MsgBox Err.Description
End Sub
Sub init()
    top = 0
    left = 0
    Ok = True
    FillGrid
    flexGrid.Editable = flexEDKbdMouse
End Sub

Private Sub FlexGrid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrorHandler
With flexGrid
    If .TextMatrix(Row, ColZoneName) <> oldZoneName Then
        If MsgBox("”Ì „  €Ì— ≈”„ «·„‰ŸﬁÂ" + Chr(13) + "Â· √‰  „ √ﬂœ", vbQuestion + vbDefaultButton2 + vbYesNo, " €ÌÌ— ≈”„ «·„‰ÿﬁÂ") = vbYes Then
            mainDataService.ChangeTheZoneName .TextMatrix(Row, ColZoneName), .TextMatrix(Row, ColZoneNo)
        Else
            .TextMatrix(Row, ColZoneName) = oldZoneName
        End If
    End If
End With
Exit Sub
ErrorHandler:
flexGrid.TextMatrix(Row, ColZoneName) = oldZoneName
MsgBox Err.Description
End Sub

Private Sub FlexGrid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, cancel As Boolean)
If Col <> ColZoneName Then cancel = True
oldZoneName = flexGrid.TextMatrix(Row, ColZoneName)
End Sub

Private Sub FlexGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim FirstRow As Integer, LastRow As Integer, Vrow As Integer
With flexGrid
If KeyCode = vbKeyDelete Then
    If MsgBox("Â· √‰  „ √ﬂœ „‰ ⁄„·Ì… «·Õ–›", vbYesNo + vbDefaultButton2, "Õ–› «·„‰«ÿﬁ «·„Õœœ…") = vbYes Then
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
                  If DeleteRow(flexGrid, Vrow, ColZoneNo, "CoZone", "ZoneNo") Then
                    .RemoveItem Vrow
                End If
            End If
        Next
        .Col = ColCallNo
        
        .SetFocus
    End If
End If
End With
End Sub

Private Sub Form_Load()
init
End Sub

Private Sub TxtZoneName_Change()
FillGrid TxtZoneName.Text
End Sub

Private Sub TxtZoneName_GotFocus()
ChangeToArabic
End Sub

Sub AddZoneToTheGrid(ZoneName As String, ZoneNo As Long, flexGrid As VSFlexGrid)
On Error GoTo ErrorHandler
Dim Vrow As Integer
With flexGrid
    .AddItem ""
    Vrow = .Rows - 1
    .TextMatrix(Vrow, ColZoneNo) = ZoneNo
    .TextMatrix(Vrow, ColZoneName) = ZoneName
    .TextMatrix(Vrow, ColZoneClientsCount) = 0
End With
Exit Sub
ErrorHandler:
MsgBox Err.Description
End Sub
Private Sub TxtZoneName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If flexGrid.Rows = 1 Then
        Dim ZoneNo As Long
        ZoneNo = mainDataService.InsertNewZone(TxtZoneName.Text)
        If ZoneNo <> -1 Then
            AddZoneToTheGrid TxtZoneName.Text, ZoneNo, flexGrid
            FillFormating 2, flexGrid
            TxtZoneName.SetFocus
            SendKeys "{home}+{end}"
        End If
    End If
End If
End Sub
