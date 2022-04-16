VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form FrmCoFeesPrices 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " —„Ì“ «·«ÃÊ—"
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
   Begin VB.TextBox TxtfeesName 
      Alignment       =   1  'Right Justify
      Height          =   405
      Left            =   60
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   30
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
      Caption         =   "≈”„ «·«ÃÊ—"
      Height          =   195
      Left            =   5805
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   90
      Width           =   720
   End
End
Attribute VB_Name = "FrmCoFeesPrices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Ok As Boolean, Flag As Boolean, Pos As Integer
Dim oldFeesName As String
Dim oldClientFeesPrice As Double

Const ColFeesNo = 1
Const ColfeesName = 2
Const ColFeesClientPrice = 3

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
    fs = "|>" + "FeesNo"
    fs = fs + "|>" + "«·«Ã—"
    fs = fs + "|>" + "”⁄— «·„” Â·ﬂ"
    With FlexGrid
        .FormatString = fs
        .Cols = 4
        
        .ColWidth(ColFeesNo) = 0
        SetColWidths ColfeesName, FlexGrid
        SetColWidths ColFeesClientPrice, FlexGrid
        
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
On Error GoTo errorhandler
Dim rsFees As New ADODB.Recordset
    sqlText = "select FeesId , feesname , CliPriceafterdiscount as Price from CoMaintFees "
    sqlText = sqlText & " where  FeesId <> -1"
    
    If Not IsMissing(searchFormula) Then
        sqlText = sqlText & " and feesname like " & LikeExpression(searchFormula)
    End If
    sqlText = sqlText & " order by FeesId"
    Set rsFees = de.con.Execute(sqlText)
    
    Set FlexGrid.DataSource = rsFees
    FillFormating 2, FlexGrid
Exit Sub
errorhandler:
MsgBox Err.Description
End Sub
Sub init()
    top = 0
    left = 0
    Ok = True
    FillGrid
    FlexGrid.Editable = flexEDKbdMouse
End Sub

Private Sub FlexGrid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo errorhandler
With FlexGrid
    If .TextMatrix(Row, ColfeesName) <> oldFeesName Or .TextMatrix(Row, ColFeesClientPrice) <> oldClientFeesPrice Then
        If MsgBox("”Ì „  €Ì— „⁄·Ê„«  «·«ÕÊ—" + Chr(13) + "Â· √‰  „ √ﬂœ", vbQuestion + vbDefaultButton2 + vbYesNo, " €ÌÌ— „⁄·Ê„«  «·«ÃÊ—") = vbYes Then
            mainDataService.ChangeTheFeesName .TextMatrix(Row, ColfeesName), .TextMatrix(Row, ColFeesClientPrice), .TextMatrix(Row, ColFeesNo)
        Else
            .TextMatrix(Row, ColfeesName) = oldFeesName
            .TextMatrix(Row, ColFeesClientPrice) = oldClientFeesPrice
        End If
    End If
End With
FillFormating 2, FlexGrid
Exit Sub
errorhandler:
FlexGrid.TextMatrix(Row, ColfeesName) = oldFeesName
FlexGrid.TextMatrix(Row, ColFeesClientPrice) = oldClientFeesPrice
MsgBox Err.Description
End Sub

Private Sub FlexGrid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, cancel As Boolean)

oldFeesName = FlexGrid.TextMatrix(Row, ColfeesName)
oldClientFeesPrice = Val(FlexGrid.TextMatrix(Row, ColFeesClientPrice))
End Sub

Private Sub FlexGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim FirstRow As Integer, LastRow As Integer, Vrow As Integer
With FlexGrid
If KeyCode = vbKeyDelete Then
    If MsgBox("Â· √‰  „ √ﬂœ „‰ ⁄„·Ì… «·Õ–›", vbYesNo + vbDefaultButton2, "Õ–› «·«ÕÊ— «·„Õœœ…") = vbYes Then
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
                  If DeleteRow(FlexGrid, Vrow, ColFeesNo, "CoMaintFees", "FeesId") Then
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

Private Sub TxtfeesName_Change()
FillGrid TxtfeesName.text
End Sub

Private Sub TxtfeesName_GotFocus()
ChangeToArabic
End Sub

Sub AddFeesToTheGrid(FeesName As String, FeesId As Long, FlexGrid As VSFlexGrid)
On Error GoTo errorhandler
Dim Vrow As Integer
With FlexGrid
    .AddItem ""
    Vrow = .Rows - 1
    .TextMatrix(Vrow, ColFeesNo) = FeesId
    .TextMatrix(Vrow, ColfeesName) = FeesName
    .TextMatrix(Vrow, ColFeesClientPrice) = 0
End With
Exit Sub
errorhandler:
MsgBox Err.Description
End Sub
Private Sub TxtfeesName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If FlexGrid.Rows = 1 Then
        Dim FessId As Long
        FessId = mainDataService.InsertNewFees(TxtfeesName.text)
        If FessId <> -1 Then
            AddFeesToTheGrid TxtfeesName.text, FessId, FlexGrid
            FillFormating 2, FlexGrid
            TxtfeesName.SetFocus
            Sendkeys "{home}+{end}"
        End If
    End If
End If
End Sub
