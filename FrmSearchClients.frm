VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form FrmSearchClients 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "«·“»«∆‰"
   ClientHeight    =   7245
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11655
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   7245
   ScaleWidth      =   11655
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtTop 
      Alignment       =   1  'Right Justify
      Height          =   405
      Left            =   30
      RightToLeft     =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Text            =   "5"
      Top             =   390
      Width           =   945
   End
   Begin VB.TextBox txtClientName 
      Alignment       =   1  'Right Justify
      Height          =   405
      Left            =   1020
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   390
      Width           =   10605
   End
   Begin VSFlex8Ctl.VSFlexGrid FlexGrid 
      Height          =   6345
      Left            =   30
      TabIndex        =   4
      Top             =   840
      Width           =   11595
      _cx             =   20452
      _cy             =   11192
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
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "⁄œœ «·”Ã·« "
      Height          =   195
      Left            =   150
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   60
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "«·“»Ê‰"
      Height          =   195
      Left            =   11220
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   90
      Width           =   420
   End
End
Attribute VB_Name = "FrmSearchClients"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Ok As Boolean, Flag As Boolean, Pos As Integer


Const ColClientNo = 1
Const ColClientName = 2
Const ColClientPhoneNBr = 3
Const ColClientMobilePhoneNbr = 4
Const ColClientAddress = 5




Dim clientsToSearch As FiledColumns

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
    fs = "|>" + "ClientNo"
    fs = fs + "|>" + "≈”„ «·“»Ê‰"
    fs = fs + "|>" + "—ﬁ„ «·Â« ›"
    fs = fs + "|>" + "—ﬁ„ «·„Ê»«Ì·"
    fs = fs + "|>" + "⁄‰Ê«‰ «·“»Ê‰"

    
  
    
    With FlexGrid
        .FormatString = fs
        .Cols = 6
        
        .ColWidth(ColClientNo) = 0
        SetColWidths ColClientName, FlexGrid
        SetColWidths ColClientPhoneNBr, FlexGrid
        SetColWidths ColClientMobilePhoneNbr, FlexGrid
        SetColWidths ColClientAddress, FlexGrid

        
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


Sub FillGrid(VTop As Integer, searchFields As FiledColumns, Optional searchFormula As String)
On Error GoTo ErrorHandler
Dim sqlText As String
FillFieldsCollectionToSearch

Dim rsClients As New ADODB.Recordset

 sqlText = mainDataService.GeSearchClients(VTop, searchFields, searchFormula)
Set FlexGrid.DataSource = de.con.Execute(sqlText)

FillFormating 2, FlexGrid
AddDefaultCheckBoxesIntoTheGrid
Exit Sub
ErrorHandler:
MsgBox Err.Description
End Sub
Sub AddDefaultCheckBoxesIntoTheGrid()



If clientsToSearch Is Nothing Then
    With FlexGrid
        .Cell(flexcpChecked, 0, 1, 0, .Cols - 2) = True
        .Cell(flexcpChecked, 0, 1, 0, .Cols - 2) = flexUnchecked
        .Cell(flexcpChecked, 0, ColClientName, 0, ColClientName) = True
        .Cell(flexcpChecked, 0, ColClientPhoneNBr, 0, ColClientPhoneNBr) = True
        .Cell(flexcpChecked, 0, ColClientMobilePhoneNbr, 0, ColClientMobilePhoneNbr) = True
    End With
Else
    Dim fieldColumn As fieldColumn
    With FlexGrid
        For i = 1 To clientsToSearch.Count
            If clientsToSearch(i).isChecked Then
                .Cell(flexcpChecked, 0, clientsToSearch(i).colHeaderName, 0, clientsToSearch(i).colHeaderName) = flexChecked
            Else
                .Cell(flexcpChecked, 0, clientsToSearch(i).colHeaderName, 0, clientsToSearch(i).colHeaderName) = flexUnchecked
            End If
         Next i
    End With
End If

End Sub
Sub init()
    top = 0
    left = 0
    Ok = True
        
    AddDefaultCheckBoxesIntoTheGrid
    FlexGrid.Rows = 1
    FlexGrid.Editable = flexEDKbdMouse
End Sub



Private Sub FlexGrid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
If FlexGrid.Cell(flexcpChecked, Row, Col) = flexChecked Then
   FillGrid Val(txtTop.Text), clientsToSearch, txtClientName.Text
End If

End Sub

Private Sub FlexGrid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, cancel As Boolean)
If Row <> 0 Then cancel = True
End Sub

Private Sub FlexGrid_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button And vbRightButton Then
            If Gettag(empNo, 42) Then
                PopupMenu mnu
            End If
    End If
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
With FlexGrid
If KeyCode = vbKeyEscape Then
    customerNumber = 0
    searchClientIsAllow = False
    Unload Me
End If
End With
End Sub

Private Sub Form_Load()
init
End Sub


 Sub FillFieldsCollectionToSearch()
Set clientsToSearch = New FiledColumns

Dim fieldToAdd As fieldColumn



With FlexGrid
'    If FlexGrid.Cell(flexcpChecked, 0, ColClientName) = flexChecked Then
        Set fieldToAdd = New fieldColumn
        fieldToAdd.filedId = "AdhamName"
        fieldToAdd.colHeaderName = ColClientName
        fieldToAdd.isChecked = IIf(FlexGrid.Cell(flexcpChecked, 0, ColClientName) = flexChecked, True, False)
        clientsToSearch.Add fieldToAdd.filedId, fieldToAdd.colHeaderName, fieldToAdd.isChecked, fieldToAdd.filedId
'    End If
    
'    If FlexGrid.Cell(flexcpChecked, 0, ColClientPhoneNBr) = flexChecked Then
        Set fieldToAdd = New fieldColumn
        fieldToAdd.filedId = "adhamPhon"
        fieldToAdd.colHeaderName = ColClientPhoneNBr
        fieldToAdd.isChecked = IIf(FlexGrid.Cell(flexcpChecked, 0, ColClientPhoneNBr) = flexChecked, True, False)
        clientsToSearch.Add fieldToAdd.filedId, fieldToAdd.colHeaderName, fieldToAdd.isChecked, fieldToAdd.filedId
'    End If
'    If FlexGrid.Cell(flexcpChecked, 0, ColClientMobilePhoneNbr) = flexChecked Then
        Set fieldToAdd = New fieldColumn
        fieldToAdd.filedId = "MobilePhone"
        fieldToAdd.colHeaderName = ColClientMobilePhoneNbr
        fieldToAdd.isChecked = IIf(FlexGrid.Cell(flexcpChecked, 0, ColClientMobilePhoneNbr) = flexChecked, True, False)
        clientsToSearch.Add fieldToAdd.filedId, fieldToAdd.colHeaderName, fieldToAdd.isChecked, fieldToAdd.filedId
'    End If
'    If FlexGrid.Cell(flexcpChecked, 0, ColClientAddress) = flexChecked Then
        Set fieldToAdd = New fieldColumn
        fieldToAdd.filedId = "adhamAdress"
        fieldToAdd.colHeaderName = ColClientAddress
        fieldToAdd.isChecked = IIf(FlexGrid.Cell(flexcpChecked, 0, ColClientAddress) = flexChecked, True, False)
        clientsToSearch.Add fieldToAdd.filedId, fieldToAdd.colHeaderName, fieldToAdd.isChecked, fieldToAdd.filedId
'    End If

End With

End Sub

Private Sub Form_Unload(cancel As Integer)
    customerNumber = 0
    searchClientIsAllow = False
End Sub

Private Sub txtClientName_Change()

FillGrid Val(txtTop.Text), clientsToSearch, txtClientName.Text

End Sub

Private Sub TxtClientName_GotFocus()
SendKeys "{end}"
ChangeToArabic
End Sub


Private Sub txtClientName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then MoveCursor KeyCode, FlexGrid

End Sub

Private Sub txtClientName_KeyPress(KeyAscii As Integer)
    With FlexGrid
        If KeyAscii = 13 Then
            If .Row <= 0 Or Trim(txtClientName.Text) = "" Then
                customerNumber = 0
                customerName = txtClientName.Text
            Else
                customerNumber = .TextMatrix(.Row, ColClientNo)
                customerName = .TextMatrix(.Row, ColClientName)
            End If
            Unload Me
        End If
    End With
End Sub

Private Sub txtTop_Change()
    FillFieldsCollectionToSearch
    FillGrid Val(txtTop.Text), clientsToSearch, txtClientName.Text
End Sub

