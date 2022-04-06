VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmClientReparation 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ﬁ—Ì— «·‘ﬂ«ÊÏ Ê«·«’·«Õ« "
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   4320
   ScaleWidth      =   5415
   Begin Crystal.CrystalReport cr1 
      Left            =   1050
      Top             =   1170
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      DragMode        =   1  'Automatic
      Height          =   2835
      Left            =   90
      TabIndex        =   3
      Top             =   1410
      Visible         =   0   'False
      Width           =   5295
      _cx             =   9340
      _cy             =   5001
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
   Begin VB.TextBox TxtCustomerName 
      Alignment       =   1  'Right Justify
      Height          =   525
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   810
      Width           =   4605
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   30
      Top             =   870
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
            Picture         =   "FrmClientReparation.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClientReparation.frx":26D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClientReparation.frx":5533
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClientReparation.frx":7CD8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClientReparation.frx":A5D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClientReparation.frx":CAF9
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClientReparation.frx":F2B1
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClientReparation.frx":11CC5
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClientReparation.frx":14617
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClientReparation.frx":1738B
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClientReparation.frx":19BDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClientReparation.frx":1CA81
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClientReparation.frx":1F7DB
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClientReparation.frx":2217B
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClientReparation.frx":250DD
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClientReparation.frx":27B05
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClientReparation.frx":2A4BF
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClientReparation.frx":2CDF0
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClientReparation.frx":2F6F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClientReparation.frx":3215E
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClientReparation.frx":3510F
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClientReparation.frx":37A38
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClientReparation.frx":3A36D
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClientReparation.frx":3CAAD
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClientReparation.frx":3F41C
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClientReparation.frx":41CA4
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClientReparation.frx":43EB3
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClientReparation.frx":46810
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClientReparation.frx":4903A
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClientReparation.frx":4BAEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClientReparation.frx":4E52A
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClientReparation.frx":51040
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClientReparation.frx":53FDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClientReparation.frx":56DA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClientReparation.frx":59A20
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClientReparation.frx":5C8B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClientReparation.frx":5F4F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClientReparation.frx":6209F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   690
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   5415
      _ExtentX        =   9551
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
            Object.ToolTipText     =   "√—‘›… «·»Ì«‰« "
            ImageIndex      =   37
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "«·“»Ê‰"
      Height          =   195
      Left            =   4920
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   900
      Width           =   420
   End
End
Attribute VB_Name = "FrmClientReparation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FieldsArrayToSearch(4) As New MapClientField
Dim Cols(4) As New MapClientField

Dim Ok As Boolean, Flag As Boolean, Pos As Integer
'Dim currentGridVisbileState As Boolean
'Dim currentSenderLength As Integer
Dim prevSendrLength As Integer
 
Dim prevDataIsNull As Boolean

Sub SetColWidths(ByVal ColNo As Integer, FlexGrid As VSFlexGrid)
    With FlexGrid
        .AutoSize (ColNo)
    End With
End Sub

Function GetDataMemebers(listField() As MapClientField) As String

Dim fields As String
fields = ""
For i = 1 To UBound(listField)
    fields = fields & "," & listField(i).filedId
Next
GetDataMemebers = Mid(fields, 2)
End Function


Sub FillFormating(ByVal i As Integer, FlexGrid As VSFlexGrid)
If i = 1 Then
    fs = ""
    For i = 1 To UBound(Cols)
        fs = fs + "|>" + Cols(i).FiledName
    Next

    With FlexGrid
        .FormatString = fs
        .Cols = UBound(Cols) + 1
        For i = 1 To UBound(Cols)
            If Cols(i).IsVisible Then
                SetColWidths Cols(i).Order, FlexGrid
            Else
                .ColWidth(i) = 0
            End If
        Next
    End With
End If
End Sub

Private Sub Grid_RowColChange()
On Error GoTo ErrorHandler
If Flag Then
    Ok = False
    With Grid
       Select Case Pos
        Case 1
            TxtCustomerName.Tag = .TextMatrix(.Row, 1)
            TxtCustomerName.Text = .TextMatrix(.Row, 2)
       End Select
    End With
    Ok = True
End If
Exit Sub
ErrorHandler:
MsgBox Err.Description
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
    Case 1
        PrintRep
    Case 3
        Unload Me
End Select
End Sub

Sub PrintRep()
On Error GoTo ErrorHandler

If Val(TxtCustomerName.Tag) = 0 Then
    MsgBox "·„ Ì „ ≈œŒ«· —ﬁ„ «·“»Ê‰", vbExclamation, "«·“»Ê‰ €Ì— „⁄—Ê›"
    Exit Sub
End If
Dim sqlText As String
sqlText = "SELECT  adhamname,RepPrice,Bdate,Edate,Notes,CallNo"
sqlText = sqlText & " ,RegestDate,ZoneName,TeamName "
sqlText = sqlText & " From AdhamClientRepView "
sqlText = sqlText & " where CliNo= " & TxtCustomerName.Tag
sqlText = sqlText & " order by repDate asc "
With cr1
    .Connect = ConnectName("")
    .SQLQuery = sqlText
    .ReportFileName = App.Path & "\Reports\RepClientReparation.rpt"
    .DiscardSavedData = True
    .WindowState = crptMaximized
    .Action = 1
End With

Exit Sub
ErrorHandler:
MsgBox Err.Description
End Sub
Private Sub TxtCustomerName_Change()
Search TxtCustomerName, 1, "AdhamView7", FieldsArrayToSearch, "AdhamNo", True, 0, 0
End Sub

Private Sub TxtCustomerName_GotFocus()
ChangeToArabic
Pos = 1
End Sub

Private Sub TxtCustomerName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then MoveCursor KeyCode, Grid
End Sub

Private Sub TxtCustomerName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        If Grid.Visible Then
            Ok = False
            TxtCustomerName.Tag = IIf(Grid.SelectedRows = 0, Grid.TextMatrix(1, 1), Grid.TextMatrix(Grid.Row, 1))
            TxtCustomerName.Text = IIf(Grid.SelectedRows = 0, Grid.TextMatrix(1, 2), Grid.TextMatrix(Grid.Row, 2))
            Ok = True
            Grid.Visible = False
        End If
        
End If
End Sub

Sub ChangeCursor(sender As Control, Optional top As Variant, Optional left As Variant)

    With sender
       Grid.top = top + .top + .Height
       Grid.left = left + .left
       Grid.Width = .Width
    End With

End Sub

Sub FillArray(Index As Integer, arr() As MapClientField)

If Index = 1 Then
    Set arr(1) = New MapClientField
    Set arr(2) = New MapClientField
    Set arr(3) = New MapClientField
    Set arr(4) = New MapClientField
    
    arr(1).filedId = "AdhamNo"
    arr(1).FiledName = "—ﬁ„ «·“»Ê‰"
    arr(1).IsVisible = False
    arr(1).IsListColumn = True
    arr(1).IsListField = False
    arr(1).Order = 1
    
    arr(2).filedId = "AdhamName"
    arr(2).FiledName = "≈”„ «·“»Ê‰"
    arr(2).IsVisible = True
    arr(2).IsListColumn = False
    arr(2).IsListField = True
    arr(2).Order = 2
    
    arr(3).filedId = "adhamphon"
    arr(3).FiledName = "—ﬁ„ «·Â« ›"
    arr(3).IsVisible = True
    arr(3).IsListColumn = False
    arr(3).IsListField = False
    arr(3).Order = 3
    
    arr(4).filedId = "MobilePhone"
    arr(4).FiledName = "—ﬁ„ «·„Ê»«Ì·"
    arr(4).IsVisible = True
    arr(4).IsListColumn = False
    arr(4).IsListField = False
    arr(4).Order = 4

End If
End Sub

Sub FillArrays()
    FillArray 1, FieldsArrayToSearch
    FillArray 1, Cols

End Sub

Sub init()
    top = 0
    left = 0
    Grid.Rows = 1
    Grid.SelectionMode = flexSelectionListBox
    Ok = True
    FillArrays

End Sub
Private Sub Form_Load()
init
End Sub

Sub Search(sender As Control, Pos As Integer, tableName As String, listField() As MapClientField, dataMember As String, Optional isChangeCursor = True, Optional top As Variant, Optional left As Variant)
On Error GoTo ErrorHandler
Dim RsSearch As New ADODB.Recordset



If Ok Then
    sender.Tag = ""

'    If sender.Text = "" Or (currentSenderLength <= Len(sender.Text) And Not currentGridVisbileState) Then
    If sender.Text = "" Or (prevDataIsNull And prevSendrLength <= Len(sender.Text)) Then
        Grid.Visible = False
        Exit Sub
    End If
    Flag = False
    
    sqlText = "Select top 5 " & GetDataMemebers(listField) & " From " & tableName & " Where "
 
    For i = 1 To UBound(listField)
         sqlText = sqlText & listField(i).filedId & " Like " & LikeExpression(sender.Text)
         If i <> UBound(listField) Then
                sqlText = sqlText & " Or "
         End If
    Next
'    sqlText = sqlText & dataMember & " Like " & LikeExpression(sender.Text) & " Or " & listField & " Like " & LikeExpression(sender.Text)
    Set RsSearch = de.con.Execute(sqlText)
    If RsSearch.RecordCount > 0 Then
        Set Grid.DataSource = RsSearch
        'Grid.Row = 0
        FillFormating Pos, Grid
        If isChangeCursor Then ChangeCursor sender, top, left
        Grid.Visible = True
        prevDataIsNull = False

    Else
        sender.Tag = 0
        Grid.Visible = False
        prevDataIsNull = True
        prevSendrLength = Len(sender.Text)

    End If
    Flag = True
End If


   Exit Sub
ErrorHandler:
MsgBox Err.Description
End Sub

