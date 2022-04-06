VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form FrmChooseTypes 
   Caption         =   "ÊÍÏíÏ ÃäæÇÚ ÇáÈíÇäÇÊ"
   ClientHeight    =   3795
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5760
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   3795
   ScaleWidth      =   5760
   StartUpPosition =   1  'CenterOwner
   Begin VSFlex8Ctl.VSFlexGrid MsflexGrid1 
      Height          =   3165
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   5715
      _cx             =   10081
      _cy             =   5583
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
   Begin Threed.SSFrame SSFrame1 
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   3180
      Width           =   5715
      _ExtentX        =   10081
      _ExtentY        =   1085
      _Version        =   131074
      Begin Threed.SSCommand CmdExit 
         Height          =   555
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   979
         _Version        =   131074
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "FrmChooseTypes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const ColChk = 1
Const ColTypeId = 2
Const ColTypeName = 3
Dim RsByanType As New ADODB.Recordset
Function FillByanType() As String
ByanNo = ""
With MsflexGrid1
    For i = 1 To .Rows - 1
    .Row = i
    If .RowData(.Row) = 1 Then
        ByanNo = ByanNo & "-" & .TextMatrix(.Row, ColTypeId)
    End If
    Next
End With
FillByanType = Mid(ByanNo, 2)
End Function

Sub FillGrid(Rs As Recordset)
Dim Vrow As Integer
With MsflexGrid1
    .Cols = 4
    .Rows = 1
    If Rs.RecordCount > 0 Then
        Rs.MoveFirst
        Do While Not Rs.EOF
            .AddItem ""
            .Col = ColChk
            Vrow = .Rows - 1
            .Row = Vrow
            Set .CellPicture = LoadPicture(App.Path + "\ICONS\Chkoff95.bmp")
            .RowData(Vrow) = 0
            .TextMatrix(Vrow, ColTypeId) = Rs!TYpeId
            .TextMatrix(Vrow, ColTypeName) = Rs!TypeName
            Rs.MoveNext
        Loop
    End If
End With
End Sub
Sub FillFormatString()
    fs = "|Ê"
    fs = fs + "|<" + "äæÚ ÇáÈíÇä"
    fs = fs + "|<" + "ÇáÔÑÍ"
    With MsflexGrid1
        .FormatString = fs
        SetColWidths ColTypeId
        SetColWidths ColTypeName
        .ColWidth(ColChk) = 300
    End With
End Sub
Sub SetColWidths(ColNo As Integer)
    Dim i, J, s, w
    With MsflexGrid1
            s = 0
            For i = 0 To .Rows - 1
                w = TextWidth(.TextMatrix(i, ColNo))
                If w > s Then s = w
            Next i
            .ColWidth(ColNo) = s + 100
    End With
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub
Private Sub Form_Load()
'    Sqltext = "Select TYpeId , TYpeName From CoByanType where TYpeId in (4,5,7)Order By TYpeId"
    Sqltext = "Select TYpeId , TYpeName From CoByanType Order By TYpeId"
    Set RsByanType = de.con.Execute(Sqltext)
    FillGrid RsByanType
    FillFormatString
End Sub

Private Sub Form_Unload(Cancel As Integer)
ByanType = FillByanType
'InsertStrNo
End Sub

Private Sub MSFlexGrid1_Click()
With MsflexGrid1
    If .Col = ColChk Then
        If .RowData(.Row) Xor 1 Then
            Set .CellPicture = LoadPicture(App.Path + "\ICONS\ChkoN95.bmp")
            .RowData(.Row) = 1
        Else
           Set .CellPicture = LoadPicture(App.Path + "\ICONS\Chkoff95.bmp")
            .RowData(.Row) = 0
        End If
    End If
End With
End Sub
