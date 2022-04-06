VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AdhamProduct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Member0" ,"AdhamProductViewModel"
'local variable(s) to hold property value(s)
Private mvarProdNo As Double 'local copy
Private mvarProdBaCodeNo As String 'local copy
Private mvarProdDate As Variant 'local copy
Private mvarProdPurchaseDate As Variant 'local copy
Private mvarModNo As Integer 'local copy
Private mvarAdhamProductViewModel As AdhamProductViewModel


Public Property Get AdhamProductViewModel() As AdhamProductViewModel
    If mvarAdhamProductViewModel Is Nothing Then
        Set mvarAdhamProductViewModel = New AdhamProductViewModel
    End If


    Set AdhamProductViewModel = mvarAdhamProductViewModel
End Property


Public Property Set AdhamProductViewModel(vData As AdhamProductViewModel)
    Set mvarAdhamProductViewModel = vData
End Property
Private Sub Class_Terminate()
    Set mvarAdhamProductViewModel = Nothing
End Sub



Public Property Let ModNo(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ModNo = 5
    mvarModNo = vData
End Property


Public Property Get ModNo() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ModNo
    ModNo = mvarModNo
End Property



Public Property Let ProdPurchaseDate(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ProdPurchaseDate = 5
    mvarProdPurchaseDate = vData
End Property


Public Property Set ProdPurchaseDate(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.ProdPurchaseDate = Form1
    Set mvarProdPurchaseDate = vData
End Property


Public Property Get ProdPurchaseDate() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ProdPurchaseDate
    If IsObject(mvarProdPurchaseDate) Then
        Set ProdPurchaseDate = mvarProdPurchaseDate
    Else
        ProdPurchaseDate = mvarProdPurchaseDate
    End If
End Property



Public Property Let ProdDate(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ProdDate = 5
    mvarProdDate = vData
End Property


Public Property Set ProdDate(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.ProdDate = Form1
    Set mvarProdDate = vData
End Property


Public Property Get ProdDate() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ProdDate
    If IsObject(mvarProdDate) Then
        Set ProdDate = mvarProdDate
    Else
        ProdDate = mvarProdDate
    End If
End Property



Public Property Let ProdBaCodeNo(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ProdBaCodeNo = 5
    mvarProdBaCodeNo = vData
End Property


Public Property Get ProdBaCodeNo() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ProdBaCodeNo
    ProdBaCodeNo = mvarProdBaCodeNo
End Property



Public Property Let ProdNo(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ProdNo = 5
    mvarProdNo = vData
End Property


Public Property Get ProdNo() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ProdNo
    ProdNo = mvarProdNo
End Property



