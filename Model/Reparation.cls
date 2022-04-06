VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Reparation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"AdhamProduct"
Attribute VB_Ext_KEY = "Member1" ,"ReparationPieces"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member2" ,"ReparationPieces"
Attribute VB_Ext_KEY = "Member3" ,"ReparationPieces"
Attribute VB_Ext_KEY = "Member4" ,"ReparationWorks"
Attribute VB_Ext_KEY = "Member5" ,"ReparationWorks"
Attribute VB_Ext_KEY = "Member6" ,"ReparationWorks"
Attribute VB_Ext_KEY = "Member7" ,"ReparationState"
'local variable(s) to hold property value(s)
Private mvarReparationWorks As ReparationWorks 'local copy
Private mvarReparationPieces As ReparationPieces 'local copy
Private mvarRepNo As Double 'local copy
'local variable(s) to hold property value(s)
Private mvarCallNo As Double 'local copy
Private mvarTeamNo As Integer 'local copy
Private mvarProdNo As Double 'local copy
Private mvarRepDate As Variant 'local copy
Private mvarRepTimeBegin As Variant 'local copy
Private mvarRepTimeEnd As Variant 'local copy
Private mvarRepPrice As Double 'local copy
Private mvarCash As Integer 'local copy
Private mvarregestdate As Variant 'local copy
Private mvarNotes As String 'local copy
Private mvarCarried As Boolean 'local copy
Private mvarCliRecever As Integer 'local copy
Private mvarDescription As String 'local copy
Private mvarVoltBefor As Integer 'local copy
Private mvarVoltAfter As Integer 'local copy
'local variable(s) to hold property value(s)
Private mvarAdhamProduct As AdhamProduct 'local copy
'local variable(s) to hold property value(s)
Private mvarReparationState As EnumState 'local copy
Public Property Let ReparationState(ByVal vData As EnumState)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.ReparationState = Form1
     mvarReparationState = vData
End Property


Public Property Get ReparationState() As EnumState
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ReparationState
     ReparationState = mvarReparationState
End Property



Public Property Set AdhamProduct(ByVal vData As AdhamProduct)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.AdhamProduct = Form1
    Set mvarAdhamProduct = vData
End Property


Public Property Get AdhamProduct() As AdhamProduct
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.AdhamProduct
    Set AdhamProduct = mvarAdhamProduct
End Property



Public Property Let VoltAfter(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.VoltAfter = 5
    mvarVoltAfter = vData
End Property


Public Property Get VoltAfter() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.VoltAfter
    VoltAfter = mvarVoltAfter
End Property



Public Property Let VoltBefor(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.VoltBefor = 5
    mvarVoltBefor = vData
End Property


Public Property Get VoltBefor() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.VoltBefor
    VoltBefor = mvarVoltBefor
End Property



Public Property Let Description(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Description = 5
    mvarDescription = vData
End Property


Public Property Get Description() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Description
    Description = mvarDescription
End Property



Public Property Let CliRecever(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CliRecever = 5
    mvarCliRecever = vData
End Property


Public Property Get CliRecever() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CliRecever
    CliRecever = mvarCliRecever
End Property



Public Property Let Carried(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Carried = 5
    mvarCarried = vData
End Property


Public Property Get Carried() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Carried
    Carried = mvarCarried
End Property



Public Property Let Notes(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Notes = 5
    mvarNotes = vData
End Property


Public Property Get Notes() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Notes
    Notes = mvarNotes
End Property



Public Property Let regestdate(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.regestdate = 5
    mvarregestdate = vData
End Property


Public Property Set regestdate(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.regestdate = Form1
    Set mvarregestdate = vData
End Property


Public Property Get regestdate() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.regestdate
    If IsObject(mvarregestdate) Then
        Set regestdate = mvarregestdate
    Else
        regestdate = mvarregestdate
    End If
End Property



Public Property Let Cash(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Cash = 5
    mvarCash = vData
End Property


Public Property Get Cash() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Cash
    Cash = mvarCash
End Property



Public Property Let RepPrice(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RepPrice = 5
    mvarRepPrice = vData
End Property


Public Property Get RepPrice() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RepPrice
    RepPrice = mvarRepPrice
End Property



Public Property Let RepTimeEnd(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RepTimeEnd = 5
    mvarRepTimeEnd = vData
End Property


Public Property Set RepTimeEnd(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.RepTimeEnd = Form1
    Set mvarRepTimeEnd = vData
End Property


Public Property Get RepTimeEnd() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RepTimeEnd
    If IsObject(mvarRepTimeEnd) Then
        Set RepTimeEnd = mvarRepTimeEnd
    Else
        RepTimeEnd = mvarRepTimeEnd
    End If
End Property



Public Property Let RepTimeBegin(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RepTimeBegin = 5
    mvarRepTimeBegin = vData
End Property


Public Property Set RepTimeBegin(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.RepTimeBegin = Form1
    Set mvarRepTimeBegin = vData
End Property


Public Property Get RepTimeBegin() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RepTimeBegin
    If IsObject(mvarRepTimeBegin) Then
        Set RepTimeBegin = mvarRepTimeBegin
    Else
        RepTimeBegin = mvarRepTimeBegin
    End If
End Property



Public Property Let RepDate(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RepDate = 5
    mvarRepDate = vData
End Property


Public Property Set RepDate(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.RepDate = Form1
    Set mvarRepDate = vData
End Property


Public Property Get RepDate() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RepDate
    If IsObject(mvarRepDate) Then
        Set RepDate = mvarRepDate
    Else
        RepDate = mvarRepDate
    End If
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



Public Property Let TeamNo(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TeamNo = 5
    mvarTeamNo = vData
End Property


Public Property Get TeamNo() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TeamNo
    TeamNo = mvarTeamNo
End Property



Public Property Let CallNo(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CallNo = 5
    mvarCallNo = vData
End Property


Public Property Get CallNo() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CallNo
    CallNo = mvarCallNo
End Property



Public Property Let repNo(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RepNo = 5
    mvarRepNo = vData
End Property


Public Property Get repNo() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RepNo
    repNo = mvarRepNo
End Property



Public Property Set ReparationPieces(ByVal vData As ReparationPieces)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.ReparationPieces = Form1
    Set mvarReparationPieces = vData
End Property


Public Property Get ReparationPieces() As ReparationPieces
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ReparationPieces
    Set ReparationPieces = mvarReparationPieces
End Property



Public Property Set ReparationWorks(ByVal vData As ReparationWorks)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.ReparationWorks = Form1
    Set mvarReparationWorks = vData
End Property


Public Property Get ReparationWorks() As ReparationWorks
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ReparationWorks
    Set ReparationWorks = mvarReparationWorks
End Property



