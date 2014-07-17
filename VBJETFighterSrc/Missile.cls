VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Missile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarX As Integer 'local copy
Private mvarY As Integer 'local copy
'local variable(s) to hold property value(s)
Private mvarVisible As Boolean 'local copy
Public Property Let Visible(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Visible = 5
    mvarVisible = vData
End Property


Public Property Get Visible() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Visible
    Visible = mvarVisible
End Property



Public Property Let Y(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Y = 5
    mvarY = vData
End Property


Public Property Get Y() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Y
    Y = mvarY
End Property



Public Property Let X(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.X = 5
    mvarX = vData
End Property


Public Property Get X() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.X
    X = mvarX
End Property



