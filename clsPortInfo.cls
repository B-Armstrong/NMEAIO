VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPortInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "ability to store and retrieve mscomm port data"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarSpeed As Variant 'local copy
Private mvarParity As Variant 'local copy

Public Property Get varParity() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.varParity
    If IsObject(mvarvarParity) Then
        Set varParity = mvarParity
    Else
        varParity = mvarParity
    End If
End Property

Public Property Set varSpeed(ByVal vData As Variant)
mvarSpeed = vData
End Property

Public Property Get varSpeed() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.varSpeed
    If IsObject(mvarvarSpeed) Then
        Set varSpeed = mvarSpeed
    Else
        varSpeed = mvarSpeed
    End If
End Property



Public Sub setPortData(varSpeed As String, varParity As String)
mvarSpeed = varSpeed
mvarParity = varParity
End Sub

