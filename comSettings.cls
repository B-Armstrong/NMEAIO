VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "comSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private mvparity As String

Public Function parity(ByVal parityBits As String) As String
Select Case parityBits
    Case "None"
        parity = "N"
       
    Case "Odd"
        parity = "O"
       
    Case "Even"
        parity = "E"
        
    Case Else
        parity = "N"
End Select
mvparity = parity
End Function

Public Function stopBits(ByVal bits As Variant) As Byte
End Function

Public Function dataBits(ByVal bits As Variant) As Byte
End Function

Public Function speed(ByVal baudRate As Variant) As Integer

End Function

