VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUTC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private mvarUTC As String
Public Property Get getUTC() As String
getUTC = mvarUTC
End Property

Public Function UTCTime(tTime) As String
On Error GoTo error:
tTime = Split(tTime, ":")
tTime(0) = CInt(tTime(0)) + 5
Select Case tTime(0)
            Case 24
             tTime(0) = "0"
            Case 25
             tTime(0) = "1"
            Case 26
             tTime(0) = "2"
            Case 27
             tTime(0) = "3"
            Case 28
             tTime(0) = "4"
End Select
If tTime(0) < 10 Then
    tTime(0) = "0" + tTime(0)
End If
mvarUTC = tTime(0) & ":" & tTime(1) & "." & tTime(2)
error:
End Function

