VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLatLon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private mvarLat As String
Private mvarLon As String
Private mvarLatCoord As String
Private mvarLonCoord As String

Public Property Get getLatCoord() As String
getLatCoord = mvarLatCoord
End Property

Public Property Get getLonCoord() As String
getLonCoord = mvarLonCoord
End Property
Public Property Get getLat() As String
getLat = mvarLat
End Property
Public Property Get getLon() As String
getLon = mvarLon
End Property
Public Function latitude(vLat As String) As String
Dim tempData As Variant
Dim latDeg As String
Dim latMin As String
On Error GoTo error:
tempData = vLat
tempData = Split(vLat, ".")
If Left(tempData(0), 1) = "-" Then
    latDeg = Mid(tempData(0), 2, Len(tempData(0)) - 1)
    mvarLatCoord = "S"
Else
    latDeg = tempData(0)
    mvarLatCoord = "N"
End If
latMin = "." & tempData(1)
mvarLat = Format(latDeg & (latMin * 60), "0000.0000")
error:
End Function

Public Function longitude(vLon As String) As String
Dim tempData As Variant
Dim lonDeg As String
Dim lonCord As String
Dim lonMin As String
On Error GoTo error:
tempData = vLon
tempData = Split(vLon, ".")
If Left(tempData(0), 1) = "-" Then
    If Len(tempData(0)) = 4 Then
        lonDeg = Mid(tempData(0), 2, Len(tempData(0)) - 1)
    Else
        lonDeg = "0" & Mid(tempData(0), 2, Len(tempData(0)) - 1)
    End If
mvarLonCoord = "W"
Else
    If Len(tempData(0)) = 3 Then
        lonDeg = tempData(0)
    Else
        lonDeg = "0" & tempData(0)
        mvarLonCoord = "E"
    End If
End If
lonMin = ("." & tempData(1)) * 60
If Len(lonMin) < 4 Then
    lonMin = "0" & lonMin
End If
mvarLon = Format(lonDeg & lonMin, "00000.0000")
error:

End Function

