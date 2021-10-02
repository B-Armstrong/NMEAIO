VERSION 5.00
Begin VB.Form frmDataOUt 
   Caption         =   "Data Output"
   ClientHeight    =   2115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5205
   LinkTopic       =   "Form1"
   ScaleHeight     =   2115
   ScaleWidth      =   5205
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   3840
      TabIndex        =   13
      Top             =   1320
      Width           =   855
   End
   Begin VB.TextBox txtBearing 
      Height          =   285
      Left            =   3720
      TabIndex        =   12
      Top             =   600
      Width           =   975
   End
   Begin VB.TextBox txtRange 
      Height          =   285
      Left            =   3720
      TabIndex        =   10
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox txtHeading 
      Height          =   285
      Left            =   1440
      TabIndex        =   8
      Top             =   1560
      Width           =   975
   End
   Begin VB.TextBox txtLong 
      Height          =   285
      Left            =   1440
      TabIndex        =   6
      Top             =   1080
      Width           =   975
   End
   Begin VB.TextBox txtLat 
      Height          =   285
      Left            =   1440
      TabIndex        =   4
      Top             =   600
      Width           =   975
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Left            =   5040
      Top             =   -120
   End
   Begin VB.ComboBox cboOutput 
      Height          =   315
      ItemData        =   "Form1.frx":0000
      Left            =   1440
      List            =   "Form1.frx":000D
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   495
      Left            =   2760
      TabIndex        =   0
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label lblBearing 
      Alignment       =   2  'Center
      Caption         =   "Bearing"
      Height          =   255
      Left            =   2760
      TabIndex        =   11
      Top             =   600
      Width           =   855
   End
   Begin VB.Label lblRange 
      Alignment       =   2  'Center
      Caption         =   "Range"
      Height          =   255
      Left            =   2760
      TabIndex        =   9
      Top             =   120
      Width           =   855
   End
   Begin VB.Label lblHeading 
      Alignment       =   2  'Center
      Caption         =   "Heading"
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label lblLong 
      Alignment       =   2  'Center
      Caption         =   "Longitude"
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label lblLat 
      Alignment       =   2  'Center
      Caption         =   "Latitude"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label lblOutputType 
      Alignment       =   1  'Right Justify
      Caption         =   "Output Sentence"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmDataOUt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
NMEAIO.Show
Unload Me

End Sub

Private Sub cmdOK_Click()
On Error GoTo error:
NMEAIO.Show
If NMEAIO.MSComm1.PortOpen = False Then
    NMEAIO.MSComm1.CommPort = port
    NMEAIO.MSComm1.Settings = speed & "," & Left(parity, 1) & "," & dataBits & "," & stopBits
    NMEAIO.MSComm1.PortOpen = True
Else
    NMEAIO.MSComm1.Settings = speed & "," & Left(parity, 1) & "," & dataBits & "," & stopBits
End If
Timer2.Enabled = True
Timer2.Interval = 1000
Timer2_Timer
Me.Hide
NMEAIO.cmdMonitor.Enabled = False
NMEAIO.cmdSend.Enabled = False
NMEAIO.cboPort.Enabled = False
NMEAIO.txtData = "!OUTPUTTING NMEA DATA!"
error:
End Sub

Private Sub Form_Load()
cboOutput = "GGA"
txtLat = "30.24"
txtLong = "88.14"
txtHeading = "45.3"

End Sub

Public Sub GGA_OUT()
On Error GoTo error:
Dim output As String
UTCtime (Time$)
output = "$GPGGA," & UTC & "," & latitude & "," & latCord & "," & longitude & "," & lonCord & ",2,04,2.0,0.0,M,0.0,M,3.1,1111*" & _
checksum("GPGGA," & UTC & "," & latitude & "," & latCord & "," & longitude & "," & lonCord & ",2,04,2.0,0.0,M,0.0,M,3.1,1111")
NMEAIO.MSComm1.output = output


'NMEAIO.txtData = output
error:
End Sub

Function UTCtime(ttime)
On Error GoTo error:
ttime = Split(ttime, ":")
ttime(0) = CInt(ttime(0)) + 5
Select Case ttime(0)
            Case 24
             ttime(0) = "0"
            Case 25
             ttime(0) = "1"
            Case 26
             ttime(0) = "2"
            Case 27
             ttime(0) = "3"
            Case 28
             ttime(0) = "4"
End Select
If ttime(0) < 10 Then
    ttime(0) = "0" + ttime(0)
End If
UTC = ttime(0) & ttime(1) & ttime(2) & ".00"
error:
End Function

Private Sub Form_Unload(Cancel As Integer)
Timer2.Enabled = False
If NMEAIO.MSComm1.PortOpen = True Then
    NMEAIO.MSComm1.PortOpen = False
End If

End Sub

Private Sub Timer2_Timer()
On Error GoTo error:
Select Case cboOutput

            Case "GGA"
                latlong
                GGA_OUT
            Case "HDT"
                HDT_OUT
End Select
If NMEAIO.txtData.ForeColor = vbBlack Then
    NMEAIO.txtData.ForeColor = vbGreen
Else
    NMEAIO.txtData.ForeColor = vbBlack
End If
error:
End Sub

Public Function checksum(datastring)
Dim csum As Integer
Dim n As Integer
On Error GoTo error:
'datastring = "GPGGA,UTC,2327.406800,N,09027.406800,W,2,08,2.0,0.0,M,0.0,M,3.1,1111"
csum = 0
For n = 1 To Len(datastring)
    If csum = 0 Then
        csum = Asc(Mid(datastring, n, 1))
    Else
        csum = csum Xor Asc(Mid(datastring, n, 1))
    End If
Next n
checksum = CStr(Hex(csum))
error:
End Function

Public Sub latlong()
Dim tempData
Dim latDeg As String
Dim latMin As String
Dim lonDeg As String
Dim lonMin As String
Dim length As Integer
On Error GoTo error:
tempData = txtLat
tempData = Split(txtLat, ".")
If Left(tempData(0), 1) = "-" Then
    latDeg = Mid(tempData(0), 2, Len(tempData(0)) - 1)
    latCord = "S"
Else
    latDeg = tempData(0)
    latCord = "N"
End If
latMin = "." & tempData(1)
latitude = Format(latDeg & (latMin * 60), "0000.0000")


tempData = txtLong
tempData = Split(txtLong, ".")
If Left(tempData(0), 1) = "-" Then
    If Len(tempData(0)) = 4 Then
        lonDeg = Mid(tempData(0), 2, Len(tempData(0)) - 1)
    Else
        lonDeg = "0" & Mid(tempData(0), 2, Len(tempData(0)) - 1)
    End If
lonCord = "W"
Else
    If Len(tempData(0)) = 3 Then
        lonDeg = tempData(0)
    Else
        lonDeg = "0" & tempData(0)
        lonCord = "E"
    End If
End If
lonMin = ("." & tempData(1)) * 60
If Len(lonMin) < 4 Then
    lonMin = "0" & lonMin
End If
longitude = Format(lonDeg & lonMin, "00000.0000")
error:

End Sub


Public Sub HDT_OUT()
Dim output As String

output = "$HEHDT," & txtHeading & ",T*" & _
checksum("HEHDT," & txtHeading & ",T")
NMEAIO.MSComm1.output = output
End Sub
