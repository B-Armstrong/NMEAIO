VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form NMEAIO 
   Caption         =   "NMEA I/O PANEL"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   6975
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cboPort 
      Height          =   315
      ItemData        =   "NMEAIO.frx":0000
      Left            =   4440
      List            =   "NMEAIO.frx":0010
      TabIndex        =   13
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send &Data"
      Height          =   375
      Left            =   5640
      TabIndex        =   11
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "&Stop"
      Height          =   375
      Left            =   2880
      TabIndex        =   10
      Top             =   2520
      Width           =   1215
   End
   Begin VB.ComboBox cboStop 
      Height          =   315
      ItemData        =   "NMEAIO.frx":0020
      Left            =   2880
      List            =   "NMEAIO.frx":002A
      TabIndex        =   9
      Top             =   600
      Width           =   615
   End
   Begin VB.ComboBox cboData 
      Height          =   315
      ItemData        =   "NMEAIO.frx":0034
      Left            =   2880
      List            =   "NMEAIO.frx":003E
      TabIndex        =   7
      Top             =   120
      Width           =   615
   End
   Begin VB.ComboBox cboParity 
      Height          =   315
      ItemData        =   "NMEAIO.frx":0048
      Left            =   960
      List            =   "NMEAIO.frx":0055
      TabIndex        =   5
      Top             =   600
      Width           =   855
   End
   Begin VB.CommandButton cmdMonitor 
      Caption         =   "&Monitor"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   2520
      Width           =   1215
   End
   Begin VB.ComboBox cboSpeed 
      Height          =   315
      ItemData        =   "NMEAIO.frx":006A
      Left            =   960
      List            =   "NMEAIO.frx":0083
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   6360
      Top             =   -120
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   5520
      Top             =   -120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.TextBox txtData 
      Height          =   1215
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1080
      Width           =   6735
   End
   Begin VB.Label lblPort 
      Caption         =   "Com Port:"
      Height          =   255
      Left            =   3600
      TabIndex        =   12
      Top             =   120
      Width           =   735
   End
   Begin VB.Label lblStop 
      Alignment       =   1  'Right Justify
      Caption         =   "Stop Bits:"
      Height          =   255
      Left            =   2040
      TabIndex        =   8
      Top             =   600
      Width           =   735
   End
   Begin VB.Label lblData 
      Alignment       =   1  'Right Justify
      Caption         =   "Data Bits:"
      Height          =   255
      Left            =   2040
      TabIndex        =   6
      Top             =   120
      Width           =   735
   End
   Begin VB.Label lblParity 
      Alignment       =   1  'Right Justify
      Caption         =   "Parity:"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   600
      Width           =   615
   End
   Begin VB.Label lblSpeed 
      Caption         =   "Baud Rate:"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "NMEAIO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim X As New clsPortInfo
Dim counter As Integer

Private Sub cboData_Click()
dataBits = cboData
MSComm1.Settings = speed & "," & Left(parity, 1) & "," & dataBits & "," & stopBits
End Sub

Private Sub cboParity_Click()
parity = cboParity
MSComm1.Settings = speed & "," & Left(parity, 1) & "," & dataBits & "," & stopBits
End Sub

Private Sub cboPort_Click()
port = cboPort

End Sub

Private Sub cboSpeed_Click()
speed = cboSpeed
MSComm1.Settings = speed & "," & Left(parity, 1) & "," & dataBits & "," & stopBits
End Sub

Private Sub cboStop_Click()
stopBits = cboStop
MSComm1.Settings = speed & "," & Left(parity, 1) & "," & dataBits & "," & stopBits
End Sub

Private Sub cmdMonitor_Click()
 On Error GoTo error
  
If MSComm1.PortOpen = False Then
    MSComm1.CommPort = port
    MSComm1.Settings = speed & "," & Left(parity, 1) & "," & dataBits & "," & stopBits
    MSComm1.PortOpen = True
Else
    MSComm1.Settings = cboSpeed & "," & Left(cboParity, 1) & "," & cboData & "," & cboStop
End If
txtData.ForeColor = vbBlack
txtData = ""
Timer1.Interval = 1000
Timer1.Enabled = True
Timer1_Timer
cboPort.Enabled = False
cboSpeed.Enabled = False
cboParity.Enabled = False
cboData.Enabled = False
cboStop.Enabled = False
cmdSend.Enabled = False
cmdSend.Enabled = False
error:
End Sub


Private Sub cmdSend_Click()
frmDataOUt.Show
Me.Hide
End Sub

Private Sub cmdStop_Click()
On Error GoTo error:
If MSComm1.PortOpen = True Then
MSComm1.PortOpen = False
End If
Timer1.Enabled = False
frmDataOUt.Timer2.Enabled = False
txtData.ForeColor = vbBlue
txtData = "SYSTEM HALTED"
counter = 0
cboPort.Enabled = True
cboSpeed.Enabled = True
cboParity.Enabled = True
cboData.Enabled = True
cboStop.Enabled = True
cmdSend.Enabled = True
cmdMonitor.Enabled = True

error:
End Sub


Private Sub Form_Load()
On Error GoTo error:

cboSpeed = 4800
speed = cboSpeed
cboParity = "None"
parity = cboParity
cboData = 8
dataBits = cboData
cboStop = 1
stopBits = cboStop
cboPort = 1
port = cboPort
counter = 0
Timer1.Enabled = False





error:
End Sub



Private Sub Form_Unload(Cancel As Integer)
If MSComm1.PortOpen = True Then
    MSComm1.PortOpen = False
End If
Unload frmDataOUt

End Sub



Private Sub Timer1_Timer()
On Error GoTo error:
If MSComm1.PortOpen = False Then Exit Sub
    If MSComm1.InBufferCount = 0 Then
        counter = counter + 1
        If counter = 2 Then
            txtData.ForeColor = vbRed
            txtData = ""
            txtData = "!!!NO DATA INPUT!!! check connections. then press monitor."
            Timer1.Enabled = False
            counter = 0
        End If
    Else
txtData = txtData + MSComm1.Input
txtData.SelStart = Len(txtData)
End If
error:

End Sub
