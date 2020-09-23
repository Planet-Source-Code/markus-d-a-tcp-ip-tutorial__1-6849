VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   Caption         =   "Fluid's TCP/IP Tutorial"
   ClientHeight    =   5910
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7125
   LinkTopic       =   "Form1"
   ScaleHeight     =   5910
   ScaleWidth      =   7125
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Height          =   735
      Left            =   3480
      TabIndex        =   15
      Top             =   2280
      Width           =   1455
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   5760
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txtIn 
      Height          =   1815
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   11
      Text            =   "Form1.frx":0000
      Top             =   3720
      Width           =   6855
   End
   Begin VB.TextBox txtOut 
      Height          =   1095
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   10
      Top             =   2280
      Width           =   3255
   End
   Begin VB.TextBox txtLocalPort 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3480
      TabIndex        =   9
      Text            =   "3434"
      Top             =   960
      Width           =   855
   End
   Begin VB.TextBox txtRemotePort 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3480
      TabIndex        =   6
      Text            =   "54321"
      Top             =   600
      Width           =   855
   End
   Begin VB.TextBox txtRemoteIP 
      Height          =   285
      Left            =   2640
      TabIndex        =   4
      Text            =   "127.0.0.1"
      Top             =   240
      Width           =   1695
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CommandButton cmdListen 
      Caption         =   "Listen"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton cmdDisconnect 
      Caption         =   "Disconnect"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1455
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Connect"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "Incoming"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "Outgoing:"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   2040
      Width           =   2295
   End
   Begin VB.Label lblStatus 
      Caption         =   "Disconnected"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   5640
      Width           =   4455
   End
   Begin VB.Label Label3 
      Caption         =   "Local Port:"
      Height          =   255
      Left            =   1680
      TabIndex        =   8
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Remote Port:"
      Height          =   255
      Left            =   1680
      TabIndex        =   7
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Remote IP:"
      Height          =   255
      Left            =   1680
      TabIndex        =   5
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' This program was made by Fluid <fluid@hack3r.com>
' As a tutorial to show people how TCP/IP is done
' Any questions may be emailed to fluid@hack3r.com
' If you use this code, cool, but I'd love to see what you do with it
' You can email me telling about your new programs you've made, I don't bite ;)
'
'Fluid
'http://fluid.hack3r.com

Option Explicit

Private Sub cmdConnect_Click()
lblStatus.Caption = "Connection to: " & txtRemoteIP.Text & ":" & txtRemotePort.Text ' Sets the status
Winsock1.Connect txtRemoteIP.Text, Val(txtRemotePort.Text) ' Connects the client to ip, val(port)
End Sub

Private Sub cmdDisconnect_Click()
Winsock1.Close ' Closes the connection
cmdConnect.Enabled = True ' Changes the buttons status
cmdDisconnect.Enabled = False
cmdListen.Enabled = True
cmdSend.Enabled = False
lblStatus.Caption = "Disconnected" ' Set the new status message
End Sub

Private Sub cmdExit_Click()
Winsock1.Close ' Closes the connection
End ' Exits
End Sub

Private Sub cmdListen_Click()
Winsock1.LocalPort = Val(txtLocalPort.Text) ' Sets the port to listen on
Winsock1.Listen ' Opens port
lblStatus.Caption = "Listening on port: " & txtLocalPort.Text ' Updates status
cmdConnect.Enabled = False ' Changes the buttons status
cmdDisconnect.Enabled = True
cmdListen.Enabled = False
cmdSend.Enabled = False
End Sub

Private Sub cmdSend_Click()
Winsock1.SendData txtOut.Text ' Sends the data to the other end
End Sub

Private Sub Form_Load()
cmdDisconnect.Enabled = False ' Sets the buttons
cmdSend.Enabled = False
End Sub

Private Sub txtIn_Change()
txtIn.SelStart = Len(txtIn.Text) ' Makes it so it scrolls down when more text is added
End Sub

Private Sub Winsock1_Connect()
lblStatus.Caption = "Connected: " & txtRemoteIP.Text & ":" & txtRemotePort.Text ' Sets status
cmdConnect.Enabled = False ' Updates buttons
cmdDisconnect.Enabled = True
cmdListen.Enabled = False
cmdSend.Enabled = True
End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
If (Winsock1.State <> sckClosed) Then Winsock1.Close ' If not closed then call close method to cleanup socket status
Winsock1.LocalPort = 0 ' Clear the localport
Winsock1.Accept requestID ' Accept the incoming connection
txtRemoteIP.Text = Winsock1.RemoteHostIP ' Tells you their ip
txtRemotePort.Text = Winsock1.RemotePort ' Tells you their port
lblStatus.Caption = "Connected: " & txtRemoteIP.Text & ":" & txtRemotePort.Text ' Updates status

End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim INCOMING ' Sets the varible
Winsock1.GetData INCOMING, vbString ' Puts the incoming data into the varible
txtIn.Text = txtIn.Text & INCOMING ' Adds the varible to the text box
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
lblStatus.Caption = "Error: " & Description ' Descibes the error in the status box
End Sub
