VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtSendData 
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   120
      Width           =   3255
   End
   Begin MSWinsockLib.Winsock tcpServer 
      Left            =   1440
      Top             =   2280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   840
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
tcpServer.Protocol = sckTCPProtocol
tcpServer.LocalPort = 1002
tcpServer.Listen

End Sub

Private Sub tcpServer_ConnectionRequest(ByVal requestID As Long)
If tcpServer.State <> sckClosed Then _
tcpServer.Close
' Accept the request with the requestID
' parameter.
tcpServer.Accept requestID
End If
End Sub

Private Sub tcpServer_DataArrival(ByVal bytesTotal As Long)
(ByVal bytesTotal As Long)
' Declare a variable for the incoming data.
' Invoke the GetData method and set the Text
' property of a TextBox named txtOutput to
' the data.
Dim strData As String
tcpServer.GetData strData
txtOutput.Text = strData

End Sub

Private Sub txtSendData_Change()
' The TextBox control named txtSendData
' contains the data to be sent. Whenever the user
' types into the  textbox, the  string is sent
' using the SendData method.

tcpServer.SendData txtSendData.Text

End Sub
