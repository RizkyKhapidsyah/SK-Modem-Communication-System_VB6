VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "msinet.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00CAF4FD&
   Caption         =   "Modem Communication System"
   ClientHeight    =   1560
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5430
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1560
   ScaleWidth      =   5430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "HH:mm"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   4
      EndProperty
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   375
   End
   Begin VB.Timer Timer3 
      Interval        =   100
      Left            =   2160
      Top             =   120
   End
   Begin VB.TextBox Text2 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "HH:mm"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   4
      EndProperty
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   120
      Width           =   375
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   4080
      Top             =   240
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   255
      Left            =   2520
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   855
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   0
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AccessType      =   1
      Protocol        =   4
      URL             =   "http://"
   End
   Begin MSWinsockLib.Winsock WS 
      Left            =   4080
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Left            =   4680
      Top             =   0
   End
   Begin VB.OptionButton Command4 
      BackColor       =   &H00CAF4FD&
      Caption         =   "Disconnect"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   600
      Width           =   1815
   End
   Begin VB.OptionButton Command3 
      BackColor       =   &H00CAF4FD&
      Caption         =   "Connect"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   600
      Width           =   1815
   End
   Begin MSCommLib.MSComm Comm 
      Left            =   5160
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      Handshaking     =   1
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Dis Connect"
      Height          =   375
      Left            =   2640
      TabIndex        =   6
      Top             =   600
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Connect"
      Height          =   375
      Left            =   720
      TabIndex        =   5
      Top             =   600
      Width           =   1815
   End
   Begin VB.Image Image8 
      Height          =   480
      Left            =   4560
      Picture         =   "Form1.frx":000C
      Stretch         =   -1  'True
      Top             =   600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image7 
      Height          =   480
      Left            =   4560
      Picture         =   "Form1.frx":044E
      Stretch         =   -1  'True
      Top             =   600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image6 
      Height          =   480
      Left            =   4560
      Picture         =   "Form1.frx":0890
      Stretch         =   -1  'True
      Top             =   600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image5 
      Height          =   480
      Left            =   4560
      Picture         =   "Form1.frx":0CD2
      Stretch         =   -1  'True
      Top             =   600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image4 
      Height          =   480
      Left            =   4560
      Picture         =   "Form1.frx":1114
      Stretch         =   -1  'True
      Top             =   600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   4560
      Picture         =   "Form1.frx":1556
      Stretch         =   -1  'True
      Top             =   600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   4560
      Picture         =   "Form1.frx":1998
      Stretch         =   -1  'True
      Top             =   600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   4560
      Picture         =   "Form1.frx":1DDA
      Stretch         =   -1  'True
      Top             =   600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Left            =   1200
      TabIndex        =   7
      Top             =   120
      Width           =   105
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Public Declare Function WNetConnectionDialog Lib "mpr.dll" (ByVal hwnd As Long, ByVal dwType As Long) As Long
'Public Declare Function WNetDisconnectDialog Lib "mpr.dll" (ByVal hwnd As Long, ByVal dwType As Long) As Long
'Public Declare Function WNetGetConnection Lib "mpr.dll" Alias "WNetGetConnectionA" ()
'Public Declare Function WNetAddConnection Lib "mpr.dll" Alias "WNetAddConnectionA" (ByVal lpszNetPath As String, ByVal lpszPassword As String, ByVal lpszLocalName As String) As Long
Dim flag As Boolean
Dim fl As String
Public timedial As String

Private Sub Comm_OnComm()
MsgBox "Hi"

'Select Case Comm.CommEvent
'    Case comEvReceive
'        Dim buffer As Variant
'        buffer = Comm.Input
'
End Sub

Private Sub Command1_Click()
On Error GoTo MdmERR:

Comm.CommPort = 2
Comm.Settings = "9600,N,8,1"
Comm.PortOpen = True
'num = InputBox("Enter ISP Number", "Communication", 172341)

Comm.Output = "ATDT 172306" & vbCr
Exit Sub
MdmERR:
    MsgBox Err.Description
    Exit Sub
End Sub

Private Sub Command2_Click()
On Error GoTo MdmERR:

Comm.Output = "ATH" & vbCr
Comm.PortOpen = False
MsgBox "Discommected"
MdmERR:
    MsgBox Err.Description
    Exit Sub

End Sub


Private Sub Command3_Click()
''On errot GoTo MdmERR:
''If Command3.Value = True Then
''Comm.CommPort = 1
''WS.LocalPort = 1
''Comm.Settings = "9600,N,8,1"
''Comm.PortOpen = True
''num = InputBox("Enter ISP Number", "Communication", 172341)
'''Comm.CTSHolding = True
'''Comm.DSRHolding = True
''Comm.DTREnable = True
''Comm.Handshaking = comRTSXOnXOff
''Comm.NullDiscard = True
''Comm.Output = "ATDT 0," & num & vbCr
''
''
''
''
''Timer1.Interval = 1000
''Exit Sub
''End If
''
''Command3.Value = False
''MdmERR:
''    MsgBox Err.Description
''    Exit Sub

If Command3.Value = True Then
Dim X
Dim cname
io:

cname = InputBox("Enter Connection Name", "Gujarat Techno World", "Dial Dishnet")
If cname = "" Then
    GoTo io:
End If
If cname = "" Then
    Command3.Value = False
    X = Shell("rundll32.exe rnaui.dll,RnaDial" & "Dial Dishnet", vbMaximizedFocus)
    SendKeys "{enter}", True
    
    Me.SetFocus
ElseIf cname <> "" Then
    Command3.Value = False
    X = Shell("rundll32.exe rnaui.dll,RnaDial " & cname, vbMaximizedFocus)
    SendKeys "{enter}", True
    Command3.Value = False
    Me.SetFocus
End If
End If
'DoEvents

'''''DoEvents

'SendKeys "{Alt} + {F4}", True
Exit Sub

End Sub

Private Sub Command4_Click()
On Error GoTo MdmERR:
Dim X
'Timer1.Interval = 12000
If Command4.Value = True Then
 SendKeys "{escape}", True
 
End If
Command4.Value = False
MdmERR:
    'MsgBox Err.Description
    Exit Sub
End Sub

Private Sub Form_Load()
flag = True
fl = "A"
End Sub

Private Sub Text1_LostFocus()
If Format(Time, "HH") > 12 Then
    timedial = Val(Text1.Text) + 12
Else
    timedial = Val(Text1.Text)
End If
End Sub

Private Sub Text2_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    timedial = timedial & ":" & Text2.Text
    
End If
End Sub

Private Sub Timer1_Timer()
'SendKeys "{escape}", True
End Sub

Private Sub Timer2_Timer()
Dim timed As String

timed = Format(Time, "HH:MI")

If Format(Time, "H:M") = Format(timedial, "H:M") Then
    If flag = False Then
        Call Dialer
        flag = True
    End If
    
Else
    flag = False
End If
End Sub
Public Sub Dialer()
Dim X
Dim cname
cname = InputBox("Enter Connection Name", "Gujarat Techno World", "icenet")
io:

cname = InputBox("Enter Connection Name", "Gujarat Techno World", "icenet")
If cname = "" Then
    GoTo io:
End If

If cname = "" Then
    Command3.Value = False
    X = Shell("rundll32.exe rnaui.dll,RnaDial" & "icenet", vbHide)
    SendKeys "{enter}", True
    
    Me.SetFocus
ElseIf cname <> "" Then
    Command3.Value = False
    X = Shell("rundll32.exe rnaui.dll,RnaDial " & cname, vbHide)
    SendKeys "{enter}", True
    Command3.Value = False
    Me.SetFocus
End If

End Sub

Private Sub Timer3_Timer()
Dim kcounter As Integer

Dim imgstr As String
imgstr = "Moon"
Image1.Picture = Icon
If fl = "A" Then
    Call VIS
    Image8.Visible = True
    fl = "B"
ElseIf fl = "B" Then
    Call VIS
    Image8.Visible = True
    fl = "C"
ElseIf fl = "C" Then
    Call VIS
    Image7.Visible = True
    fl = "D"
ElseIf fl = "D" Then
    Call VIS
    Image6.Visible = True
    fl = "E"
ElseIf fl = "E" Then
    Call VIS
    Image5.Visible = True
    fl = "F"
ElseIf fl = "F" Then
    Call VIS
    Image4.Visible = True
    fl = "G"
ElseIf fl = "G" Then
    Call VIS
    Image3.Visible = True
    fl = "H"
ElseIf fl = "H" Then
    Call VIS
    Image2.Visible = True
    fl = "A"
''ElseIf fl = "I" Then
''    Call VIS
''    Image7.Visible = True
''    fl = "J"
''ElseIf fl = "J" Then
''    Call VIS
''    Image6.Visible = True
''    fl = "K"
''ElseIf fl = "K" Then
''    Call VIS
''    Image5.Visible = True
''    fl = "L"
''ElseIf fl = "L" Then
''    Call VIS
''    Image4.Visible = True
''    fl = "M"
''ElseIf fl = "M" Then
''    Call VIS
''    Image3.Visible = True
''    fl = "N"
''ElseIf fl = "N" Then
''    Call VIS
''    Image2.Visible = True
''    fl = "A"
End If


End Sub
Private Sub VIS()
Image1.Visible = False
Image2.Visible = False
Image3.Visible = False
Image4.Visible = False
Image5.Visible = False
Image6.Visible = False
Image7.Visible = False
Image8.Visible = False
End Sub
