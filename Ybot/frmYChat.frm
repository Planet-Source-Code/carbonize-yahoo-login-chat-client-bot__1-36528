VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmYbot 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Y!Bot"
   ClientHeight    =   2805
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   5175
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdJoin 
      Caption         =   "Join Room"
      Height          =   495
      Left            =   0
      TabIndex        =   6
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox txtLog 
      Height          =   2775
      Left            =   1320
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   0
      Width           =   3855
   End
   Begin MSWinsockLib.Winsock WS 
      Left            =   840
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton CmdLogin 
      Caption         =   "Login to Server"
      Height          =   495
      Left            =   0
      TabIndex        =   4
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox txtRoom 
      Height          =   285
      Left            =   0
      TabIndex        =   3
      Text            =   "Web Design:1"
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton CmdGetCookie 
      Caption         =   "Get Cookie"
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox txtPass 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   0
      PasswordChar    =   "*"
      TabIndex        =   1
      Text            =   "password"
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox txtUser 
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Text            =   "Username"
      Top             =   0
      Width           =   1215
   End
End
Attribute VB_Name = "frmYbot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim GotCookie As Boolean
Dim LoggedIn As Boolean
Dim Connected As Boolean
Public HEADER As String
Dim Cookie As String
Dim Data As String
Public Function AddLog(Log As String)
  txtLog.Text = txtLog.Text & Log & vbCrLf
  txtLog.SelStart = Len(txtLog.Text)
End Function
Private Function GetCookie(User As String, Pass As String)
Dim Luser As String
  GotCookie = False
  
  WS.Close
  
  AddLog "Connecting to [msg.edit.yahoo.com:80]"
  WS.Connect "msg.edit.yahoo.com", 80
  Connected = False

  Do While Connected = False
    DoEvents
  Loop
  Luser = txtUser.Text
  Luser = Replace(Luser, " ", "%20")
  
  AddLog "Connected .. Requesting Cookie..."
  
  WS.SendData "GET /config/ncclogin?.src=bl&login=" + Luser + "&passwd=" + Pass + "&n=1 HTTP/1.0" & _
        vbCrLf & "Accept: */*" & _
            vbCrLf & "Accept: text/html" & vbCrLf & vbCrLf

End Function
Private Function LoginToChatServer(User As String, Server As String, Port As Integer)

  LoggedIn = False
  
  WS.Close
  
  AddLog "Connecting to [" & Server & ":" & Port & "]"
  WS.Connect Server, Port
  Connected = False

  Do While Connected = False
    DoEvents
  Loop
  
  AddLog "Sending Information... "
  
    WS.SendData HEADER & Chr(1) & Chr(0) & Chr(0) & _
        CalcSize(Len(User) + Len(Cookie) + 1) & _
            User & Chr(1) & Cookie
            
  AddLog "Sent..."
  
  Do While LoggedIn = False
    DoEvents
  Loop
            
  AddLog "Confirmation Recieved. Ready to Join Room."
  
End Function
Private Sub CmdGetCookie_Click()
  GetCookie txtUser.Text, txtPass.Text
End Sub

Private Sub CmdJoin_Click()
Join txtRoom.Text
End Sub

Private Sub CmdLogin_Click()
  LoginToChatServer txtUser, "cs7.chat.yahoo.com", 8001
End Sub

Private Sub Command1_Click()
PM txtMess.Text
End Sub

Private Sub Form_Load()
If WS.State <> sckClosed Then WS.Close
HEADER = "YCHT" & Chr(0) & Chr(0) & Chr(1) & Chr(0) & Chr(0) & Chr(0) & Chr(0)
End Sub



Private Sub WS_Connect()

  Connected = True
  
  If GotCookie = False Then
    AddLog "Connected to [msg.edit.yahoo.com:80]"
  Else
    AddLog "Connected to [cs7.chat.yahoo.com:8001]"
  End If

End Sub

Private Sub WS_DataArrival(ByVal bytesTotal As Long)

Dim CookieStart As Integer
Dim CookieEnd As Integer
Dim DataType As Integer

  WS.GetData Data, vbString
  'Debug.Print Data
  
  If GotCookie = False Then       '/ If gotcookie = false then data recieved must be cookie
  If InStr(1, Data, "ERROR: Invalid NCC Login") Then AddLog "Invalid ID / Pass": WS.Close: Exit Sub
    
    AddLog "Cookie Recieved"

    Cookie1Start = InStr(1, Data, "Y=v=")                            '/
    Cookie1End = InStr(Cookie1Start + 1, Data, ";") + 1                '/ PARSE COOKIE
    Cookie1 = Mid(Data, Cookie1Start, Cookie1End - Cookie1Start)      '/
    Cookie2Start = InStr(1, Data, "T=z=")                            '/
    Cookie2End = InStr(Cookie2Start + 1, Data, ";")                 '/ PARSE COOKIE
    Cookie2 = Mid(Data, Cookie2Start, Cookie2End - Cookie2Start)      '/
    Cookie = Cookie1 & " " & Cookie2
    MsgBox Cookie
    AddLog "Cookie Parsed"
    
    GotCookie = True
    
    WS.Close
    
    AddLog "Connection Closed. Login to Chat Server"
    
    Exit Sub
    
  End If
  
  DataType = Asc(Mid(Data, Len(HEADER) + 1, 1))
     
  HandleData DataType, Data
  
End Sub
Private Function HandleData(DataType As Integer, Data As String)

Debug.Print DataType & ": " & Data

    Select Case DataType
    
      Case 1            'LOGIN
        LoggedIn = True
        
      Case 17           'USER ENTERS/JOINED ROOM
      
      Case 18           'USER LEAVES
      
      Case 23           'Invites
      
      Case 33           'USER AWAY
      
      Case 65           'SPEECH
      
      Case 67           'EMOTION
      
      Case 68           'ADVERTISMENT

      Case 69           'PM
      
      Case 104          'Buddylist message
      
      Case 105          'Yahoo Mail
    End Select
    
End Function
