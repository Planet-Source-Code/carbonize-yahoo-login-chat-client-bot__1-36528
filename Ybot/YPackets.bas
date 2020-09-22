Attribute VB_Name = "YPackets"
Public Function SendD(Data As String)
With frmYbot
  .WS.SendData .HEADER & Chr(65) & Chr(0) & Chr(0) & CalcSize(Len(.txtRoom.Text) + Len(Data) + 16) & .txtRoom.Text & Chr(1) & Data & "À€KyUaFhCoHoA!T" 'Cheeta: "À€[$m" Yahelite: '"À€[~w2~tJDLINFHANFCOAAAA~m"
End With
End Function
Public Function Away(Number As Integer)
With frmYbot
  .WS.SendData .HEADER & Chr(33) & Chr(0) & Chr(0) & CalcSize(Len(.txtUser)) & Chr(192) & Chr(128) & Number & Chr(192) & Chr(128) 'À€
End With
End Function
Public Function CustomAway(Message)
With frmYbot
  .WS.SendData .HEADER & Chr(33) & Chr(0) & Chr(0) & CalcSize(Len(.txtUser) + Len(Message)) & Chr(192) & Chr(128) & "10" & Chr(192) & Chr(128) & Message 'À€
End With
End Function
Public Function Think(Data As String)
With frmYbot
  .WS.SendData .HEADER & Chr(67) & Chr(0) & Chr(0) & CalcSize(Len(.txtRoom.Text) + Len(Data) + 24) & .txtRoom.Text & Chr(1) & ". o 0 (" & Data & ")" & "À€KyUaFhCoHoA!T" 'Cheeta: "À€[$m" Yahelite: '"À€[~w2~tJDLINFHANFCOAAAA~m"
End With
End Function
Public Function Emote(Data As String)
With frmYbot
  .WS.SendData .HEADER & Chr(67) & Chr(0) & Chr(0) & CalcSize(Len(.txtRoom.Text) + Len(Data) + 16) & .txtRoom.Text & Chr(1) & Data & "À€KyUaFhCoHoA!T" 'Cheeta: "À€[$m" Yahelite: '"À€[~w2~tJDLINFHANFCOAAAA~m"
End With
End Function
Public Function Join(Room As String)
  With frmYbot
    .WS.SendData .HEADER & Chr(17) & Chr(0) & Chr(0) & CalcSize(Len(.txtRoom.Text)) & Room
  End With
End Function
Public Function GotoUser(User As String)
  With frmYbot
    .WS.SendData .HEADER & Chr(37) & Chr(0) & Chr(0) & CalcSize(Len(User)) & User
  End With
End Function
Public Function Ping()
  With frmYbot
    .WS.SendData .HEADER & Chr(98) & Chr(0) & Chr(0) & Chr(0) & Chr(0)
  End With
End Function
Public Function LogOut()
  With frmYbot
    .WS.SendData .HEADER & Chr(2) & Chr(0) & Chr(0)
  End With
End Function

Public Function Invite(User As String)
  With frmYbot
    .WS.SendData .HEADER & Chr(23) & Chr(0) & Chr(0) & CalcSize(Len(User)) & User
  End With
End Function
Public Function PM(User2PM As String, Data As String)
  With frmYbot
    .WS.SendData .HEADER & Chr(69) & Chr(0) & Chr(0) & CalcSize(Len(.txtUser) + Len(User2PM) + Len(Data) + 2) & .txtUser & Chr(1) & User2PM & Chr(1) & Data
    End With
End Function
Public Function Leave()
  With frmYbot
  .WS.SendData .HEADER & Chr(18) & Chr(0) & Chr(0) & CalcSize(Len(.txtRoom.Text)) & Room
  End With
End Function
Public Function CalcSize(PckLen As Integer) As String
Dim FstNum As String
  FstNum = 0
    Do While PckLen > 255
      FstNum = FstNum + 1
      PckLen = PckLen - 256
    Loop
  CalcSize = Chr$(FstNum) & Chr$(PckLen)
End Function

