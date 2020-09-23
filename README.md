<div align="center">

## Upload a file to an FTP using Winsock


</div>

### Description

It Uploads a file to an FTP using Winsock
 
### More Info
 
Two winsock controls (winsock1 and Winsock2), a command button (command1) and a labal (label1).


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Kristian Trenskow](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/kristian-trenskow.md)
**Level**          |Unknown
**User Rating**    |5.0 (25 globes from 5 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Internet/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-html__1-34.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/kristian-trenskow-upload-a-file-to-an-ftp-using-winsock__1-1783/archive/master.zip)

### API Declarations

```
Type Com
  Reply As String
  BackCommand As String
End Type
```


### Source Code

```
Dim Commun(5) As Com
Dim CommunState As Integer
Dim Site As String
Dim Username As String
Dim Password As String
Dim Remotefile As String
Dim Localfile As String
Dim Buffersize As Long
Dim CloseAfterSend As Boolean
Private Sub Command1_Click()
Site = ""
Username = ""
Password = ""
Localfile = "c:\windows\desktop\view.exe"
Remotefile = "/view.exe"
Commun(0).Reply = "220"
Commun(0).BackCommand = "USER " + Username
Commun(1).Reply = "331"
Commun(1).BackCommand = "PASS " + Password
Commun(2).Reply = "230"
Commun(2).BackCommand = "TYPE I"
Commun(3).Reply = "200"
Commun(3).BackCommand = "PORT"
Commun(4).Reply = "200"
Commun(4).BackCommand = "STOR " + Remotefile
Commun(5).Reply = ""
Commun(5).BackCommand = ""
Buffersize = 2920
Dim Nr1 As Integer
Dim Nr2 As Integer
Dim LocalIP As String
LocalIP = Winsock1.LocalIP
Do Until InStr(LocalIP, ".") = 0
LocalIP = Left(LocalIP, InStr(LocalIP, ".") - 1) + "," + Right(LocalIP, Len(LocalIP) - InStr(LocalIP, "."))
Loop
Randomize Timer
Nr1 = Int(Rnd * 12) + 5
Nr2 = Int(Rnd * 254) + 1
Commun(3).BackCommand = "PORT " + LocalIP + "," + Trim(Str(Nr1)) + "," + Trim(Str(Nr2))
Winsock2.Close
Do Until Winsock2.State = 0
DoEvents
Loop
Winsock2.LocalPort = (Nr1 * 256) + Nr2
Winsock2.Listen
Winsock1.Close
Do Until Winsock1.State = 0
DoEvents
Loop
Winsock1.RemoteHost = Site
Winsock1.RemotePort = 21
Winsock1.Connect
CommunState = 0
Do Until Winsock1.State = 7 Or Winsock1.State = 9
DoEvents
Loop
Select Case Winsock1.State
Case 9
MsgBox "Couldn't reach server " + Site + ".", vbOKOnly + vbInformation, "FTP Upper"
Case 7
Open Localfile For Binary As #1
End Select
End Sub
Private Sub Form_Load()
End Sub
Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim tmpS As String
Winsock1.GetData tmpS, , bytesTotal
Debug.Print tmpS;
Select Case Left(tmpS, 3)
Case Commun(CommunState).Reply
Winsock1.SendData Commun(CommunState).BackCommand + Chr(13) + Chr(10)
Debug.Print Commun(CommunState).BackCommand
CommunState = CommunState + 1
Case "150"
Do Until Winsock2.State = 7
DoEvents
Loop
SendNextData
Case "226"
Winsock1.Close
Do Until Winsock1.State = 0
DoEvents
Loop
MsgBox "Transfer complete.", vbOKOnly + vbInformation, "FTP Upper"
Case Else
MsgBox "Bad reply: " + Left(tmpS, Len(tmpS) - 2), vbOKOnly + vbInformation, "FTP Upper"
End Select
End Sub
Private Sub Winsock2_ConnectionRequest(ByVal requestID As Long)
Winsock2.Close
Do Until Winsock2.State = 0
DoEvents
Loop
Winsock2.Accept requestID
Do Until Winsock2.State = 7
DoEvents
Loop
End Sub
Sub SendNextData()
Dim Take As Long
Dim Buffer As String
If LOF(1) - Seek(1) < Buffersize Then Take = LOF(1) - Seek(1) + 1 Else Take = Buffersize
Buffer = Input(Take, 1)
Winsock2.SendData Buffer
If Take < Buffersize Then
Close #1
CloseAfterSend = True
End If
On Error Resume Next
Label1 = Trim(Str(Seek(1))) + "/" + Trim(Str(LOF(1)))
On Error GoTo 0
End Sub
Private Sub Winsock2_SendComplete()
If CloseAfterSend = True Then
Winsock2.Close
Do Until Winsock2.State = 0
DoEvents
Loop
CloseAfterSend = False
Else
SendNextData
End If
End Sub
```

