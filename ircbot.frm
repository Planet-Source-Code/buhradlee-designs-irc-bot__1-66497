VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "Mswinsck.ocx"
Begin VB.Form ircbott 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "#Channel Bot by Sniper"
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   8340
   Icon            =   "ircbot.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   314
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   556
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   1455
      Left            =   4200
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   8
      Top             =   3000
      Width           =   3855
   End
   Begin VB.ListBox List1 
      Height          =   840
      ItemData        =   "ircbot.frx":08CA
      Left            =   1320
      List            =   "ircbot.frx":08CC
      TabIndex        =   7
      Top             =   2520
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   5400
      TabIndex        =   6
      Top             =   2520
      Width           =   1815
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   360
      Top             =   3000
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "google.com"
      RemotePort      =   80
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Log Raw"
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Log Chan"
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   720
      Width           =   1095
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   120
      Top             =   1440
   End
   Begin VB.TextBox txtBuffer 
      Height          =   2295
      Index           =   1
      Left            =   1200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   0
      Width           =   7095
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Chan Text"
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   2
      Top             =   360
      Width           =   1095
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Raw Text"
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1095
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   480
      Top             =   1920
   End
   Begin VB.TextBox txtBuffer 
      Height          =   2295
      Index           =   0
      Left            =   1200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   7095
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   120
      Top             =   1920
   End
   Begin VB.Menu main 
      Caption         =   "main"
      Visible         =   0   'False
      Begin VB.Menu show 
         Caption         =   "show"
      End
      Begin VB.Menu exit 
         Caption         =   "exit"
      End
   End
End
Attribute VB_Name = "ircbott"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
      Private Type NOTIFYICONDATA
       cbSize As Long
       hwnd As Long
       uId As Long
       uFlags As Long
       uCallBackMessage As Long
       hIcon As Long
       szTip As String * 64
      End Type
      Private Const NIM_ADD = &H0
      Private Const NIM_MODIFY = &H1
      Private Const NIM_DELETE = &H2
      Private Const NIF_MESSAGE = &H1
      Private Const NIF_ICON = &H2
      Private Const NIF_TIP = &H4
      Private Const WM_MOUSEMOVE = &H200
      Private Const WM_LBUTTONDOWN = &H201     'Button down
      Private Const WM_LBUTTONUP = &H202
      Private Const WM_LBUTTONDBLCLK = &H203
      Private Const WM_RBUTTONDOWN = &H204
      Private Const WM_RBUTTONUP = &H205
      Private Const WM_RBUTTONDBLCLK = &H206

      Private Declare Function SetForegroundWindow Lib "user32" _
      (ByVal hwnd As Long) As Long
      Private Declare Function Shell_NotifyIcon Lib "shell32" _
      Alias "Shell_NotifyIconA" _
      (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
      Private nid As NOTIFYICONDATA
Private Type InData
   Nick As String
   Host As String
   Act As String
   To As String
   msg As String
   Err As Boolean
End Type
Private Type QOCINFO
  dwSize As Long
  dwFlags As Long
  dwInSpeed As Long
  dwOutSpeed As Long
End Type
Private Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Integer
Private Declare Function InternetConnect Lib "wininet.dll" Alias "InternetConnectA" (ByVal hInternetSession As Long, ByVal sServerName As String, ByVal nServerPort As Integer, ByVal sUserName As String, ByVal sPassword As String, ByVal lService As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
Private Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
Private Declare Function InternetOpenUrl Lib "wininet.dll" Alias "InternetOpenUrlA" (ByVal hInternetSession As Long, ByVal sURL As String, ByVal sHeaders As String, ByVal lHeadersLength As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
Private Declare Function InternetReadFile Lib "wininet.dll" (ByVal hFile As Long, ByVal sBuffer As String, ByVal lNumBytesToRead As Long, lNumberOfBytesRead As Long) As Integer
Private Declare Function IsNetworkAlive Lib "SENSAPI.DLL" (ByRef lpdwFlags As Long) As Long
Private Declare Function IsDestinationReachable Lib "SENSAPI.DLL" Alias "IsDestinationReachableA" (ByVal lpszDestination As String, ByRef lpQOCInfo As QOCINFO) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function FindWindowX Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SetFileAttributes Lib "kernel32.dll" Alias "SetFileAttributesA" (ByVal lpFileName As String, ByVal dwFileAttributes As Long) As Long
Private Const BUFFER_LEN = 256
Private Const PassiveConnection As Boolean = True
Public IRCCHANNEL As String
Public IRCCHANNEL2 As String
Public IRCCHANNEL3 As String
Public IRCCHANNEL4 As String
Public IRCCHANNEL5 As String
Public IRCSERVER As String
Public IRCUSERNAME As String
Public IRCCHANNELPASS As String
Public IRCNICK As String
Public NSPASS As String
Public CPass As String
Dim c As InData
Dim minute As Integer
Dim Complete As Integer
Dim WithEvents sckIRC As CSocketMaster
Attribute sckIRC.VB_VarHelpID = -1
Dim Parts() As String
Dim ball(5) As String
Dim chan As String
Dim GChan As String

Private Sub Check1_Click()
If Check1.Value = 1 Then
Timer2.Enabled = True
ElseIf Check1.Value = 0 Then
Timer2.Enabled = False
End If
End Sub
Private Sub Check2_Click()
If Check2.Value = 1 Then
Timer3.Enabled = True
ElseIf Check2.Value = 0 Then
Timer3.Enabled = False
End If
End Sub
Private Sub Form_Load()
    IRCSERVER = ReadIniFile(App.Path & "\Server.ini", "IrcBot", "Server", "")
    IRCCHANNEL = ReadIniFile(App.Path & "\Server.ini", "IrcBot", "Channel", "")
    IRCCHANNEL2 = ReadIniFile(App.Path & "\Server.ini", "IrcBot", "Channel2", "")
    IRCCHANNEL3 = ReadIniFile(App.Path & "\Server.ini", "IrcBot", "Channel3", "")
    IRCCHANNEL4 = ReadIniFile(App.Path & "\Server.ini", "IrcBot", "Channel4", "")
    IRCCHANNEL5 = ReadIniFile(App.Path & "\Server.ini", "IrcBot", "Channel5", "")
    IRCCHANNELPASS = ReadIniFile(App.Path & "\Server.ini", "IrcBot", "Chanpass", "")
    CPass = ReadIniFile(App.Path & "\Server.ini", "IrcBot", "Cpass", "")
    IRCNICK = ReadIniFile(App.Path & "\Server.ini", "IrcBot", "Nick", "")
    IRCUSERNAME = ReadIniFile(App.Path & "\Server.ini", "IrcBot", "User", "")
    NSPASS = ReadIniFile(App.Path & "\Server.ini", "IrcBot", "Npass", "")
    Set sckIRC = New CSocketMaster
          With sckIRC
             .RemoteHost = IRCSERVER
             .RemotePort = 6667
             .Connect
          End With
           Me.show
       Me.Refresh
       With nid
        .cbSize = Len(nid)
        .hwnd = Me.hwnd
        .uId = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallBackMessage = WM_MOUSEMOVE
        .hIcon = Me.Icon
        .szTip = "Ircbot" & vbNullChar
       End With
       Shell_NotifyIcon NIM_ADD, nid
       Me.WindowState = vbMinimized
          End Sub


Private Sub BotTriggers(irctext As String, name As String, bottype As Integer)
    If bottype = 0 Then ' Normal Text Response
       Dim pname As String
       pname = "." & IRCNICK
       If Left$(irctext, Len(pname)) = pname Then 'check for individual command
          If Len(irctext) > Len(pname) + 1 Then
             Dim tempstring As String
             tempstring = Right(irctext, Len(irctext) - Len(pname) - 1)
             irctext = tempstring
          End If
       End If
    End If
    Dim channel As String
    Dim nickname As String
    nickname = c.Nick
    channel = c.To
If Left$(LCase(irctext), 7) = "!admins" Then
Dim a3 As String
a3 = ReadIniFile(App.Path & "\Server.ini", "IrcBot", "Admins", "")
If name = a3 Then
sckIRC.SendData "PRIVMSG " & channel & " " & a3 & vbCrLf
    End If
    End If
If Left$(LCase(irctext), 9) = "!commands" Then
Dim cmd As String
cmd = "!say, !google, !en, !dictionary, !thes, !8ball, !basic <number>, !real <number>, !app <number>, !flag <entry> <reason>, +k50 <YourTextHere>, -k50 <ž¡ž>"
sckIRC.SendData "PRIVMSG " & channel & " " & cmd & vbCrLf
End If
If Left$(LCase(irctext), 5) = "!join" Then
Dim a999 As String
a999 = ReadIniFile(App.Path & "\Server.ini", "IrcBot", "Admins", "")
If a999 = nickname Then
If Len(irctext) > 6 Then
Dim something As String
something = Right(irctext, Len(irctext) - 6)
sckIRC.SendData "JOIN " & something & vbCrLf
End If
End If
End If
If Left$(LCase(irctext), 8) = "!killbot" Then
If nickname = "Sniper" Then
End
Else
sckIRC.SendData "PRIVMSG " & channel & " your not an admin so ugh no!" & a3 & vbCrLf
End If
End If
If Left$(LCase(irctext), 4) = "!raw" Then
Dim a99 As String
a99 = ReadIniFile(App.Path & "\Server.ini", "IrcBot", "Admins", "")
If a99 = nickname Then
 If Len(irctext) > 5 Then
                Dim rawtext As String
                rawtext = Right(irctext, Len(irctext) - 5)
                sckIRC.SendData rawtext & vbCrLf
             End If
             End If
             End If
             If Left$(LCase(irctext), 4) = "!say" Then
 If Len(irctext) > 5 Then
                Dim rawr As String
                rawr = Right(irctext, Len(irctext) - 5)
               sckIRC.SendData "PRIVMSG " & channel & " " & rawr & vbCrLf
             End If
             End If
If Left$(LCase(irctext), 7) = "!google" Then
  If Len(irctext) > 8 Then
    Dim gogo As String
    gogo = Right(irctext, Len(irctext) - 8)
    Text1.Text = gogo
    If (Winsock1.State <> sckClosed) Then Winsock1.Close
  Winsock1.Connect
    GChan = channel
  End If
End If
If Left$(LCase(irctext), 11) = "!dictionary" Then
If Len(irctext) > 12 Then
Dim dic As String
Dim dictionary As String
dic = Right(irctext, Len(irctext) - 12)
dic = Replace(dic, " ", "+")
dictionary = "http://dictionary.reference.com/search?q=" & dic
SendText (dictionary)
Else
SendText (nickname & ", do you have something to lookup in the dictionary?")
End If
End If
If Left$(LCase(irctext), 5) = "!thes" Then
If Len(irctext) > 6 Then
Dim dik As String
Dim dictonary As String
dik = Right(irctext, Len(irctext) - 6)
dik = Replace(dik, " ", "+")
dictonary = "http://thesaurus.reference.com/search?q=" & dik
SendText (dictonary)
Else
SendText (nickname & ", !thes is short for thesarus...")
End If
End If
If Left$(LCase(irctext), 3) = "!en" Then
If Len(irctext) > 4 Then
Dim dikl As String
Dim dictonar As String
dikl = Right(irctext, Len(irctext) - 4)
dikl = Replace(dikl, " ", "+")
dictonar = "http://www.reference.com/search?q=" & dikl
SendText (dictonar)
Else
SendText (nickname & ", do you have something to lookup in the encyclopedia?")
End If
End If
If Left$(LCase(irctext), 3) = "!me" Then
If Len(irctext) > 4 Then
Dim ttt As String
ttt = Right(irctext, Len(irctext) - 4)
sckIRC.SendData "PRIVMSG " & channel & " :ACTION " & ttt & vbCrLf
End If
End If
If Left$(LCase(irctext), 6) = "!8ball" Then
If Len(irctext) > 7 Then
Dim ba As String
Dim B As Integer
ba = Right(irctext, Len(irctext) - 7)
B = Int(Rnd * 5)
          ball(0) = nickname & ", Yes"
          ball(1) = nickname & ", No"
          ball(2) = nickname & ", maybe"
          ball(3) = nickname & ", Hell No"
          ball(4) = nickname & ", Hell yes"
          ball(5) = nickname & ", i can't answer that"
Dim nnn As String
nnn = ball(B)
SendText (nnn)
End If
End If
If Left$(LCase(irctext), 6) = "!rules" Then
Dim a10 As String
a10 = ReadIniFile(App.Path & "\Server.ini", "IrcBot", "Rules", "")
SendText (a10)
End If
If Left$(LCase(irctext), 6) = "!basic" Then
If Len(irctext) > 7 Then
Dim bas As String
Dim klkl As String
klkl = Right(irctext, Len(irctext) - 7)
bas = ReadIniFile(App.Path & "\Server.ini", "Basic", klkl, "")
SendText (bas)
Else
SendText (nickname & ", the syntax must be entered as !basic mission number, for example !basic 1")
End If
End If
If Left$(LCase(irctext), 5) = "!real" Then
If Len(irctext) > 6 Then
Dim ko As String
Dim kp As String
kp = Right(irctext, Len(irctext) - 6)
ko = ReadIniFile(App.Path & "\Server.ini", "Real", kp, "")
SendText (ko)
Else
SendText (nickname & ", the syntax must be entered as !real mission number, for example !real 1")
End If
End If
If Left$(LCase(irctext), 4) = "!app" Then
If Len(irctext) > 5 Then
Dim bas2 As String
Dim klkl2 As String
klkl2 = Right(irctext, Len(irctext) - 5)
bas2 = ReadIniFile(App.Path & "\Server.ini", "APP", klkl, "")
SendText (bas2)
Else
SendText (nickname & ", the syntax must be entered as !basic mission number, for example !app 1")
End If
End If
If Left$(LCase(irctext), 5) = "!flag" Then
If Len(irctext) > 6 Then
Dim lll As String
lll = Right(irctext, Len(irctext) - 6)
WriteIniFile App.Path & "\Server.ini", "Flag", nickname, lll
SendText (nickname & ", your comment has been noted and will be open for review soon.")
Else
SendText ("Syntax goes as follows, !flag <entry> <reason> minus the <> of course")
End If
End If

If Left$(LCase(irctext), 4) = "+k50" Then
If Len(irctext) > 5 Then
Dim encript As String
encript = Right(irctext, Len(irctext) - 5)
Call encrypt(encript)
SendText (nickname & ", the text you wanted encrypted is " & encript)
Else
SendText (nickname & ", +k50 is a descent encryption, example, +k50 <YourTextHere>")
End If
End If
If Left$(LCase(irctext), 4) = "-k50" Then
If Len(irctext) > 5 Then
Dim decript As String
decript = Right(irctext, Len(irctext) - 5)
Call decrypt(decript)
SendText (nickname & ", the encrypted text means, " & decript)
Else
SendText (nickname & ", -k50 is for decrypting the k50 encryption... example, -k50 ž¡ž")
End If
End If
If Left$(LCase(irctext), 3) = "!op" Then
If Len(irctext) > 4 Then
Dim opp As String
opp = Right(irctext, Len(irctext) - 4)
sckIRC.SendData "MODE " & channel & " +oooooooooooooooo " & opp & vbCrLf
End If
End If
If Left$(LCase(irctext), 6) = "!owner" Then
If Len(irctext) > 7 Then
Dim own As String
own = Right(irctext, Len(irctext) - 7)
sckIRC.SendData "MODE " & channel & " +qqqqqqqqqqqqqqq " & own & vbCrLf
End If
End If
If Left$(LCase(irctext), 8) = "!protect" Then
If Len(irctext) > 9 Then
Dim pro As String
pro = Right(irctext, Len(irctext) - 9)
sckIRC.SendData "MODE " & channel & " +aaaaaaaaaaaaaaa " & pro & vbCrLf
End If
End If
If Left$(LCase(irctext), 6) = "!voice" Then
If Len(irctext) > 7 Then
Dim voi As String
voi = Right(irctext, Len(irctext) - 7)
sckIRC.SendData "MODE " & channel & " +vvvvvvvvvvvvvvvv " & voi & vbCrLf
End If
End If
If Left$(LCase(irctext), 7) = "!!halfop" Then
If Len(irctext) > 8 Then
Dim halfo As String
halfo = Right(irctext, Len(irctext) - 8)
sckIRC.SendData "MODE " & channel & " +hhhhhhhhhhhhhhhhhh " & halfo & vbCrLf
End If
End If
If Left$(LCase(irctext), 8) = "!devoice" Then
If Len(irctext) > 9 Then
Dim devoi As String
devoi = Right(irctext, Len(irctext) - 9)
sckIRC.SendData "MODE " & channel & " -vvvvvvvvvvvvvvvvvvvvvv " & devoi & vbCrLf
End If
End If
If Left$(LCase(irctext), 9) = "!dehalfop" Then
If Len(irctext) > 10 Then
Dim deh As String
deh = Right(irctext, Len(irctext) - 10)
sckIRC.SendData "MODE " & channel & " -hhhhhhhhhhhhhhhhhhhhh " & deh & vbCrLf
End If
End If
If Left$(LCase(irctext), 10) = "!deprotect" Then
If Len(irctext) > 11 Then
Dim depro As String
depro = Right(irctext, Len(irctext) - 11)
sckIRC.SendData "MODE " & channel & " -aaaaaaaaaaaaaaaaaaaaa " & depro & vbCrLf
End If
End If
If Left$(LCase(irctext), 5) = "!deop" Then
If Len(irctext) > 6 Then
Dim deop As String
deop = Right(irctext, Len(irctext) - 6)
sckIRC.SendData "MODE " & channel & " -oooooooooooooooooooooo " & deop & vbCrLf
End If
End If
If Left$(LCase(irctext), 8) = "!deowner" Then
If Len(irctext) > 9 Then
Dim deowner As String
deowner = Right(irctext, Len(irctext) - 9)
sckIRC.SendData "MODE " & channel & " -qqqqqqqqqqqqqqq " & deowner & vbCrLf
End If
End If
If Left$(LCase(irctext), 5) = "!kick" Then
If Len(irctext) > 6 Then
Dim kick As String
kick = Right(irctext, Len(irctext) - 6)
sckIRC.SendData "KICK " & channel & " " & kick & vbCrLf
End If
End If
If Left$(LCase(irctext), 4) = "!ban" Then
If Len(irctext) > 5 Then
Dim ban3 As String
ban3 = Right(irctext, Len(irctext) - 5)
sckIRC.SendData "MODE " & channel & " +bbbbbbbbbbbbbbb " & ban3 & vbCrLf
End If
End If
If Left$(LCase(irctext), 3) = "!kb" Then
If Len(irctext) > 4 Then
Dim kokl As String
Dim ban1 As String
kokl = Right(irctext, Len(irctext) - 4)
sckIRC.SendData "MODE " & channel & " +bbbbbbbbbbbbbbb " & kokl & vbCrLf
sckIRC.SendData "KICK " & channel & " " & kokl & vbCrLf
End If
End If
If Left$(LCase(irctext), 6) = "!unban" Then
If Len(irctext) > 7 Then
Dim ban2 As String
ban2 = Right(irctext, Len(irctext) - 7)
sckIRC.SendData "MODE " & channel & " -bbbbbbbbbbbbbbbbbbbb " & ban2 & vbCrLf
End If
End If
If Left$(LCase(irctext), 6) = "!login" Then
If Len(irctext) > 7 Then
Dim cc As String
cc = Right(irctext, Len(irctext) - 7)
sckIRC.SendData "IDENTIFY " & cc & " " & CPass & vbCrLf
End If
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    sckIRC.CloseSck
    Shell_NotifyIcon NIM_DELETE, nid
Savechanlog
Saverawlog
Unload Me
End
End Sub
Private Sub Option1_Click(Index As Integer)
    If Option1(0).Value = True Then
       txtBuffer(0).Visible = True
       txtBuffer(1).Visible = False
    Else
       txtBuffer(1).Visible = True
       txtBuffer(0).Visible = False
    End If
End Sub
Private Sub sckIRC_Connect()
    With sckIRC
        .SendData "NICK " & IRCNICK & " " & vbCrLf
        .SendData "USER " & IRCUSERNAME & " " & sckIRC.LocalHostName & " " & UCase(sckIRC.LocalHostName & ":" & sckIRC.LocalPort & "/0") & ": " & IRCUSERNAME & vbCrLf
        .SendData "IDENTIFY " & NSPASS & " " & vbCrLf
    End With
    Timer1.Enabled = True
End Sub
Private Sub sckIRC_CloseSck()
          With sckIRC
             .RemoteHost = IRCSERVER
             .RemotePort = 6667
             .Connect
          End With
End Sub
Private Function Unspool(rawirc As String) As Variant
Dim items() As String
items = Split(rawirc, vbCrLf)
Unspool = items
End Function
Private Sub sckIRC_DataArrival(ByVal bytesTotal As Long)
    Dim sRecv As String
    sckIRC.GetData sRecv
    ' Need to unspool grouped commands
    Dim i As Integer, Search As Integer, LineCount As Integer, LastSearch As Integer
    For i = 1 To Len(sRecv)
       Search = InStr(i, sRecv, vbCrLf, 0)
       If Search <> 0 And Search < Len(sRecv) Then
          LineCount = LineCount + 1
          i = Search
          LastSearch = Search
       End If
    Next i
    If LineCount > 0 Then
       Dim rawtext() As String
       rawtext = Unspool(sRecv)
       For i = 0 To LineCount
          Process_IRC (rawtext(i))
       Next i
    Else
       Process_IRC (sRecv)
    End If

End Sub
Private Sub Process_IRC(rawirc As String)
    On Error Resume Next

    Dim sRecv As String
    sRecv = rawirc
    sRecv = Replace(sRecv, vbCrLf, vbNullString)
    c = ParseInData(sRecv)
    If LCase(c.To) <> LCase(IRCCHANNEL) Then
       If InStr(sRecv, "PING") <> 0 Then
          sckIRC.SendData Replace(sRecv, "PING", "PONG") & vbCrLf
          Exit Sub
       End If
    End If
    txtBuffer(0).Text = txtBuffer(0).Text & sRecv & vbCrLf 'rawtext
    txtBuffer(0).SelStart = Len(txtBuffer(0).Text)
    If c.Err = True Then Exit Sub
    Dim TotalText As String
    Dim nickname As String
    Dim channel As String
    Dim Server As String
    Dim Action As String
    TotalText = c.msg
    nickname = c.Nick
    channel = c.To
    Server = c.Host
    Action = c.Act
    If LCase(channel) = LCase(IRCCHANNEL) Then 'channel text
       txtBuffer(1).Text = txtBuffer(1).Text & nickname & ": " & TotalText & vbCrLf
       txtBuffer(1).SelStart = Len(txtBuffer(1).Text) & vbCrLf
       End If
    If LCase(channel) = LCase(channel) Then
          Call BotTriggers(TotalText, vbNullString, 0)
       End If
    
            If LCase(c.To) <> LCase(IRCCHANNEL) Then
       If InStr(sRecv, "JOIN") <> 0 Then
       Dim admin007 As String
       admin007 = ReadIniFile(App.Path & "\Server.ini", "IrcBot", "Admins", "")
          Exit Sub
       End If
    End If
               If LCase(c.Nick) <> LCase(IRCCHANNEL) Then
       If InStr(sRecv, "VERSION") <> 0 Then
          sckIRC.SendData "NOTICE " & nickname & " im version 2 of Snipers #vb bot" & vbCrLf
          Exit Sub
       End If
    End If
       If LCase(c.Nick) <> LCase(IRCCHANNEL) Then
       If InStr(sRecv, "TIME") <> 0 Then
          sckIRC.SendData "NOTICE " & nickname & " " & Time & vbCrLf
          Exit Sub
       End If
    End If
           If LCase(c.Nick) <> LCase(IRCCHANNEL) Then
       If InStr(sRecv, "!time") <> 0 Then
          sckIRC.SendData "NOTICE " & nickname & " " & Time & vbCrLf
          Exit Sub
       End If
    End If
       If LCase(c.Nick) <> LCase(IRCCHANNEL) Then
       If InStr(sRecv, "DCC SEND") <> 0 Then
          sckIRC.SendData "PRIVMSG " & nickname & " I dont accept your shit, now piss off." & vbCrLf
          sckIRC.SendData "DCC DENY "
          Exit Sub
       End If
    End If
           If LCase(c.Nick) <> LCase(IRCCHANNEL) Then
       If InStr(sRecv, "DCC CHAT") <> 0 Then
          sckIRC.SendData "NOTICE " & nickname & " psh you can talk to me in the channel bitch." & vbCrLf
          Exit Sub
       End If
    End If
              If LCase(c.Nick) <> LCase(IRCCHANNEL) Then
       If InStr(sRecv, "FINGER") <> 0 Then
          sckIRC.SendData "NOTICE " & nickname & " zomg you touched me..... do it again XD!" & vbCrLf
          Exit Sub
       End If
    End If
    'sckIRC.SendData "IDENTIFY " & cc & " " & CPass & vbCrLf
              If LCase(c.Nick) <> LCase(IRCCHANNEL) Then
       If InStr(sRecv, "!login") <> 0 Then
       If Len(InStr(sRecv, Nothing)) > 7 Then
       Dim login As String
       login = InStr(sRecv, Nothing) - 7
        sckIRC.SendData "IDENTIFY " & login & " " & CPass & vbCrLf
        Exit Sub
       End If
    End If
    End If
                  If LCase(c.Nick) <> LCase(IRCCHANNEL) Then
       If InStr(sRecv, "!raw") <> 0 Then
       If Len(InStr(sRecv, Nothing)) > 4 Then
       Dim rawtext As String
      rawtext = InStr(sRecv, Nothing) - 4
        sckIRC.SendData rawtext & vbCrLf
        Exit Sub
       End If
    End If
    End If
End Sub
Private Sub Timer1_Timer()
                If sckIRC.State = sckConnected Then
       With sckIRC
         .SendData "identify " & IRCCHANNEL2 & " " & CPass & " " & vbCrLf
       End With
    End If
    If sckIRC.State = sckConnected Then
       With sckIRC
         .SendData "JOIN " & IRCCHANNEL & " " & IRCCHANNELPASS & " " & vbCrLf
       End With
    End If
        If sckIRC.State = sckConnected Then
       With sckIRC
         .SendData "JOIN " & IRCCHANNEL2 & " " & IRCCHANNELPASS & " " & vbCrLf
       End With
    End If
            If sckIRC.State = sckConnected Then
       With sckIRC
         .SendData "JOIN " & IRCCHANNEL3 & " " & IRCCHANNELPASS & " " & vbCrLf
       End With
    End If
            If sckIRC.State = sckConnected Then
       With sckIRC
         .SendData "JOIN " & IRCCHANNEL4 & " " & IRCCHANNELPASS & " " & vbCrLf
       End With
    End If
            If sckIRC.State = sckConnected Then
       With sckIRC
         .SendData "JOIN " & IRCCHANNEL5 & " " & IRCCHANNELPASS & " " & vbCrLf
       End With
    End If
    Timer1.Enabled = False
End Sub
Private Function SendText(FLOODTEXT As String)
    Dim channel As String
    channel = c.To
Dim name As String
   If sckIRC.State = sckConnected Then
      sckIRC.SendData "PRIVMSG " & name & channel & " :" & FLOODTEXT & vbCrLf
   End If
End Function
Private Function sendme(flood As String)
   If sckIRC.State = sckConnected Then
      sckIRC.SendData "PRIVMSG " & IRCCHANNEL & " :ACTION " & flood & vbCrLf
   End If
End Function
Private Function ParseInData(dta As String) As InData
    On Error GoTo ABigError
    Dim a As String
    Dim B As String
    If Mid$(dta, 1, 1) = ":" Then
       dta = Mid$(dta, 2)
       a = Mid$(dta, 1, InStr(dta, " ") - 1)
       If a Like "*!*@*" Then
          ParseInData.Host = Trim$(Mid$(a, InStr(a, "!") + 1))
          ParseInData.Nick = Trim$(Mid$(a, 1, InStr(a, "!") - 1))
          ParseInData.msg = Mid$(dta, InStr(dta, ":") + 1)
          B = Trim$(Left$(Mid$(dta, InStr(InStr(dta, " "), dta, " ")), InStr(Mid$(dta, InStr(InStr(dta, " "), dta, " ")), ":") - 1))
          If B = "JOIN" Then
                   Call BotTriggers(vbNullString, ParseInData.Nick, 1)
                Exit Function
          End If
          ParseInData.Act = Trim$(Mid$(B, 1, InStr(B, " ") - 1))
          ParseInData.To = Trim$(Mid$(B, InStr(B, " ") + 1))
       Else
          ParseInData.Host = a
          ParseInData.Nick = a
          ParseInData.msg = Mid$(dta, InStr(dta, ":") + 1)
          B = Trim$(Left$(Mid$(dta, InStr(InStr(dta, " "), dta, " ")), InStr(Mid$(dta, InStr(InStr(dta, " "), dta, " ")), ":") - 1))
          ParseInData.Act = Trim$(Mid$(B, 1, InStr(B, " ") - 1))
          ParseInData.To = Trim$(Mid$(B, InStr(B, " ") + 1))
       End If
    ElseIf Mid$(dta, 1, 6) = "ERROR:" Then
       ParseInData.Err = True
    Else
       ParseInData.Err = True
    End If
    Exit Function
ABigError:
    ParseInData.Err = True
    
End Function
Public Sub Savechanlog()
Dim InFile
Dim strmonitoring As String
Dim strsavename As String
strmonitoring = strmonitoring & vbCrLf & vbCrLf & vbCrLf
strmonitoring = txtBuffer(1).Text
On Error Resume Next
InFile = FreeFile
strsavename = "C:\" & "chanlog " & Format(Date, " mm-dd-yy") & ".txt"
Open strsavename For Output As InFile
    Print #InFile, strmonitoring
Close InFile
End Sub
Public Sub Saverawlog()
Dim InFile
Dim strmonitoring As String
Dim strsavename As String
strmonitoring = strmonitoring & vbCrLf & vbCrLf & vbCrLf
strmonitoring = txtBuffer(0).Text
On Error Resume Next
InFile = FreeFile
strsavename = "C:\" & "rawlog " & Format(Date, " mm-dd-yy") & ".txt"
Open strsavename For Output As InFile
    Print #InFile, strmonitoring
Close InFile
End Sub
Private Sub Timer2_Timer()
Savechanlog
End Sub
      Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
      Dim result As Long
      Dim msg As Long
       If Me.ScaleMode = vbPixels Then
        msg = x
       Else
        msg = x / Screen.TwipsPerPixelX
       End If
       Select Case msg
        Case WM_LBUTTONUP
         Me.WindowState = vbNormal
         result = SetForegroundWindow(Me.hwnd)
         Me.show
        Case WM_LBUTTONDBLCLK
         Me.WindowState = vbNormal
         result = SetForegroundWindow(Me.hwnd)
         Me.show
        Case WM_RBUTTONUP
         result = SetForegroundWindow(Me.hwnd)
         Me.PopupMenu main
       End Select
      End Sub
  Private Sub Form_Resize()
       If Me.WindowState = vbMinimized Then Me.Hide
      End Sub
Private Sub show_click()
   Dim result As Long
       Me.WindowState = vbNormal
       result = SetForegroundWindow(Me.hwnd)
       Me.show
End Sub
Private Sub exit_click()
Unload Me
End
End Sub
Private Function encrypt(k As String)
    Dim n As Integer, i As Integer
    n = 155
    For i = 1 To Len(k)
        Mid(k, i, 1) = Chr((Asc(Mid(k, i, 1)) + n) Mod 255)
    Next i
    encrypt = k
End Function
Function decrypt(k As String)
    Dim n As Integer, i As Integer
    n = 155
    For i = 1 To Len(k)
        Mid(k, i, 1) = Chr((Asc(Mid(k, i, 1)) - n) Mod 255)
    Next i
    decrypt = k
End Function

Private Sub Timer3_Timer()
Saverawlog
End Sub
Public Function stripshit(x As String)
x = Replace(x, "<b>", "")
x = Replace(x, "</b>", "")
stripshit = x
End Function
Public Function getcontent(ByVal Html As String) As String
Dim sc As CStrCat
Dim m As Match
Dim xxi As Integer
Dim regex As RegExp
Set sc = New CStrCat
Set regex = New RegExp
regex.IgnoreCase = True
regex.Global = True
regex.Pattern = "<a title=\x22[^>]*\x22 href=([^>]*)>(.*?)</a>"
sc.MaxLength = Len(Html)
For Each m In regex.Execute(Html)
  List1.AddItem stripshit(m.SubMatches(1)) & " :: " & stripshit(m.SubMatches(0))
  sc.AddStr stripshit(m.SubMatches(1)) & " :: " & stripshit(m.SubMatches(0)) & vbCrLf
Next
getcontent = sc
End Function
Private Sub Winsock1_Connect()
  Winsock1.SendData _
  "GET /ie?q=" & Replace(Text1, " ", "+") & " HTTP/1.1" & vbCrLf & _
  "Host: www.google.com" & vbCrLf & _
  "User-Agent: Mozilla/5.0" & vbCrLf & _
  "Keep-Alive: 300" & vbCrLf & _
  "Connection: Keep -Alive" & vbCrLf & vbCrLf
End Sub
Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
  Dim Data As String
  Winsock1.GetData Data
  Data = Replace(Data, vbCrLf, vbCrLf)
     If getcontent(Data) = vbNullString Then
     Exit Sub
     Else
sckIRC.SendData "PRIVMSG " & GChan & " :" & List1.List(0) & vbCrLf
End If
List1.Clear
End Sub


