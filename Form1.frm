VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "IP Address Sniffer"
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   4815
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   4815
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   1935
      Left            =   0
      TabIndex        =   12
      Top             =   1320
      Width           =   4815
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   1920
         TabIndex        =   19
         Text            =   "700"
         Top             =   840
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Start Scan"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1200
         Width           =   2775
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Pause Scan"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   1560
         Width           =   2775
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   480
         Width           =   2775
      End
      Begin VB.ListBox List1 
         Height          =   1230
         Left            =   3000
         TabIndex        =   13
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "Timeout (atleast 700):"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   870
         Width           =   1695
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "Current IP Address:"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "Servers found at:"
         Height          =   255
         Left            =   3000
         TabIndex        =   14
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4815
      Begin MSWinsockLib.Winsock Winsock1 
         Left            =   960
         Top             =   120
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Left            =   1440
         Top             =   120
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   3360
         TabIndex        =   2
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   3960
         TabIndex        =   1
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "(xxxxx)"
         Height          =   255
         Left            =   3960
         TabIndex        =   11
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "(xxx . xxx . xxx) . (xxx)"
         Height          =   255
         Left            =   2160
         TabIndex        =   10
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "(xxx . xxx . xxx . xxx)"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Start at:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "To"
         Height          =   255
         Left            =   1800
         TabIndex        =   7
         Top             =   540
         Width           =   255
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "End at: (up to 254)"
         Height          =   255
         Left            =   2160
         TabIndex        =   6
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Port:"
         Height          =   255
         Left            =   3960
         TabIndex        =   5
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Comparison
Dim Comparison2
Dim Comparison3
Dim Comparison4
Dim IPToScan As Integer
Dim IPToStopOn As Integer
Dim Paused
Dim TimeOut
Dim StartString

Private Sub Command1_Click()
Winsock1.Close
Text1.Enabled = True
Text2.Enabled = True
Text1.Text = Text2.Text & "." & (IPToScan - 1)
IPToScan = IPToScan - 1
Text3.Enabled = True
Text4.Enabled = True
Text6.Enabled = True
Command2.Enabled = True
Text5.Text = "Paused"
Timer1.Enabled = False
Command1.Enabled = False
Paused = 1
End Sub

Private Sub Command2_Click()
Winsock1.Close
Winsock1.RemotePort = Text4.Text
Timer1.Enabled = True
IPToStopOn = Text3.Text
StartString = Text1.Text
Comparison4 = Text1.Text Like "*[.]*"
Do Until Comparison4 = False
Text1.SelStart = 0
Text1.SelLength = 1
Text1.SelText = ""
Comparison4 = Text1.Text Like "*[.]*"
Loop
If Paused <> 1 Then
IPToScan = Text1.Text
End If
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Text6.Enabled = False
Command2.Enabled = False
Command1.Enabled = True
If Paused <> 1 Then
List1.Clear
End If
If Paused <> 1 Then
Text5.Text = "Starting..."
End If
Paused = 0
Text1.Text = StartString
End Sub

Private Sub Form_Load()
Timer1.Interval = 1000
Form1.Caption = "IP Address Sniffer " & App.Major & "." & App.Minor & App.Revision & " by TheLemon"
End Sub

Private Sub mnuAbout_Click()
MsgBox "This program scans a given range of IP addresses for working servers" & vbCrLf & _
       "on any specified port.  For example, if you enter 127.0.0.1 as the start" & vbCrLf & _
       "address, 127.0.0.254 as the end IP address, and 12345 as the port, then" & vbCrLf & _
       "this program will scan 127.0.0.1, 127.0.0.254, and every address between" & vbCrLf & _
       "(127.0.0.2, 127.0.0.3, 127.0.0.4, etc...) for servers running on the port" & vbCrLf & _
       "12345.  The timeout option should be above 700 if you are going to scan" & vbCrLf & _
       "all 254 addresses, and/or if you have an unclean connection to the net." & vbCrLf & _
       "However, I have had luck scanning up to 65 ports with the timeout at 250." & vbCrLf & _
       "Increase the timeout setting if you experience performance problems.", _
       vbOKOnly, "IP Address Sniffer " & App.Major & "." & App.Minor & App.Revision
End Sub

Private Sub mnuExit_Click()
End
End Sub

Private Sub Text1_Change()
If Timer1.Enabled = False Then
Comparison = Text1.Text Like "*.*.*.*"
 If Comparison = False Then
 Text2.Text = Text1.Text
 End If
Comparison2 = Text2.Text Like "*.*.*[!.]"
Comparison3 = Text2.Text Like "*.*.*.*"
 Do While Comparison3 = True
  Do Until Comparison2 = True
  Text2.SelStart = Len(Text2.Text) - 1
  Text2.SelLength = 1
  Text2.SelText = ""
  Comparison2 = Text2.Text Like "*.*.*[!.]"
  Loop
 Loop
End If
End Sub

Private Sub Text6_Change()
If Text6.Text <> "" Then
Timer1.Interval = Text6.Text
End If
End Sub

Private Sub Text6_Click()
Text6.SelStart = 0
Text6.SelLength = Len(Text6.Text)
End Sub

Private Sub Timer1_Timer()
Text5.Text = Text2.Text & "." & IPToScan

If TimeOut = 1 Then
 If Winsock1.State = sckConnected Then
 Winsock1.Close
 List1.AddItem Text2.Text & "." & IPToScan
 Beep
 TimeOut = 0
 IPToScan = IPToScan + 1
 Else
 Winsock1.Close
 TimeOut = 0
 IPToScan = IPToScan + 1
 End If
End If

If TimeOut <> 1 Then
 If IPToScan <= IPToStopOn Then
 Winsock1.RemoteHost = Text2.Text & "." & IPToScan
 Winsock1.Connect
 TimeOut = 1
 Else
 Text1.Enabled = True
 Text2.Enabled = True
 Text3.Enabled = True
 Text4.Enabled = True
 Text6.Enabled = True
 Command2.Enabled = True
 Text5.Text = "Done"
 Command1.Enabled = False
 Timer1.Enabled = False
 End If
End If
End Sub
