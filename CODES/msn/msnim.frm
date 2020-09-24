VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form2 
   Caption         =   "Instant Message"
   ClientHeight    =   4200
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6135
   Icon            =   "msnim.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   6135
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   1200
      TabIndex        =   12
      Top             =   3480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   2040
      TabIndex        =   11
      Top             =   3480
      Visible         =   0   'False
      Width           =   1335
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   2160
      Top             =   3600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   600
      TabIndex        =   10
      Top             =   240
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   3960
      TabIndex        =   7
      Top             =   240
      Width           =   1935
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   3945
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7752
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Close"
      Height          =   375
      Left            =   5040
      TabIndex        =   5
      Top             =   3480
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Send"
      Default         =   -1  'True
      Height          =   375
      Left            =   3960
      TabIndex        =   4
      Top             =   3480
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ignore User"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   3480
      Width           =   975
   End
   Begin RichTextLib.RichTextBox rtb2 
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   2760
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   1085
      _Version        =   393217
      TextRTF         =   $"msnim.frx":0442
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   390
      Left            =   0
      TabIndex        =   2
      Top             =   2280
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   688
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox rtb 
      Height          =   1455
      Left            =   0
      TabIndex        =   1
      Top             =   720
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   2566
      _Version        =   393217
      ScrollBars      =   2
      TextRTF         =   $"msnim.frx":04C4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label2 
      Caption         =   "To"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   240
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "From"
      Height          =   255
      Left            =   3360
      TabIndex        =   8
      Top             =   240
      Width           =   495
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim try
Dim done As Boolean
Private Sub Command2_Click()


If Command3.Tag = "hotmail" Then
Dim mess As String
Dim mess2 As String
mess = "MIME-Version: 1.0" & vbCrLf & "Content-Type: text/plain; charset=UTF-8" & vbCrLf & "X-MMS-IM-Format: FN=MS%20Shell%20Dlg; EF=; CO=0; CS=0; PF=0" & vbCrLf & vbCrLf & rtb2.Text

mess2 = "MSG " & try & " N " & Len(mess) & vbCrLf & mess
try = try + 1

'Form1.Text3.Text = Form1.Text3.Text & vbCrLf & mess2
Winsock1.SendData mess2
rtb.Text = rtb.Text & vbCrLf & Text2.Text & " : " & rtb2.Text
rtb2.Text = ""
rtb.SelStart = Len(rtb)
End If
'msg = "YCHT" & Chr(0) & Chr(0) & Chr(1) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(&H45) & Chr(0) & Chr(0) & Chr(0) & Chr(Len(Text1.Text) + Len(Text2.Text) + 2 + Len(rtb2.Text)) & Text2.Text & Chr(1) & Text1.Text & Chr(1) & rtb2.Text '& Chr(1) '& "-1:-1" & Chr(1)
'Form1.Winsock2.SendData msg
'rtb.Text = rtb.Text & vbCrLf & Text2.Text & " : " & rtb2.Text
'rtb2.Text = ""
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Form_load()
If done = False Then
If Command3.Tag = "hotmail" Then
Command1.Visible = False
Command2.Enabled = False

End If
done = True
End If
try = 1
End Sub

Private Sub rtb2_KeyDown(KeyCode As Integer, Shift As Integer)
'If Command3.Tag = "hotmail" Then
'mess = "MIME-Version: 1.0" & vbCrLf & "Content-Type: text/x-msmsgscontrol" & vbCrLf & "TypingUser: venkat13@ hotmail.com" & vbCrLf & vbCrLf & vbCrLf
'Winsock1.SendData "MSG " & try & " U " & Len(mess) & vbCrLf & mess
'try = try + 1
'End If

End Sub

Private Sub Winsock1_Close()
Unload Me
End Sub

Private Sub Winsock1_Connect()
StatusBar1.Panels(1).Text = "INITIALIZING CONTACT"
If Command3.Tag = "hotmail" Then
If Command1.Tag = "called" Then
Winsock1.SendData "USR " & try & " " & Form1.Text1.Text & " " & Text3.Text & vbCrLf
try = try + 1
Exit Sub
End If
'mkr = mkr + 1
Winsock1.SendData "ANS " & try & " " & Form1.Text1.Text & " " & Text3.Text & " " & Text4.Text & vbCrLf
try = try + 1
End If
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
If Command3.Tag = "hotmail" Then
Dim str As String
Winsock1.GetData str
If InStr(str, "JOI") Then
Command2.Enabled = True
StatusBar1.Panels(1).Text = " Contacted "
End If
If InStr(str, "USR") Then
Winsock1.SendData "CAL " & try & " " & Command2.Tag & vbCrLf
try = try + 1
Command1.Tag = ""
End If
Debug.Print "im " & str
If InStr(str, "MSG") Then
mn = InStr(str, "text/plain")
If mn = 0 Then StatusBar1.Panels(1).Text = Text1.Text & " is typing a message"
'
'StatusBar1.Panels.Item.Text = Text1.Text & " is typing a message"

If mn > 0 Then
'StatusBar1.SimpleText = ""
StatusBar1.Panels(1).Text = ""
Do While InStr(str, vbCrLf & vbCrLf)
m = InStr(str, vbCrLf & vbCrLf)
cut = Right(str, Len(str) - m - 3)
If InStr(cut, "MSG") Then
If InStr(cut, "MIME-") Then
m = InStr(cut, vbCrLf & vbCrLf)
cut = Right(cut, Len(cut) - m - 3)
End If
End If
If InStr(cut, "TypingUser") Then
mp = InStr(cut, "MSG")
cut = Left(cut, mp - 1)
If cut <> vbCrLf And cut <> "" Then
mkp = InStr(cut, "MIME-")
If mkp = 0 Then rtb.Text = rtb.Text & vbCrLf & Text1.Text & " : " & cut
rtb.SelStart = Len(rtb.Text)
pp = True
End If
End If
'If InStr(cut, "MSG") Then cut = Left(cut, mp - 1)
str = Right(str, Len(str) - m - 1)
If pp = False Then
If cut <> vbCrLf And cut <> "" Then
mkp = InStr(cut, "MIME-")
If mkp = 0 Then rtb.Text = rtb.Text & vbCrLf & Text1.Text & " : " & cut
rtb.SelStart = Len(rtb.Text)
pp = True
End If
End If
'mn = InStr(str, "text/plain")
'If mn = 0 Then Exit Sub
'MsgBox Len(cut)
'mr = InStrRev(cut, vbCrLf & vbCrLf)
'mk = Left(cut, mr - 1)
'mp = InStrRev(cut, vbCrLf & "MIME")
'mt = Right(cut, Len(cut) - mp + 4)
'mt = Left(mt, 3)
'Debug.Print cut
'test = Right(mk, 25)
'tre = Left(str, m - 1)
'b = InStr(tre, vbCrLf)
'mer = Left(tre, b - 1)
'mer = Right(tre, Len(tre) - b)
'mer = Right(mer, 3)
'MsgBox Len(str)
'bre = Right(tre, 1)
'If pp = False Then
'If cut <> vbCrLf And cut <> "" Then rtb.Text = rtb.Text & vbCrLf & Text1.Text & " : " & cut
'End If
'str = Left(str, m - 1)
Loop
End If
End If
End If
End Sub
