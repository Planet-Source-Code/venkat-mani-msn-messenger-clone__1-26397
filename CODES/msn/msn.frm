VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "MSN Demo"
   ClientHeight    =   3600
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4950
   LinkTopic       =   "Form1"
   ScaleHeight     =   3600
   ScaleWidth      =   4950
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   2655
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   4683
      _Version        =   393217
      Style           =   7
      Appearance      =   1
   End
   Begin MSWinsockLib.Winsock Winsock3 
      Left            =   2280
      Top             =   5400
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock2 
      Left            =   2400
      Top             =   4800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   6360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   2520
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   3360
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1920
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      Top             =   1320
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Login"
      Height          =   495
      Left            =   3480
      TabIndex        =   0
      Top             =   2640
      Width           =   1095
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   0
      Top             =   1680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "64.4.13.55"
      RemotePort      =   1863
   End
   Begin VB.Label Label3 
      Caption         =   "Chat Friends"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "password"
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "username"
      Height          =   255
      Left            =   2520
      TabIndex        =   3
      Top             =   1440
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim m
Dim comm As Boolean
Dim tr As Long
Dim cki
Dim ter
Dim term As Boolean
Dim mkr As Long
Dim name4
Dim name2
Dim name3
Dim x(50) As New Form2
Dim mpt As Boolean
'Dim exists1 As Boolean
Dim qwert(30)
Private Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal fEnable As Long) As Long
    Const sBASE_64_CHARACTERS = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
    Dim exists1(50) As Boolean
Private Sub Command1_Click()
Text3.Text = ""

Winsock1.Close
Winsock1.Connect
m = 0
End Sub



Private Sub Command2_Click()
'comm = True
'Winsock1.SendData Text4.Text & vbCrLf
'MsgBox MD5String(Text4.Text)
'm = Left(MD5String(Text4.Text), 1)
'm = Base64encode(MD5String(Text4.Text))
'i = Asc("i")
'Dim p As String
'p = i Xor Asc(m)

End Sub

Private Sub Command3_Click()

  'm = Base64decode(Text4.Text)
 'm = Base64encode(Text4.Text)
End Sub

'Private Sub Command3_Click()
'winsock1.Close
'winsock1.Connect Text5.Text, Text6.Text
'End Sub

'Private Sub Command4_Click()
'Winsock2.SendData "CAL 3 " & Text1.Text & vbCrLf
'End Sub

Private Sub Form_load()
'exists1 = False
tr = 1
mkr = 1

End Sub







Private Sub TreeView1_DblClick()
'For y = 0 To 50
'If x(y).Winsock1.Tag = TreeView1.SelectedItem.Text Then
Winsock1.SendData "XFR " & tr & " SB" & vbCrLf
tr = tr + 1
mkr = mkr + 1
For xx = 1 To 50
If x(xx).Tag = "" Then
Load x(xx)
x(xx).Text1.Text = TreeView1.SelectedItem.Text
x(xx).Text2.Text = name2
x(xx).Command3.Tag = "hotmail"
x(xx).Command1.Tag = "called"
x(xx).Command2.Tag = TreeView1.SelectedItem.Key
'Form2.Text1.Text = TreeView1.SelectedItem.Text
x(xx).Tag = TreeView1.SelectedItem.Text
x(xx).Visible = True
Exit Sub
End If
Next xx
End Sub



Private Sub Winsock1_Close()
If mpt = False Then
MsgBox "You have been Logged out ,Login Again"
mpt = True
TreeView1.Nodes.Clear
Text2.Text = ""
End If
End Sub

Private Sub Winsock1_Connect()
Winsock1.SendData "VER " & tr & " MSNP5 MSNP4 CVRO" & vbCrLf
 tr = tr + 1
m = 0
mpt = False
 End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next

Dim str As String
If comm = True Then
Winsock1.GetData str

comm = False
Exit Sub
End If
If m = o Then
Winsock1.GetData str
Debug.Print str
we = InStr(str, "XFR")
If we > 0 Then
rt = InStrRev(str, ":")
et = Left(str, rt - 1)

End If
Winsock1.SendData "INF " & tr & vbCrLf
Debug.Print "INF " & tr & vbCrLf
tr = tr + 1
'Text3.Text = Text3.Text & vbCrLf & str
'Text3.SelStart = Len(Text3)
m = m + 1
Exit Sub
End If

If m = 1 Then

Winsock1.GetData str
Debug.Print str
we = InStr(str, "XFR")
If we > 0 Then
rt = InStrRev(str, ":")
et = Left(str, rt - 1)

End If
'Text3.Text = Text3.Text & vbCrLf & str
'text3.SelStart = Len(Text3)
Winsock1.SendData "USR " & tr & " MD5 I " & Text1.Text & vbCrLf
Debug.Print "USR " & tr & " MD5 I " & Text1.Text & vbCrLf
tr = tr + 1
m = m + 1
Exit Sub
End If
If m = 2 Then
m = m + 1
Winsock1.GetData str
Debug.Print str
we = InStr(str, "XFR")
If we > 0 Then
rt = InStrRev(str, ":")
et = Left(str, rt - 1)
mp = InStrRev(et, " ")
ipo = Right(et, Len(et) - mp)
mm = 0
Winsock1.Close
Winsock1.Connect ipo, 1863
Exit Sub
End If
mt = InStrRev(str, "S")
xt = Right(str, Len(str) - mt)
xt = Left(xt, Len(xt) - 2)
xt = Right(xt, Len(xt) - 1)
Winsock1.SendData "USR " & tr & " MD5 S " & MD5String(xt & Text2.Text) & vbCrLf
Debug.Print "USR " & tr & " MD5 S " & MD5String(xt & Text2.Text) & vbCrLf
tr = tr + 1
'Text3.Text = Text3.Text & vbCrLf & str
'Text3.SelStart = Len(Text3)
Exit Sub
End If
If m = 3 Then
Winsock1.GetData str
Debug.Print str
mkw = InStr(str, vbCrLf)
name5 = Left(str, mkw - 1)
mpl = InStrRev(str, " ")
name2 = Right(name5, Len(name5) - mpl)
mw = InStr(name2, "%20")

Do While mw > 0
left3 = Left(name2, mw - 1)
right3 = Right(name2, Len(name2) - mw - 2)
name2 = left3 & " " & right3
mw = InStr(name2, "%20")
Loop
Winsock1.SendData "CHG " & tr & " NLN" & vbCrLf
Debug.Print "CHG " & tr & " NLN" & vbCrLf
tr = tr + 1
'Text3.Text = Text3.Text & vbCrLf & str
'Text3.SelStart = Len(Text3)
m = m + 1
Exit Sub
End If
If m = 4 Then
Winsock1.GetData str

'Text3.Text = Text3.Text & vbCrLf & str
'Text3.SelStart = Len(Text3)
m = m + 1
Exit Sub
End If
If m = 5 Then
zo = 0
Winsock1.GetData str
Debug.Print "data at " & str
ad = InStr(str, "NLN")
naw = Right(str, Len(str) - ad)
ad = InStr(naw, "NLN")
naw = Right(naw, Len(naw) - ad)
Do While ad > 0

naw = Right(nawn, Len(nawn) - ad)
'mopq = InStr(naw, "NLN")
'If mopq > 0 Then
pqw = InStr(naw, vbCrLf)
nawe = Left(naw, pqw - 1)
nawn = Right(naw, Len(naw) - pqw)
mpop = InStrRev(nawe, " ")
name11 = Right(nawe, Len(nawe) - mpop)
mw = InStr(name11, "%20")
Do While mw > 0
left3 = Left(name11, mw - 1)
right3 = Right(name11, Len(name11) - mw - 2)
name11 = left3 & " " & right3
mw = InStr(name11, "%20")
DoEvents
Loop
qwert(zo) = name11
zo = zo + 1
'End If
ad = InStr(nawn, "NLN")
DoEvents
Loop
Winsock1.SendData "LST " & tr & " RL" & vbCrLf
ter = tr
tr = tr + 1
'Text3.Text = Text3.Text & vbCrLf & str
'Text3.SelStart = Len(Text3)
m = m + 1
Exit Sub
End If
If m = 6 Then
'winsock1.GetData str

'tr = tr + 1
'Text3.Text = Text3.Text & vbCrLf & str
'Text3.SelStart = Len(Text3)
'm = m + 1
'Exit Sub
'End If
'If m = 7 Then
Winsock1.GetData str
Debug.Print "frm here " & str

If InStr(str, "NLN") Then
p = InStrRev(str, " ")
name5 = Right(str, Len(str) - p)
name5 = Left(name5, Len(name5) - 2)
mw = InStr(name5, "%20")
Do While mw > 0
left3 = Left(name5, mw - 1)
right3 = Right(name5, Len(name5) - mw - 2)
name5 = left3 & " " & right3
mw = InStr(name5, "%20")
Loop
r = 0
'pr = TreeView1.Nodes.Count
mr = TreeView1.Nodes.Count
For xr = 1 To mr
'TreeView1.Nodes.Item(1).Child.Index = xr
test1 = TreeView1.Nodes.Item(xr).Text
If test1 = name5 Then
'If TreeView1.Nodes.Item(xr).Bold = False Then
TreeView1.Nodes.Item(xr).Bold = True
'TreeView1.Nodes.Item(xr).Image = 3
'Exit Sub
'End If
'If TreeView1.Nodes.Item(xr).Bold = True Then
'TreeView1.Nodes.Item(xr).Bold = False
'TreeView1.Nodes.Item(xr).Image = 7
'Exit Sub
'End If
End If
Next xr

'For k = 0 To TreeView1.Nodes.Item(2).Children
' TreeView1.Nodes.Item(2).Child.Index = k
' test = TreeView1.Nodes.Item(2).Child.Text
'If test = name5 Then TreeView1.Nodes.Item(2).Child.Bold = True
'If test = name5 Then TreeView1.Nodes.Item(2).Child.Image = 3
'Next k
End If
If InStr(str, "FLN") Then
p = InStrRev(str, " ")
name5 = Right(str, Len(str) - p)
name5 = Left(name5, Len(name5) - 2)
mw = InStr(name5, "%20")
Do While mw > 0
left3 = Left(name5, mw - 1)
right3 = Right(name5, Len(name5) - mw - 2)
name5 = left3 & " " & right3
mw = InStr(name5, "%20")
Loop
r = 0
'pr = TreeView1.Nodes.Count

mr = TreeView1.Nodes.Count
For xr = 1 To mr
'TreeView1.Nodes.Item(1).Child.Index = xr
test1 = TreeView1.Nodes.Item(xr).Key
If test1 = name5 Then
'If TreeView1.Nodes.Item(xr).Bold = False Then
'TreeView1.Nodes.Item(xr).Bold = True
'TreeView1.Nodes.Item(xr).Image = 3
'Exit Sub
'End If
'If TreeView1.Nodes.Item(xr).Bold = True Then
TreeView1.Nodes.Item(xr).Bold = False
'TreeView1.Nodes.Item(xr).Image = 7
'Exit Sub
'End If
End If
Next xr
'For k = 0 To TreeView1.Nodes.Item(2).Children
' TreeView1.Nodes.Item(2).Child.Index = k
' test = TreeView1.Nodes.Item(2).Child.Key
'If test = name5 Then TreeView1.Nodes.Item(2).Child.Bold = False
'If test = name5 Then TreeView1.Nodes.Item(2).Child.Image = 7

'Next k
End If
If InStr(str, "RNG") Then
mwq = InStr(str, "CKI")
mew = InStr(str, ":")
ip = Left(str, mew - 1)
isw = InStrRev(ip, " ")
ip = Right(ip, Len(ip) - isw)
cals = Right(str, Len(str) - mwq - 3)
ca = InStr(cals, " ")
cals = Left(cals, ca - 1)
fr = InStrRev(str, " ")
name3 = Right(str, Len(str) - fr)
name3 = Left(name3, Len(name3) - 2)
mw = InStr(name3, "%20")
Do While mw > 0
left3 = Left(name3, mw - 1)
right3 = Right(name3, Len(name3) - mw - 2)
name3 = left3 & " " & right3
mw = InStr(name3, "%20")
Loop
mop = Right(str, Len(str) - 4)
map = InStr(mop, " ")
sid = Left(mop, map - 1)


For xx = 0 To 50
If x(xx).Tag = name3 Then
exists1(t) = True
mew = t
GoTo 165
End If
Next xx
For xx = 0 To 50
If x(xx).Tag = "" Then
exists1(xx) = False
mew = xx
GoTo 165
End If
Next xx

165
If exists1(mew) = False Then
 Load x(mew)
 x(mew).Tag = name3
 x(mew).Command3.Tag = "hotmail"
 exists1(mew) = True
 End If
 x(mew).Text3.Text = cals
x(mew).Text4.Text = sid
x(mew).Winsock1.Close
x(mew).Winsock1.Connect ip, 1863
x(mew).Show
x(mew).Text1.Text = name3
x(mew).Text2.Text = name2
Exit Sub
End If
End If
If InStr(str, "MSG") Then
mn = InStr(str, "TypingUser")
If mn = 0 Then
ap = InStrRev(str, " ")
art = Left(str, ap - 1)
ar = InStrRev(art, " ")
art1 = Right(art, Len(art) - ar)
art2 = Left(art, ar - 1)
For xx = 0 To 50
If x(xx).Tag = art1 Then
If x(xx).Visible = False Then x(xx).Visible = True
If x(xx).Visible = True Then
m = InStrRev(str, vbCrLf)
tre = Right(str, Len(str) - m - 1)
x(xx).Visible = True
x(xx).rtb.Text = x(xx).rtb.Text & vbCrLf & tre
Exit Sub
End If
End If
Next xx
'm = InStrRev(str, vbCrLf)
'tre = Right(str, Len(str) - m - 1)
'Form2.rtb.Text = Form2.rtb.Text & vbCrLf & Form2.Text1.Text & " : " & tre
'b = 1
End If
End If
'If term = False Then
'winsock1.SendData "XFR " & tr & " SB" & vbCrLf
'tr = tr + 1
'term = True
'End If
If InStr(str, "LST " & ter) Then
'Text3.Text = Text3.Text & vbCrLf & str
'Debug.Print "ipo " & str
ae = InStr(str, vbCrLf)
Do While ae > 0
left1 = Left(str, ae - 1)
left2 = InStrRev(left1, " ")
name1 = Right(left1, Len(left1) - left2)
emal = Left(left1, left2 - 1)
mo = InStrRev(emal, " ")
emal = Right(emal, Len(emal) - mo)
mw = InStr(name1, "%20")
Do While mw > 0
left3 = Left(name1, mw - 1)
right3 = Right(name1, Len(name1) - mw - 2)
name1 = left3 & " " & right3
mw = InStr(name1, "%20")
Loop
TreeView1.Nodes(1).Bold = True
TreeView1.Nodes.Add , , emal, name1
'List2.AddItem name1
right1 = Right(str, Len(str) - ae)
str = right1
ae = InStr(str, vbCrLf)
DoEvents
Loop
For pk = 0 To zo
mr = TreeView1.Nodes.Count
For xr = 1 To mr
'TreeView1.Nodes.Item(1).Child.Index = xr
test1 = TreeView1.Nodes.Item(xr).Text
If test1 = qwert(pk) Then
'If TreeView1.Nodes.Item(xr).Bold = False Then
TreeView1.Nodes.Item(xr).Bold = True
'TreeView1.Nodes.Item(xr).Image = 3
'Exit Sub
'End If
'If TreeView1.Nodes.Item(xr).Bold = True Then
'TreeView1.Nodes.Item(xr).Bold = False
'TreeView1.Nodes.Item(xr).Image = 7
'Exit Sub
'End If
End If

Next xr
Next pk
StatusBar1.Panels(2).Text = "Connected to MSN"
TreeView1.Nodes.Item(2).Expanded = True
'Text3.Text = Text3.Text & vbCrLf & str
'Text3.SelStart = Len(Text3)
ElseIf InStr(str, "XFR") Then
'Text3.Text = Text3.Text & vbCrLf & str
'Text3.SelStart = Len(Text3)
asl = InStr(str, ":")
ip = Left(str, asl - 1)
aw = InStrRev(ip, " ")
ip = Right(ip, Len(ip) - aw)
mp = InStrRev(str, " ")
auth = Right(str, Len(str) - mp)
auth = Left(auth, Len(auth) - 2)
cki = auth
For xx = 0 To 50
If x(xx).Tag = TreeView1.SelectedItem.Text Then
exists1(xx) = True
mew = xx
GoTo 1455
End If
Next xx
1455
If exists1(mew) = True Then
x(xx).Visible = True
x(xx).Winsock1.Close
x(xx).Winsock1.Connect ip, 1863
x(xx).Text3.Text = cki
exists1(xx) = False
'End If
End If

'x (xx)
End If
'End If
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
m = 1
End Sub


