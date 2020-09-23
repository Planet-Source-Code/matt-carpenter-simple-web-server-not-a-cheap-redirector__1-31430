VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   Caption         =   "Simple Web Server"
   ClientHeight    =   5565
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4290
   LinkTopic       =   "Form1"
   ScaleHeight     =   5565
   ScaleWidth      =   4290
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Custom 404 message..."
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   5040
      Width           =   2055
   End
   Begin VB.Frame Frame2 
      Caption         =   "Activity Log"
      Height          =   1455
      Left            =   120
      TabIndex        =   7
      Top             =   3480
      Width           =   3975
      Begin VB.ListBox List1 
         Height          =   1035
         ItemData        =   "frmmain.frx":0000
         Left            =   120
         List            =   "frmmain.frx":0002
         TabIndex        =   8
         Top             =   240
         Width           =   3735
      End
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   495
      Left            =   2400
      TabIndex        =   0
      Top             =   2640
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmmain.frx":0004
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1800
      Top             =   2520
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   1080
      Top             =   2640
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Height          =   3255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3975
      Begin VB.FileListBox File1 
         Height          =   1455
         Left            =   1440
         Pattern         =   "*.html;*.jpg;*.htm;*.gif"
         TabIndex        =   6
         Top             =   960
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Refresh"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   2880
         Width           =   975
      End
      Begin ComctlLib.ListView ListView1 
         Height          =   1935
         Left            =   240
         TabIndex        =   4
         Top             =   840
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   3413
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   327682
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Files"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.TextBox Text1 
         Height          =   350
         Left            =   1440
         TabIndex        =   3
         Text            =   "C:\"
         Top             =   360
         Width           =   2295
      End
      Begin ComctlLib.ImageList ImageList1 
         Left            =   0
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   327682
         BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
            NumListImages   =   2
            BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmmain.frx":00D9
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmmain.frx":03F3
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Label Label1 
         Caption         =   "Root Directory"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   440
         Width           =   1095
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dImage As Long
Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Dim sResponse As String
Dim HomePageFile As String
Dim msg404 As String



Private Sub Command1_Click()
Text1_Change

End Sub

Private Sub Command2_Click()
msg = InputBox("404 message HTML:")
Open "C:\404.html" For Output As #1
Print #1, msg
Close #1
msg404 = msg

End Sub

Private Sub Form_Load()
Winsock1.LocalPort = 80
Winsock1.Listen
HomePageFile = Text1.Text & "index.html"
File1.Path = Text1.Text
filecount = File1.ListCount

For i = 1 To filecount - 1
  If UCase(Right(File1.List(i), 1)) = "L" Or UCase(Right(File1.List(i), 1)) = "M" Then ListView1.ListItems.Add i, "", File1.List(i), , 1
  If UCase(Right(File1.List(i), 1)) = "F" Or UCase(Right(File1.List(i), 1)) = "G" Then ListView1.ListItems.Add i, "", File1.List(i), , 2
  
  
 Next i
 On Error GoTo new404
 Open "C:\404.html" For Input As #1
 Do While Not EOF(1)
 Input #1, msg404
 Loop
 Close #1
 Exit Sub
new404:
 Open "C:\404.html" For Output As #1
 Print #1, "<h1><i>404 Page could not be found</i></h1>"
 Close #1
 
End Sub

Private Sub ListView1_Click()
Form2.Show
Form2.RichTextBox1.LoadFile Text1.Text & ListView1.SelectedItem.Text

End Sub

Private Sub Text1_Change()
On Error Resume Next
ListView1.ListItems.Clear


File1.Path = Text1.Text
filecount = File1.ListCount
For i = 1 To filecount - 1
  If UCase(Right(File1.List(i), 1)) = "L" Or UCase(Right(File1.List(i), 1)) = "M" Then ListView1.ListItems.Add i, "", File1.List(i), , 1
  If UCase(Right(File1.List(i), 1)) = "F" Or UCase(Right(File1.List(i), 1)) = "G" Then ListView1.ListItems.Add i, "", File1.List(i), , 2
  
  
 Next i
 HomePageFile = Text1.Text & "index.html"
End Sub

Private Sub Timer1_Timer()
'This is the one second timeout
'if the client doesn't request a new document in less than a second
'then it is probably done. If it is, dis-connect so somebody else can connect

Winsock1.Close
Winsock1.Listen

End Sub

Private Sub Winsock1_Close()
Winsock1.Close
Winsock1.Listen

End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
If Winsock1.State <> sckClosed Then Winsock1.Close
Winsock1.Accept requestID
List1.AddItem Winsock1.RemoteHostIP & " connected at " & Time

End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Winsock1.GetData Data, vbString, bytesTotal



Timer1.Enabled = False

If Left(Data, 14) = "GET / HTTP/1.1" Then   'First Connection
RichTextBox1.LoadFile HomePageFile
sResponse = "HTTP/1.1 200 OK" & vbNewLine & _
"Date: Sat, 02 Feb 2002 15:57:05 GMT" & vbNewLine & _
"Server: GWS/1.11" & vbNewLine & _
"Content-Type: text/html" & vbNewLine & _
"Content-Length: " & Len(RichTextBox1.Text) & vbNewLine & _
"Cache-Control: private"
List1.AddItem Winsock1.RemoteHostIP & " requested homepage"
Winsock1.SendData sResponse & vbNewLine & vbNewLine & RichTextBox1.Text & vbNewLine



End If

On Error GoTo show404
arydata = Split(Data, vbNewLine, -1, vbBinaryCompare)
For i = 1 To Len(arydata(0)) - 3
  If UCase(Mid(arydata(0), i, 3)) = "GIF" Or UCase(Mid(arydata(0), i, 3)) = "JPG" Then 'User is requesting a file
    
    arydata2 = Split(arydata(0), " ", -1, vbBinaryCompare)
    imagepath = Text1.Text & Right(arydata2(1), Len(arydata2(1)) - 1)
    RichTextBox1.LoadFile imagepath
    Winsock1.SendData RichTextBox1.Text
    List1.AddItem Winsock1.RemoteHostIP & " requested " & imagepath
  End If
Next i

'Get HTML files
For i = 1 To Len(arydata(0)) - 3
  If UCase(Mid(arydata(0), i, 4)) = "HTML" Or UCase(Mid(arydata(0), i, 3)) = "HTM" Then 'User is requesting a file
    
    arydata2 = Split(arydata(0), " ", -1, vbBinaryCompare)
    imagepath = Text1.Text & Right(arydata2(1), Len(arydata2(1)) - 1)
    Open imagepath For Input As #1
    Do While Not EOF(1)
    Input #1, test
    Loop
    Close #1
    RichTextBox1.LoadFile imagepath
    Winsock1.SendData sResponse & vbNewLine & vbNewLine & RichTextBox1.Text
     List1.AddItem Winsock1.RemoteHostIP & " requested " & imagepath
  End If
Next i


Timer1.Enabled = True
Exit Sub
show404:

sResponse = "HTTP/1.1 200 OK" & vbNewLine & _
"Date: Sat, 02 Feb 2002 15:57:05 GMT" & vbNewLine & _
"Server: GWS/1.11" & vbNewLine & _
"Content-Type: text/html" & vbNewLine & _
"Content-Length: " & Len(msg404) & vbNewLine & _
"Cache-Control: private"
Winsock1.SendData sResponse & vbNewLine & vbNewLine & msg404 & vbNewLine
List1.AddItem Winsock1.RemoteHostIP & " got 404 error at " & Time




End Sub

