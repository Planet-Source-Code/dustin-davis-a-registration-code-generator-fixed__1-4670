VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Registration code generator"
   ClientHeight    =   1770
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6090
   LinkTopic       =   "Form1"
   ScaleHeight     =   1770
   ScaleWidth      =   6090
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Clear"
      Height          =   375
      Left            =   4440
      TabIndex        =   10
      Top             =   1320
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Copy code to Clipboard"
      Height          =   375
      Left            =   2520
      TabIndex        =   9
      Top             =   1320
      Width           =   1815
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   1560
      TabIndex        =   8
      Text            =   "7"
      Top             =   1320
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Get Code"
      Height          =   255
      Left            =   4440
      TabIndex        =   4
      Top             =   480
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Check Code"
      Height          =   255
      Left            =   4440
      TabIndex        =   3
      Top             =   960
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1080
      TabIndex        =   2
      Top             =   960
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Top             =   480
      Width           =   3255
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Enter Code Length"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Enter Code"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   795
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Enter name"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Reg Code:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   765
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Regcode Generator v1.0
'Author: Dustin DAvis
'Bootleg Software Inc.
'http://www.warpnet.org/bsi
'
'I created this code and put it up to help people create
'registration codes for their programs. It is very simple
'to use and to figure out!
'Do not steal this code!! Please give me proper credit for it
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Amritanshu Gupta (tanshu@i.am) -- I Changed the text1_keypress event
'Removed the bug -- Pl. give me some credit too
Dim regCode As String

Private Sub Command1_Click()
Dim temp As String
Dim length As Integer
length = CInt(Text4.Text)
temp = Left$(regCode, length)
If Text2.Text = temp Then
    MsgBox "Pass"
Else
    MsgBox "Failed"
End If
End Sub

Private Sub Command2_Click()
Dim temp As String
Dim length As Integer
length = CInt(Text4.Text)
temp = Left$(regCode, length)
Label1.Caption = "Reg Code for:" & Text1.Text & "-" & temp
End Sub

Private Sub Command3_Click()
Dim temp As String
Dim length As Integer
length = CInt(Text4.Text)
temp = Left$(regCode, length)
Clipboard.SetText temp
End Sub

Private Sub Command4_Click()
Clipboard.SetText ""
regCode = ""
Text1.Text = ""
Text2.Text = ""
End Sub

Private Sub Form_Load()
regCode = ""
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
'Rewrote it -- Works much better and bugfree
If KeyAscii <= 26 Then Exit Sub
Dim I As Integer
regCode = ""
For I = 1 To Len(Text1.Text)
    regCode = regCode & Asc(Mid(Text1.Text, I, 1))
Next
regCode = regCode & KeyAscii
If (Len(Text1.Text) + 1) <> 0 Then Text4.Text = (Len(Text1.Text) + 1) Else Text4.Text = 7
Label1.Caption = regCode
End Sub

Private Sub Text4_Change()
Dim temp As String
Dim length As Integer

If Text4.Text <> "" Then
    length = CInt(Text4.Text)
    temp = Left$(regCode, length)
    Label1.Caption = "Reg Code for:" & Text1.Text & "-" & temp
Else
    DoEvents
End If
End Sub
