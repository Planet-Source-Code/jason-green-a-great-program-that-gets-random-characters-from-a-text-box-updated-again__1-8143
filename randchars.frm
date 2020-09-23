VERSION 5.00
Begin VB.Form frmRand 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Take random letters from text"
   ClientHeight    =   2580
   ClientLeft      =   1650
   ClientTop       =   1545
   ClientWidth     =   7110
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   7110
   Begin VB.CheckBox chkInfo 
      Caption         =   "Show Info"
      Height          =   195
      Left            =   120
      TabIndex        =   19
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox txtnum 
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Go!"
      Height          =   495
      Left            =   3480
      TabIndex        =   1
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox txtInput 
      Height          =   975
      Left            =   15
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   480
      Width           =   6960
   End
   Begin VB.Label lblWords 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   2280
      TabIndex        =   18
      Top             =   2640
      Width           =   45
   End
   Begin VB.Label lblVow 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   5760
      TabIndex        =   17
      Top             =   2640
      Width           =   45
   End
   Begin VB.Label lblCon 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   2280
      TabIndex        =   16
      Top             =   2880
      Width           =   45
   End
   Begin VB.Label lblNumcon 
      AutoSize        =   -1  'True
      Caption         =   "Number of consonants:"
      Height          =   195
      Left            =   120
      TabIndex        =   15
      Top             =   2880
      Width           =   1650
   End
   Begin VB.Label lblNumvow 
      AutoSize        =   -1  'True
      Caption         =   "Number of vowels:"
      Height          =   195
      Left            =   3600
      TabIndex        =   14
      Top             =   2640
      Width           =   1320
   End
   Begin VB.Label lblWordnum 
      AutoSize        =   -1  'True
      Caption         =   "Number of words:"
      Height          =   195
      Left            =   120
      TabIndex        =   13
      Top             =   2640
      Width           =   1245
   End
   Begin VB.Label lblNumchar 
      AutoSize        =   -1  'True
      Caption         =   "Number of characters:"
      Height          =   195
      Left            =   3600
      TabIndex        =   12
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Label lblChars 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   5760
      TabIndex        =   11
      Top             =   3120
      Width           =   45
   End
   Begin VB.Label lblNumoth 
      AutoSize        =   -1  'True
      Caption         =   "Number of other characters:"
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   3120
      Width           =   1980
   End
   Begin VB.Label lblOth 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   2280
      TabIndex        =   9
      Top             =   3120
      Width           =   45
   End
   Begin VB.Label lblNumspa 
      AutoSize        =   -1  'True
      Caption         =   "Number of spaces:"
      Height          =   195
      Left            =   3600
      TabIndex        =   8
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Label lblSpa 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   5760
      TabIndex        =   7
      Top             =   2880
      Width           =   45
   End
   Begin VB.Label lblWordis 
      Caption         =   "The word is:"
      Height          =   495
      Left            =   4920
      TabIndex        =   6
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label lblChar 
      Caption         =   "How many characters do you want taken?"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   3135
   End
   Begin VB.Label lblAbout 
      AutoSize        =   -1  'True
      Caption         =   "Just type or copy/paste text into the field below.  Made by: Jason Green"
      Height          =   195
      Left            =   960
      TabIndex        =   3
      Top             =   120
      Width           =   5070
   End
   Begin VB.Label lblWord 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   6000
      TabIndex        =   2
      Top             =   1680
      Width           =   45
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmRand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkInfo_Click()
If chkInfo.Value = 1 Then
frmRand.Height = 3795
Else
frmRand.Height = 3000
End If
End Sub

Private Sub txtInput_Change()
On Error Resume Next
lblWords = 0
lblVow = 0
lblCon = 0
lblChars = 0
lblOth = 0
lblSpa = 0
If txtInput <> "" Then
If Mid(txtInput.Text, 1, 1) <> " " Then
For i = 1 To Len(txtInput.Text)
If Mid(txtInput.Text, i, 1) = " " Then
lblWords.Caption = Val(lblWords.Caption) + 1
End If
Next
For i = 1 To Len(txtInput.Text)
If Mid(txtInput.Text, i, 2) = "  " Then
lblWords.Caption = lblWords.Caption - 1
End If
Next
lblWords = lblWords + 1
For i = 1 To Len(txtInput.Text)
If Mid(txtInput.Text, i, 1) = "a" Or Mid(txtInput.Text, i, 1) = "e" Or Mid(txtInput.Text, i, 1) = "i" Or Mid(txtInput.Text, i, 1) = "o" Or Mid(txtInput.Text, i, 1) = "u" Then
lblVow.Caption = Val(lblVow.Caption) + 1
ElseIf Mid(txtInput.Text, i, 1) = "b" Or Mid(txtInput.Text, i, 1) = "c" Or Mid(txtInput.Text, i, 1) = "d" Or Mid(txtInput.Text, i, 1) = "f" Or Mid(txtInput.Text, i, 1) = "g" Or Mid(txtInput.Text, i, 1) = "h" Or Mid(txtInput.Text, i, 1) = "j" Or Mid(txtInput.Text, i, 1) = "k" Or Mid(txtInput.Text, i, 1) = "l" Or Mid(txtInput.Text, i, 1) = "m" Or Mid(txtInput.Text, i, 1) = "n" Or Mid(txtInput.Text, i, 1) = "p" Or Mid(txtInput.Text, i, 1) = "q" Or Mid(txtInput.Text, i, 1) = "r" Or Mid(txtInput.Text, i, 1) = "s" Or Mid(txtInput.Text, i, 1) = "t" Or Mid(txtInput.Text, i, 1) = "v" Or Mid(txtInput.Text, i, 1) = "w" Or Mid(txtInput.Text, i, 1) = "x" Or Mid(txtInput.Text, i, 1) = "y" Or Mid(txtInput.Text, i, 1) = "z" Then
lblCon.Caption = Val(lblCon.Caption) + 1
ElseIf Mid(txtInput.Text, i, 1) = " " Then
lblSpa = lblSpa + 1
Else
lblOth = lblOth + 1
End If
Next
If Mid(txtInput.Text, Len(txtInput.Text), 1) = " " Then
lblWords = lblWords - 1
End If
End If
End If
If txtInput.Text = "" Then
lblWords = 0
End If
lblChars = Len(txtInput.Text)
End Sub



Private Sub Command1_Click()
On Error Resume Next
lblWord.Caption = ""
Randomize Timer
If txtInput.Text <> "" Then
For X = 1 To Val(txtnum.Text)
random = Int((Rnd * Len(txtInput.Text)) + 1)
lblWord = lblWord + Mid(txtInput.Text, random, 1)
Next X
End If
End Sub


Private Sub Text2_Change()

End Sub


