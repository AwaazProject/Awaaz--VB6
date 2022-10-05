VERSION 5.00
Begin VB.Form Awaaz 
   BackColor       =   &H008080FF&
   Caption         =   "Awaaz"
   ClientHeight    =   5235
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7770
   LinkTopic       =   "Form1"
   ScaleHeight     =   5235
   ScaleWidth      =   7770
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSpace 
      Caption         =   "Add Space"
      Height          =   495
      Left            =   5520
      TabIndex        =   7
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset"
      Height          =   495
      Left            =   5520
      TabIndex        =   6
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "Select"
      Height          =   495
      Left            =   5520
      TabIndex        =   5
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next"
      Height          =   495
      Left            =   5520
      TabIndex        =   4
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox txtCurrent 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   3
      Text            =   "A"
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton cmdSpeak 
      BackColor       =   &H00000000&
      Caption         =   "Speak"
      Height          =   735
      Left            =   600
      TabIndex        =   1
      Top             =   3360
      Width           =   3615
   End
   Begin VB.TextBox txtSpeak 
      Height          =   1095
      Left            =   720
      TabIndex        =   0
      Top             =   960
      Width           =   4335
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "START TYPING..."
      Height          =   255
      Left            =   2040
      TabIndex        =   8
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Current Word"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   2
      Top             =   2400
      Width           =   1575
   End
End
Attribute VB_Name = "Awaaz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim currentWord, selectedWord As String 'getting, setting and updating variables'
Dim alphabets 'declaring the array to store the alphabets'
Dim i As Integer 'declaring a counter'

Private Sub cmdNext_Click() 'the command to iterate through the alphabets variable'
    i = i + 1 'íncrementing the counter'
    alphabets = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z") 'alphabets array'
    If i <= 25 Then 'iterating the alphabets variable'
        currentWord = alphabets(i) 'setting the value of the variable to current location of the alphabets'
    Else
        i = -1 'when end is reached, counter sets to -1'
    End If
    txtCurrent.Text = currentWord 'setting the word'
End Sub

Private Sub cmdReset_Click()
    txtCurrent.Text = "A" 'sets the value to A'
    i = 0 'setting counter back to 0'
End Sub

Private Sub cmdSelect_Click()
    selectedWord = txtCurrent.Text 'the current letter is set to the current text box'
    txtSpeak.Text = txtSpeak.Text + selectedWord 'the letter is set in the main text box'
End Sub

Private Sub cmdSpace_Click()
    txtSpeak.Text = txtSpeak.Text + " " 'adding space'
End Sub

Private Sub cmdSpeak_Click() 'main speaking function'
    Set voice = CreateObject("SAPI.spvoice")
    voice.speak (txtSpeak.Text)
End Sub

