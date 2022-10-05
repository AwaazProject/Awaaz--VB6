VERSION 5.00
Begin VB.Form Output 
   Caption         =   "Output"
   ClientHeight    =   5820
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7455
   LinkTopic       =   "Form1"
   ScaleHeight     =   5820
   ScaleWidth      =   7455
   StartUpPosition =   3  'Windows Default
   Begin VB.Label txtOutput 
      Alignment       =   2  'Center
      Height          =   3615
      Left            =   120
      TabIndex        =   1
      Top             =   2040
      Width           =   7335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "What was Typed"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1920
      TabIndex        =   0
      Top             =   360
      Width           =   3615
   End
End
Attribute VB_Name = "Output"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Text1_Change()

End Sub

