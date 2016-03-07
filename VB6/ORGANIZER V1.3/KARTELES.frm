VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H80000013&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "йаятекес еяцасиас"
   ClientHeight    =   7125
   ClientLeft      =   3390
   ClientTop       =   1110
   ClientWidth     =   8505
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7125
   ScaleWidth      =   8505
   Begin VB.CommandButton Command4 
      BackColor       =   &H00E9C5AD&
      Caption         =   "епистяожг сто аявийо лемоу"
      Height          =   855
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3960
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00E9C5AD&
      Caption         =   "пяовеияо"
      Height          =   855
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2520
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E9C5AD&
      Caption         =   "йаятека 2"
      Height          =   855
      Left            =   3270
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2520
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E9C5AD&
      Caption         =   "йаятека 1"
      Height          =   855
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2520
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000013&
      Caption         =   "паяайакы епикенте лиа апо тис йаятекес еяцасиас ╧ то пяовеияо тым сглеиысеым"
      Height          =   495
      Left            =   840
      TabIndex        =   1
      Top             =   1320
      Width           =   6855
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000013&
      Caption         =   "епикоцг йаятекас"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1680
      TabIndex        =   0
      Top             =   480
      Width           =   5295
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error GoTo er:
EPIL_KARTELAS = 1
EPIL_KARTELAS_GRAMMA = 0
Load h
h.Show
Form5.Hide
Unload Form5
GoTo TELOS:

er:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
End Sub

Private Sub Command2_Click()
On Error GoTo er:
EPIL_KARTELAS = 2
EPIL_KARTELAS_GRAMMA = 0
Load h
h.Show
Form5.Hide
Unload Form5
GoTo TELOS:

er:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
End Sub

Private Sub Command3_Click()
On Error GoTo er:
EPIL_KARTELAS = 0
Load PROXEIRO
PROXEIRO.Show
Form5.Hide
Unload Form5
GoTo TELOS:

er:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:

End Sub

Private Sub Command4_Click()
On Error GoTo er:
EPIL_KARTELAS_GRAMMA = 0
EPIL_KARTELAS = 0
Form5.Hide
Unload Form5
GoTo TELOS:

er:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
End Sub

Private Sub Form_Load()
On Error GoTo er:
EPIL_KARTELAS_GRAMMA = 0
EPIL_KARTELAS = 0
GoTo TELOS:

er:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
End Sub
