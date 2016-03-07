VERSION 5.00
Begin VB.Form Form9 
   BackColor       =   &H80000013&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "диавеияисг амтицяажым"
   ClientHeight    =   6840
   ClientLeft      =   2295
   ClientTop       =   2580
   ClientWidth     =   10875
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   10875
   Begin VB.CommandButton Command4 
      BackColor       =   &H00E9C5AD&
      Caption         =   "епистяожг сто пяогцоулемо лемоу"
      Height          =   735
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5880
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00E9C5AD&
      Caption         =   "етаияиым йаи тилокоциым"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4350
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4440
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E9C5AD&
      Caption         =   "глеяокоциоу"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4350
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2880
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E9C5AD&
      Caption         =   "тгкежымийоу йатакоцоу"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4350
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1320
      Width           =   2295
   End
   Begin VB.Line Line1 
      Visible         =   0   'False
      X1              =   5497
      X2              =   5497
      Y1              =   0
      Y2              =   6720
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2010
      TabIndex        =   0
      Top             =   240
      Width           =   6975
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error GoTo er:
If Form9.Label1.Caption = "диацяажг амтицяажым" Then
    Load Form10THL
    Form10THL.Show
End If
If Form9.Label1.Caption = "диавеияисг амтицяажым" Then
    Load Form11THL
    Form11THL.Show
End If
If Form9.Label1.Caption = "амтийатастасг ле амтицяажа" Then
    Load Form12THL
    Form12THL.Show
End If
Form9.Enabled = False
GoTo TELOS:

er:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
End Sub

Private Sub Command2_Click()
On Error GoTo er:
If Form9.Label1.Caption = "диацяажг амтицяажым" Then
    Load Form10HMER
    Form10HMER.Show
End If
If Form9.Label1.Caption = "диавеияисг амтицяажым" Then
    Load Form11HMER
    Form11HMER.Show
End If
If Form9.Label1.Caption = "амтийатастасг ле амтицяажа" Then
    Load Form12HMER
    Form12HMER.Show
End If
Form9.Enabled = False
GoTo TELOS:

er:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
End Sub

Private Sub Command3_Click()
On Error GoTo er:
If Form9.Label1.Caption = "диацяажг амтицяажым" Then
    Load Form10ETAIRIES
    Form10ETAIRIES.Show
    
End If
If Form9.Label1.Caption = "диавеияисг амтицяажым" Then
    Load Form11ETAIRIES
    Form11ETAIRIES.Show
End If
If Form9.Label1.Caption = "амтийатастасг ле амтицяажа" Then
    Load Form12ETAIRIES
    Form12ETAIRIES.Show
End If
Form9.Enabled = False
GoTo TELOS:

er:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
End Sub

Private Sub Command4_Click()
Form9.Hide
Unload Form9
BACKUP.Enabled = True
BACKUP.Show
End Sub

Private Sub Form_Load()
' "диацяажг амтицяажым"
'"амтийатастасг ле амтицяажа"
'"диавеияисг амтицяажым"

Label1.Caption = LABEL_FOR_BACKUP

End Sub

Private Sub Form_Unload(Cancel As Integer)
BACKUP.Enabled = True
BACKUP.Show
End Sub
