VERSION 5.00
Begin VB.Form Form1_б 
   BackColor       =   &H80000013&
   Caption         =   "аккес кеитоуяциес"
   ClientHeight    =   8640
   ClientLeft      =   5445
   ClientTop       =   780
   ClientWidth     =   4515
   LinkTopic       =   "Form10"
   ScaleHeight     =   8640
   ScaleWidth      =   4515
   Begin VB.CommandButton Command5 
      BackColor       =   &H00E9C5AD&
      Caption         =   "енодос"
      Height          =   855
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6480
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00E9C5AD&
      Caption         =   "яухлисг паяалетяым пяоцяаллатос"
      Height          =   855
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5040
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00E9C5AD&
      Caption         =   "тфияои"
      Height          =   855
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3600
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E9C5AD&
      Caption         =   "дглиоуяциа амтицяажым"
      Height          =   855
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2160
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E9C5AD&
      Caption         =   "тгкежымийос йатакоцос"
      Height          =   855
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   720
      Width           =   1575
   End
   Begin VB.Line Line11 
      Visible         =   0   'False
      X1              =   360
      X2              =   1560
      Y1              =   6480
      Y2              =   6480
   End
   Begin VB.Line Line10 
      Visible         =   0   'False
      X1              =   480
      X2              =   1560
      Y1              =   5880
      Y2              =   5880
   End
   Begin VB.Line Line9 
      Visible         =   0   'False
      X1              =   840
      X2              =   840
      Y1              =   5880
      Y2              =   6480
   End
   Begin VB.Line Line8 
      Visible         =   0   'False
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   600
   End
   Begin VB.Line Line7 
      Visible         =   0   'False
      X1              =   1440
      X2              =   720
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Line Line6 
      Visible         =   0   'False
      X1              =   1440
      X2              =   720
      Y1              =   5040
      Y2              =   5040
   End
   Begin VB.Line Line5 
      Visible         =   0   'False
      X1              =   960
      X2              =   960
      Y1              =   4440
      Y2              =   5040
   End
   Begin VB.Line Line4 
      Visible         =   0   'False
      X1              =   960
      X2              =   960
      Y1              =   1560
      Y2              =   2160
   End
   Begin VB.Line Line3 
      Visible         =   0   'False
      X1              =   1440
      X2              =   720
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Line Line2 
      Visible         =   0   'False
      X1              =   1440
      X2              =   720
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Line Line1 
      Visible         =   0   'False
      X1              =   2235
      X2              =   2235
      Y1              =   0
      Y2              =   7200
   End
End
Attribute VB_Name = "Form1_б"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Load Form2
Form2.Show
Form1_б.Hide
Unload Form1_б
End Sub



Private Sub Command2_Click()
Load BACKUP
BACKUP.Show
Form1_б.Hide
Unload Form1_б
End Sub

Private Sub Command3_Click()
Load TZIROI
TZIROI.Show
Form1_б.Hide
Unload Form1_б
End Sub

Private Sub Command4_Click()
Form1_б.Hide
Unload Form1_б
End Sub

Private Sub Command5_Click()
Form1_б.Hide
Unload Form1_б
End Sub

Private Sub Form_Unload(Cancel As Integer)
Form1_б.Hide
Unload Form1_б
End Sub
