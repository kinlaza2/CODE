VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form1 
   BackColor       =   &H80000013&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "аявийг"
   ClientHeight    =   11085
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   16020
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   11085
   ScaleMode       =   0  'User
   ScaleWidth      =   20050.06
   ShowInTaskbar   =   0   'False
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   600
      Top             =   7920
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   120
      TabIndex        =   13
      Text            =   "3"
      Top             =   4560
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   3720
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   3120
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00E9C5AD&
      Caption         =   "аккес кеитоуяциес"
      Height          =   855
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   9600
      Width           =   1575
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00E9C5AD&
      Caption         =   "упокоцислос тилым"
      Height          =   855
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   9600
      Width           =   1575
   End
   Begin VB.Timer Timer2 
      Interval        =   4000
      Left            =   480
      Top             =   2280
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00E9C5AD&
      Caption         =   "енодос"
      Height          =   855
      Left            =   11160
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   9600
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   500
      Left            =   13800
      MouseIcon       =   "Form1.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":030A
      ScaleHeight     =   465
      ScaleWidth      =   330
      TabIndex        =   7
      Top             =   360
      Width           =   360
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00E9C5AD&
      Caption         =   "йаятекес"
      Height          =   855
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   9600
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00E9C5AD&
      Caption         =   "глеяокоцио"
      Height          =   855
      Left            =   11160
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   8160
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00E9C5AD&
      Caption         =   "пяосжояес пяоиомтым"
      Height          =   855
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   8160
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E9C5AD&
      Caption         =   "тилокоциа"
      Height          =   855
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8160
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E9C5AD&
      Caption         =   "етаияиес"
      Height          =   855
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8160
      Width           =   1575
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1560
      Top             =   10320
   End
   Begin VB.Line Line5 
      BorderColor     =   &H000000FF&
      Visible         =   0   'False
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   600
   End
   Begin VB.Line Line8 
      BorderColor     =   &H000000FF&
      Visible         =   0   'False
      X1              =   15976.22
      X2              =   15976.22
      Y1              =   8760
      Y2              =   10560
   End
   Begin VB.Line Line7 
      BorderColor     =   &H000000FF&
      Visible         =   0   'False
      X1              =   3172.716
      X2              =   3172.716
      Y1              =   9000
      Y2              =   9600
   End
   Begin VB.Line Line4 
      Visible         =   0   'False
      X1              =   10513.14
      X2              =   10513.14
      Y1              =   8040
      Y2              =   11040
   End
   Begin VB.Line Line2 
      Visible         =   0   'False
      X1              =   16599.5
      X2              =   16599.5
      Y1              =   7680
      Y2              =   11040
   End
   Begin VB.Line Line1 
      Visible         =   0   'False
      X1              =   2553.191
      X2              =   2553.191
      Y1              =   7440
      Y2              =   11040
   End
   Begin VB.Line Line3 
      Visible         =   0   'False
      X1              =   9536.92
      X2              =   9536.92
      Y1              =   10
      Y2              =   8040
   End
   Begin VB.Image Image2 
      Height          =   2955
      Left            =   5040
      Picture         =   "Form1.frx":0659
      Top             =   2880
      Width           =   5475
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000013&
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000013&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   7680
      Left            =   2040
      Picture         =   "Form1.frx":382A
      Top             =   360
      Width           =   11280
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Load Form3
Form3.Show
End Sub



Private Sub Command2_Click()
Load Form4
Form4.Show

End Sub

Private Sub Command3_Click()
Load proionta1
proionta1.Show
End Sub

Private Sub Command4_Click()
Load HMEROLOGIO
HMEROLOGIO.Show
End Sub

Private Sub Command5_Click()
Load Form5
Form5.Show
EPIL_KARTELAS = 0

End Sub

Private Sub Command6_Click()
End
End Sub

Private Sub Command7_Click()
Load YPOL_TIMON
YPOL_TIMON.Show
End Sub

Private Sub Command8_Click()
Load Form1_б 'пяосовг б еккгмийо
Form1_б.Show

End Sub



Private Sub Form_Load()
Dim A As Integer
Dim F1, F2, F3, F4, F5, F6, F7, F8, F9, F10, F11, F12 As String
' ELEGXOS AN YPARXEI O FAKELOS ORGANIZER_BACKUP STON FAKELO WINDOWS.
' AN OXI DHMIOYRGIA AYTOY TOY FAKELOY
If Dir("C:\WINDOWS\ORGANIZER_BACKUP\", 55) = "" Then
    A = 0
Else
    A = 1
End If
If A = 1 Then

Else
    F1 = "C:\WINDOWS\ORGANIZER_BACKUP"
    F2 = F1 & "\BACKUP_THL"
    F3 = F1 & "\BACKUP_HMER"
    F4 = F1 & "\BACKUP_ETAIRIES"
    F5 = F3 & "\BACKUP_HMER"
    F6 = F3 & "\BACKUP_HMER_ETOS"
    F7 = F4 & "\BACKUP_ETAIRIES"
    F8 = F4 & "\BACKUP_ETAIRIES_ETOS"
    F9 = F7 & "\ETAIRIES"
    F10 = F7 & "\TXTS"
    F11 = F8 & "\ETAIRIES"
    F12 = F8 & "\TXTS"
    MkDir F1
    MkDir F2
    MkDir F3
    MkDir F4
    MkDir F5
    MkDir F6
    MkDir F7
    MkDir F8
    MkDir F9
    MkDir F10
    MkDir F11
    MkDir F12
End If
            
Label1.Caption = Date
Label2.Caption = Time
End Sub

Private Sub Picture1_Click()
Call Shell(App.Path & "\EFARMOGES\calc.exe", vbNormalFocus)
End Sub

Private Sub Timer1_Timer()
Label2.Caption = Time
End Sub

Private Sub Timer2_Timer()
Image2.Visible = False

End Sub
