VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form5 
   Caption         =   "FM-NOTES"
   ClientHeight    =   7755
   ClientLeft      =   6060
   ClientTop       =   450
   ClientWidth     =   12075
   LinkTopic       =   "Form5"
   ScaleHeight     =   7755
   ScaleWidth      =   12075
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2160
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6735
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "h.frx":0000
      Top             =   600
      Width           =   11775
   End
   Begin VB.Image Image3 
      Height          =   360
      Left            =   7080
      Picture         =   "h.frx":0006
      Stretch         =   -1  'True
      Top             =   120
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   345
      Left            =   6240
      Picture         =   "h.frx":08D0
      Top             =   120
      Width           =   360
   End
   Begin VB.Image Image1 
      Height          =   345
      Left            =   5400
      Picture         =   "h.frx":0F3A
      Top             =   120
      Width           =   360
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub OLE1_Updated(Code As Integer)

End Sub

Private Sub CommonDialog1_Click()

End Sub

Private Sub Form_Load()
On Error GoTo ERROR:
'Text1.Height = 5790
'Text1.Top = 360
'Text1.Left = 165
'Text1.Width = 7160
'Text1.Visible = True

Wrap$ = Chr$(13) + Chr$(10)
CommonDialog1.FileName = App.Path & "\TXTS\" & Form1.Text1.Text & ".TXT"
Open CommonDialog1.FileName For Input As #1
Do Until EOF(1)
Line Input #1, LINEOFTEXT$
ALLTEXT$ = ALLTEXT$ & LINEOFTEXT$ & Wrap$
Loop
Text1.Text = ALLTEXT$
Close #1
GoTo TELOS:

ERROR:
MsgBox ("поку лецакг сглеиысг"), vbCritical, "пяосовг!!"

TELOS:
End Sub

Private Sub Form_Unload(Cancel As Integer)
Form5.Hide
Unload Form5
End Sub

Private Sub Image1_Click()
On Error GoTo ERROR:
Text1.Text = Trim(Text1.Text)
CommonDialog1.FileName = App.Path & "\TXTS\" & Form1.Text1.Text & ".TXT"
Open CommonDialog1.FileName For Output As #1
Print #1, Text1.Text
Close #1
Form5.Hide
Unload Form5
GoTo TELOS:

ERROR:
MsgBox ("поку лецакг сглеиысH. пкгйтяокоцисте кицотеяа"), vbCritical, "пяосовг!!"

TELOS:
End Sub

Private Sub Image2_Click()
Close #1
Form5.Hide
Unload Form5
End Sub

Private Sub Image3_Click()
On Error GoTo ERROR:
Text1.Text = Text2.Text
CommonDialog1.FileName = App.Path & "\TXTS\" & Form1.Text1.Text & ".TXT"
Open CommonDialog1.FileName For Output As #1
Print #1, Text1.Text
Close #1
Form5.Hide
Unload Form5
GoTo TELOS:

ERROR:
MsgBox ("поку лецакг сглеиысH. пкгйтяокоцисте кицотеяа"), vbCritical, "пяосовг!!"

TELOS:
End Sub

