VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form TZIROI 
   Caption         =   "тфияои"
   ClientHeight    =   10485
   ClientLeft      =   60
   ClientTop       =   465
   ClientWidth     =   15270
   LinkTopic       =   "Form10"
   ScaleHeight     =   10485
   ScaleWidth      =   15270
   Begin VB.Frame Frame3 
      Caption         =   "тфияои ама вяомийг пеяиодо"
      Height          =   9615
      Left            =   12600
      TabIndex        =   3
      Top             =   240
      Width           =   2535
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   360
         TabIndex        =   13
         Text            =   "Combo3"
         Top             =   3480
         Width           =   1935
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   360
         TabIndex        =   11
         Text            =   "Combo2"
         Top             =   2280
         Width           =   1935
      End
      Begin VB.CommandButton Command3 
         Caption         =   "тяевом етос"
         Height          =   375
         Left            =   360
         TabIndex        =   10
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label Label4 
         Caption         =   "лгмас"
         Height          =   375
         Left            =   360
         TabIndex        =   14
         Top             =   3000
         Width           =   1935
      End
      Begin VB.Label Label3 
         Caption         =   "етос"
         Height          =   375
         Left            =   360
         TabIndex        =   12
         Top             =   1800
         Width           =   1935
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "цягцояг амафгтгсг"
      Height          =   9375
      Left            =   6720
      TabIndex        =   2
      Top             =   240
      Width           =   5655
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000013&
      Caption         =   "тфияои ама етаияиа"
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
      Height          =   9495
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   6375
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   375
         Left            =   4080
         Top             =   8640
         Visible         =   0   'False
         Width           =   2055
         _ExtentX        =   3625
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
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   3960
         TabIndex        =   9
         Text            =   "епикоцг етоус"
         Top             =   2640
         Width           =   2295
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Height          =   855
         Left            =   4080
         TabIndex        =   6
         Top             =   4200
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   3960
         TabIndex        =   5
         Top             =   1200
         Width           =   2295
      End
      Begin VB.ListBox List1 
         Height          =   8835
         ItemData        =   "TZIROI.frx":0000
         Left            =   240
         List            =   "TZIROI.frx":0002
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   360
         Width           =   3495
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H80000013&
         Caption         =   "етос"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3960
         TabIndex        =   8
         Top             =   2400
         Width           =   2295
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H80000013&
         Caption         =   "омола етаияиас"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3960
         TabIndex        =   7
         Top             =   960
         Width           =   2295
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E9C5AD&
      Caption         =   "енодос"
      Height          =   735
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   9720
      Width           =   1575
   End
End
Attribute VB_Name = "TZIROI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Click()
'On Error GoTo ER:
List1.Clear
Dim DB1 As New ADODB.Connection
Dim DB2 As New ADODB.Connection
Dim RS1 As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset
Dim C As Integer
Dim DATABASE_F As String
C = 1
If IsNumeric(Combo1.Text) = False Then GoTo ER1:
'ELEGXOS AN TO YEAR POY YPARXEI STO COMBO BOX 1 EINAI TO TREXON H BACKUP
If Combo1.Text = Year(Date) Then
    DATABASE_F = "\databases\ETAIRIES.mdb"
Else
    ' SYNDESI ME BASH HELP_BACKUP.mdb
    DB2.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
    "Data Source=" & "\databases\BACK_UPS\HELP_BACKUP.mdb" & ";" & _
    "Persist Security Info=False"
    DB2.Open App.Path & "\databases\BACK_UPS\HELP_BACKUP.mdb"
    RS2.Open "[BACKUP_YEAR_ETAIRIES]", DB2, adOpenDynamic, adLockBatchOptimistic
    ' ELEGXOS AN YPARXEI BACK UP GIA TO SYGKEKRIMENO ETOS
    If RS2.BOF = RS2.EOF Then GoTo NN:
    RS2.MoveFirst
NN:
    Do While Not RS2.EOF
        If RS2![ETOS] = Combo1.Text Then
            If RS2![FLAG] = 1 Then C = 2
        End If
    RS2.MoveNext
    Loop
    If C = 1 Then
        MsgBox ("дем евете йяатгсеи амтицяажо циа то суцйейяилемо етос"), vbCritical, "пяосовг !!!"
        GoTo TELOS:
    Else
        DATABASE_F = "\databases\BACK_UPS\BACKUP_ETAIRIES\BACKUP_ETAIRIES_ETOS\" & Combo1.Text & "_ETAIRIES.mdb"
    End If
End If

DB1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & DATABASE_F & ";" & _
"Persist Security Info=False"
DB1.Open App.Path & DATABASE_F
RS1.Open "[ONOMATA_ETAIRION_ABCDEF]", DB1, adOpenDynamic, adLockBatchOptimistic
If RS1.BOF = RS1.EOF Then GoTo H:
RS1.MoveFirst
H:
Do While Not RS1.EOF
    List1.AddItem RS1![омолата_етаияиым]
    RS1.MoveNext
Loop

GoTo TELOS:

ER1:
MsgBox ("дем дысате сыста то етос"), vbCritical, "пяосовг !!!"
GoTo TELOS:

ER:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
End Sub


Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
'On Error GoTo ER:
List1.Clear
Dim DB1 As New ADODB.Connection
Dim DB2 As New ADODB.Connection
Dim RS1 As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset
Dim C As Integer
Dim DATABASE_F As String
C = 1
If IsNumeric(Combo1.Text) = False Then GoTo ER1:
'ELEGXOS AN TO YEAR POY YPARXEI STO COMBO BOX 1 EINAI TO TREXON H BACKUP
If Combo1.Text = Year(Date) Then
    DATABASE_F = "\databases\ETAIRIES.mdb"
Else
    ' SYNDESI ME BASH HELP_BACKUP.mdb
    DB2.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
    "Data Source=" & "\databases\BACK_UPS\HELP_BACKUP.mdb" & ";" & _
    "Persist Security Info=False"
    DB2.Open App.Path & "\databases\BACK_UPS\HELP_BACKUP.mdb"
    RS2.Open "[BACKUP_YEAR_ETAIRIES]", DB2, adOpenDynamic, adLockBatchOptimistic
    ' ELEGXOS AN YPARXEI BACK UP GIA TO SYGKEKRIMENO ETOS
    If RS2.BOF = RS2.EOF Then GoTo NN:
    RS2.MoveFirst
NN:
    Do While Not RS2.EOF
        If RS2![ETOS] = Combo1.Text Then
            If RS2![FLAG] = 1 Then C = 2
        End If
    RS2.MoveNext
    Loop
    If C = 1 Then
        MsgBox ("дем евете йяатгсеи амтицяажо циа то суцйейяилемо етос"), vbCritical, "пяосовг !!!"
        GoTo TELOS:
    Else
        DATABASE_F = "\databases\BACK_UPS\BACKUP_ETAIRIES\BACKUP_ETAIRIES_ETOS\" & Combo1.Text & "_ETAIRIES.mdb"
    End If
End If

DB1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & DATABASE_F & ";" & _
"Persist Security Info=False"
DB1.Open App.Path & DATABASE_F
RS1.Open "[ONOMATA_ETAIRION_ABCDEF]", DB1, adOpenDynamic, adLockBatchOptimistic
If RS1.BOF = RS1.EOF Then GoTo H:
RS1.MoveFirst
H:
Do While Not RS1.EOF
    List1.AddItem RS1![омолата_етаияиым]
    RS1.MoveNext
Loop
End If
GoTo TELOS:

ER1:
MsgBox ("дем дысате сыста то етос"), vbCritical, "пяосовг !!!"
GoTo TELOS:

ER:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
End Sub

Private Sub Command1_Click()
TZIROI.Hide
Unload TZIROI
End Sub


Private Sub Command2_Click()
Load TZIROI_1
TZIROI_1.Show
End Sub

Private Sub Form_Load()
'On Error GoTo ER:

'gemisma ton combo box me hmeromhnies
'COMBO1 KAI COMBO 2
Dim COMB12 As Integer
COMB12 = 2005
For I = 0 To 35
    Combo1.AddItem COMB12 + I
    Combo2.AddItem COMB12 + I
Next I
'Combo1.Text = Year(Date) - 1
'Combo2.Text = Year(Date)



GoTo TELOS:

ER:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
End Sub

Private Sub List1_DblClick()
Text1.Text = List1.Text
End Sub
