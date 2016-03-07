VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form2 
   BackColor       =   &H80000013&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "тгкежымийос йатакоцос"
   ClientHeight    =   10515
   ClientLeft      =   90
   ClientTop       =   450
   ClientWidth     =   15210
   LinkTopic       =   "Form2"
   ScaleHeight     =   11455.49
   ScaleMode       =   0  'User
   ScaleTop        =   1
   ScaleWidth      =   19800.23
   Begin VB.PictureBox Picture1 
      Height          =   418
      Left            =   4920
      Picture         =   "Form2.frx":0000
      ScaleHeight     =   360
      ScaleWidth      =   360
      TabIndex        =   45
      Top             =   120
      Width           =   422
   End
   Begin VB.TextBox Text8 
      Height          =   375
      Left            =   4680
      TabIndex        =   44
      Text            =   "Text8"
      Top             =   3360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   1560
      TabIndex        =   42
      Top             =   1560
      Width           =   2535
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   4800
      TabIndex        =   41
      Text            =   "HELP TXT"
      Top             =   4200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command29 
      BackColor       =   &H00E9C5AD&
      Caption         =   "еуяесг"
      Height          =   735
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   2280
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   615
      Left            =   4800
      Top             =   1080
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1085
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   10207
      Left            =   6120
      TabIndex        =   39
      Top             =   120
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   17992
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1032
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1032
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         ScrollBars      =   2
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command28 
      BackColor       =   &H00E9C5AD&
      Caption         =   "епистяожг сто аявийо MENU"
      Height          =   855
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   9480
      Width           =   1455
   End
   Begin VB.CommandButton Command27 
      BackColor       =   &H00E9C5AD&
      Caption         =   "ы"
      Height          =   375
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   8040
      Width           =   375
   End
   Begin VB.CommandButton Command26 
      BackColor       =   &H00E9C5AD&
      Caption         =   "ь"
      Height          =   375
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   8040
      Width           =   375
   End
   Begin VB.CommandButton Command25 
      BackColor       =   &H00E9C5AD&
      Caption         =   "в"
      Height          =   375
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   8040
      Width           =   375
   End
   Begin VB.CommandButton Command24 
      BackColor       =   &H00E9C5AD&
      Caption         =   "ж"
      Height          =   375
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   8040
      Width           =   375
   End
   Begin VB.CommandButton Command23 
      BackColor       =   &H00E9C5AD&
      Caption         =   "у"
      Height          =   375
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   8040
      Width           =   375
   End
   Begin VB.CommandButton Command22 
      BackColor       =   &H00E9C5AD&
      Caption         =   "т"
      Height          =   375
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   8040
      Width           =   375
   End
   Begin VB.CommandButton Command21 
      BackColor       =   &H00E9C5AD&
      Caption         =   "с"
      Height          =   375
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   8040
      Width           =   375
   End
   Begin VB.CommandButton Command20 
      BackColor       =   &H00E9C5AD&
      Caption         =   "я"
      Height          =   375
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   8040
      Width           =   375
   End
   Begin VB.CommandButton Command19 
      BackColor       =   &H00E9C5AD&
      Caption         =   "п"
      Height          =   375
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   7320
      Width           =   375
   End
   Begin VB.CommandButton Command18 
      BackColor       =   &H00E9C5AD&
      Caption         =   "о"
      Height          =   375
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   7320
      Width           =   375
   End
   Begin VB.CommandButton Command17 
      BackColor       =   &H00E9C5AD&
      Caption         =   "н"
      Height          =   375
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   7320
      Width           =   375
   End
   Begin VB.CommandButton Command16 
      BackColor       =   &H00E9C5AD&
      Caption         =   "м"
      Height          =   375
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   7320
      Width           =   375
   End
   Begin VB.CommandButton Command15 
      BackColor       =   &H00E9C5AD&
      Caption         =   "л"
      Height          =   375
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   7320
      Width           =   375
   End
   Begin VB.CommandButton Command14 
      BackColor       =   &H00E9C5AD&
      Caption         =   "к"
      Height          =   375
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   7320
      Width           =   375
   End
   Begin VB.CommandButton Command13 
      BackColor       =   &H00E9C5AD&
      Caption         =   "й"
      Height          =   375
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   7320
      Width           =   375
   End
   Begin VB.CommandButton Command12 
      BackColor       =   &H00E9C5AD&
      Caption         =   "и"
      Height          =   375
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   7320
      Width           =   375
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H00E9C5AD&
      Caption         =   "х"
      Height          =   375
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   6600
      Width           =   375
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H00E9C5AD&
      Caption         =   "г"
      Height          =   375
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   6600
      Width           =   375
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00E9C5AD&
      Caption         =   "ф"
      Height          =   375
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   6600
      Width           =   375
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00E9C5AD&
      Caption         =   "е"
      Height          =   375
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   6600
      Width           =   375
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00E9C5AD&
      Caption         =   "д"
      Height          =   375
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   6600
      Width           =   375
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00E9C5AD&
      Caption         =   "ц"
      Height          =   375
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6600
      Width           =   375
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00E9C5AD&
      Caption         =   "B"
      Height          =   375
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   6600
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00E9C5AD&
      Caption         =   "а"
      Height          =   375
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   6600
      Width           =   375
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000013&
      Height          =   2655
      Left            =   240
      TabIndex        =   10
      Top             =   5880
      Width           =   5655
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   "елжамисг йатакацоу ле басг то пяыто цяалла тоу епымулоу"
         Top             =   240
         Width           =   5415
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000013&
      Height          =   2415
      Left            =   240
      TabIndex        =   8
      Top             =   3360
      Width           =   3615
      Begin VB.TextBox Text4 
         Height          =   495
         Left            =   120
         TabIndex        =   11
         Top             =   1080
         Width           =   3135
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00E9C5AD&
         Caption         =   "амафгтгсг"
         Height          =   495
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1800
         Width           =   1935
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H80000013&
         Caption         =   "пкгйтяокоцгсте то епымуло тоу атолоу пяос амафгтгсг"
         Height          =   495
         Left            =   120
         TabIndex        =   38
         Top             =   360
         Width           =   3375
      End
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E9C5AD&
      Caption         =   "диацяажг"
      Height          =   735
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E9C5AD&
      Caption         =   "еццяажг "
      Height          =   735
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2280
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1560
      TabIndex        =   5
      Top             =   1080
      Width           =   2535
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1560
      TabIndex        =   4
      Top             =   600
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000013&
      Caption         =   "жан"
      Height          =   255
      Left            =   360
      TabIndex        =   43
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000013&
      Caption         =   "тгкежымо"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000013&
      Caption         =   "омола"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000013&
      Caption         =   "епихето"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
On Error GoTo er:

' ftiaksimo ton texts
Text1.Text = Trim(Text1.Text)
Text2.Text = Trim(Text2.Text)
Text3.Text = Trim(Text3.Text)
Text7.Text = Trim(Text7.Text)

Dim t1, t2, t3, t7 As String
Dim ddd1, ddd2, ddd3, ddd7 As Integer

ddd1 = Len(Text1.Text)
ddd2 = Len(Text2.Text)
ddd3 = Len(Text3.Text)
ddd7 = Len(Text4.Text)

If ddd1 > 50 Then
    t1 = Mid(Text1.Text, 1, 50)
Else
    t1 = Text1.Text
End If

If ddd2 > 50 Then
    t2 = Mid(Text2.Text, 1, 50)
Else
    t2 = Text2.Text
End If

If ddd3 > 50 Then
    t3 = Mid(Text3.Text, 1, 50)
Else
    t3 = Text3.Text
End If
    
If ddd7 > 50 Then
    t7 = Mid(Text7.Text, 1, 50)
Else
    t7 = Text7.Text
End If
Text1.Text = t1
Text2.Text = t2
Text3.Text = t3
Text7.Text = t7

If rs.STATE = 1 Then rs.Close
If db.STATE = 1 Then db.Close

If Text3.Text = "" Then GoTo KENO:

db.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & "\databases\telephone.mdb" & ";" & _
"Persist Security Info=False"
db.Open App.Path & "\databases\telephone.mdb"
rs.Open "[TEL_PIN]", db, adOpenDynamic, adLockBatchOptimistic

Dim STATEMENT As String
Dim C As Integer
Dim dd, dd1, count, m, i, n, count1 As Integer
count = 1
count1 = 1
m = 1
n = 1
C = 1
'********* PSAKSIMO AN YPARXEI HDH TO THLEFONO *********
Do While Not rs.EOF
    If Text3.Text <> rs![тгкежымо] Then
        rs.MoveNext
    Else
        C = C + 1
        rs.MoveNext
    End If
Loop

'*** AN DEN YPARXEI ********************
'ELEGXOS AN THLEFONO MONO APO ARITHMITIKA PSIFIA
If C = 1 Then

    dd = Len(Text3.Text)
    For i = 1 To dd ' * ELEGXOS AN YPARXEI MH ARITHMITIKO PSIFIO СТО ТГКЕЖЫМО*
        If Asc(Mid(Text3.Text, m, m)) < 48 Or Asc(Mid(Text3.Text, m, m)) > 57 Then
            count = count + 1
        End If
        m = m + 1
    Next i
    
'ELEGXOS AN FAX MONO APO ARITHMITIKA PSIFIA
    dd1 = Len(Text7.Text)
    For i = 1 To dd1 ' * ELEGXOS AN YPARXEI MH ARITHMITIKO PSIFIO СТО ТГКЕЖЫМО*
        If Asc(Mid(Text7.Text, n, n)) < 48 Or Asc(Mid(Text7.Text, n, n)) > 57 Then
            count1 = count1 + 1
        End If
        n = n + 1
    Next i
    
    If count1 = 1 Then

    Else
        GoTo LATHOS_FAX:
    End If
    
    If count = 1 Then  'AN TO THL EINAI OK
        STATEMENT = "INSERT INTO TEL_PIN (тгкежымо,епихето,омола,жан) VALUES (" & _
        "'" & Text3.Text & "'," & _
        "'" & UCase(Text1.Text) & "', " & _
        "'" & UCase(Text2.Text) & "'," & _
        "'" & Text7.Text & "'" & _
         ")"
        db.Execute STATEMENT
        rs.Fields.Refresh
        rs.Close
        db.Close
        Text1.Text = ""
        Text2.Text = ""
        Text3.Text = ""
        Text7.Text = ""
        Dim DATABASE_FILE As String
        DATABASE_FILE = App.Path & "/DATABASES/TELEPHONE.MDB"
        Adodc1.ConnectionString = _
        "PROVIDER=Microsoft.Jet.OLEDB.4.0;" & _
        "Data Source=" & DATABASE_FILE & ";"
        Adodc1.RecordSource = "SELECT * FROM TEL_PIN ORDER BY епихето"
        ' Bind the ADODC to the DataGrid.
        DataGrid1.Refresh
        Adodc1.Refresh
        Set DataGrid1.DataSource = Adodc1
        Text8.Text = Adodc1.Recordset.RecordCount
        If Text8.Text <= 33 Then
            DataGrid1.Height = 327.059 + (CInt(Text8.Text) * 327.059)
        Else
            DataGrid1.Height = 11120
        End If
    Else
        MsgBox ("пкгйтяокоцгсате кахос то тгкежымо"), vbCritical, "пяосовг !!!"
    End If
Else
    MsgBox ("о аяихлос тгкежымоу поу дысате упаявеи гдг"), vbCritical, "пяосовг !!!"
End If

'Text1.Text = ""
'Text2.Text = ""
'Text3.Text = ""
'Text7.Text = ""
GoTo TELOS:

KENO:
 MsgBox ("дем дысате йамема тгкежымо"), vbCritical, "пяосовг !!!"
 GoTo TELOS:
 
LATHOS_FAX:
MsgBox ("пкгйтяокоцгсате кахос то жан"), vbCritical, "пяосовг !!!"
GoTo TELOS:

er:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
If rs.STATE = 1 Then rs.Close
If db.STATE = 1 Then db.Close
End Sub

Private Sub Command10_Click()
On Error GoTo er:

ELEGXOSEYRESHSTHL "г"
GoTo TELOS:

er:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
End Sub

Private Sub Command11_Click()
On Error GoTo er:
ELEGXOSEYRESHSTHL "х"
GoTo TELOS:

er:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
End Sub

Private Sub Command12_Click()
On Error GoTo er:

ELEGXOSEYRESHSTHL "и"
GoTo TELOS:

er:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
End Sub

Private Sub Command13_Click()
On Error GoTo er:

ELEGXOSEYRESHSTHL "й"
GoTo TELOS:

er:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
End Sub

Private Sub Command14_Click()
On Error GoTo er:

ELEGXOSEYRESHSTHL "к"
GoTo TELOS:

er:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
End Sub

Private Sub Command15_Click()
On Error GoTo er:

ELEGXOSEYRESHSTHL "л"
GoTo TELOS:

er:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
End Sub

Private Sub Command16_Click()
On Error GoTo er:

ELEGXOSEYRESHSTHL "м"
GoTo TELOS:

er:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
End Sub

Private Sub Command17_Click()
On Error GoTo er:

ELEGXOSEYRESHSTHL "н"
GoTo TELOS:

er:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
End Sub

Private Sub Command18_Click()
On Error GoTo er:

ELEGXOSEYRESHSTHL "о"
GoTo TELOS:

er:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
End Sub

Private Sub Command19_Click()
On Error GoTo er:

ELEGXOSEYRESHSTHL "п"
GoTo TELOS:

er:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
End Sub

Private Sub Command2_Click()
On Error GoTo er:
Text3.Text = Trim(Text3.Text)

Dim C As Integer
C = 0

If rs.STATE = 1 Then rs.Close
If db.STATE = 1 Then db.Close

Dim STATEMENT As String
db.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & "\databases\telephone.mdb" & ";" & _
"Persist Security Info=False"
db.Open App.Path & "\databases\telephone.mdb"
rs.Open "[TEL_PIN]", db, adOpenDynamic, adLockBatchOptimistic
'****** ELEGXOI ************************************
If Text3.Text = "" Then GoTo KENO:
' ELEGXOS AN YPARXEI HDH TO THLEFONO *******
If rs.BOF = rs.EOF Then GoTo NIK:
rs.MoveFirst
NIK:
Do While Not rs.EOF
    If rs![тгкежымо] = Text3.Text Then
        C = C + 1
        rs.MoveNext
    Else
        rs.MoveNext
    End If
Loop
If C = 0 Then
    GoTo lathos:
Else
    
End If

' **** DIAGRAFH EGRAFHS APO BASH ************
If MsgBox("хекете ма пяовыягсете се диацяажг  тгс епикецлемгс еццяажгс", vbOKCancel, "") = vbOK Then
    STATEMENT = "DELETE from TEL_PIN " & _
    "WHERE тгкежымо=" & "'" & Text3.Text & "'"
    db.Execute STATEMENT
    rs.Fields.Refresh
    rs.Close
    db.Close

'************ EMFANISH META APO DIAGRAFH *********
    Dim DATABASE_FILE As String
    DATABASE_FILE = App.Path & "/DATABASES/TELEPHONE.MDB"
    Adodc1.ConnectionString = _
    "PROVIDER=Microsoft.Jet.OLEDB.4.0;" & _
    "Data Source=" & DATABASE_FILE & ";"
    Adodc1.RecordSource = "SELECT * FROM TEL_PIN ORDER BY епихето"
    ' Bind the ADODC to the DataGrid.
    DataGrid1.Refresh
    Adodc1.Refresh
    Set DataGrid1.DataSource = Adodc1
    Text8.Text = Adodc1.Recordset.RecordCount
    If Text8.Text <= 33 Then
        DataGrid1.Height = 327.059 + (CInt(Text8.Text) * 327.059)
    Else
        DataGrid1.Height = 11120
    End If
    Dim ASD As String
    ASD = Text4.Text
    ELEGXOSEYRESHSTHL (ASD)
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text7.Text = ""
Else
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text7.Text = ""
End If
GoTo TELOS:

KENO:
MsgBox (" дем дысате йамема аяихло тгкежымоу пяос диацяажг"), vbCritical, "пяосовг !!!"
GoTo TELOS:


lathos:
MsgBox ("аяихлос тгкежымоу поу дысате дем упаявеи"), vbCritical, "пяосовг !!!"
GoTo TELOS:

er:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
If rs.STATE = 1 Then rs.Close
If db.STATE = 1 Then db.Close
End Sub

Private Sub Command20_Click()
On Error GoTo er:

ELEGXOSEYRESHSTHL "я"
GoTo TELOS:

er:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
End Sub

Private Sub Command21_Click()
On Error GoTo er:

ELEGXOSEYRESHSTHL "с"
GoTo TELOS:

er:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
End Sub

Private Sub Command22_Click()
On Error GoTo er:

ELEGXOSEYRESHSTHL "т"
GoTo TELOS:

er:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
End Sub

Private Sub Command23_Click()
On Error GoTo er:

ELEGXOSEYRESHSTHL "у"
GoTo TELOS:

er:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
End Sub

Private Sub Command24_Click()
On Error GoTo er:

ELEGXOSEYRESHSTHL "ж"
GoTo TELOS:

er:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
End Sub

Private Sub Command25_Click()
On Error GoTo er:

ELEGXOSEYRESHSTHL "в"
GoTo TELOS:

er:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
End Sub

Private Sub Command26_Click()
On Error GoTo er:

ELEGXOSEYRESHSTHL "ь"
GoTo TELOS:

er:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
End Sub

Private Sub Command27_Click()
On Error GoTo er:

ELEGXOSEYRESHSTHL "ы"
GoTo TELOS:

er:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
End Sub

Private Sub Command28_Click()
On Error GoTo er:

Form2.Hide
Unload Form2
GoTo TELOS:

er:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
End Sub

Private Sub Command29_Click()
On Error GoTo er:

Text1.Text = Trim(Text1.Text)
Text2.Text = Trim(Text2.Text)
Text3.Text = Trim(Text3.Text)
Text7.Text = Trim(Text7.Text)

Dim STATE1, STATE2, STATE3, STATE4 As String
Dim C, index, ARE As Integer
C = 1
index = 2
ARE = 0

If db.STATE = 1 Then db.Close
If rs.STATE = 1 Then rs.Close
db.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & "\databases\telephone.mdb" & ";" & _
"Persist Security Info=False"
db.Open App.Path & "\databases\telephone.mdb"
rs.Open "[TEL_PIN]", db, adOpenDynamic, adLockBatchOptimistic
'********** AN TO KOYMPI EINAI EYRESH ****************
If Command29.Caption = "еуяесг" Then
    If Text3.Text = "" Then GoTo NOEYRESH:
    If IsNumeric(Text3.Text) = False Then GoTo nonum:
    If rs.BOF = rs.EOF Then GoTo NNIK:
        rs.MoveFirst
NNIK:
    Do While Not rs.EOF
        If rs![тгкежымо] = Text3.Text Then
            Text1.Text = rs![епихето]
            Text2.Text = rs![омола]
            Text7.Text = rs![жан]
            Text3.Text = rs![тгкежымо]
            Text6.Text = rs![тгкежымо]
            C = C + 1
            Command29.Caption = "диояхысг"
            Command1.Enabled = False
            Command2.Enabled = False
            rs.MoveNext
        Else
            rs.MoveNext
        End If
    Loop
        
    If C = 1 Then
        MsgBox ("дем упаявеи еццяажг ле том аяихло тгкежымоу поу дысате"), vbCritical, "пяосовг!!!"
    End If
    
Else ' ******* AN KOYMPI EINAI DIORTHOSI *************
        ' ftiaksimo ton texts
        Text1.Text = Trim(Text1.Text)
        Text2.Text = Trim(Text2.Text)
        Text3.Text = Trim(Text3.Text)
        Text7.Text = Trim(Text7.Text)

        Dim t1, t2, t3, t7 As String
        Dim ddd1, ddd2, ddd3, ddd7 As Integer

        ddd1 = Len(Text1.Text)
        ddd2 = Len(Text2.Text)
        ddd3 = Len(Text3.Text)
        ddd7 = Len(Text4.Text)

        If ddd1 > 50 Then
            t1 = Mid(Text1.Text, 1, 50)
        Else
            t1 = Text1.Text
        End If

        If ddd2 > 50 Then
            t2 = Mid(Text2.Text, 1, 50)
        Else
            t2 = Text2.Text
        End If

        If ddd3 > 50 Then
            t3 = Mid(Text3.Text, 1, 50)
        Else
            t3 = Text3.Text
        End If
    
        If ddd7 > 50 Then
            t7 = Mid(Text7.Text, 1, 50)
        Else
            t7 = Text7.Text
        End If
        Text1.Text = t1
        Text2.Text = t2
        Text3.Text = t3
        Text7.Text = t7

        If Text1.Text = "" Then GoTo EPITHETO1:
        If Text3.Text = "" Then GoTo THL1:
        If IsNumeric(Text3.Text) = False Then GoTo THL2:
        If Text7.Text = "" Then
        
        Else
            If IsNumeric(Text7.Text) = False Then GoTo fax:
        End If
'********* ELEGXOS AN TO NOYMERO POY THELEI NA TOY DOSEI MESO UPDATE EPITREPETAI ****
        If rs.BOF = rs.EOF Then GoTo NNN:
        rs.MoveFirst
NNN:
        Do While Not rs.EOF
            If rs![тгкежымо] = Text3.Text Then
                ARE = ARE + 1
                rs.MoveNext
            Else
                rs.MoveNext
            End If
        Loop
        
        If ARE <> 0 Then
            If Text3.Text = Text6.Text Then
            
            Else
                MsgBox ("о аяихлос тгкежымоу поу дысате упаявеи гдг"), vbCritical, "пяосовг !!!"
                GoTo TELOS:
            End If
         End If
            
        If MsgBox("хекете ма пяовыягсете се аккацг тгс еццяажгс", vbOKCancel, "") = vbOK Then
            STATE1 = "UPDATE TEL_PIN SET епихето='" & UCase(Text1.Text) & "' WHERE тгкежымо='" & Text6.Text & "'"
            STATE2 = "UPDATE TEL_PIN SET омола='" & UCase(Text2.Text) & "' WHERE тгкежымо='" & Text6.Text & "'"
            STATE3 = "UPDATE TEL_PIN SET тгкежымо=" & Text3.Text & " WHERE тгкежымо='" & Text6.Text & "'"
            STATE4 = "UPDATE TEL_PIN SET жан='" & Text7.Text & "' WHERE тгкежымо='" & Text6.Text & "'"
            db.Execute STATE4
            db.Execute STATE1
            db.Execute STATE2
            db.Execute STATE3
            
                rs.Fields.Refresh
                rs.Close
                db.Close
            MsgBox ("г диояхысг окойкгяыхгйе"), , ""
            Command29.Caption = "еуяесг"
            Text1.Text = ""
            Text2.Text = ""
            Text3.Text = ""
            Text7.Text = ""
            Command1.Enabled = True
            Command2.Enabled = True
            
            Adodc1.Refresh
            Dim DATABASE_FILE As String
            DATABASE_FILE = App.Path & "/DATABASES/TELEPHONE.MDB"
            Adodc1.ConnectionString = _
            "PROVIDER=Microsoft.Jet.OLEDB.4.0;" & _
            "Data Source=" & DATABASE_FILE & ";"
            Adodc1.RecordSource = "SELECT * FROM TEL_PIN ORDER BY епихето"
            ' Bind the ADODC to the DataGrid.
            Set DataGrid1.DataSource = Adodc1
            Text8.Text = Adodc1.Recordset.RecordCount
            If Text8.Text <= 33 Then
                DataGrid1.Height = 327.059 + (CInt(Text8.Text) * 327.059)
            Else
                DataGrid1.Height = 11120
            End If
        Else
            Command29.Caption = "еуяесг"
            Text1.Text = ""
            Text2.Text = ""
            Text3.Text = ""
            Text7.Text = ""
            Command1.Enabled = True
            Command2.Enabled = True
        End If
    End If

GoTo TELOS:

NOEYRESH:
MsgBox ("дем дысате аяихло тгкежымоу пяойеилемоу ма цимеи г еуяесг"), vbCritical, "пяосовг!!!"
index = 32
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text7.Text = ""

nonum:
If index = 32 Then
    GoTo TELOS:
Else
    MsgBox ("дем дысате сыста том аяихло тгкежымоу пяойеилемоу ма цимеи г еуяесг"), vbCritical, "пяосовг!!!"
    index = 32
End If

EPITHETO1:
If index = 32 Then
    GoTo TELOS:
Else
    MsgBox ("дем дысате епихето"), vbCritical, "пяосовг!!!"
    index = 32
End If


THL1:
If index = 32 Then
    GoTo TELOS:
Else
    MsgBox ("дем дысате аяихло тгкежымоу"), vbCritical, "пяосовг!!!"
    index = 32
End If

THL2:
If index = 32 Then
    GoTo TELOS:
Else
    MsgBox ("дем дысате сыста том аяихло тгкежымоу"), vbCritical, "пяосовг!!!"
    index = 32
End If

fax:
If index = 32 Then
    GoTo TELOS:
Else
    MsgBox ("дем дысате сыста том аяихло тоу жан"), vbCritical, "пяосовг!!!"
    index = 32
End If


er:
If index = 32 Then
    GoTo TELOS:
Else
    MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"
    index = 32
End If

TELOS:
If db.STATE = 1 Then db.Close
If rs.STATE = 1 Then rs.Close
End Sub

Private Sub Command3_Click()
On Error GoTo er:

Dim ASD As String
ASD = Text4.Text
ELEGXOSEYRESHSTHL (ASD)
GoTo TELOS:

er:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
End Sub

Private Sub Command4_Click()
On Error GoTo er:

ELEGXOSEYRESHSTHL "а"
GoTo TELOS:

er:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
End Sub

Private Sub Command5_Click()
On Error GoTo er:

ELEGXOSEYRESHSTHL "б"
GoTo TELOS:

er:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
End Sub

Private Sub Command6_Click()
On Error GoTo er:

ELEGXOSEYRESHSTHL "ц"
GoTo TELOS:

er:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
End Sub

Private Sub Command7_Click()
On Error GoTo er:

ELEGXOSEYRESHSTHL "д"
GoTo TELOS:

er:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
End Sub

Private Sub Command8_Click()
On Error GoTo er:

ELEGXOSEYRESHSTHL "е"
GoTo TELOS:

er:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
End Sub

Private Sub Command9_Click()
On Error GoTo er:

ELEGXOSEYRESHSTHL "ф"
GoTo TELOS:

er:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
End Sub

Private Sub DataGrid1_Click()
Text3.Text = DataGrid1.Columns(2).Text

End Sub

Private Sub Form_Load()
On Error GoTo er:


DataGrid1.Font.Size = 10
DataGrid1.DefColWidth = 2700
DataGrid1.HeadFont.Bold = True
DataGrid1.HeadFont.Size = 10

Dim A As String
Dim DATABASE_FILE As String
DATABASE_FILE = App.Path & "/DATABASES/TELEPHONE.MDB"
Adodc1.ConnectionString = _
"PROVIDER=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & DATABASE_FILE & ";"
Adodc1.RecordSource = "SELECT * FROM TEL_PIN ORDER BY епихето"
' Bind the ADODC to the DataGrid.
Set DataGrid1.DataSource = Adodc1
Text8.Text = Adodc1.Recordset.RecordCount
If Text8.Text <= 33 Then
    DataGrid1.Height = 327.059 + (CInt(Text8.Text) * 327.059)
Else
    DataGrid1.Height = 11120
End If
GoTo TELOS:
    
er:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"
    
TELOS:
End Sub

Private Sub Label6_Click()

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo er:

Form2.Hide
Unload Form2
GoTo TELOS:

er:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
End Sub

Private Sub Picture1_Click()
        Text1.Text = ""
        Text2.Text = ""
        Text3.Text = ""
        Text7.Text = ""
        Command29.Caption = "еуяесг"
        Command1.Enabled = True
        Command2.Enabled = True
End Sub
