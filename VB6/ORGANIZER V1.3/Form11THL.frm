VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form11THL 
   BackColor       =   &H80000013&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "диавеияисг амтицяажым тгкежымийоу йатакоцоу"
   ClientHeight    =   10500
   ClientLeft      =   90
   ClientTop       =   450
   ClientWidth     =   15180
   LinkTopic       =   "Form10"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   10500
   ScaleWidth      =   15180
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   2880
      TabIndex        =   8
      Text            =   "Text2"
      Top             =   9720
      Visible         =   0   'False
      Width           =   1455
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   4800
      Top             =   9840
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
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
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E9C5AD&
      Caption         =   "диавеияисг"
      Height          =   855
      Left            =   12840
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5400
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   12240
      TabIndex        =   4
      Top             =   4920
      Width           =   2655
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   8100
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   14288
      _Version        =   393216
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
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E9C5AD&
      Caption         =   "епистяожг сто пяогцоулемо лемоу"
      Height          =   735
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   9600
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000013&
      Caption         =   "епикенте апо том паяайаты пимайа то амтицяажо поу хекете ма деите"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   720
      Width           =   9375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
      Caption         =   "амтицяажо пяос диавеияисг"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12240
      TabIndex        =   6
      Top             =   4440
      Width           =   2655
   End
   Begin VB.Line Line2 
      Visible         =   0   'False
      X1              =   12
      X2              =   15000
      Y1              =   5505
      Y2              =   5505
   End
   Begin VB.Line Line1 
      Visible         =   0   'False
      X1              =   7650
      X2              =   7650
      Y1              =   0
      Y2              =   10320
   End
   Begin VB.Label Label2 
      Caption         =   "епикенте апо том паяайаты пимайа поио амтицяажо тгкежымийоу йатакоцоу хекете ма деите:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   4215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
      Caption         =   "диавеияисг амтицяажым тгкежымийоу йатакоцоу"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   2550
      TabIndex        =   0
      Top             =   120
      Width           =   10215
   End
End
Attribute VB_Name = "Form11THL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error GoTo er:

Form11THL.Hide
Unload Form11THL
Form9.Enabled = True
BACKUP.Show
Form9.Show
GoTo TELOS:

er:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
End Sub

Private Sub Command2_Click()
On Error GoTo NIK:
Dim D As New ADODB.Connection
Dim DS As New ADODB.Recordset
Dim C As Integer
C = 1
D.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & "\databases\BACK_UPS\HELP_BACKUP.mdb" & ";" & _
"Persist Security Info=False"
D.Open App.Path & "\databases\BACK_UPS\HELP_BACKUP.mdb"
DS.Open "[BACKUP_THL]", D, adOpenDynamic, adLockBatchOptimistic
DS.MoveFirst
Do While Not DS.EOF
    If DS![омола_аявеиоу] <> Text1.Text Then
        DS.MoveNext
    Else
        C = C + 1
        DS.MoveNext
    End If
Loop
If Text1.Text = "" Then
    MsgBox ("дем пкгйтяокоцгсате йамема омола аявеиоу"), vbCritical, "пяосовг"
    GoTo TELOS:
End If
If C = 1 Then
    MsgBox ("дем упаявеи амтицяажо ле то омола аявеиоу поу дысате"), vbCritical, "пяосовг!!"
Else
    Load ZTHL
    ZTHL.Show
    Form11THL.Enabled = False
End If
GoTo TELOS:

NIK:
MsgBox ("то омола аявеиоу поу дысате дем упаявеи 'г циа йапоио коцо дем еимаи думатом ма диабастеи"), vbCritical, "пяосовг!!"

TELOS:
End Sub

Private Sub DataGrid1_DblClick()
Text1.Text = DataGrid1.Columns(1).Text
End Sub

Private Sub Form_Load()
On Error GoTo er:

DataGrid1.HeadFont.Size = 10
DataGrid1.HeadFont.Bold = True
DataGrid1.Font.Size = 10
DataGrid1.DefColWidth = 5640

Dim STATEMENT As String
STATEMENT = "SELECT * FROM BACKUP_THL"
Adodc1.ConnectionString = "PROVIDER=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & App.Path & "\databases\BACK_UPS\HELP_BACKUP.mdb" & ";"
Adodc1.RecordSource = STATEMENT
Set DataGrid1.DataSource = Adodc1
Text2.Text = Adodc1.Recordset.RecordCount
If CInt(Text2.Text) < 26 Then
    DataGrid1.Height = (CInt(Text2.Text) + 1) * 300
Else
    DataGrid1.Height = 8100
End If
Adodc1.Refresh
GoTo TELOS:

er:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo er:

Form11THL.Hide
Unload Form11THL
Form9.Enabled = True
BACKUP.Show
Form9.Show
GoTo TELOS:

er:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
End Sub
