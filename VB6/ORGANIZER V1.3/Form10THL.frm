VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form10THL 
   BackColor       =   &H80000013&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "диацяажг амтицяажым тгкежымийоу йатакоцоу"
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
      Left            =   2160
      TabIndex        =   7
      Text            =   "Text2"
      Top             =   9480
      Visible         =   0   'False
      Width           =   2055
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   0
      Top             =   8880
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
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E9C5AD&
      Caption         =   "епистяожг сто пяогцоулемо лемоу"
      Height          =   975
      Left            =   120
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   9360
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E9C5AD&
      Caption         =   "диацяажг"
      Height          =   855
      Left            =   12840
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5400
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   12240
      TabIndex        =   2
      Top             =   4920
      Width           =   2655
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   8100
      Left            =   120
      TabIndex        =   0
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
   Begin VB.Line Line2 
      Visible         =   0   'False
      X1              =   10
      X2              =   15200
      Y1              =   5505
      Y2              =   5505
   End
   Begin VB.Line Line1 
      Visible         =   0   'False
      X1              =   7650
      X2              =   7650
      Y1              =   0
      Y2              =   10440
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000013&
      Caption         =   "диацяажг амтицяажым тгкежымийоу йатакоцоу"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3570
      TabIndex        =   6
      Top             =   120
      Width           =   8175
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
      Caption         =   "амтицяажо пяос диацяажг"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12240
      TabIndex        =   5
      Top             =   4560
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000013&
      Caption         =   "епикенте апо том паяайаты пимайа, поио апо та амтицяажа хекете ма диацяажеи"
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
      TabIndex        =   1
      Top             =   720
      Width           =   10335
   End
End
Attribute VB_Name = "Form10THL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error GoTo er:

Dim DBH As New ADODB.Connection
Dim RSH As New ADODB.Recordset
Dim DESTINATION, STATEMENT, temp As String
Dim C As Integer
C = 1

DBH.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & "\databases\BACK_UPS\HELP_BACKUP.mdb" & ";" & _
"Persist Security Info=False"
DBH.Open App.Path & "\databases\BACK_UPS\HELP_BACKUP.mdb"
RSH.Open "[BACKUP_THL]", DBH, adOpenDynamic, adLockBatchOptimistic

If Text1.Text = "" Then GoTo ER1:
If RSH.BOF = RSH.EOF Then GoTo NIK:
RSH.MoveFirst
NIK:
Do While Not RSH.EOF
    If Text1.Text <> RSH![омола_аявеиоу] Then
        RSH.MoveNext
    Else
        C = C + 1
        temp = RSH![глеяолгмиа_дглиоуяциас_амтицяажоу]
        RSH.MoveNext
    End If
Loop

If C = 1 Then
    MsgBox ("дем упаявеи аявеио ле то омола поу дысате"), vbCritical, "пяосовг!!"
    Text1.Text = ""
Else
    If MsgBox("пяойеите ма диацяаьете то аявеио амтицяажоу тгкежымийоу йатакоцоу поу дглиоуяцгхгйе стгс :" & temp, vbOKCancel, "пяосовг") = vbOK Then
    If MsgBox("хекете ма пяовыягсете стгм диацяажг", vbOKCancel, "пяосовг") = vbOK Then
    DESTINATION = App.Path & "\databases\BACK_UPS\BACKUP_THL\" & Text1.Text & ".MDB"
    Kill DESTINATION
    STATEMENT = "DELETE FROM BACKUP_THL WHERE омола_аявеиоу='" & Text1.Text & "'"
    DBH.Execute STATEMENT
    MsgBox ("г диацяажг пяацлатопоигхгйе"), , "ой"
    End If
    End If
    Text1.Text = ""
End If
RSH.Fields.Refresh
RSH.Close
DBH.Close
Adodc1.Refresh
Text2.Text = Adodc1.Recordset.RecordCount
If CInt(Text2.Text) < 26 Then
    DataGrid1.Height = (CInt(Text2.Text) + 1) * 300
Else
    DataGrid1.Height = 8100
End If
GoTo TELOS:


ER1:
MsgBox ("дем дысате йамема аявеио пяос диацяажг"), vbCritical, "пяосовг!!"
GoTo TELOS:

er:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
If DBH.STATE = 1 Then DBH.Close
If RSH.STATE = 1 Then RSH.Close
End Sub

Private Sub Command2_Click()
On Error GoTo er:

Form10THL.Hide
Unload Form10THL
Form9.Enabled = True
BACKUP.Show
Form9.Show
GoTo TELOS:

er:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

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

Dim DATABASE_FILE As String
DATABASE_FILE = App.Path & "\databases\BACK_UPS\HELP_BACKUP.mdb"
Adodc1.ConnectionString = _
"PROVIDER=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & DATABASE_FILE & ";"
Adodc1.RecordSource = _
"SELECT глеяолгмиа_дглиоуяциас_амтицяажоу,омола_аявеиоу FROM BACKUP_THL ORDER BY HMEROMHNIA"
' Bind the ADODC to the DataGrid.
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

Form10THL.Hide
Unload Form10THL
Form9.Enabled = True
BACKUP.Show
Form9.Show
GoTo TELOS:

er:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
End Sub
