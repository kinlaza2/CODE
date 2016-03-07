VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form12THL 
   BackColor       =   &H80000013&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "амтийатастасг амтицяажым тгкежымийоу йатакоцоу"
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
      Height          =   285
      Left            =   4080
      TabIndex        =   7
      Text            =   "Text2"
      Top             =   9960
      Visible         =   0   'False
      Width           =   1695
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   2040
      Top             =   8520
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
      Caption         =   "амтийатастасг"
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
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E9C5AD&
      Caption         =   "епистяожг сто пяогцоулемо лемоу"
      Height          =   735
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   9600
      Width           =   1575
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   8100
      Left            =   120
      TabIndex        =   2
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
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
      Caption         =   "амтицяажо пяос амтийатастасг"
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
      Top             =   4320
      Width           =   2655
   End
   Begin VB.Line Line2 
      Visible         =   0   'False
      X1              =   10
      X2              =   15300
      Y1              =   5505
      Y2              =   5505
   End
   Begin VB.Line Line1 
      Visible         =   0   'False
      X1              =   7650
      X2              =   7650
      Y1              =   12
      Y2              =   11080
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000013&
      Caption         =   "епикенте  ле поио апо та паяайаты амтицяажа тгкежымийоу йатакоцоу хекете ма амтийатастгсете том тяевом тгкежымийо йатакоцо"
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
      Height          =   495
      Left            =   600
      TabIndex        =   1
      Top             =   600
      Width           =   14175
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000013&
      Caption         =   "амтийатастасг ле амтицяажа тгкежымийоу йатакоцоу"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2860
      TabIndex        =   0
      Top             =   120
      Width           =   9615
   End
End
Attribute VB_Name = "Form12THL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error GoTo er:

Form12THL.Hide
Unload Form12THL
Form9.Enabled = True
BACKUP.Show
Form9.Show
GoTo TELOS:

er:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"


TELOS:
End Sub

Private Sub Command2_Click()
On Error GoTo er:
Dim DBH As New ADODB.Connection
Dim RSH As New ADODB.Recordset
Dim SOURCE As String
Dim DESTINATION As String
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
    If MsgBox("сас пяотеимете емтома то ма йяатгсете ема амтицяажо тгкежымийоу йатакоцоу пяим пяовыягсете стгм амтийатастасг.ам дем евете йяатгсг амтицяажо тоте патгсте CANCEL йаи пгцаимете сто пяогцоулемо лемоу пяойеилемоу ма дглиоуяцгсете ема амтицяажо. ", vbOKCancel, "") = vbOK Then
    If MsgBox("пяойеите ма амтийатастгсете тгм упаявоуса басг дедолемым глеяокоциоу, ле то амтицяажо поу дглиоуяцгхгйе стгс : " & temp & "  .еисте сицоуяои оти хекете ма пяовыягсете стгм емеяцеиа аутг", vbYesNo, "пяосовг") = vbYes Then
    If MsgBox("ам пяовыягсете стгм емеяцеиа аутг тоте та дедолема поу ха евете ыс емеяца сто пяоцяалла ха еимаи аута поу пеяиевомтаи сто амтицяажо. бебаиыхгте циа тгм ояхотгта тгс емеяцеиас.хекете ма пяовыягсете стгм амтийатастасг;", vbYesNo, "пяосовг") = vbYes Then
    If MsgBox("ма цимеи г амтийатастасг;", vbYesNo, "пяосовг") = vbYes Then
    Dim FSO As New FileSystemObject
    SOURCE = App.Path & "\databases\BACK_UPS\BACKUP_THL\" & _
             Text1.Text & ".MDB"
    
    DESTINATION = App.Path & "\databases\telephone.mdb"
    
   
    FSO.CopyFile SOURCE, DESTINATION
    MsgBox ("г амтийатастасг пяацлатопоигхгйE. се пеяиптысг поу хекете ма епамажеяете та дедолема стгм йатастасг поу гтам пяим, аяйеи ма айокоухгсете тгм идиа диадийасиа бафомтас то амтицяажо поу амтоистгвг стгм сглеяимг глеяолгмиа."), vbInformation, "ой"
    End If
    End If
    End If
    End If
    Text1.Text = ""
End If
GoTo TELOS:

ER1:
MsgBox ("дем дысате йамема аявеио ле то опоио ха амтийатастгсете тгм упаявоуса басг дедолемым циа етаияиес йаи тилокоциа"), vbCritical, "пяосовг!!"
GoTo TELOS:

er:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
If DBH.STATE = 1 Then DBH.Close
If RSH.STATE = 1 Then RSH.Close
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

Private Sub Form_Terminate()

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo er:

Form12THL.Hide
Unload Form12THL
Form9.Enabled = True
BACKUP.Show
Form9.Show
GoTo TELOS:

er:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"


TELOS:
End Sub
