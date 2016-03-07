VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form11HMER 
   BackColor       =   &H80000013&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "диавеияисг амтицяажым глеяокоциоу"
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   495
      Left            =   10920
      Top             =   9960
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
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
      Caption         =   "Adodc2"
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   3960
      Top             =   9720
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
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
   Begin VB.Frame Frame2 
      BackColor       =   &H80000013&
      Caption         =   "етгсиа амтицяажа глеяокоциоу"
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
      Height          =   8895
      Left            =   7800
      TabIndex        =   2
      Top             =   600
      Width           =   7215
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   480
         TabIndex        =   13
         Text            =   "Text4"
         Top             =   7920
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00E9C5AD&
         Caption         =   "диавеияисг етгсиым амтицяажым"
         Height          =   735
         Left            =   2950
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   8040
         Width           =   1335
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   2360
         TabIndex        =   10
         Top             =   7560
         Width           =   2535
      End
      Begin MSDataGridLib.DataGrid DataGrid2 
         Height          =   6550
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   6990
         _ExtentX        =   12330
         _ExtentY        =   11562
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
      Begin VB.Line Line3 
         Visible         =   0   'False
         X1              =   3620
         X2              =   3620
         Y1              =   8880
         Y2              =   120
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H80000013&
         Caption         =   "епикенте апо том паяайаты пимайа то етгсио амтицяажо поу хекете ма деите"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   6495
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000013&
      Caption         =   "амтицяажа глеяокоциоу"
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
      Height          =   8895
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   7335
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   600
         TabIndex        =   12
         Text            =   "Text3"
         Top             =   8040
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00E9C5AD&
         Caption         =   "диавеияисг амтицяажоу"
         Height          =   735
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   8040
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   2400
         TabIndex        =   8
         Top             =   7560
         Width           =   2535
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   6550
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   11562
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
         X1              =   3677
         X2              =   3677
         Y1              =   8880
         Y2              =   120
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H80000013&
         Caption         =   "епикенте апо том паяайаты пимайа то амтицяажо глеяокоциоу поу хекете ма деите"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   6975
      End
   End
   Begin VB.Line Line1 
      X1              =   7650
      X2              =   7650
      Y1              =   15
      Y2              =   11920
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
      Caption         =   "диавеияисг амтицяажым глеяокоциоу"
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
      Left            =   4530
      TabIndex        =   0
      Top             =   120
      Width           =   6255
   End
End
Attribute VB_Name = "Form11HMER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error GoTo er:

Form11HMER.Hide
Unload Form11HMER
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
Dim DESTINATION, STATEMENT, temp As String
Dim C As Integer
C = 1
DBH.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & "\databases\BACK_UPS\HELP_BACKUP.mdb" & ";" & _
"Persist Security Info=False"
DBH.Open App.Path & "\databases\BACK_UPS\HELP_BACKUP.mdb"
RSH.Open "[BACKUP_HMER]", DBH, adOpenDynamic, adLockBatchOptimistic
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
    ZETAIRIES_BASHS_BACKUP_DIAX = Form11HMER.Text1.Text & ".MDB"
    ZETAIRIES_DIADROMHS_BACKUP_DIAX = App.Path & "\databases\BACK_UPS\BACKUP_HMER\BACKUP_HMER\" & Form11HMER.Text1.Text & ".MDB"
    ZETAIRIES_DIADROMHS_BACKUP_DIAX_1 = "\databases\BACK_UPS\BACKUP_HMER\BACKUP_HMER\" & Form11HMER.Text1.Text & ".MDB"
    Dim temp_name_all As String
    temp_name_all = Right(Text1.Text, 9)
    ETOSBACKUP = Left(temp_name_all, 4)
    Load ZHMEROLOGIO
    ZHMEROLOGIO.Show
    Form11HMER.Enabled = False
    
End If

GoTo TELOS:

ER1:
MsgBox ("дем дысате йамема аявеио пяос диацяажг"), vbCritical, "пяосовг!!"
GoTo TELOS:

er:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:

End Sub

Private Sub Command3_Click()
On Error GoTo er:

Dim DBH As New ADODB.Connection
Dim RSH As New ADODB.Recordset
Dim DESTINATION, STATEMENT, temp As String
Dim TEMP1 As Integer
Dim C As Integer
C = 1
DBH.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & "\databases\BACK_UPS\HELP_BACKUP.mdb" & ";" & _
"Persist Security Info=False"
DBH.Open App.Path & "\databases\BACK_UPS\HELP_BACKUP.mdb"
RSH.Open "[BACKUP_YEAR_HMER]", DBH, adOpenDynamic, adLockBatchOptimistic
If Text2.Text = "" Then GoTo ER1:
If RSH.BOF = RSH.EOF Then GoTo NIK:
RSH.MoveFirst
NIK:
Do While Not RSH.EOF
    If Text2.Text <> RSH![омола_аявеиоу] Then
        RSH.MoveNext
    Else
        C = C + 1
        temp = RSH![ETOS]
        TEMP1 = RSH![FLAG]
        RSH.MoveNext
    End If
Loop
If C = 1 Then
    MsgBox ("дем упаявеи аявеио ле то омола поу дысате"), vbCritical, "пяосовг!!"
    Text2.Text = ""
Else
    If TEMP1 = 0 Then
        MsgBox ("дем евете йяатгсг амтицяажы глеяокоциоу циа то етос:" & temp), , "пяосовг"
        GoTo TELOS:
    End If
    If TEMP1 = 2 Then
        MsgBox ("евете диацяаьг то амтицяажы глеяокоциоу циа то етос:" & temp), , "пяосовг"
        GoTo TELOS:
    End If
    If TEMP1 = 1 Then
    ZETAIRIES_BASHS_BACKUP_DIAX = Form11HMER.Text2.Text & ".MDB"
    ZETAIRIES_DIADROMHS_BACKUP_DIAX = App.Path & "\databases\BACK_UPS\BACKUP_HMER\BACKUP_HMER_ETOS\" & Form11HMER.Text2.Text & ".MDB"
    ZETAIRIES_DIADROMHS_BACKUP_DIAX_1 = "\databases\BACK_UPS\BACKUP_HMER\BACKUP_HMER_ETOS\" & Form11HMER.Text2.Text & ".MDB"
    ETOSBACKUP = CStr(temp)
    Load ZHMEROLOGIO
    ZHMEROLOGIO.Show
    Form11HMER.Enabled = False
    End If
End If

GoTo TELOS:

ER1:
MsgBox ("дем дысате йамема аявеио пяос диацяажг"), vbCritical, "пяосовг!!"
GoTo TELOS:

er:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:

End Sub

Private Sub DataGrid1_DblClick()
Text1.Text = DataGrid1.Columns(1).Text
End Sub

Private Sub DataGrid2_DblClick()
Text2.Text = DataGrid2.Columns(1).Text
End Sub

Private Sub Form_Load()
On Error GoTo er:

DataGrid1.Font.Size = 10
DataGrid2.Font.Size = 10
DataGrid1.DefColWidth = 3235
DataGrid2.DefColWidth = 3200

Dim STATEMENT As String
STATEMENT = "SELECT глеяолгмиа_дглиоуяциас_амтицяажоу,омола_аявеиоу FROM BACKUP_HMER"
Adodc1.ConnectionString = "PROVIDER=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & App.Path & "\databases\BACK_UPS\HELP_BACKUP.mdb" & ";"
Adodc1.RecordSource = STATEMENT
Set DataGrid1.DataSource = Adodc1
Text3.Text = Adodc1.Recordset.RecordCount
If CInt(Text3.Text) < 21 Then
    DataGrid1.Height = 250 + (300 * Text3.Text)
Else
    DataGrid1.Height = 6550
End If
Adodc1.Refresh


Adodc2.ConnectionString = _
"PROVIDER=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & App.Path & "\databases\BACK_UPS\HELP_BACKUP.mdb" & ";"
Adodc2.RecordSource = _
"SELECT ETOS,омола_аявеиоу FROM BACKUP_YEAR_HMER WHERE FLAG=1 ORDER BY INDEX"
' Bind the ADODC to the DataGrid.
Set DataGrid2.DataSource = Adodc2
Text4.Text = Adodc2.Recordset.RecordCount
If CInt(Text3.Text) < 21 Then
    DataGrid2.Height = 250 + (300 * Text4.Text)
Else
    DataGrid2.Height = 6550
End If
Adodc2.Refresh
GoTo TELOS:

er:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"


TELOS:
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo er:

Form11HMER.Hide
Unload Form11HMER
Form9.Enabled = True
BACKUP.Show
Form9.Show
GoTo TELOS:

er:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
End Sub
