VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form10ETAIRIES 
   BackColor       =   &H80000013&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "диацяажг амтицяажым етаияиым"
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
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E9C5AD&
      Caption         =   "епистяожг сто пяогцоулемо лемоу"
      Height          =   735
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   9600
      Width           =   1815
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000013&
      Caption         =   "диацяажг етгсиым амтицяажым"
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
      TabIndex        =   1
      Top             =   600
      Width           =   7215
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   720
         TabIndex        =   13
         Text            =   "Text4"
         Top             =   8040
         Visible         =   0   'False
         Width           =   975
      End
      Begin MSAdodcLib.Adodc Adodc2 
         Height          =   330
         Left            =   5880
         Top             =   8400
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   582
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
      Begin VB.CommandButton Command3 
         BackColor       =   &H00E9C5AD&
         Caption         =   "диацяажг"
         Height          =   615
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   8160
         Width           =   1335
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   2340
         TabIndex        =   10
         Top             =   7680
         Width           =   2655
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
         X1              =   3667
         X2              =   3667
         Y1              =   120
         Y2              =   8880
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H80000013&
         Caption         =   "епикенте то етгсио амтицяажо етаияиым йаи тилокоциым поу хекете ма диацяажг апо том паяайаты пимайа"
         Height          =   495
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   6855
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000013&
      Caption         =   "диацяажг глеягсиым амтицяажым"
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
      TabIndex        =   0
      Top             =   600
      Width           =   7335
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   360
         TabIndex        =   12
         Text            =   "Text3"
         Top             =   7680
         Visible         =   0   'False
         Width           =   1095
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   330
         Left            =   6000
         Top             =   8400
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   582
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
         Caption         =   "диацяажг"
         Height          =   615
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   8160
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   2420
         TabIndex        =   8
         Top             =   7680
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
         Y1              =   240
         Y2              =   8820
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H80000013&
         Caption         =   "епикенте то амтцяажо етаияиым йаи тилокоциым поу хекете ма диацяажг апо том паяайаты пимайа"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   7095
      End
   End
   Begin VB.Line Line1 
      Visible         =   0   'False
      X1              =   7650
      X2              =   7650
      Y1              =   5
      Y2              =   11000
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
      Caption         =   "диацяажг амтицяажым етаияиым йаи тилокоциым"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3460
      TabIndex        =   3
      Top             =   120
      Width           =   8415
   End
End
Attribute VB_Name = "Form10ETAIRIES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error GoTo er:

Form10ETAIRIES.Hide
Unload Form10ETAIRIES
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
RSH.Open "[BACKUP_ETAIRIES]", DBH, adOpenDynamic, adLockBatchOptimistic
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
    If MsgBox("пяойеите ма диацяаьете то аявеио амтицяажоу етаияиым йаи тилокоциым поу дглиоуяцгхгйе стгс :" & temp, vbOKCancel, "пяосовг") = vbOK Then
    If MsgBox("хекете ма пяовыягсете стгм диацяажг", vbOKCancel, "пяосовг") = vbOK Then
    DESTINATION = App.Path & "\databases\BACK_UPS\BACKUP_ETAIRIES\BACKUP_ETAIRIES\" & Text1.Text & ".MDB"
    Kill DESTINATION
    STATEMENT = "DELETE FROM BACKUP_ETAIRIES WHERE омола_аявеиоу='" & Text1.Text & "'"
    DBH.Execute STATEMENT
    '*** DIAGRAFH TXTS KAI ETAIRIES *****************************
    Dim FSO As New FileSystemObject
    Dim PATH_TXT, PATH_ETAIR As String
    
    PATH_TXT = App.Path & "\databases\BACK_UPS\BACKUP_ETAIRIES\" & _
    "BACKUP_ETAIRIES\TXTS\" & Text1.Text & "_TXTS"
    
    PATH_ETAIR = App.Path & "\databases\BACK_UPS\BACKUP_ETAIRIES\" & _
    "BACKUP_ETAIRIES\ETAIRIES\" & Text1.Text
    
    FSO.DeleteFolder PATH_TXT
    FSO.DeleteFolder PATH_ETAIR
    MsgBox ("г диацяажг пяацлатопоигхгйе"), , "ой"
    End If
    End If
    Text1.Text = ""
End If
RSH.Fields.Refresh
RSH.Close
DBH.Close
Adodc1.Refresh
Text3.Text = Adodc1.Recordset.RecordCount
If CInt(Text3.Text) < 21 Then
    DataGrid1.Height = 250 + (300 * Text3.Text)
Else
    DataGrid1.Height = 6550
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

Private Sub Command3_Click()
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
RSH.Open "[BACKUP_YEAR_ETAIRIES]", DBH, adOpenDynamic, adLockBatchOptimistic
RSH.MoveFirst
Do While Not RSH.EOF
    If RSH![омола_аявеиоу] = Text2.Text Then
        C = C + 1
        RSH.MoveNext
    Else
        RSH.MoveNext
    End If
Loop
If C = 1 Then
    MsgBox ("то омола тоу аявеиоу поу дысате дем еимаи сысто.паяайакы намадысте то омола аявеиоу"), vbCritical, "пяосовг!!!"
Else
    RSH.MoveFirst
    Do While Not RSH.EOF
        If RSH![омола_аявеиоу] = Text2.Text Then
            If RSH![FLAG] = 0 Then
                MsgBox ("дем евете йяатгсг амтицяажо циа то суцйейяилемо етос"), vbCritical, "пяосовг!!!"
                RSH.MoveNext
            End If
            If RSH![FLAG] = 2 Then
                MsgBox ("евете гдг диацяаьеи то суцйейяилемо аявеио"), vbCritical, "пяосовг!!!"
                RSH.MoveNext
            End If
            If RSH![FLAG] = 1 Then
               If MsgBox("пяойеите ма диацяаьете то аявеио амтицяажоу циа то етос:" & RSH![ETOS] & ". пяотеимете ма лгм диацяаьете то аявеио циа опоиодгпоте етос, диоти ауто ха вахеи циа памта йаи етси дем ха евете дедолема циа то аявеио сас. пяотеимете ма патгсете то CANCEL пяойеилемоу ма амаияесете тгм кеитоуяциа диацяажгс.ам хекете ма пяовыягсете стгм диацяажг тоу аявеиоу патгсте ой", vbOKCancel, "пяосовг!!!") = vbOK Then
                  If MsgBox("еисте сицоуяои оти хекете ма пяовыягсете стгм диацяажг тоу аявеиоу", vbOKCancel, "пяосовг!!!") = vbOK Then
                    DESTINATION = App.Path & "\databases\BACK_UPS\BACKUP_ETAIRIES\BACKUP_ETAIRIES_ETOS\" & Text2.Text & ".MDB"
                    Kill DESTINATION
                    STATEMENT = "UPDATE BACKUP_YEAR_ETAIRIES SET FLAG=2 WHERE омола_аявеиоу='" & Text2.Text & "'"
                    DBH.Execute STATEMENT
                    '*** DIAGRAFH TXTS KAI ETAIRIES *****************************
    Dim FSO As New FileSystemObject
    Dim PATH_TXT, PATH_ETAIR As String
    
    PATH_TXT = App.Path & "\databases\BACK_UPS\BACKUP_ETAIRIES\" & _
    "BACKUP_ETAIRIES_ETOS\TXTS\" & "TXTS_" & Left(Text2.Text, 4)
    
    PATH_ETAIR = App.Path & "\databases\BACK_UPS\BACKUP_ETAIRIES\" & _
    "BACKUP_ETAIRIES_ETOS\ETAIRIES\" & Text2.Text
    
    FSO.DeleteFolder PATH_TXT
    FSO.DeleteFolder PATH_ETAIR
    
                    MsgBox ("г диацяажг пяацлатопоигхгйе"), , ""
                    End If
               End If
               
            RSH.MoveNext
            End If
        Else
            RSH.MoveNext
        End If
    Loop
End If
Text2.Text = ""
RSH.Fields.Refresh
RSH.Close
DBH.Close
Adodc2.Refresh
Text4.Text = Adodc2.Recordset.RecordCount
If CInt(Text3.Text) < 21 Then
    DataGrid2.Height = 250 + (300 * Text4.Text)
Else
    DataGrid2.Height = 6550
End If
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

Private Sub DataGrid2_DblClick()
Text2.Text = DataGrid2.Columns(1).Text
End Sub

Private Sub Form_Load()
On Error GoTo er:

DataGrid1.Font.Size = 10
DataGrid2.Font.Size = 10
DataGrid1.DefColWidth = 3235
DataGrid2.DefColWidth = 3200

Dim DATABASE_FILE As String
DATABASE_FILE = App.Path & "\databases\BACK_UPS\HELP_BACKUP.mdb"
Adodc1.ConnectionString = _
"PROVIDER=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & DATABASE_FILE & ";"
Adodc1.RecordSource = _
"SELECT глеяолгмиа_дглиоуяциас_амтицяажоу,омола_аявеиоу FROM BACKUP_ETAIRIES ORDER BY HMEROMHNIA"
' Bind the ADODC to the DataGrid.
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
"Data Source=" & DATABASE_FILE & ";"
Adodc2.RecordSource = _
"SELECT ETOS,омола_аявеиоу FROM BACKUP_YEAR_ETAIRIES WHERE FLAG=1 ORDER BY ETOS"
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

Form10ETAIRIES.Hide
Unload Form10ETAIRIES
Form9.Enabled = True
BACKUP.Show
Form9.Show
GoTo TELOS:

er:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
End Sub
