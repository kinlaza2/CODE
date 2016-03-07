VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form3 
   BackColor       =   &H80000013&
   Caption         =   "етаияиес"
   ClientHeight    =   10485
   ClientLeft      =   2130
   ClientTop       =   465
   ClientWidth     =   11175
   LinkTopic       =   "Form3"
   ScaleHeight     =   10485
   ScaleWidth      =   11175
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   4005
      Left            =   5280
      TabIndex        =   7
      Top             =   120
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   7064
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   -2147483634
      DefColWidth     =   400
      HeadLines       =   1
      RowHeight       =   15
      RowDividerStyle =   4
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
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   2640
      TabIndex        =   12
      Text            =   "Text4"
      Top             =   7560
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   2160
      Top             =   7200
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
   Begin MSDataGridLib.DataGrid DataGrid2 
      Height          =   735
      Left            =   2640
      TabIndex        =   11
      Top             =   6480
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1296
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
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   0
      TabIndex        =   10
      Text            =   "HELP_TXT"
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   4400
      Width           =   3015
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00E9C5AD&
      Caption         =   "епистяожг сто аявийо лемоу"
      Height          =   855
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   9360
      Width           =   1815
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   4080
      Top             =   120
      Width           =   1200
      _ExtentX        =   2117
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
      BackColor       =   -2147483629
      ForeColor       =   -2147483629
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
   Begin VB.CommandButton Command4 
      BackColor       =   &H00E9C5AD&
      Caption         =   "амафгтгсг"
      Height          =   495
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5000
      Width           =   2655
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00E9C5AD&
      Caption         =   "диацяажг етаияиас"
      Height          =   615
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2000
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E9C5AD&
      Caption         =   "еуяесг"
      Height          =   615
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2000
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E9C5AD&
      Caption         =   "еццяажг етаияиас"
      Height          =   615
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2000
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1160
      Width           =   3975
   End
   Begin VB.Line Line10 
      Visible         =   0   'False
      X1              =   3720
      X2              =   5280
      Y1              =   3155
      Y2              =   3155
   End
   Begin VB.Line Line9 
      Visible         =   0   'False
      X1              =   3720
      X2              =   5280
      Y1              =   2840
      Y2              =   2840
   End
   Begin VB.Line Line8 
      Visible         =   0   'False
      X1              =   3720
      X2              =   5280
      Y1              =   2525
      Y2              =   2525
   End
   Begin VB.Line Line7 
      Visible         =   0   'False
      X1              =   3600
      X2              =   5280
      Y1              =   2210
      Y2              =   2210
   End
   Begin VB.Line Line6 
      Visible         =   0   'False
      X1              =   3840
      X2              =   5280
      Y1              =   1890
      Y2              =   1890
   End
   Begin VB.Line Line5 
      Visible         =   0   'False
      X1              =   3960
      X2              =   5280
      Y1              =   1580
      Y2              =   1580
   End
   Begin VB.Line Line4 
      Visible         =   0   'False
      X1              =   4320
      X2              =   5280
      Y1              =   1260
      Y2              =   1260
   End
   Begin VB.Line Line3 
      Visible         =   0   'False
      X1              =   4080
      X2              =   5280
      Y1              =   950
      Y2              =   950
   End
   Begin VB.Line Line2 
      Visible         =   0   'False
      X1              =   4200
      X2              =   5280
      Y1              =   625
      Y2              =   625
   End
   Begin VB.Line Line1 
      Visible         =   0   'False
      X1              =   4200
      X2              =   5280
      Y1              =   330
      Y2              =   330
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000013&
      Caption         =   "аяихлос сумеяцафолемым етаияиым :"
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
      Left            =   120
      TabIndex        =   13
      Top             =   7200
      Width           =   4575
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000013&
      Caption         =   "амафгтгсг етаияиас:"
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
      Left            =   120
      TabIndex        =   4
      Top             =   3800
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000013&
      Caption         =   "омола етаияиас:"
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
      Left            =   120
      TabIndex        =   0
      Top             =   675
      Width           =   2055
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'On Error GoTo ER:

ELX (Text1.Text)
Form3.Text1.Text = Form1.Text1.Text
If CInt(Form1.Text2.Text) > 50 Then GoTo ER_MHKOS:

If DB1.STATE = 1 Then DB1.Close
If RS1.STATE = 1 Then RS1.Close


Dim STATEMENT, statement_B As String
Dim SOURCE, DESTINATION, DESTINATION_B, S2, D2 As String
Dim IND, C As Integer
IND = 1
C = 1

DB1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
       "Data Source=" & "\databases\ETAIRIES.mdb" & ";" & _
      "Persist Security Info=False"
DB1.Open App.Path & "\databases\ETAIRIES.mdb"
RS1.Open "[ONOMATA_ETAIRION_ABCDEF]", DB1, adOpenDynamic, adLockBatchOptimistic

If Text1.Text = "" Then
    IND = 2
    GoTo KENO:
End If

If RS1.BOF = RS1.EOF Then GoTo NIK:
RS1.MoveFirst
NIK:
Do While Not RS1.EOF
    If RS1![омолата_етаияиым] = UCase(Text1.Text) Then
        C = C + 1
        RS1.MoveNext
    Else
        RS1.MoveNext
    End If
Loop

If C <> 1 Then
    MsgBox ("то омола етаияиас поу дысате упаявеи гдг."), vbCritical, "пяосовг !!!"
Else
    If MsgBox("хекете ма пяовыягсете се еццяажг тгс етаияиас:  " & Text1.Text, vbOKCancel, "") = vbOK Then
        STATEMENT = " create table " & UCase(Text1.Text) & "(" & _
        "аяихлос_тилокоциоу varchar(20) ," & _
        "тупос varchar(20) ," & _
        "тупос_пистытийоу varchar(20) ," & _
        "глеяолгмиа_ейдосгс DATE ," & _
        "еныжкгсг  BIT  , " & _
        "посо  FLOAT  , " & _
        "глеяолгмиа_еныжкгсгс DATE ," & _
        "аяихлос_епитацгс VARCHAR(30) ," & _
        "вяеысг  FLOAT , " & _
        "пистысг  FLOAT , " & _
        "упокоипо  FLOAT , " & _
        "PRIMARY KEY (аяихлос_тилокоциоу) )"
        DB1.Execute STATEMENT
        statement_B = " insert into ONOMATA_ETAIRION_ABCDEF(омолата_етаияиым)" & _
        " values ( '" & UCase(Text1.Text) & "' )"
        DB1.Execute statement_B
        DB1.Close
        MkDir App.Path & "\ETAIRIES\" & UCase(Text1.Text)
        MkDir App.Path & "\ETAIRIES\" & UCase(Text1.Text) & "\пкгяылема_" & UCase(Text1.Text)
        MkDir App.Path & "\ETAIRIES\" & UCase(Text1.Text) & "\апкгяыта_" & UCase(Text1.Text)
        SOURCE = App.Path & "\ind.jpg"
        DESTINATION = App.Path & "\ETAIRIES\" & UCase(Text1.Text) & "\пкгяылема_" & UCase(Text1.Text) & "\index.jpg"
        DESTINATION_B = App.Path & "\ETAIRIES\" & UCase(Text1.Text) & "\апкгяыта_" & UCase(Text1.Text) & "\index.jpg"
        FileCopy SOURCE, DESTINATION
        FileCopy SOURCE, DESTINATION_B
        S2 = App.Path & "\TXTS\TEMP.TXT"
        D2 = App.Path & "\TXTS\" & UCase(Text1.Text) & ".TXT"
        FileCopy S2, D2
        Text1.Text = ""
        If DB1.STATE = 1 Then DB1.Close
        ANAZHTHSH_ETAIRION (" ")
    Else
        Text1.Text = ""
    End If
End If
GoTo TELOS:

KENO:
If IND = 1 Then
    GoTo TELOS:
Else
    MsgBox ("дем пкгйтяокоцгсате йамема омола етаияиас"), vbCritical, "пяосовг !!!"
    IND = 2
    If DB1.STATE = 1 Then DB1.Close
End If


ER_MHKOS:
If IND = 2 Then
    GoTo TELOS:
Else
    MsgBox ("то омола етаияиас дем лпояеи ма пеяиевеи пеяиссотеяоус апо 50 ваяайтгяес"), vbCritical, "пяосовг !!!"
    If DB1.STATE = 1 Then DB1.Close
End If

ER:
If IND = 2 Then
    GoTo TELOS:
Else
    MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"
    If DB1.STATE = 1 Then DB1.Close
End If

TELOS:
If DB1.STATE = 1 Then DB1.Close
ANAZHTHSH_ETAIRION ("")
If Text4.Text <= 31 Then
    DataGrid1.Height = (Text4.Text * 320) + 215
Else
    DataGrid1.Height = 10135
End If
End Sub

Private Sub Command2_Click()
On Error GoTo ER:

If RS1.STATE = 1 Then RS1.Close
If DB1.STATE = 1 Then DB1.Close
DB1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
       "Data Source=" & "\databases\ETAIRIES.mdb" & ";" & _
      "Persist Security Info=False"
DB1.Open App.Path & "\databases\ETAIRIES.mdb"
RS1.Open "[ONOMATA_ETAIRION_ABCDEF]", DB1, adOpenDynamic, adLockBatchOptimistic
Dim C, C1 As Integer
C = 1
C1 = 1

If Text1.Text = "" Then
    MsgBox ("дем пкгйтяокоцгсате йамема омола етаияиас"), vbCritical, "пяосовг !!!"
    Command2.Caption = "еуяесг"
    GoTo TELOS:
End If
If Command2.Caption = "еуяесг" Then
    If RS1.BOF = RS1.EOF Then GoTo NIK:
    RS1.MoveFirst
NIK:
    Do While Not RS1.EOF
        If RS1![омолата_етаияиым] = UCase(Text1.Text) Then
            C = C + 1
            RS1.MoveNext
        Else
            RS1.MoveNext
        End If
    Loop
    
    If C = 1 Then
        MsgBox ("дем евеи йатавыягхг етаияиа ле то омола поу дысате. паяайакы екецнте то омола етаияиас."), vbCritical, "пяосовг !!!"
        GoTo TELOS:
    Else
        Text3.Text = UCase(Text1.Text)
        If MsgBox("ам хекете ма пяовыягсете се диояхысг тым стоивеиым тгс етаияиас тоте патгсте ой, пкгйтяокоцгсте то мео омола етаияиас йаи патгсте то пкгйтяо 'диояхысг'. диажояетийа патгсте CANCEL.", vbOKCancel, "") = vbOK Then
            Command2.Caption = "диояхысг"
            Text1.SetFocus
        Else
            Text1.Text = ""
            GoTo TELOS:
        End If
    End If
Else 'AN TO CAPTION EINAI диояхысг
    'ELEGXOS MHN EXEI DOTHEI TO IDIO ONOMA
    If UCase(Text1.Text) = Text3.Text Then
        MsgBox ("пкгйтяокоцгсате то идио омола етаияиас"), vbCritical, "пяосовг!!!"
        Command2.Caption = "еуяесг"
        Text1.Text = ""
        GoTo TELOS:
    End If
    If RS1.BOF = RS1.EOF Then GoTo NIK1:
        RS1.MoveNext
NIK1:
        Do While Not RS1.EOF
            If RS1![омолата_етаияиым] = Text1.Text Then
                C1 = C1 + 1
                RS1.MoveNext
            Else
                RS1.MoveNext
            End If
        Loop
        If C1 <> 1 Then
            MsgBox ("TO ONOMA етаияиас поу дысате упаявеи гдг.паяайакы диояхысте"), vbCritical, "пяосовг !!!"
            Command2.Caption = "еуяесг"
            GoTo TELOS:
        Else
            
        End If
    If MsgBox("еисте бебаиои оти хекете ма пяовыягсете се диояхысг тым стоивеиым тгс етаияиас?", vbOKCancel, "") = vbOK Then
        
        
        Dim TXT_APO, FOLDER_APO_1, FOLDER_APO_2, FOLDER_APO_3 As String
        Dim TXT_SE, FOLDER_SE_1, FOLDER_SE_2, FOLDER_SE_3 As String
        Dim PROTASH, STATEMENT, STATEMENT1 As String
        
        
        
        ' METONOMASIA ARXEIOY TXT ETAIRIAS
        TXT_APO = App.Path & "\TXTS\" & UCase(Text3.Text) & ".TXT"
        TXT_SE = App.Path & "\TXTS\" & UCase(Text1.Text) & ".TXT"
        Name TXT_APO As TXT_SE
        
        'летомоласиа жайеком етаияиас циа жыто
        FOLDER_APO_1 = App.Path & "\ETAIRIES\" & UCase(Text3.Text)
        FOLDER_APO_2 = App.Path & "\ETAIRIES\" & UCase(Text3.Text) & "\пкгяылема_" & UCase(Text3.Text)
        FOLDER_APO_3 = App.Path & "\ETAIRIES\" & UCase(Text3.Text) & "\апкгяыта_" & UCase(Text3.Text)
        FOLDER_SE_1 = App.Path & "\ETAIRIES\" & UCase(Text1.Text)
        FOLDER_SE_2 = App.Path & "\ETAIRIES\" & UCase(Text3.Text) & "\пкгяылема_" & UCase(Text1.Text)
        FOLDER_SE_3 = App.Path & "\ETAIRIES\" & UCase(Text3.Text) & "\апкгяыта_" & UCase(Text1.Text)
        Name FOLDER_APO_3 As FOLDER_SE_3
        Name FOLDER_APO_2 As FOLDER_SE_2
        Name FOLDER_APO_1 As FOLDER_SE_1
        
        ' аккацг омолатос етаияиас стом пимайа ONOMATA_ETAIRION_ABCDEF
        PROTASH = "UPDATE ONOMATA_ETAIRION_ABCDEF SET омолата_етаияиым='" & UCase(Text1.Text) & _
        "' WHERE омолата_етаияиым='" & UCase(Text3.Text) & "'"
        DB1.Execute PROTASH
        
        'аккацг омолатос пимайа поу амтоистгвг стгм етаияиа
        STATEMENT = " select * into " & UCase(Text1.Text) & " from " & Text3.Text
        DB1.Execute STATEMENT
        STATEMENT1 = "DROP TABLE " & UCase(Text3.Text)
        DB1.Execute STATEMENT1
        MsgBox ("г летомоласиа тоу омолатос етаияиас ециме ле епитувиа"), , ""
        Adodc1.Refresh
        DataGrid1.Refresh
        Text1.Text = ""
        Command2.Caption = "еуяесг"
        
    Else
        Command2.Caption = "еуяесг"
        Text1.Text = ""
        GoTo TELOS:
    End If
    
End If
GoTo TELOS:

ER:
    MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
If RS1.STATE = 1 Then RS1.Close
If DB1.STATE = 1 Then DB1.Close
End Sub

Private Sub Command3_Click()
If DB1.STATE = 1 Then DB1.Close
'On Error GoTo ER:
Dim DBB As New ADODB.Connection
Dim RSB As New ADODB.Recordset
Dim STATEMENT, statement_B As String
Dim IND, index As Integer
index = 1
IND = 1
DB1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
       "Data Source=" & "\databases\ETAIRIES.mdb" & ";" & _
      "Persist Security Info=False"
DB1.Open App.Path & "\databases\ETAIRIES.mdb"
RS1.Open "[ONOMATA_ETAIRION_ABCDEF]", DB1, adOpenDynamic, adLockBatchOptimistic

If Text1.Text = "" Then
    GoTo KENO:
End If

If RS1.BOF = RS1.EOF Then GoTo NIK:
    RS1.MoveFirst
NIK:
    Do While Not RS1.EOF
        If RS1![омолата_етаияиым] = UCase(Text1.Text) Then
            index = index + 1
            RS1.MoveNext
        Else
            RS1.MoveNext
        End If
    Loop
    If index = 1 Then
        GoTo ER_NAME:
    Else
        If MsgBox("хекете ма пяовыягсете се диацяажг тгс етаияиас :  " & Text1.Text, vbOKCancel, "") = vbOK Then
            STATEMENT = " drop table " & UCase(Text1.Text)
            DB1.Execute STATEMENT

            statement_B = " delete from ONOMATA_ETAIRION_ABCDEF " & _
            " where омолата_етаияиым = '" & UCase(Text1.Text) & "'"
            DB1.Execute statement_B

            Kill App.Path & "\ETAIRIES\" & UCase(Text1.Text) & "\пкгяылема_" & UCase(Text1.Text) & "\*.*"
            Kill App.Path & "\ETAIRIES\" & UCase(Text1.Text) & "\апкгяыта_" & UCase(Text1.Text) & "\*.*"
            RmDir App.Path & "\ETAIRIES\" & UCase(Text1.Text) & "\пкгяылема_" & UCase(Text1.Text)
            RmDir App.Path & "\ETAIRIES\" & UCase(Text1.Text) & "\апкгяыта_" & UCase(Text1.Text)
            RmDir App.Path & "\ETAIRIES\" & UCase(Text1.Text)
            Kill App.Path & "\TXTS\" & UCase(Text1.Text) & ".TXT"
            DB1.Close
            '************************* SBHSIMO ETAIRIAS APO TZIROYS ******************
            Dim CDE As Integer
            CDE = 1
            DBB.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
            "Data Source=" & "\databases\ETAIRIES.mdb" & ";" & _
            "Persist Security Info=False"
            DBB.Open App.Path & "\databases\ETAIRIES.mdb"
            RSB.Open "[TZIROI]", DBB, adOpenDynamic, adLockBatchOptimistic
            If RSB.EOF = RSB.BOF Then GoTo GG:
            RSB.MoveFirst
GG:
            Do While Not RSB.EOF
                If RSB![ETAIRIA] = Trim(UCase(Text1.Text)) Then
                    CDE = 2
                End If
                RSB.MoveNext
            Loop
            If CDE = 1 Then
    
            Else
                SAS = " DELETE FROM TZIROI WHERE ETAIRIA='" & Text1.Text & "'"
                DBB.Execute SAS
            End If
           '*************************************************************************
            Text1.Text = ""
            If DB1.STATE = 1 Then DB1.Close
            ANAZHTHSH_ETAIRION (" ")
        Else
            Text1.Text = ""
        End If
    End If
    
GoTo TELOS:

KENO:
If IND = 1 Then
    MsgBox ("дем пкгйтяокоцгсате йамема омола етаияиас"), vbCritical, "пяосовг !!!"
    IND = 22
    GoTo TELOS:
End If
    
ER_NAME:
If IND = 1 Then
    MsgBox ("дем евеи йатавыягхг етаияиа ле то омола поу дысате. паяайакы екецнте то омола етаияиас."), vbCritical, "пяосовг !!!"
    IND = 22
    GoTo TELOS:
Else
    GoTo TELOS:
End If

ER:
If IND = 22 Then
    GoTo TELOS:
Else
    MsgBox ("дем упаявеи йалиа етаияиа йатавыяглемг ╧ йапоио амапамтево пяобкгла елжамистгйе. ам евете есты йаи лиа етаияиа йатавыяглемг етаияиа тоте йапоио апяосдойгто кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс."), vbCritical, ""
    GoTo TELOS:
End If


ANAZHTHSH_ETAIRION ("")

TELOS:
ANAZHTHSH_ETAIRION ("")
If DB1.STATE = 1 Then DB1.Close
If Text4.Text <= 31 Then
    DataGrid1.Height = (Text4.Text * 320) + 215
Else
    DataGrid1.Height = 10135
End If
End Sub

Private Sub Command4_Click()
On Error GoTo ER:

Dim ASD As String
ASD = Trim(UCase(Text2.Text))
ANAZHTHSH_ETAIRION (UCase(Text2.Text))
GoTo TELOS:

ER:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
If Text4.Text <= 31 Then
    DataGrid1.Height = (Text4.Text * 320) + 210
Else
    DataGrid1.Height = 10135
End If
End Sub

Private Sub Command5_Click()
On Error GoTo ER:

Form3.Hide
Unload Form3
GoTo TELOS:

ER:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
End Sub

Private Sub DataGrid1_Click()
Text1.Text = DataGrid1.Columns(0).Text
Text1.SetFocus
End Sub

Private Sub Form_Load()
On Error GoTo ER:

Dim B As String
DataGrid1.Font.Bold = True
DataGrid1.Font.Size = 10
DataGrid1.DefColWidth = 4950
Dim A As String
Dim DATABASE_FILE As String
DATABASE_FILE = App.Path & "/DATABASES/ETAIRIES.MDB"

Adodc1.ConnectionString = _
"PROVIDER=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & DATABASE_FILE & ";"
Adodc1.RecordSource = "SELECT * FROM ONOMATA_ETAIRION_ABCDEF ORDER BY омолата_етаияиым"
' Bind the ADODC to the DataGrid.
Set DataGrid1.DataSource = Adodc1

Adodc2.ConnectionString = _
"PROVIDER=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & DATABASE_FILE & ";"
Adodc2.RecordSource = "SELECT COUNT(омолата_етаияиым)FROM ONOMATA_ETAIRION_ABCDEF"
' Bind the ADODC to the DataGrid.
Set DataGrid2.DataSource = Adodc2
Text4.Text = DataGrid2.Text

If Text4.Text <= 31 Then
    DataGrid1.Height = (Text4.Text * 320) + 210
Else
    DataGrid1.Height = 10135
End If
Label3.Caption = "аяихлос сумеяцафолемым етаияиым : " & Text4.Text

GoTo TELOS:

ER:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ER:

Form3.Hide
Unload Form3
GoTo TELOS:

ER:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
End Sub

Private Sub Text1_Change()
Text1.Text = Trim(Text1.Text)
End Sub

Private Sub Text2_Change()
Text2.Text = Trim(Text2.Text)
End Sub

Private Sub Text3_Change()
Text3.Text = Trim(Text3.Text)
End Sub
