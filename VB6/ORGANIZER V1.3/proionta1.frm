VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form proionta1 
   BackColor       =   &H80000013&
   Caption         =   "пяосжояес"
   ClientHeight    =   10485
   ClientLeft      =   105
   ClientTop       =   465
   ClientWidth     =   15195
   LinkTopic       =   "Form10"
   ScaleHeight     =   10485
   ScaleWidth      =   15195
   Begin VB.TextBox Text6 
      Height          =   615
      Left            =   5880
      TabIndex        =   14
      Text            =   "Text6"
      Top             =   9600
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      Height          =   495
      Left            =   7440
      Picture         =   "proionta1.frx":0000
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   13
      Top             =   9840
      Width           =   495
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Text            =   "Text5"
      Top             =   9120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   13920
      TabIndex        =   11
      Text            =   "Text4"
      Top             =   7680
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   6120
      TabIndex        =   10
      Text            =   "Text3"
      Top             =   7800
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   7565
      Left            =   120
      TabIndex        =   9
      Top             =   0
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   13335
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   7920
      Top             =   7680
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
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
      Height          =   375
      Left            =   0
      Top             =   7680
      Visible         =   0   'False
      Width           =   1920
      _ExtentX        =   3387
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
   Begin VB.CommandButton Command6 
      BackColor       =   &H00F1C896&
      Caption         =   "елжамисг пяоиомтос"
      Height          =   975
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   9360
      Width           =   2775
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00F1C896&
      Caption         =   "елжамисг етаияиас"
      Height          =   975
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   8280
      Width           =   2775
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   9720
      TabIndex        =   6
      Top             =   7680
      Width           =   3615
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00E9C5AD&
      Caption         =   "диацяажг"
      Height          =   735
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   8280
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00E9C5AD&
      Caption         =   "еуяесг"
      Height          =   735
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   8280
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E9C5AD&
      Caption         =   "еццяажг"
      Height          =   735
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8280
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   7680
      Width           =   3855
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Height          =   7560
      Left            =   7800
      TabIndex        =   1
      Top             =   0
      Width           =   7305
      _ExtentX        =   12885
      _ExtentY        =   13335
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
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E9C5AD&
      Caption         =   "епистяожг сто аявийо лемоу"
      Height          =   735
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   9600
      Width           =   1815
   End
   Begin VB.Line Line3 
      Visible         =   0   'False
      X1              =   11475
      X2              =   11475
      Y1              =   10
      Y2              =   11000
   End
   Begin VB.Line Line2 
      Visible         =   0   'False
      X1              =   3825
      X2              =   3825
      Y1              =   15
      Y2              =   11000
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      Visible         =   0   'False
      X1              =   7650
      X2              =   7650
      Y1              =   0
      Y2              =   10800
   End
End
Attribute VB_Name = "proionta1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error GoTo ER:
proionta1.Hide
Unload proionta1
GoTo TELOS:

ER:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

GoTo TELOS:

TELOS:
End Sub

Private Sub Command2_Click()
On Error GoTo ER:
Dim dbp As New ADODB.Connection
Dim rsp As New ADODB.Recordset
Dim L, i As Integer
Dim STATEMENT, STATEMENT1 As String

If rsp.STATE = 1 Then rsp.Close
If dbp.STATE = 1 Then dbp.Close
dbp.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & "\databases\PROION.mdb" & ";" & _
"Persist Security Info=False"
dbp.Open App.Path & "\databases\PROION.mdb"
rsp.Open "[PROIONTA_ABC]", dbp, adOpenDynamic, adLockBatchOptimistic

'****************** ELEGXOI **************************************************
If Text1.Text = "" Then GoTo KENO:
Text1.Text = Trim(Text1.Text)
Text1.Text = UCase(Text1.Text)
L = Len(Text1.Text)
If L > 49 Then GoTo BIG:
For i = 1 To L
If Mid(Text1.Text, i, i) = "." Or Mid(Text1.Text, i, i) = "!" Or Mid(Text1.Text, i, i) = "'" _
Or Mid(Text1.Text, i, i) = "]" Or Mid(Text1.Text, i, i) = "[" Then GoTo XARA:
Next i
If rsp.BOF = rsp.EOF Then GoTo NIK:
rs1.MoveFirst
NIK:
Do While Not rsp.EOF
If Text1.Text = rsp![пяоиомта] Then
    GoTo ER1:
Else
    rsp.MoveNext
End If
Loop
'********************* LEITOYRGIA ******************************************
 If MsgBox("хекете ма пяовыягсете стгм йатавыягсг тоу пяоиомтос ле омола :  " & Text1.Text, vbOKCancel, "") = vbOK Then
    STATEMENT = " create table " & Text1.Text & " (" & _
        "омола_етаияиас varchar(50) ," & _
        "тилг FLOAT ," & _
        "ейптысг FLOAT ," & _
        "сглеиысг varchar(250)," & _
        "PRIMARY KEY (омола_етаияиас) )"
    dbp.Execute STATEMENT
    STATEMENT1 = "INSERT INTO PROIONTA_ABC values ( '" & UCase(Text1.Text) & "' )"
    dbp.Execute STATEMENT1
    rsp.Fields.Refresh
    If rsp.STATE = 1 Then rsp.Close
    If dbp.STATE = 1 Then dbp.Close
    Adodc1.Refresh
    Text3.Text = Adodc1.Recordset.RecordCount
    If CInt(Text3.Text) < 24 Then
        DataGrid1.Height = 355 + (302.5 * CInt(Text3.Text))
    Else
        DataGrid1.Height = 7565
    End If
    Adodc1.Refresh
End If
Text1.Text = ""
GoTo TELOS:
'***************************************************************************

'******** ANTIMETOPISIS ELEGXON ***********************************************
KENO:
MsgBox ("дем дысате йамема омола"), vbCritical, "пяосовг !!!"
GoTo TELOS:

BIG:
MsgBox ("то омола поу дысате еимаи поку лецако. паяайакы дысте омола пяоиомтос поу ма пеяиевеи кицотеяоус апо 50 ваяайтгяес"), vbCritical, "пяосовг !!!"
GoTo TELOS:

XARA:
MsgBox ("то омола пяоиомтос дем лпояеи ма пеяиевеи тоус ваяайтгяес . ! ' [ ] паяайакы диояхысте то омола поу дысате "), vbCritical, "пяосовг !!!"
GoTo TELOS:

ER1:
MsgBox ("то омола пяоиомтос поу дысате упаявеи гдг"), vbCritical, "пяосовг !!!"
GoTo TELOS:


ER:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"
    

TELOS:
If rsp.STATE = 1 Then rsp.Close
If dbp.STATE = 1 Then dbp.Close
End Sub
 
Private Sub Command3_Click()
On Error GoTo ER:
Dim dbp As New ADODB.Connection
Dim rsp As New ADODB.Recordset
Dim C, L As Integer
C = 1
Dim STATEMENT, STATEMENT1, STATEMENT2 As String

If rsp.STATE = 1 Then rsp.Close
If dbp.STATE = 1 Then dbp.Close
dbp.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & "\databases\PROION.mdb" & ";" & _
"Persist Security Info=False"
dbp.Open App.Path & "\databases\PROION.mdb"
rsp.Open "[PROIONTA_ABC]", dbp, adOpenDynamic, adLockBatchOptimistic

If Command3.Caption = "еуяесг" Then ' ** еуяесг *********************
    If Text1.Text = "" Then GoTo KENO:
    If rsp.BOF = rsp.EOF Then GoTo NIK1:
    rsp.MoveFirst
NIK1:
    Do While Not rsp.EOF
        If Text1.Text = rsp![пяоиомта] Then
            Text5.Text = rsp![пяоиомта]
            C = C + 1
            rsp.MoveNext
        Else
            rsp.MoveNext
        End If
    Loop
    If C = 1 Then
        MsgBox ("дем бяехгйе пяоиом ле то омола поу дысате"), vbCritical, "пяосовг !!!"
        Command3.Caption = "еуяесг"
        GoTo TELOS:
    Else
        Text1.Text = Text5.Text
        MsgBox ("то омола пяоиомтос поу дысате бяехгйе.ам хекете ма пяовыягсете се диояхысг тоу пяоиомтос цяаьте то мео омола йаи патгсте диояхысг диажояетийа патгсте сто еийомидио "), , ""
        Command3.Caption = "диояхысг"
        GoTo TELOS:
    End If
Else  ' ****************************  диояхысг ****************************
     If MsgBox("хекете ма пяовыягсете се диояхысг тоу омолатос тоу пяоиомтос", vbOKCancel, "") = vbOK Then
        '************ ELEGXOI ***********
        If Text1.Text = "" Then GoTo KENO:
        Text1.Text = Trim(Text1.Text)
        Text1.Text = UCase(Text1.Text)
        L = Len(Text1.Text)
        If L > 49 Then GoTo BIG:
        For i = 1 To L
        If Mid(Text1.Text, i, i) = "." Or Mid(Text1.Text, i, i) = "!" Or Mid(Text1.Text, i, i) = "'" _
        Or Mid(Text1.Text, i, i) = "]" Or Mid(Text1.Text, i, i) = "[" Then GoTo XARA:
        Next i
        If rsp.BOF = rsp.EOF Then GoTo NIK:
        rs1.MoveFirst
NIK:
        Do While Not rsp.EOF
            If Text1.Text = rsp![пяоиомта] Then
            GoTo ER1:
        Else
            rsp.MoveNext
        End If
        Loop
        '*************************  LEITOYRGIA  ***************************
        STATEMENT1 = "select * into " & UCase(Text1.Text) & " from " & Text5.Text
        dbp.Execute STATEMENT1
        STATEMENT2 = "DROP TABLE " & UCase(Text5.Text)
        dbp.Execute STATEMENT2
        STATEMENT = "UPDATE PROIONTA_ABC SET пяоиомта ='" & Text1.Text & "' WHERE пяоиомта = '" & Text5.Text & "'"
        dbp.Execute STATEMENT
        'DIORTHOSI SE ETAIRIES TOY ONOMATOS PROIONTOS.SE KATHE PROION
        'DIORTHONETE TO ONOMA PROIONTOS AN YPARXEI
        Dim STATEM As String
        Dim rsp11 As New ADODB.Recordset
        If rsp.STATE = 1 Then rsp.Close
        If dbp.STATE = 1 Then dbp.Close
        dbp.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
        "Data Source=" & "\databases\PROION_ETAIRIA.mdb" & ";" & _
        "Persist Security Info=False"
        dbp.Open App.Path & "\databases\PROION_ETAIRIA.mdb"
        rsp.Open "[ETAIRIES_ABC]", dbp, adOpenDynamic, adLockBatchOptimistic
        
        If rsp.EOF = rsp.BOF Then GoTo NH:
        rsp.MoveFirst
NH:
        Do While Not rsp.EOF
            rsp11.Open "[" & rsp![етаияиес] & "]", dbp, adOpenDynamic, adLockBatchOptimistic
            If rsp11.EOF = rsp11.BOF Then GoTo jj:
            rsp11.MoveFirst
jj:
            Do While Not rsp11.EOF
                Text6.Text = rsp11![омола_пяоиомтос]
                If Text6.Text <> Text5.Text Then
                   rsp11.MoveNext
                Else
                    STATEM = " UPDATE " & rsp![етаияиес] & _
                    " SET омола_пяоиомтос='" & Text1.Text & _
                    "' WHERE омола_пяоиомтос='" & Text6.Text & "'"
                    dbp.Execute STATEM
                    rsp11.MoveNext
                End If
            Loop
            If rsp11.STATE = 1 Then rsp11.Close
            rsp.MoveNext
        Loop
        Adodc1.Refresh
        DataGrid1.Refresh
        MsgBox ("г диояхысг тоу омолатос тоу пяоиомтос ециме ле епитувиа"), vbOKOnly, ""
        Command3.Caption = "еуяесг"
        Text1.Text = ""
        Text5.Text = ""
        GoTo TELOS:
     Else
        Command3.Caption = "еуяесг"
        Text1.Text = ""
        GoTo TELOS:
     End If
End If
    '*********************************************************************
    
    '******** ANTIMETOPISIS ELEGXON ***********************************************
KENO:
    MsgBox ("дем дысате йамема омола"), vbCritical, "пяосовг !!!"
    GoTo TELOS:

BIG:
    MsgBox ("то омола поу дысате еимаи поку лецако. паяайакы дысте омола пяоиомтос поу ма пеяиевеи кицотеяоус апо 50 ваяайтгяес"), vbCritical, "пяосовг !!!"
    GoTo TELOS:

XARA:
    MsgBox ("то омола пяоиомтос дем лпояеи ма пеяиевеи тоус ваяайтгяес . ! ' [ ] паяайакы диояхысте то омола поу дысате "), vbCritical, "пяосовг !!!"
    GoTo TELOS:

ER1:
    MsgBox ("то омола пяоиомтос поу дысате упаявеи гдг"), vbCritical, "пяосовг !!!"
    GoTo TELOS:

ER:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"
    


TELOS:
If rsp11.STATE = 1 Then rsp.Close
If rsp.STATE = 1 Then rsp.Close
If dbp.STATE = 1 Then dbp.Close
End Sub

Private Sub Command4_Click()
On Error GoTo ER:
Dim dbp As New ADODB.Connection
Dim rsp As New ADODB.Recordset
Dim STATEMENT, STATEMENT1 As String
Dim C As Integer
C = 1
Text1.Text = Trim(Text1.Text)
Text1.Text = UCase(Text1.Text)


If rsp.STATE = 1 Then rsp.Close
If dbp.STATE = 1 Then dbp.Close
dbp.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & "\databases\PROION.mdb" & ";" & _
"Persist Security Info=False"
dbp.Open App.Path & "\databases\PROION.mdb"
rsp.Open "[PROIONTA_ABC]", dbp, adOpenDynamic, adLockBatchOptimistic

'**************************** ELEGXOI *******************************************
If Text1.Text = "" Then GoTo KENO:
If rsp.BOF = rsp.EOF Then GoTo NIK:
rs1.MoveFirst
NIK:
Do While Not rsp.EOF
If Text1.Text <> rsp![пяоиомта] Then
    rsp.MoveNext
Else
    C = C + 1
    rsp.MoveNext
End If
Loop
If C <> 1 Then
    If MsgBox("хекете ма пяовыягсете се диацяажг тоу пяоиомтос", vbOKCancel, "пяосовг") = vbOK Then
' ****** DIAGRAFH PINAKA PROIONTOS KAI ONOMATOS APO PROIONTA_ABC
' STHN BASH PROIONTA.MDB *********************
        STATEMENT = " drop table " & Text1.Text
        STATEMENT1 = "delete from PROIONTA_ABC" & _
        " where пяоиомта = '" & Text1.Text & "'"
        dbp.Execute STATEMENT
        dbp.Execute STATEMENT1
        rsp.Fields.Refresh
        If rsp.STATE = 1 Then rsp.Close
        If dbp.STATE = 1 Then dbp.Close
        Adodc1.Refresh
        Text3.Text = Adodc1.Recordset.RecordCount
        If CInt(Text3.Text) < 24 Then
            DataGrid1.Height = 355 + (302.5 * CInt(Text3.Text))
        Else
            DataGrid1.Height = 7565
        End If
        Adodc1.Refresh
        
        
'**** DIAGRAFH APO KATHE ETAIRIA TOY PROINTOS.AN STHN ETAIRIA PLEON
'     DEN YPARXEI  ALLO PROION TOTE NA SBHMEI KAI H ETAIRIA ***************************
        Dim dbp11 As New ADODB.Connection
        Dim rsp11 As New ADODB.Recordset
        dbp11.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
        "Data Source=" & "\databases\PROION_ETAIRIA.mdb" & ";" & _
        "Persist Security Info=False"
        dbp11.Open App.Path & "\databases\PROION_ETAIRIA.mdb"
        rsp.Open "[ETAIRIES_ABC]", dbp11, adOpenDynamic, adLockBatchOptimistic
        Dim stat, stat2, SKSK, SKSK2, SKSK3 As String
        Dim coun As Integer
        coun = 0
        If rsp.BOF = rsp.EOF Then GoTo NIKKK:
        rsp.MoveFirst
NIKKK:
        Do While Not rsp.EOF
           rsp11.Open "[" & rsp![етаияиес] & "]", dbp11, adOpenDynamic, adLockBatchOptimistic
           If rsp11.EOF = rsp11.BOF Then GoTo KKK:
           rsp11.MoveFirst
KKK:
           Do While Not rsp11.EOF
                If rsp11![омола_пяоиомтос] = Text1.Text Then
                    stat = " delete from " & rsp![етаияиес] & _
                    " where омола_пяоиомтос='" & rsp11![омола_пяоиомтос] & _
                    "'"
                    dbp11.Execute stat
                    rsp11.MoveNext
                Else
                    coun = coun + 1
                    rsp11.MoveNext
                End If
            Loop
            ' KATAGRAFH SE PINAKA TEMP_ABC TON ETAIRION POY DEN EXOYN KAMIA EGRAFH
            If coun = 0 Then
                stat2 = "INSERT INTO TEMP_ABC values ( '" & rsp![етаияиес] & "' )"
                dbp11.Execute stat2
            End If
            If rsp11.STATE = 1 Then rsp11.Close
            rsp.MoveNext
            coun = 0
        Loop
        'DIAGRAFH ETAIRION POY EIANAI MESA STON TEMP_ABC
        If rsp11.STATE = 1 Then rsp11.Close
        rsp11.Open "[TEMP_ABC]", dbp11, adOpenDynamic, adLockBatchOptimistic
        If rsp11.EOF = rsp11.BOF Then GoTo NHNH:
        rsp11.MoveFirst
NHNH:
        Do While Not rsp11.EOF
            SKSK = "DROP TABLE " & rsp11![DIA_ETAIRION]
            SKSK2 = "DELETE FROM ETAIRIES_ABC " & _
            "WHERE етаияиес='" & rsp11![DIA_ETAIRION] & "'"
            dbp11.Execute SKSK
            dbp11.Execute SKSK2
            rsp11.MoveNext
        Loop
        SKSK3 = "DELETE * FROM TEMP_ABC"
        dbp11.Execute SKSK3
        ' REFRESH TOY ADODC2
        rsp.Fields.Refresh
        rsp11.Fields.Refresh
        If rsp.STATE = 1 Then rsp.Close
        If rsp11.STATE = 1 Then rsp11.Close
        If dbp.STATE = 1 Then dbp.Close
        If dbp11.STATE = 1 Then dbp11.Close
        Adodc2.Refresh
        Text4.Text = Adodc2.Recordset.RecordCount
        If CInt(Text4.Text) < 24 Then
            DataGrid2.Height = 355 + (302.5 * CInt(Text4.Text))
        Else
            DataGrid2.Height = 7565
        End If
        Adodc2.Refresh
        Text1.Text = ""
        GoTo TELOS:
    Else
        Text1.Text = ""
        GoTo TELOS:
    End If
Else
    MsgBox ("дем бяехгйе йатавыягсг пяоиомтос ле то омола поу дысате"), vbCritical, "пяосовг !!!"
    GoTo TELOS:
End If

KENO:
MsgBox ("дем дысате йамема омола пяоиомтос пяос диацяажг"), vbCritical, "пяосовг !!!"
GoTo TELOS:

ER:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"
    


TELOS:
If rsp.STATE = 1 Then rsp.Close
If rsp11.STATE = 1 Then rsp.Close
If dbp.STATE = 1 Then dbp.Close
If dbp11.STATE = 1 Then dbp.Close
End Sub

Private Sub Command5_Click()
On Error GoTo ER:
Dim dbp As New ADODB.Connection
Dim rsp As New ADODB.Recordset
Text2.Text = Trim(Text2.Text)
Text2.Text = UCase(Text2.Text)
Dim C As Integer
C = 1

If rsp.STATE = 1 Then rs1.Close
If dbp.STATE = 1 Then db1.Close
dbp.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & "\databases\PROION_ETAIRIA.mdb" & ";" & _
"Persist Security Info=False"
dbp.Open App.Path & "\databases\PROION_ETAIRIA.mdb"
rsp.Open "[ETAIRIES_ABC]", dbp, adOpenDynamic, adLockBatchOptimistic

'**************************** ELEGXOI *******************************************
If Text2.Text = "" Then
    MsgBox ("дем дысате йамема омола етаияиас"), vbCritical, "пяосовг !!!"
    GoTo TELOS:
End If
If rsp.BOF = rsp.EOF Then GoTo NIK:
rsp.MoveFirst
NIK:
Do While Not rsp.EOF
    If rsp![етаияиес] = Text2.Text Then
        C = C + 1
        rsp.MoveNext
    Else
        rsp.MoveNext
    End If
Loop
If C = 1 Then
    MsgBox ("дем упаявеи етаияиа ле то омола поу дысате"), vbCritical, "пяосовг !!!"
    GoTo TELOS:
End If
' **************** LEITOYRGIA *****************************************
proionta3.Text5.Text = Text2.Text
Load proionta3
proionta3.Show
Text2.Text = ""
GoTo TELOS:

ER:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"


TELOS:
If rsp.STATE = 1 Then rsp.Close
If dbp.STATE = 1 Then dbp.Close
End Sub

Private Sub Command6_Click()
On Error GoTo ER:
Dim dbp As New ADODB.Connection
Dim rsp As New ADODB.Recordset
Text1.Text = Trim(Text1.Text)
Text1.Text = UCase(Text1.Text)
Dim i, L, C As Integer
C = 1

If rsp.STATE = 1 Then rsp.Close
If dbp.STATE = 1 Then dbp.Close
dbp.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & "\databases\PROION.mdb" & ";" & _
"Persist Security Info=False"
dbp.Open App.Path & "\databases\PROION.mdb"
rsp.Open "[PROIONTA_ABC]", dbp, adOpenDynamic, adLockBatchOptimistic

'**************************** ELEGXOI *******************************************
If Text1.Text = "" Then GoTo KENO:
L = Len(Text1.Text)
If L > 49 Then GoTo BIG:
For i = 1 To L
If Mid(Text1.Text, i, i) = "." Or Mid(Text1.Text, i, i) = "!" Or Mid(Text1.Text, i, i) = "'" _
Or Mid(Text1.Text, i, i) = "]" Or Mid(Text1.Text, i, i) = "[" Then GoTo XARA:
Next i
If rsp.BOF = rsp.EOF Then GoTo NIK:
rs1.MoveFirst
NIK:
Do While Not rsp.EOF
If Text1.Text <> rsp![пяоиомта] Then
    rsp.MoveNext
Else
    C = C + 1
    rsp.MoveNext
End If
Loop

'*************************** LEITOYRGIA ******************************************
If C <> 1 Then
    Load proionta2
    proionta2.Show
    GoTo TELOS:
Else
    GoTo ER1:
End If

'**************************************************************************************

'******** ANTIMETOPISIS ELEGXON ***********************************************
KENO:
MsgBox ("дем дысате йамема омола"), vbCritical, "пяосовг !!!"
GoTo TELOS:

BIG:
MsgBox ("то омола поу дысате еимаи поку лецако. паяайакы дысте омола пяоиомтос поу ма пеяиевеи кицотеяоус апо 50 ваяайтгяес"), vbCritical, "пяосовг !!!"
GoTo TELOS:

XARA:
MsgBox ("то омола пяоиомтос дем лпояеи ма пеяиевеи тоус ваяайтгяес . ! ' [ ] паяайакы диояхысте то омола поу дысате "), vbCritical, "пяосовг !!!"
GoTo TELOS:

ER1:
MsgBox ("то омола пяоиомтос поу дысате упаявеи гдг"), vbCritical, "пяосовг !!!"
GoTo TELOS:

ER:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"
    


TELOS:
Text1.Text = ""
If rsp.STATE = 1 Then rsp.Close
If dbp.STATE = 1 Then dbp.Close
End Sub

Private Sub Command7_Click()

End Sub

Private Sub DataGrid1_Click()
Text1.Text = DataGrid1.Columns(0).Text
End Sub

Private Sub DataGrid2_Click()
Text2.Text = DataGrid2.Columns(0).Text
End Sub

Private Sub Form_Load()
On Error GoTo ER:
Dim DATABASE_FILE, DATABASE_FILE1 As String
DATABASE_FILE = App.Path & "\databases\PROION.mdb"
DATABASE_FILE1 = App.Path & "\databases\PROION_ETAIRIA.mdb"

DataGrid1.DefColWidth = 6980
DataGrid1.Font.Size = 10
DataGrid1.HeadFont.Size = 12
DataGrid1.HeadFont.Bold = True
DataGrid2.DefColWidth = 6980
DataGrid2.Font.Size = 10
DataGrid2.HeadFont.Size = 12
DataGrid2.HeadFont.Bold = True

Adodc1.ConnectionString = _
"PROVIDER=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & DATABASE_FILE & ";"
Adodc1.RecordSource = _
"SELECT пяоиомта FROM PROIONTA_ABC ORDER BY пяоиомта"
Set DataGrid1.DataSource = Adodc1
Text3.Text = Adodc1.Recordset.RecordCount
If CInt(Text3.Text) < 24 Then
    DataGrid1.Height = 355 + (302.5 * CInt(Text3.Text))
Else
    DataGrid1.Height = 7565
End If
Adodc1.Refresh

Adodc2.ConnectionString = _
"PROVIDER=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & DATABASE_FILE1 & ";"
Adodc2.RecordSource = _
"SELECT етаияиес FROM ETAIRIES_ABC ORDER BY етаияиес"
Set DataGrid2.DataSource = Adodc2
Text4.Text = Adodc2.Recordset.RecordCount
If CInt(Text4.Text) < 24 Then
    DataGrid2.Height = 355 + (302.5 * CInt(Text4.Text))
Else
    DataGrid2.Height = 7565
End If
Adodc2.Refresh

GoTo TELOS:

ER:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"
    
TELOS:
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ER:
proionta1.Hide
Unload proionta1
GoTo TELOS:

ER:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

GoTo TELOS:

TELOS:
End Sub

Private Sub Picture1_Click()
On Error GoTo ER:
Text1.Text = ""
Text2.Text = ""
Command3.Caption = "еуяесг"
GoTo TELOS:

ER:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"
    

TELOS:
End Sub
