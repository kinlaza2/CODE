VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form proionta3 
   BackColor       =   &H80000013&
   Caption         =   "етаияиа"
   ClientHeight    =   10485
   ClientLeft      =   105
   ClientTop       =   465
   ClientWidth     =   15180
   LinkTopic       =   "Form10"
   ScaleHeight     =   10485
   ScaleWidth      =   15180
   Begin MSDataGridLib.DataGrid DataGrid3 
      Height          =   495
      Left            =   1680
      TabIndex        =   20
      Top             =   8640
      Visible         =   0   'False
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   873
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
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   495
      Left            =   240
      Top             =   8760
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
      Caption         =   "Adodc3"
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
   Begin VB.TextBox Text8 
      Height          =   495
      Left            =   2640
      TabIndex        =   19
      Text            =   "Text8"
      Top             =   600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Height          =   375
      Left            =   2160
      TabIndex        =   18
      Top             =   9000
      Visible         =   0   'False
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   495
      Left            =   2160
      Top             =   9600
      Visible         =   0   'False
      Width           =   1920
      _ExtentX        =   3387
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
   Begin VB.PictureBox Picture1 
      Height          =   375
      Left            =   3960
      Picture         =   "proionta3.frx":0000
      ScaleHeight     =   315
      ScaleWidth      =   435
      TabIndex        =   17
      Top             =   9960
      Width           =   495
   End
   Begin VB.TextBox Text7 
      Height          =   375
      Left            =   1320
      TabIndex        =   16
      Text            =   "Text7"
      Top             =   2160
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   240
      TabIndex        =   15
      Text            =   "Text6"
      Top             =   840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   690
      Left            =   3000
      Top             =   8640
      Visible         =   0   'False
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   1217
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
   Begin VB.TextBox Text5 
      Height          =   495
      Left            =   240
      TabIndex        =   14
      Text            =   "Text5"
      Top             =   240
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00E9C5AD&
      Caption         =   "епистяожг сто пяогцоулемо лемоу"
      Height          =   855
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   9480
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00E9C5AD&
      Caption         =   "диацяажг"
      Height          =   735
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   7800
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E9C5AD&
      Caption         =   "еуяесг"
      Height          =   735
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7800
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E9C5AD&
      Caption         =   "еццяажг"
      Height          =   735
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7800
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   9375
      Left            =   4560
      TabIndex        =   8
      Top             =   840
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   16536
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
   Begin VB.TextBox Text4 
      Height          =   1335
      Left            =   120
      TabIndex        =   7
      Top             =   6000
      Width           =   4335
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   4560
      Width           =   4335
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   3120
      Width           =   4335
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1680
      Width           =   4335
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
      Caption         =   "етаияиа :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1560
      TabIndex        =   9
      Top             =   120
      Width           =   12255
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      Visible         =   0   'False
      X1              =   7650
      X2              =   7650
      Y1              =   0
      Y2              =   11000
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000013&
      Caption         =   "сглеиысг"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   5520
      Width           =   4335
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000013&
      Caption         =   "ейптысг"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   4080
      Width           =   4335
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000013&
      Caption         =   "тилг"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   2640
      Width           =   4335
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000013&
      Caption         =   "омола пяоиомтос"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   4335
   End
   Begin VB.Line Line1 
      Visible         =   0   'False
      X1              =   4440
      X2              =   4440
      Y1              =   240
      Y2              =   8280
   End
End
Attribute VB_Name = "proionta3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error GoTo ER:
Dim dbp As New ADODB.Connection
Dim rsp As New ADODB.Recordset
Dim D As Integer
D = 0
Dim STATEMENT As String
Text1.Text = UCase(Trim(Text1.Text))
Text4.Text = UCase(Trim(Text4.Text))
If dbp.STATE = 1 Then dbp.Close
dbp.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
       "Data Source=" & "\databases\PROION_ETAIRIA.mdb" & ";" & _
      "Persist Security Info=False"
dbp.Open App.Path & "\databases\PROION_ETAIRIA.mdb"
rsp.Open "[" & Text5.Text & "]", dbp, adOpenDynamic, adLockBatchOptimistic
'************************  ELEGXOI  ************************************
' AN YPARXEI HDH H EGRAFH
If rsp.EOF = rsp.BOF Then GoTo NNIK:
    rsp.MoveFirst
NNIK:
    Do While Not rsp.EOF
        If rsp![омола_пяоиомтос] = Text1.Text Then
            D = D + 1
            rsp.MoveNext
        Else
            rsp.MoveNext
        End If
    Loop
' TEXT1
If Text1.Text = "" Then GoTo TEXT1KENO:
If Len(Text1.Text) > 49 Then GoTo TEXT1BIGLEN:
For i = 1 To Len(Text1.Text)
    If Mid(Text1.Text, i, i) = "." Or Mid(Text1.Text, i, i) = "!" Or _
    Mid(Text1.Text, i, i) = "[" Or Mid(Text1.Text, i, i) = "]" Then GoTo TEXT1WCHAR:
Next i
' TEXT2
If Text2.Text = "" Then GoTo TEXT2KENO:
If IsNumeric(Text2.Text) = False Then GoTo TEXT2NONUM:
' TEXT3
If Text3.Text = "" Then Text3.Text = 0
If IsNumeric(Text3.Text) = False Then GoTo TEXT3NONUM:
' TEXT1
If Len(Text4.Text) > 249 Then GoTo TEXT4BIGLEN:

'*************************** пяоцяаллатислос *****************************
If MsgBox("хекете ма пяовыягсете стгм еццяажг тгс пяосжояас", vbOKCancel, "") = vbOK Then
    If D = 0 Then
        
    Else
       MsgBox ("циа тгм етаияиа : " & Text5.Text & "  упаявеи гдг пяосжояа циа то пяоиом :  " & Text1.Text & " паяайакы екецнте"), vbCritical, "пяосовг !!!"
       GoTo TELOS:
    End If
    STATEMENT = "INSERT INTO " & Text5.Text & " (" & _
    "омола_пяоиомтос,тилг,ейптысг,сглеиысг) VALUES (" & _
    "'" & UCase(Text1.Text) & "'," & _
    Text2.Text & "," & _
    Text3.Text & "," & _
    "'" & UCase(Text4.Text) & "'" & _
    ")"
    dbp.Execute STATEMENT
    'RSH.Fields.Refresh
    'RSH.Close
    If dbp.STATE = 1 Then dbp.Close
    Adodc1.Refresh
    Text6.Text = Adodc1.Recordset.RecordCount
    If CInt(Text6.Text) < 24 Then
            DataGrid1.Height = 355 + (302.5 * CInt(Text6.Text))
    Else
            DataGrid1.Height = 7565
    End If
    '************************************ PERASMA DEDOMENON SE PINAKA пяоиомтос******
    '*******************************************************************************
    Dim count As Integer
    count = 0
    Dim SSS, SSS1, SSS2 As String
    If dbp.STATE = 1 Then dbp.Close
    If rsp.STATE = 1 Then dbp.Close

    dbp.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
    "Data Source=" & "\databases\PROION.mdb" & ";" & _
    "Persist Security Info=False"
    dbp.Open App.Path & "\databases\PROION.mdb"
    rsp.Open "[PROIONTA_ABC]", dbp, adOpenDynamic, adLockBatchOptimistic
    ' ANIXNEYSH AN YPARXEI пяоиом KAI AN OXI DHMIOYRGIA
    If rsp.EOF = rsp.BOF Then GoTo HH:
    rsp.MoveFirst
HH:
    Do While Not rsp.EOF
        If Text1.Text = rsp![пяоиомта] Then
            count = count + 1
            rsp.MoveNext
        Else
            rsp.MoveNext
        End If
    Loop
    If count = 0 Then  ' AN OXI
        SSS = " create table " & Text1.Text & " (" & _
            "омола_етаияиас varchar(50)," & _
            "тилг FLOAT," & _
            "ейптысг FLOAT," & _
            "сглеиысг varchar(250)," & _
            "PRIMARY KEY (омола_етаияиас) )"
        dbp.Execute SSS
    
        SSS2 = "INSERT INTO PROIONTA_ABC (" & _
        "пяоиомта) VALUES (" & _
        "'" & UCase(Text1.Text) & "')"

        dbp.Execute SSS2
    End If

    SSS1 = "INSERT INTO " & Text1.Text & " (" & _
        "омола_етаияиас,тилг,ейптысг,сглеиысг) VALUES (" & _
        "'" & UCase(Text5.Text) & "'," & _
        Text2.Text & "," & _
        Text3.Text & "," & _
        "'" & UCase(Text4.Text) & "'" & _
        ")"
    dbp.Execute SSS1

'**********************************************************************************
'**********************************************************************************
End If
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
GoTo TELOS:
'*********** ANTIMETOPISI ELEGXON **************************************
TEXT1KENO:
    MsgBox ("то педио омола пяоиомтос еимаи йемо"), vbCritical, "пяосовг !!!"
    Text1.SetFocus
    GoTo TELOS:
TEXT1BIGLEN:
    MsgBox ("то омола пяоиомтос дем пяепеи ма еимаи лецакутеяо апо 50 ваяайтгяес"), vbCritical, "пяосовг !!!"
    Text1.SetFocus
    GoTo TELOS:
TEXT1WCHAR:
    MsgBox ("то омола пяоиомтос дем пяепеи ма пеяиевеи тоус ваяайтгяес [ ] . !"), vbCritical, "пяосовг !!!"
    Text1.SetFocus
    GoTo TELOS:
TEXT2KENO:
    MsgBox ("то педио тилг еимаи йемо"), vbCritical, "пяосовг !!!"
    Text2.SetFocus
    GoTo TELOS:
TEXT2NONUM:
    MsgBox ("дем дысате сыста тгм тилг тоу пяоиомтос"), vbCritical, "пяосовг !!!"
    Text2.SetFocus
    GoTo TELOS:
TEXT3NONUM:
    MsgBox ("дем дысате сыста то пососто тгс ейптысгс"), vbCritical, "пяосовг !!!"
    Text3.SetFocus
    GoTo TELOS:
TEXT4BIGLEN:
    MsgBox ("г сглеиысг дем пяепеи ма еимаи лецакутеяг апо 250 ваяайтгяес"), vbCritical, "пяосовг !!!"
    Text4.SetFocus
    GoTo TELOS:

ER:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"


TELOS:
If dbp.STATE = 1 Then dbp.Close
If rsp.STATE = 1 Then dbp.Close
End Sub

Private Sub Command2_Click()
On Error GoTo ER:
Dim dbp As New ADODB.Connection
Dim rsp As New ADODB.Recordset
Dim rrsp As New ADODB.Recordset
Dim D As Integer
D = 0
Dim STATE1, STATE2, STATE3, STATE4 As String
Text1.Text = UCase(Trim(Text1.Text))
Text4.Text = UCase(Trim(Text4.Text))
If dbp.STATE = 1 Then dbp.Close
dbp.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & "\databases\PROION_ETAIRIA.mdb" & ";" & _
"Persist Security Info=False"
dbp.Open App.Path & "\databases\PROION_ETAIRIA.mdb"
rsp.Open "[" & Text5.Text & "]", dbp, adOpenDynamic, adLockBatchOptimistic

' ********** GENIKOI ELEGXOI *********
If Text1.Text = "" Then GoTo TEXT1KENO:

'**************************** DIAXORISMOS PEIPTOSEON ********************************

If Command2.Caption = "еуяесг" Then  '*********** AN EYRESH **********
    ' AN YPARXEI H EGRAFH
    If rsp.EOF = rsp.BOF Then GoTo NNIK:
    rsp.MoveFirst
NNIK:
    Do While Not rsp.EOF
        If rsp![омола_пяоиомтос] = Text1.Text Then
            D = D + 1
            Text7.Text = rsp![омола_пяоиомтос]
            Text1.Text = rsp![омола_пяоиомтос]
            Text2.Text = rsp![тилг]
            Text3.Text = rsp![ейптысг]
            Text4.Text = rsp![сглеиысг]
            rsp.MoveNext
        Else
            rsp.MoveNext
        End If
    Loop
    If D = 0 Then
        GoTo norecord:
    Else
        Command2.Caption = "диояхысг"
    End If
    

Else        ' ******************************* AN DIORTHOSI ********************

    ' TEXT1
    If Text1.Text = "" Then GoTo TEXT1KENO:
    If Len(Text1.Text) > 49 Then GoTo TEXT1BIGLEN:
    For i = 1 To Len(Text1.Text)
        If Mid(Text1.Text, i, i) = "." Or Mid(Text1.Text, i, i) = "!" Or _
        Mid(Text1.Text, i, i) = "[" Or Mid(Text1.Text, i, i) = "]" Then GoTo TEXT1WCHAR:
    Next i
    ' TEXT2
    If Text2.Text = "" Then GoTo TEXT2KENO:
    If IsNumeric(Text2.Text) = False Then GoTo TEXT2NONUM:
    ' TEXT3
    If Text3.Text = "" Then Text3.Text = 0
    If IsNumeric(Text3.Text) = False Then GoTo TEXT3NONUM:
    ' TEXT1
    If Len(Text4.Text) > 249 Then GoTo TEXT4BIGLEN:
    ' екецвос ам то омола поу паы ма дысы упаявеи гдг
    Dim ERT As Integer
    ERT = 0
    If Text1.Text = Text7.Text Then ' AN TO ONOMA PROIONTOS DEN ALLAZEI
    
    Else ' AN TO ONOMA PROIONTOS ALLLAZEI
        If dbp.STATE = 1 Then dbp.Close
        dbp.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
        "Data Source=" & "\databases\PROION_ETAIRIA.mdb" & ";" & _
        "Persist Security Info=False"
        dbp.Open App.Path & "\databases\PROION_ETAIRIA.mdb"
        rsp.Open "[" & Text5.Text & "]", dbp, adOpenDynamic, adLockBatchOptimistic
        If rsp.EOF = rsp.BOF Then GoTo NIKKK:
        rsp.MoveFirst
NIKKK:
        Do While Not rsp.EOF
            If rsp![омола_пяоиомтос] = Text1.Text Then
                ERT = ERT + 1
                rsp.MoveNext
            Else
                rsp.MoveNext
            End If
        Loop
    End If
    If ERT <> 0 Then GoTo YPAR_ETAIR:
    
    '*************************** пяоцяаллатислос  диояхысгс *****************************
    If MsgBox("хекете ма пяовыягсете стгм диояхысг тгс еццяажгс", vbOKCancel, "") = vbOK Then
        STATE1 = " UPDATE " & Text5.Text & _
                " SET тилг=" & "'" & Text2.Text & "'" & _
                " WHERE омола_пяоиомтос=" & "'" & Text7.Text & "'"
        
        STATE2 = " UPDATE " & Text5.Text & _
                " SET ейптысг=" & "'" & Text3.Text & "'" & _
                " WHERE омола_пяоиомтос=" & "'" & Text7.Text & "'"
    
        STATE3 = " UPDATE " & Text5.Text & _
                " SET сглеиысг=" & "'" & Text4.Text & "'" & _
                " WHERE омола_пяоиомтос=" & "'" & Text7.Text & "'"
     
        STATE4 = " UPDATE " & Text5.Text & _
                " SET омола_пяоиомтос=" & "'" & Text1.Text & "'" & _
                " WHERE омола_пяоиомтос=" & "'" & Text7.Text & "'"
        dbp.Execute STATE1
        dbp.Execute STATE2
        dbp.Execute STATE3
        dbp.Execute STATE4
        
' ********************** DIORTHOSI ONOMATOS ETAIRIAS ***********************************
'***************************************************************************************
        If dbp.STATE = 1 Then dbp.Close
        If rsp.STATE = 1 Then dbp.Close
        dbp.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
        "Data Source=" & "\databases\PROION.mdb" & ";" & _
        "Persist Security Info=False"
        dbp.Open App.Path & "\databases\PROION.mdb"
        rsp.Open "[PROIONTA_ABC]", dbp, adOpenDynamic, adLockBatchOptimistic

        Dim CCOUNT As Integer
        CCOUNT = 0
        Dim SSS1, SSS2, SSS3, SSS4, SSS, DS, US As String
        ' AN DEN EXEI ALLAXTEI TO ONOMA PROIONTOS
        If Text1.Text = Text7.Text Then
                SSS1 = " UPDATE " & Text7.Text & _
                " SET тилг=" & "'" & Text2.Text & "'" & _
                " WHERE омола_етаияиас=" & "'" & Text5.Text & "'"
        
                SSS2 = " UPDATE " & Text7.Text & _
                " SET ейптысг=" & "'" & Text3.Text & "'" & _
                " WHERE омола_етаияиас=" & "'" & Text5.Text & "'"
    
                SSS3 = " UPDATE " & Text7.Text & _
                " SET сглеиысг=" & "'" & Text4.Text & "'" & _
                " WHERE омола_етаияиас=" & "'" & Text5.Text & "'"
                dbp.Execute SSS1
                dbp.Execute SSS2
                dbp.Execute SSS3
        Else ' EXEI ALLAXTHEI KAI TO ONOMA ETAIRIAS
            If rsp.EOF = rsp.BOF Then GoTo EE:
            rsp.MoveFirst
EE:
            Do While Not rsp.EOF
                If rsp![пяоиомта] = Text1.Text Then
                    CCOUNT = CCOUNT + 1
                    rsp.MoveNext
                Else
                    rsp.MoveNext
                End If
            Loop
            If CCOUNT = 0 Then 'TO NEO ONOMA DEN YPARXEI HDH
                ' DHMIOYRGIA NEOY PINAKA
                SSS = " select * into " & Text1.Text & " from " & Text7.Text
                ' UPDATE STA YPOLOIPA PEDIA
                SSS1 = " UPDATE " & Text1.Text & _
                " SET тилг=" & "'" & Text2.Text & "'" & _
                " WHERE омола_етаияиас=" & "'" & Text5.Text & "'"
        
                SSS2 = " UPDATE " & Text1.Text & _
                " SET ейптысг=" & "'" & Text3.Text & "'" & _
                " WHERE омола_етаияиас=" & "'" & Text5.Text & "'"
    
                SSS3 = " UPDATE " & Text1.Text & _
                " SET сглеиысг=" & "'" & Text4.Text & "'" & _
                " WHERE омола_етаияиас=" & "'" & Text5.Text & "'"
                
                dbp.Execute SSS
                dbp.Execute SSS1
                dbp.Execute SSS2
                dbp.Execute SSS3
                
                ' THA PREPEI STON ARXIKO PINAKA(TEXT7)(PINAKAS PRIN THN DIORTHOSI)
                ' NA SBHSO TO ONOMA TOY PROIONTOS POY MOLIS EKANA DIORTHOSI
                ' AFHNONTAS ANEPAFA TA ALLA ONOMATA. EPISIS AN AYTO HTAN TO
                ' TELEYTAIO PROION NA SBHSO TON PINAKA ETAIRIAS.
                DS = "DELETE FROM " & Text7.Text & " WHERE омола_етаияиас='" & _
                 Text5.Text & "'"
                dbp.Execute DS
                
                'AN O PINAKAS POY APOMENEI EINAI ADEIOS SBHSTON
                 '   If dbp.STATE = 1 Then dbp.Close
                 '   Dim d_file, s11, s22 As String
                 '  d_file = App.Path & "\databases\PROION.mdb"
                 '   Adodc4.ConnectionString = _
                 '   "PROVIDER=Microsoft.Jet.OLEDB.4.0;" & _
                 '   "Data Source=" & d_file & ";"
                 '   Adodc4.RecordSource = _
                 '   "SELECT * FROM " & Text7.Text
                 '   Set DataGrid4.DataSource = Adodc4
                 '   Text8.Text = Adodc4.Recordset.RecordCount
                 '   Adodc4.Refresh
                 '   If Text8.Text = 0 Then
                 '       If dbp.STATE = 0 Then dbp.Open
                 '       s11 = "DROP TABLE " & Text7.Text
                 '       s22 = " DELETE FROM PROIONTA_ABC " & _
                 '               " where пяоиомта= '" & Text7.Text & "'"
                 '       dbp.Execute s11
                 '       dbp.Execute s22
                 '  End If
               
            
            ' THA PREPEI STON NEODHMIOYRGITHENTA PINAKA (TEXT1)ETAIRIAS NA
'PERIEXETAI MONO TO ONOMA TOY PROIONTOS APO TO OPOIO PROEKYPSE H DIORTHOSI
            ' ARA KAI ONEOS PINAKAS. ETSI STON NE PINAKA THA PREPEI NA SBHSTOYN
            'OLES OI EGRAFES EKTOS TOY ENOS PROIONTOS APO TO OPOIO PROEKIPSE
            ' H DHMIOYRGIA TOY PINAKA
                If dbp.STATE = 0 Then dbp.Open
                Dim r1 As New ADODB.Recordset
                Dim KK As String
                r1.Open "[" & Text1.Text & "]", dbp, adOpenDynamic, adLockBatchOptimistic
                If r1.BOF = r1.EOF Then GoTo NN1:
                r1.MoveFirst
NN1:
                Do While Not r1.EOF
                    If r1![омола_етаияиас] = Text5.Text Then
                        r1.MoveNext
                    Else
                        KK = "DELETE FROM " & Text1.Text & " WHERE " & _
                        " омола_етаияиас='" & r1![омола_етаияиас] & "'"
                        dbp.Execute KK
                        r1.MoveNext
                    End If
                Loop
                
                US = "INSERT INTO PROIONTA_ABC (" & _
                "пяоиомта) VALUES (" & _
                "'" & UCase(Text1.Text) & "')"
                dbp.Execute US
                If r1.STATE = 1 Then r1.Close
                    
                
                
                
            Else                      ' !!!YPARXEI HDH!!!!
                Dim ctate, ctate2, ctate3 As String
                Dim c2 As Integer
                c2 = 0
                rrsp.Open "[" & Text1.Text & "]", dbp, adOpenDynamic, adLockBatchOptimistic
                If rrsp.EOF = rrsp.BOF Then GoTo eee:
                rrsp.MoveFirst
eee:
                Do While Not rrsp.EOF
                    If rrsp![омола_етаияиас] = Text5.Text Then
                        c2 = c2 + 1
                        rrsp.MoveNext
                    Else
                        rrsp.MoveNext
                    End If
                Loop
                If c2 = 0 Then
                    ctate = "INSERT INTO " & Text1.Text & " (" & _
                    "омола_етаияиас,тилг,ейптысг,сглеиысг) VALUES (" & _
                    "'" & UCase(Text5.Text) & "'," & _
                    Text2.Text & "," & _
                    Text3.Text & "," & _
                    "'" & UCase(Text4.Text) & "'" & _
                    ")"
                    ctate2 = "delete from " & Text7.Text & " where омола_етаияиас='" & Text5.Text & "'"
                    dbp.Execute ctate
                    dbp.Execute ctate2
                        
                    'AN O PINAKAS POY APOMENEI EINAI ADEIOS SBHSTON
                        
                    'Dim d_file, s11, s22 As String
                    'If dbp.STATE = 1 Then dbp.Close
                    'd_file = App.Path & "\databases\PROION.mdb"
                    'Adodc3.ConnectionString = _
                    '"PROVIDER=Microsoft.Jet.OLEDB.4.0;" & _
                    '"Data Source=" & d_file & ";"
                    'Adodc3.RecordSource = _
                    '"SELECT * FROM " & Text7.Text
                    'Set DataGrid3.DataSource = Adodc3
                    'Text8.Text = Adodc3.Recordset.RecordCount
                    'Adodc3.Refresh
                    'If Text8.Text = 0 Then
                    '    If dbp.STATE = 0 Then dbp.Open
                    '    s11 = "DROP TABLE " & Text7.Text
                    '    s22 = " DELETE FROM PROIONTA_ABC " & _
                    '            " where пяоиомта= '" & Text7.Text & "'"
                    '   dbp.Execute s11
                    '    dbp.Execute s22
                    'End If
                Else
                    GoTo QWERTY:
                End If
            End If
        End If
        
'***************************************************************************************
'***************************************************************************************
        'RSH.Fields.Refresh
        'RSH.Close
        If dbp.STATE = 1 Then dbp.Close
        Adodc1.Refresh
        Text6.Text = Adodc1.Recordset.RecordCount
        If CInt(Text6.Text) < 24 Then
            DataGrid1.Height = 355 + (302.5 * CInt(Text6.Text))
        Else
            DataGrid1.Height = 7565
        End If
        Text1.Text = ""
        Text2.Text = ""
        Text3.Text = ""
        Text4.Text = ""
        Command2.Caption = "еуяесг"
    Else
        Text1.Text = ""
        Text2.Text = ""
        Text3.Text = ""
        Text4.Text = ""
        Command2.Caption = "еуяесг"
        GoTo TELOS:
    End If
End If
 GoTo TELOS:
    
'*********** ANTIMETOPISI ELEGXON **************************************
TEXT1KENO:
    MsgBox ("то педио етаияиа еимаи йемо"), vbCritical, "пяосовг !!!"
    Text1.SetFocus
    GoTo TELOS:
TEXT1BIGLEN:
    MsgBox ("то омола тоу пяоиомтос дем пяепеи ма еимаи лецакутеяо апо 50 ваяайтгяес"), vbCritical, "пяосовг !!!"
    Text1.SetFocus
    GoTo TELOS:
TEXT1WCHAR:
    MsgBox ("то омола тоу пяоиомтос дем пяепеи ма пеяиевеи тоус ваяайтгяес [ ] . !"), vbCritical, "пяосовг !!!"
    Text1.SetFocus
    GoTo TELOS:
TEXT2KENO:
    MsgBox ("то педио тилг еимаи йемо"), vbCritical, "пяосовг !!!"
    Text2.SetFocus
    GoTo TELOS:
TEXT2NONUM:
    MsgBox ("дем дысате сыста тгм тилг тоу пяоиомтос"), vbCritical, "пяосовг !!!"
    Text2.SetFocus
    GoTo TELOS:
TEXT3NONUM:
    MsgBox ("дем дысате сыста то пососто тгс ейптысгс"), vbCritical, "пяосовг !!!"
    Text3.SetFocus
    GoTo TELOS:
TEXT4BIGLEN:
    MsgBox ("г сглеиысг дем пяепеи ма еимаи лецакутеяг апо 250 ваяайтгяес"), vbCritical, "пяосовг !!!"
    Text4.SetFocus
    GoTo TELOS:
norecord:
 MsgBox ("дем бяехгйе еццяажг ле то омола  пяоиомтос поу дысате"), vbCritical, "пяосовг !!!"
    Text4.SetFocus
    GoTo TELOS:


YPAR_ETAIR:
MsgBox ("то омола тоу пяоиомтос поу пяоспахгте ма дысете упаявеи гдг"), vbCritical, "пяосовг !!!"
GoTo TELOS:


QWERTY:
MsgBox ("то мео омола пяоиомтос поу дысате упаявеи гдг йаи евеи йатавыяглемо то пяоиом : " & Text5.Text), vbCritical, "пяосовг !!!"
GoTo TELOS:

ER:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"


TELOS:
If dbp.STATE = 1 Then dbp.Close
If rsp.STATE = 1 Then dbp.Close
If rrsp.STATE = 1 Then dbp.Close
End Sub

Private Sub Command3_Click()
On Error GoTo ER:
Dim dbp As New ADODB.Connection
Dim rsp As New ADODB.Recordset
Dim D As Integer
D = 0
Dim STATEMENT As String
Text1.Text = UCase(Trim(Text1.Text))
Text4.Text = UCase(Trim(Text4.Text))
If dbp.STATE = 1 Then dbp.Close
dbp.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
       "Data Source=" & "\databases\PROION_ETAIRIA.mdb" & ";" & _
      "Persist Security Info=False"
dbp.Open App.Path & "\databases\PROION_ETAIRIA.mdb"
rsp.Open "[" & Text5.Text & "]", dbp, adOpenDynamic, adLockBatchOptimistic

' AN YPARXEI H EGRAFH
If rsp.EOF = rsp.BOF Then GoTo NNIK:
    rsp.MoveFirst
NNIK:
    Do While Not rsp.EOF
        If rsp![омола_пяоиомтос] = Text1.Text Then
            D = D + 1
            rsp.MoveNext
        Else
            rsp.MoveNext
        End If
    Loop

'************************  ELEGXOI  ************************************
' TEXT1
If Text1.Text = "" Then GoTo TEXT1KENO:

If D = 0 Then
    GoTo ERROR:
Else  '***********************  PROGRAMMATISMOS *******************************
   If MsgBox("хекете ма пяовыягсете стгм диацяажг тгс еццяажгс", vbOKCancel, "пяосовг") = vbOK Then
            STATEMENT = " delete from " & Text5.Text & _
                        " where омола_пяоиомтос= '" & Text1.Text & "'"
            dbp.Execute STATEMENT
            'RSH.Fields.Refresh
            'RSH.Close
            If dbp.STATE = 1 Then dbp.Close
            Adodc1.Refresh
            Text6.Text = Adodc1.Recordset.RecordCount
            If CInt(Text6.Text) < 24 Then
                DataGrid1.Height = 355 + (302.5 * CInt(Text6.Text))
            Else
                DataGrid1.Height = 7565
            End If
            If dbp.STATE = 0 Then dbp.Open
            If CInt(Text6.Text) = 0 Then
                Dim A1, A2
                A1 = "DROP TABLE " & Text5.Text
                A2 = " DELETE FROM ETAIRIES_ABC " & _
                     " where етаияиес= '" & Text5.Text & "'"
                dbp.Execute A1
                dbp.Execute A2
            End If
            If dbp.STATE = 1 Then dbp.Close
                
'******** DIAGRAFH EGRAFHS APO ETAIRIA KAI AN MONADIKH EGRAFH DIAGRAFH ETAIRIAS***
'*********************************************************************************
        If dbp.STATE = 1 Then dbp.Close
        If rsp.STATE = 1 Then dbp.Close
        Dim SSS, SSS1, SS2, d_file As String
        dbp.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
        "Data Source=" & "\databases\PROION.mdb" & ";" & _
        "Persist Security Info=False"
        dbp.Open App.Path & "\databases\PROION.mdb"
        
        SSS = " DELETE FROM " & Text1.Text & _
        " where омола_етаияиас= '" & Text5.Text & "'"
        dbp.Execute SSS
            
         'd_file = App.Path & "\databases\PROION.mdb"
         'Adodc2.ConnectionString = _
         '"PROVIDER=Microsoft.Jet.OLEDB.4.0;" & _
         '"Data Source=" & d_file & ";"
         'Adodc2.RecordSource = _
         '"SELECT * FROM " & Text1.Text
         'Set DataGrid2.DataSource = Adodc2
         'Text8.Text = Adodc2.Recordset.RecordCount
         'Adodc2.Refresh
         'If Text8.Text = 1 Then
         '   SSS1 = "DROP TABLE " & Text1.Text
         '   SSS2 = " DELETE FROM PROIONTA_ABC " & _
         '   " where пяоиомта= '" & Text1.Text & "'"
         '   dbp.Execute SSS1
         '   dbp.Execute SSS2
         'End If
'*********************************************************************************
'*********************************************************************************
        Text1.Text = ""
        Text2.Text = ""
        Text3.Text = ""
        Text4.Text = ""
    Else
        GoTo TELOS:
    End If
End If
GoTo TELOS:
' ************** ANTIMETOPISI ELEGXON ****************************
TEXT1KENO:
MsgBox ("дем дысате йамема омола етаияиас"), vbCritical, "пяосовг !!!"
GoTo TELOS:

ERROR:
MsgBox ("то омола етаияиас поу дысате дем упаявеи циа то суцйейяилемо пяоиом"), vbCritical, "пяосовг !!!"
GoTo TELOS:
 
ER:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"
  

TELOS:
If dbp.STATE = 1 Then dbp.Close
If rsp.STATE = 1 Then dbp.Close
End Sub

Private Sub Command4_Click()
On Error GoTo ER:
proionta3.Hide
Unload proionta3
proionta1.Adodc1.Refresh
proionta1.Adodc2.Refresh
proionta1.Text3.Text = proionta1.Adodc1.Recordset.RecordCount
If CInt(proionta1.Text3.Text) < 24 Then
    proionta1.DataGrid1.Height = 355 + (302.5 * CInt(proionta1.Text3.Text))
Else
    proionta1.DataGrid1.Height = 7565
End If

proionta1.Text4.Text = proionta1.Adodc2.Recordset.RecordCount
If CInt(proionta1.Text4.Text) < 21 Then
    proionta1.DataGrid2.Height = 355 + (302.5 * CInt(proionta1.Text4.Text))
Else
    proionta1.DataGrid2.Height = 7565
End If
GoTo TELOS:

ER:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"
  


TELOS:

End Sub

Private Sub DataGrid1_Click()
Text1.Text = DataGrid1.Text
End Sub

Private Sub Form_Load()
On Error GoTo ER:

Text5.Text = proionta1.Text2.Text
Label5.Caption = "етаияиа :" & Text5.Text

Dim DATABASE_FILE, DATABASE_FILE1 As String
DATABASE_FILE = App.Path & "\databases\PROION_ETAIRIA.mdb"

DataGrid1.DefColWidth = 2330
DataGrid1.Font.Size = 10
DataGrid1.HeadFont.Size = 12
DataGrid1.HeadFont.Bold = True


Adodc1.ConnectionString = _
"PROVIDER=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & DATABASE_FILE & ";"
Adodc1.RecordSource = _
"SELECT * FROM " & Text5.Text & " ORDER BY тилг "
Set DataGrid1.DataSource = Adodc1
Text6.Text = Adodc1.Recordset.RecordCount
If CInt(Text6.Text) < 24 Then
    DataGrid1.Height = 355 + (302.5 * CInt(Text6.Text))
Else
    DataGrid1.Height = 7565
End If
Adodc1.Refresh
GoTo TELOS:


ER:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"
  


TELOS:

End Sub

Private Sub Text10_Change()

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ER:
proionta3.Hide
Unload proionta3
proionta1.Adodc1.Refresh
proionta1.Adodc2.Refresh
proionta1.Text3.Text = proionta1.Adodc1.Recordset.RecordCount
If CInt(proionta1.Text3.Text) < 24 Then
    proionta1.DataGrid1.Height = 355 + (302.5 * CInt(proionta1.Text3.Text))
Else
    proionta1.DataGrid1.Height = 7565
End If

proionta1.Text4.Text = proionta1.Adodc2.Recordset.RecordCount
If CInt(proionta1.Text4.Text) < 21 Then
    proionta1.DataGrid2.Height = 355 + (302.5 * CInt(proionta1.Text4.Text))
Else
    proionta1.DataGrid2.Height = 7565
End If
GoTo TELOS:

ER:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"
  


TELOS:
End Sub

Private Sub Picture1_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Command2.Caption = "еуяесг"
End Sub

Private Sub Text2_LostFocus()
Dim dd As Integer
Dim S As String

dd = Len(Text2.Text)
S = Text2.Text
For i = 1 To dd
    If Mid(S, i, 1) = "," Then
        Mid(S, i, 1) = "."
    End If
Next i
Text2.Text = S
End Sub

Private Sub Text3_LostFocus()
Dim dd As Integer
Dim S As String

dd = Len(Text3.Text)
S = Text3.Text
For i = 1 To dd
    If Mid(S, i, 1) = "," Then
        Mid(S, i, 1) = "."
    End If
Next i
Text3.Text = S
End Sub

