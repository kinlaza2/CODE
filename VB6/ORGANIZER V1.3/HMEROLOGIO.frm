VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form HMEROLOGIO 
   BackColor       =   &H80000013&
   Caption         =   "глеяокоцио"
   ClientHeight    =   10485
   ClientLeft      =   105
   ClientTop       =   465
   ClientWidth     =   15180
   LinkTopic       =   "Form1"
   ScaleHeight     =   10485
   ScaleWidth      =   15180
   Begin VB.PictureBox Picture2 
      Height          =   375
      Left            =   120
      Picture         =   "HMEROLOGIO.frx":0000
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   25
      Top             =   9000
      Width           =   375
   End
   Begin VB.TextBox Text8 
      Height          =   375
      Left            =   5880
      TabIndex        =   24
      Text            =   "Text8"
      Top             =   1200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   5880
      TabIndex        =   22
      Text            =   "Text7"
      Top             =   720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   5040
      TabIndex        =   21
      Text            =   "Text6"
      Top             =   240
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00E9C5AD&
      Caption         =   "диацяажг"
      Height          =   615
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   9480
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00E9C5AD&
      Caption         =   "еуяесг"
      Height          =   615
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   9480
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E9C5AD&
      Caption         =   "еццяажг"
      Height          =   615
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   9480
      Width           =   1455
   End
   Begin VB.TextBox Text5 
      Height          =   1335
      Left            =   9240
      MultiLine       =   -1  'True
      TabIndex        =   17
      Top             =   8760
      Width           =   5775
   End
   Begin VB.TextBox Text4 
      Height          =   405
      Left            =   6840
      TabIndex        =   16
      Top             =   8760
      Width           =   2175
   End
   Begin VB.TextBox Text3 
      Height          =   405
      Left            =   4080
      TabIndex        =   13
      Top             =   8760
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1440
      TabIndex        =   11
      Top             =   8760
      Width           =   2175
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   14400
      Picture         =   "HMEROLOGIO.frx":0442
      ScaleHeight     =   465
      ScaleWidth      =   330
      TabIndex        =   9
      Top             =   120
      Width           =   360
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E9C5AD&
      Caption         =   "епистяожг сто аявийо лемоу"
      Height          =   735
      Left            =   120
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   9600
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   6622
      TabIndex        =   2
      Top             =   0
      Width           =   2056
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   6060
      Left            =   0
      TabIndex        =   1
      Top             =   2400
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   10689
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   615
      Left            =   7560
      Top             =   1800
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
   Begin MSComCtl2.MonthView MonthView1 
      Height          =   2370
      Left            =   10080
      TabIndex        =   0
      Top             =   0
      Width           =   2805
      _ExtentX        =   4948
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      MaxSelCount     =   31
      MultiSelect     =   -1  'True
      StartOfWeek     =   20512770
      TitleBackColor  =   -2147483635
      TitleForeColor  =   -2147483634
      CurrentDate     =   38377
   End
   Begin VB.Label Label10 
      BackColor       =   &H80000013&
      Caption         =   "(*)  то педио сглеиысг лпояеи ма евеи еыс 240 ваяайтгяес"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   9240
      TabIndex        =   23
      Top             =   10200
      Width           =   5775
   End
   Begin VB.Line Line2 
      Visible         =   0   'False
      X1              =   11475
      X2              =   11475
      Y1              =   0
      Y2              =   10425
   End
   Begin VB.Line Line1 
      Visible         =   0   'False
      X1              =   7650
      X2              =   7650
      Y1              =   15
      Y2              =   10440
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
      Caption         =   "сглеиысг"
      Height          =   375
      Left            =   9600
      TabIndex        =   15
      Top             =   8520
      Width           =   2175
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
      Caption         =   "посо"
      Height          =   375
      Left            =   6840
      TabIndex        =   14
      Top             =   8520
      Width           =   2175
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
      Caption         =   "етаияиа"
      Height          =   375
      Left            =   4080
      TabIndex        =   12
      Top             =   8520
      Width           =   2175
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
      Caption         =   "аяихлос епитацгс"
      Height          =   375
      Left            =   1440
      TabIndex        =   10
      Top             =   8520
      Width           =   2175
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000013&
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
      Left            =   2880
      TabIndex        =   7
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
      Caption         =   "0"
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
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000013&
      Caption         =   "сумокийо посо епитацым пеяиодоу :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   2895
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000013&
      Caption         =   "сумокийо посо глеяас : "
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
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000013&
      Caption         =   "аяихлос епитацым глеяас :"
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
      TabIndex        =   3
      Top             =   600
      Width           =   3735
   End
End
Attribute VB_Name = "HMEROLOGIO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error GoTo ER:

HMEROLOGIO.Hide
Unload HMEROLOGIO
GoTo TELOS:

ER:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"


TELOS:
End Sub

Private Sub Command2_Click()
On Error GoTo ER:

Text2.Text = UCase(Trim(Text2.Text))
Text3.Text = UCase(Trim(Text3.Text))
Text4.Text = UCase(Trim(Text4.Text))
Text5.Text = UCase(Trim(Text5.Text))

' elegxos mhkoys text2,text3
Dim ddd1, ddd2 As Integer
Dim t1, t2 As String

ddd1 = Len(Text1.Text)
ddd2 = Len(Text2.Text)

If ddd1 > 30 Then
    t1 = Mid(Text1.Text, 1, 30)
Else
    t1 = Text1.Text
End If

If ddd2 > 30 Then
    t2 = Mid(Text2.Text, 1, 30)
Else
    t2 = Text2.Text
End If

Text1.Text = t1
Text2.Text = t2
'*******************************************************

Dim DB2 As New ADODB.Connection
Dim RS2 As New ADODB.Recordset
Dim DATABASE_FILE, STATEMENT, ASD As String
Dim index, C As Integer
Dim SUM As Double
index = 2
C = 1
ASD = Text7.Text
SUM = 0

'***SYNDESI ME BASH HMEROLOGIO.MDB**************
DATABASE_FILE = App.Path & "\databases\HMEROLOGIO.mdb"
DB2.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
       "Data Source=" & "\databases\HMEROLOGIO.mdb" & ";" & _
      "Persist Security Info=False"
DB2.Open App.Path & "\databases\HMEROLOGIO.mdb"
RS2.Open "[" & ASD & "]", DB2, adOpenDynamic, adLockBatchOptimistic

'****ELEGXOI LATHON*****************************
If Text2.Text = "" Then GoTo ER1:
If Text3.Text = "" Then GoTo ER2:
If Text4.Text = "" Then GoTo ER3A:
If IsNumeric(Text4.Text) = False Then GoTo ER3B:


If RS2.EOF = RS2.BOF Then GoTo NIK:
RS2.MoveFirst
NIK:
Do While Not RS2.EOF
    If RS2![аяихлос_епитацгс] <> Text2.Text Then
        RS2.MoveNext
    Else
        C = C + 1
        RS2.MoveNext
    End If
Loop

If C <> 1 Then
    MsgBox ("евеи йатавыягхг гдг епитацг ле том аяихло поу дысате"), vbCritical, "пяосовг!!"
'*******PROGRAMMATISMOS*****************
Else
    If MsgBox("хекете ма пяовыягсете стгм еццяажг тгс епитацгс", vbOKCancel, "") = vbOK Then
        STATEMENT = " INSERT INTO " & ASD & _
        " (аяихлос_епитацгс,етаияиа,посо,сглеиысг) VALUES " & _
        "('" & Text2.Text & "'," & _
        "'" & Text3.Text & "'," & _
        Text4.Text & ",'" & Text5.Text & "')"
        DB2.Execute STATEMENT
        MsgBox ("г еццяажг окойкгяыхгйе"), , "ой"
        Text2.Text = ""
        Text3.Text = ""
        Text4.Text = ""
        Text5.Text = ""
    End If
    
End If
    If RS2.STATE = 1 Then RS2.Close
    RS2.Open "[" & ASD & "]", DB2, adOpenDynamic, adLockBatchOptimistic
    If (RS2.BOF = True) And (RS2.EOF = True) Then
        Label2.Caption = "сумокийо посо глеяас : 0"
    Else
        RS2.MoveFirst
        Do While Not RS2.EOF
            SUM = RS2![посо] + SUM
            Label5.Caption = SUM
            RS2.MoveNext
        Loop
    End If
    
    RS2.Fields.Refresh
    RS2.Close
    DB2.Close
    Adodc1.ConnectionString = _
    "PROVIDER=Microsoft.Jet.OLEDB.4.0;" & _
    "Data Source=" & DATABASE_FILE & ";"
    Adodc1.RecordSource = "SELECT * FROM " & ASD
    ' Bind the ADODC to the DataGrid.
    Set DataGrid1.DataSource = Adodc1
    DataGrid1.Refresh
    Adodc1.Refresh
    Label1.Caption = "аяихлос епитацым глеяас : " & Adodc1.Recordset.RecordCount
    Text8.Text = Adodc1.Recordset.RecordCount
    If CInt(Text8.Text) <= 19 Then
        DataGrid1.Height = 303 + (303 * CInt(Text8.Text))
    Else
        DataGrid1.Height = 6060
    End If
    GoTo TELOS:


'******ANTIMETOPISH LATHON***********************
ER1:
MsgBox ("дем дысате аяихло епитацгс"), vbCritical, "пяосовг!!"
index = 32
GoTo TELOS:

ER2:
If index = 32 Then
    GoTo TELOS:
Else
    MsgBox ("дем дысате то омола етаияиас"), vbCritical, "пяосовг!!"
    index = 32
    GoTo TELOS:
End If

ER3A:
If index = 32 Then
    GoTo TELOS:
Else
    MsgBox ("дем дысате то посо"), vbCritical, "пяосовг!!"
    index = 32
    GoTo TELOS:
End If
   
ER3B:
If index = 32 Then
    GoTo TELOS:
Else
    MsgBox ("дем дысате сыста то посо"), vbCritical, "пяосовг!!"
    index = 32
    GoTo TELOS:
End If

ER:
    If index = 32 Then
        GoTo TELOS:
    Else
        MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"
        GoTo TELOS:
    End If
    
TELOS:
    If RS2.STATE = 1 Then RS2.Close
    If DB2.STATE = 1 Then DB2.Close
End Sub

Private Sub Command3_Click()
On Error GoTo ER:

Text2.Text = UCase(Trim(Text2.Text))
Text3.Text = UCase(Trim(Text3.Text))
Text4.Text = UCase(Trim(Text4.Text))
Text5.Text = UCase(Trim(Text5.Text))
Dim DB2 As New ADODB.Connection
Dim RS2 As New ADODB.Recordset
Dim DATABASE_FILE, STATE1, STATE2, STATE3, STATE4, temp As String
Dim index, C As Integer
Dim SUM As Double
Dim ASD As String
index = 2
C = 1
SUM = 0
ASD = Text7.Text
'***SYNDESI ME BASH HMEROLOGIO.MDB**************
DATABASE_FILE = App.Path & "\databases\HMEROLOGIO.mdb"
DB2.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
       "Data Source=" & "\databases\HMEROLOGIO.mdb" & ";" & _
      "Persist Security Info=False"
DB2.Open App.Path & "\databases\HMEROLOGIO.mdb"
RS2.Open "[" & ASD & "]", DB2, adOpenDynamic, adLockBatchOptimistic

If Command3.Caption = "еуяесг" Then
'***амивмеусг ам упаявеи г епитацг**************
    If RS2.EOF = RS2.BOF Then GoTo NIK:
    RS2.MoveFirst
NIK:
    Do While Not RS2.EOF
        If RS2![аяихлос_епитацгс] <> Text2.Text Then
            RS2.MoveNext
        Else
            C = C + 1
            Text2.Text = RS2![аяихлос_епитацгс]
            Text3.Text = RS2![етаияиа]
            Text4.Text = RS2![посо]
            Text5.Text = RS2![сглеиысг]
            Text6.Text = RS2![аяихлос_епитацгс]
            RS2.MoveNext
        End If
    Loop
'******* ELEGXOS AN BRETHIKE EPITAGH*****************
    If C = 1 Then
        MsgBox ("дем бяехгйе епитацг ле то моулеяо поу дысате"), vbCritical, "пяосовг"
    Else
        Command3.Caption = "диояхысг"
    End If
Else ' *** AN KOYMPI DIORTHOSI**********
    Text2.Text = UCase(Trim(Text2.Text))
    Text3.Text = UCase(Trim(Text3.Text))
    Text4.Text = UCase(Trim(Text4.Text))
    Text5.Text = UCase(Trim(Text5.Text))

    ' elegxos mhkoys text2,text3
    Dim ddd1, ddd2 As Integer
    Dim t1, t2 As String

    ddd1 = Len(Text1.Text)
    ddd2 = Len(Text2.Text)

    If ddd1 > 30 Then
        t1 = Mid(Text1.Text, 1, 30)
    Else
        t1 = Text1.Text
    End If

    If ddd2 > 30 Then
        t2 = Mid(Text2.Text, 1, 30)
    Else
        t2 = Text2.Text
    End If

    Text1.Text = t1
    Text2.Text = t2
    '*******************************************************
    If MsgBox("хекете ма пяовыягсете стгм диояхысг тгс епитацгс", vbOKCancel, "") = vbOK Then
        Text2.Text = UCase(Trim(Text2.Text))
        Text3.Text = UCase(Trim(Text3.Text))
        Text4.Text = UCase(Trim(Text4.Text))
        Text5.Text = UCase(Trim(Text5.Text))
        '****ELEGXOI LATHON*****************************
        If Text2.Text = "" Then GoTo ER1:
        If Text3.Text = "" Then GoTo ER2:
        If Text4.Text = "" Then GoTo ER3A:
        If IsNumeric(Text4.Text) = False Then GoTo ER3B:
        
        STATE1 = " UPDATE " & ASD & _
        " SET сглеиысг='" & Text5.Text & "'" & _
        " WHERE аяихлос_епитацгс='" & Text6.Text & "'"

        STATE2 = " UPDATE " & ASD & _
        " SET посо='" & Text4.Text & "'" & _
        " WHERE аяихлос_епитацгс='" & Text6.Text & "'"
    
        STATE3 = " UPDATE " & ASD & _
        " SET етаияиа='" & Text3.Text & "'" & _
        " WHERE аяихлос_епитацгс='" & Text6.Text & "'"
        
        STATE4 = " UPDATE " & ASD & _
        " SET аяихлос_епитацгс='" & Text2.Text & "'" & _
        " WHERE аяихлос_епитацгс='" & Text6.Text & "'"
        
        DB2.Execute STATE1
        DB2.Execute STATE2
        DB2.Execute STATE3
        DB2.Execute STATE4
        MsgBox ("г диояхысг окойкгяыхгйе"), , "ой"
        Text2.Text = ""
        Text3.Text = ""
        Text4.Text = ""
        Text5.Text = ""
    Else ' *** AN PATITHI CANCEL********
        Text2.Text = ""
        Text3.Text = ""
        Text4.Text = ""
        Text5.Text = ""
    End If
    Command3.Caption = "еуяесг"
    Adodc1.ConnectionString = _
    "PROVIDER=Microsoft.Jet.OLEDB.4.0;" & _
    "Data Source=" & DATABASE_FILE & ";"
    Adodc1.RecordSource = _
    "SELECT * FROM " & ASD
    ' Bind the ADODC to the DataGrid.
    Set DataGrid1.DataSource = Adodc1
    Adodc1.Refresh
    
    Label1.Caption = "аяихлос епитацым глеяас : " & Adodc1.Recordset.RecordCount
    Text8.Text = Adodc1.Recordset.RecordCount
    
    If CInt(Text8.Text) <= 19 Then
        DataGrid1.Height = 303 + (303 * CInt(Text8.Text))
    Else
        DataGrid1.Height = 6060
    End If
    
    If RS2.STATE = 1 Then RS2.Close
    RS2.Open "[" & ASD & "]", DB2, adOpenDynamic, adLockBatchOptimistic
    If (RS2.BOF = True) And (RS2.EOF = True) Then
        Label2.Caption = "сумокийо посо глеяас : 0"
    Else
        RS2.MoveFirst
        Do While Not RS2.EOF
            SUM = RS2![посо] + SUM
            Label5.Caption = SUM
            RS2.MoveNext
        Loop
    End If
    
End If
GoTo TELOS:


'******ANTIMETOPISH LATHON***********************
ER1:
MsgBox ("дем дысате аяихло епитацгс"), vbCritical, "пяосовг!!"
index = 32
GoTo TELOS:

ER2:
If index = 32 Then
    GoTo TELOS:
Else
    MsgBox ("дем дысате то омола етаияиас"), vbCritical, "пяосовг!!"
    index = 32
End If

ER3A:
If index = 32 Then
    GoTo TELOS:
Else
    MsgBox ("дем дысате то посо"), vbCritical, "пяосовг!!"
    index = 32
End If
   
ER3B:
If index = 32 Then
    GoTo TELOS:
Else
    MsgBox ("дем дысате сыста то посо"), vbCritical, "пяосовг!!"
    index = 32
End If

ER:
    If index = 32 Then
        GoTo TELOS:
    Else
        MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"
    End If
    
    
TELOS:
    If RS2.STATE = 1 Then RS2.Close
    If DB2.STATE = 1 Then DB2.Close
End Sub

Private Sub Command4_Click()
On Error GoTo ER:

Text2.Text = UCase(Trim(Text2.Text))

Dim DB2 As New ADODB.Connection
Dim RS2 As New ADODB.Recordset
Dim DATABASE_FILE, STATEMENT, ASD As String
Dim C As Integer
Dim SUM As Double
C = 1
SUM = 0
ASD = Text7.Text
'***SYNDESI ME BASH HMEROLOGIO.MDB**************
DATABASE_FILE = App.Path & "\databases\HMEROLOGIO.mdb"
DB2.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
       "Data Source=" & "\databases\HMEROLOGIO.mdb" & ";" & _
      "Persist Security Info=False"
DB2.Open App.Path & "\databases\HMEROLOGIO.mdb"
RS2.Open "[" & ASD & "]", DB2, adOpenDynamic, adLockBatchOptimistic
'*****ELEGXOS***********************************
If Text2.Text = "" Then GoTo KENO:
'****ELEGXOS AN YPARXEI EGRAFH*****************************
If RS2.EOF = RS2.BOF Then GoTo NIK:
RS2.MoveFirst
NIK:
Do While Not RS2.EOF
    If RS2![аяихлос_епитацгс] <> Text2.Text Then
        RS2.MoveNext
    Else
        C = C + 1
        RS2.MoveNext
    End If
Loop

If C = 1 Then
    MsgBox ("дем евеи йатавыягхг епитацг ле том аяихло поу дысате"), vbCritical, "пяосовг!!"
'*******PROGRAMMATISMOS*****************
Else
  If MsgBox("хекете ма пяовыягсете стгм диацяажг тгс епитацгс", vbOKCancel, "") = vbOK Then
        STATEMENT = "DELETE FROM  " & ASD & _
        " WHERE аяихлос_епитацгс ='" & Text2.Text & "'"
        DB2.Execute STATEMENT
        MsgBox ("г диацяажг окойкгяыхгйе"), , "ой"
        Text2.Text = ""
        Text3.Text = ""
        Text4.Text = ""
        Text5.Text = ""
    End If
    If RS2.STATE = 1 Then RS2.Close
    RS2.Open "[" & ASD & "]", DB2, adOpenDynamic, adLockBatchOptimistic
    If (RS2.BOF = True) And (RS2.EOF = True) Then
        Label2.Caption = "сумокийо посо глеяас : 0"
    Else
        RS2.MoveFirst
        Do While Not RS2.EOF
            SUM = RS2![посо] + SUM
            Label5.Caption = SUM
            RS2.MoveNext
        Loop
    End If
End If
    RS2.Fields.Refresh
    RS2.Close
    DB2.Close
    Adodc1.ConnectionString = _
    "PROVIDER=Microsoft.Jet.OLEDB.4.0;" & _
    "Data Source=" & DATABASE_FILE & ";"
    Adodc1.RecordSource = _
    "SELECT * FROM " & ASD
    ' Bind the ADODC to the DataGrid.
    Set DataGrid1.DataSource = Adodc1
    Adodc1.Refresh
    
    Label1.Caption = "аяихлос епитацым глеяас : " & Adodc1.Recordset.RecordCount
    Text8.Text = Adodc1.Recordset.RecordCount
    
    If CInt(Text8.Text) <= 19 Then
        DataGrid1.Height = 303 + (303 * CInt(Text8.Text))
    Else
        DataGrid1.Height = 6060
    End If
GoTo TELOS:

KENO:
MsgBox ("дем дысате том аяихло епитацгс поу хекете ма диацяаьете"), vbCritical, "пяосовг!!"
GoTo TELOS:

ER:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
    If RS2.STATE = 1 Then RS2.Close
    If DB2.STATE = 1 Then DB2.Close
End Sub

Private Sub DataGrid1_Click()
        Text2.Text = DataGrid1.Columns(0).Text
        Text5.Text = DataGrid1.Columns(3).Text
End Sub

Private Sub Form_Load()
On Error GoTo ER:
MonthView1.Value = Date
Text1.Text = MonthView1.Value
DataGrid1.HeadFont.Size = 12
DataGrid1.Font.Bold = True
DataGrid1.Font.Size = 10
DataGrid1.DefColWidth = 3585

Dim STATEMENT, STATEMENT2, ST1 As String
Dim SUM As Double
SUM = 0
Dim D, m, y
Dim DB2 As New ADODB.Connection
Dim RS2 As New ADODB.Recordset
Dim RS2A As New ADODB.Recordset
Dim DATABASE_FILE, ASD As String
Dim C As Integer
C = 0

Command2.Enabled = True
Command3.Enabled = True
Command4.Enabled = True
' APODOSH TIMON
D = Day(MonthView1.Value)
m = Month(MonthView1.Value)
y = Year(MonthView1.Value)
ASD = D & m & y
Text7.Text = ASD
DATABASE_FILE = App.Path & "\databases\HMEROLOGIO.mdb"
DB2.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
       "Data Source=" & "\databases\HMEROLOGIO.mdb" & ";" & _
      "Persist Security Info=False"
DB2.Open App.Path & "\databases\HMEROLOGIO.mdb"
RS2A.Open "[ONOMATA_PINAKON]", DB2, adOpenDynamic, adLockBatchOptimistic

 
If RS2A.BOF = RS2A.EOF Then GoTo NIK:
RS2A.MoveFirst
NIK:
Do While Not RS2A.EOF
    If RS2A![ONOMATA_PINAKON] = ASD Then
        C = C + 1
        RS2A.MoveNext
    Else
        RS2A.MoveNext
    End If
Loop
If RS2A.STATE = 1 Then RS2A.Close

If C = 0 Then
    ' EKTELESH AN DEN YPARXEI PINAKAS
    STATEMENT = " create table " & ASD & _
    " ( аяихлос_епитацгс VARCHAR(30), " & _
    " етаияиа VARCHAR(30), " & _
    " посо DOUBLE, " & _
    " сглеиысг VARCHAR(245) )"
    DB2.Execute STATEMENT
    
    ST1 = " INSERT INTO ONOMATA_PINAKON (ONOMATA_PINAKON)" & _
    " VALUES('" & ASD & "')"
    DB2.Execute ST1
    
    Adodc1.ConnectionString = _
    "PROVIDER=Microsoft.Jet.OLEDB.4.0;" & _
    "Data Source=" & DATABASE_FILE & ";"
    Adodc1.RecordSource = _
    " SELECT * FROM " & ASD
    ' Bind the ADODC to the DataGrid.
    Set DataGrid1.DataSource = Adodc1
    Adodc1.Refresh
    DataGrid1.Refresh
    
    Text1.Text = MonthView1.Value
    Label1.Caption = "аяихлос епитацым глеяас : " & Adodc1.Recordset.RecordCount
    Label2.Caption = "сумокийо посо глеяас : 0"
    Text8.Text = Adodc1.Recordset.RecordCount
    DB2.Close
Else
    RS2.Open "[" & ASD & "]", DB2, adOpenDynamic, adLockBatchOptimistic
    'EKTELESH AN YPARXEI PINAKAS
    Adodc1.ConnectionString = _
    "PROVIDER=Microsoft.Jet.OLEDB.4.0;" & _
    "Data Source=" & DATABASE_FILE & ";"
    Adodc1.RecordSource = _
    "SELECT * FROM " & ASD
    ' Bind the ADODC to the DataGrid.
    Set DataGrid1.DataSource = Adodc1
    Adodc1.Refresh
    DataGrid1.Refresh
    
    Text1.Text = MonthView1.Value
    Label1.Caption = "аяихлос епитацым глеяас : " & Adodc1.Recordset.RecordCount
    Text8.Text = Adodc1.Recordset.RecordCount
   
    If (RS2.BOF = True) And (RS2.EOF = True) Then
        Label2.Caption = "сумокийо посо глеяас : 0"
    Else
        RS2.MoveFirst
        Do While Not RS2.EOF
            SUM = RS2![посо] + SUM
            Label5.Caption = SUM
            RS2.MoveNext
        Loop
    End If
End If
If CInt(Text8.Text) <= 19 Then
    DataGrid1.Height = 303 + (303 * CInt(Text8.Text))
Else
    DataGrid1.Height = 6060
End If
GoTo TELOS:

ER:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
If DB2.STATE = 1 Then DB2.Close
If RS2.STATE = 1 Then RS2.Close
If RS2A.STATE = 1 Then RS2A.Close

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ER:

HMEROLOGIO.Hide
Unload HMEROLOGIO
GoTo TELOS:

ER:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"


TELOS:
End Sub

Private Sub MonthView1_DateClick(ByVal DateClicked As Date)
On Error GoTo ER:


Dim STATEMENT, STATEMENT2, ST1 As String
Dim SUM As Double
SUM = 0
Dim D, m, y
Dim DB2 As New ADODB.Connection
Dim RS2 As New ADODB.Recordset
Dim RS2A As New ADODB.Recordset
Dim DATABASE_FILE, ASD As String
Dim C As Integer
C = 0

Command2.Enabled = True
Command3.Enabled = True
Command4.Enabled = True
' APODOSH TIMON
D = Day(MonthView1.Value)
m = Month(MonthView1.Value)
y = Year(MonthView1.Value)
ASD = D & m & y
Text7.Text = ASD
DATABASE_FILE = App.Path & "\databases\HMEROLOGIO.mdb"
DB2.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
       "Data Source=" & "\databases\HMEROLOGIO.mdb" & ";" & _
      "Persist Security Info=False"
DB2.Open App.Path & "\databases\HMEROLOGIO.mdb"
RS2A.Open "[ONOMATA_PINAKON]", DB2, adOpenDynamic, adLockBatchOptimistic

 
If RS2A.BOF = RS2A.EOF Then GoTo NIK:
RS2A.MoveFirst
NIK:
Do While Not RS2A.EOF
    If RS2A![ONOMATA_PINAKON] = ASD Then
        C = C + 1
        RS2A.MoveNext
    Else
        RS2A.MoveNext
    End If
Loop
If RS2A.STATE = 1 Then RS2A.Close

If C = 0 Then
    ' EKTELESH AN DEN YPARXEI PINAKAS
    STATEMENT = " create table " & ASD & _
    " ( аяихлос_епитацгс VARCHAR(30), " & _
    " етаияиа VARCHAR(30), " & _
    " посо DOUBLE, " & _
    " сглеиысг VARCHAR(245) )"
    DB2.Execute STATEMENT
    
    ST1 = " INSERT INTO ONOMATA_PINAKON (ONOMATA_PINAKON)" & _
    " VALUES('" & ASD & "')"
    DB2.Execute ST1
    
    Adodc1.ConnectionString = _
    "PROVIDER=Microsoft.Jet.OLEDB.4.0;" & _
    "Data Source=" & DATABASE_FILE & ";"
    Adodc1.RecordSource = _
    " SELECT * FROM " & ASD
    ' Bind the ADODC to the DataGrid.
    Set DataGrid1.DataSource = Adodc1
    Adodc1.Refresh
    DataGrid1.Refresh
    
    Text1.Text = MonthView1.Value
    Label1.Caption = "аяихлос епитацым глеяас : " & Adodc1.Recordset.RecordCount
    Label2.Caption = "сумокийо посо глеяас : 0"
    Text8.Text = Adodc1.Recordset.RecordCount
    DB2.Close
Else
    RS2.Open "[" & ASD & "]", DB2, adOpenDynamic, adLockBatchOptimistic
    'EKTELESH AN YPARXEI PINAKAS
    Adodc1.ConnectionString = _
    "PROVIDER=Microsoft.Jet.OLEDB.4.0;" & _
    "Data Source=" & DATABASE_FILE & ";"
    Adodc1.RecordSource = _
    "SELECT * FROM " & ASD
    ' Bind the ADODC to the DataGrid.
    Set DataGrid1.DataSource = Adodc1
    Adodc1.Refresh
    DataGrid1.Refresh
    
    Text1.Text = MonthView1.Value
    Label1.Caption = "аяихлос епитацым глеяас : " & Adodc1.Recordset.RecordCount
    Text8.Text = Adodc1.Recordset.RecordCount
   
    If (RS2.BOF = True) And (RS2.EOF = True) Then
        Label2.Caption = "сумокийо посо глеяас : 0"
    Else
        RS2.MoveFirst
        Do While Not RS2.EOF
            SUM = RS2![посо] + SUM
            Label5.Caption = SUM
            RS2.MoveNext
        Loop
    End If
End If
If CInt(Text8.Text) <= 19 Then
    DataGrid1.Height = 303 + (303 * CInt(Text8.Text))
Else
    DataGrid1.Height = 6060
End If
GoTo TELOS:

ER:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
If DB2.STATE = 1 Then DB2.Close
If RS2.STATE = 1 Then RS2.Close
If RS2A.STATE = 1 Then RS2A.Close
End Sub

Private Sub MonthView1_SelChange(ByVal StartDate As Date, ByVal EndDate As Date, Cancel As Boolean)
On Error GoTo ER:


Dim DB2 As New ADODB.Connection
Dim RS2 As New ADODB.Recordset
Dim RS2A As New ADODB.Recordset
Dim D, m, y, XX, D1, M1, Y1, XX1
Dim ASDD, ASD1, STATEMENT, ST1 As String
Dim SUM As Double
SUM = 0
Dim C As Integer
C = 0

If DB2.STATE = 1 Then DB2.Close
If RS2.STATE = 1 Then RS2.Close
DB2.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
       "Data Source=" & "\databases\HMEROLOGIO.mdb" & ";" & _
      "Persist Security Info=False"
DB2.Open App.Path & "\databases\HMEROLOGIO.mdb"

'** DIADIKASIA DHMIOYRGIAS AN DEN YPARXOYN EPILEGMENON PINAKON ************
If MonthView1.SelStart <> MonthView1.SelEnd Then
    XX1 = MonthView1.SelStart
    I = 0
    Do While XX1 <= MonthView1.SelEnd
        D1 = Day(MonthView1.Value + I)
        M1 = Month(MonthView1.Value)
        Y1 = Year(MonthView1.Value)
        ASD1 = D1 & M1 & Y1
        RS2A.Open "[ONOMATA_PINAKON]", DB2, adOpenDynamic, adLockBatchOptimistic
        If RS2A.BOF = RS2A.EOF Then GoTo NIK1:
        RS2A.MoveFirst
NIK1:
        Do While Not RS2A.EOF
            If RS2A![ONOMATA_PINAKON] = ASD1 Then
                C = C + 1
                RS2A.MoveNext
            Else
                RS2A.MoveNext
            End If
        Loop
        If C = 0 Then
            STATEMENT = " create table " & ASD1 & _
            " ( аяихлос_епитацгс VARCHAR(30), " & _
            " етаияиа VARCHAR(30), " & _
            " посо DOUBLE, " & _
            " сглеиысг VARCHAR(245) )"
            DB2.Execute STATEMENT

            ST1 = " INSERT INTO ONOMATA_PINAKON (ONOMATA_PINAKON)" & _
            " VALUES('" & ASD1 & "')"
            DB2.Execute ST1
        Else
            'TIPOTA
        End If
        XX1 = XX1 + 1
        I = I + 1
        If RS2A.STATE = 1 Then RS2A.Close
        C = 0
    Loop
End If

'**** DIADIKASIAS YPOLOGISMOY SUM *************************
If MonthView1.SelStart <> MonthView1.SelEnd Then
    XX = MonthView1.SelStart
    I = 0
    Do While XX <= MonthView1.SelEnd
        D = Day(MonthView1.Value + I)
        m = Month(MonthView1.Value)
        y = Year(MonthView1.Value)
        ASDD = D & m & y
        RS2.Open "[" & ASDD & "]", DB2, adOpenDynamic, adLockBatchOptimistic
        If RS2.BOF = RS2.EOF Then GoTo NIK:
        RS2.MoveFirst
NIK:
        Do While Not RS2.EOF
            SUM = RS2![посо] + SUM
            RS2.MoveNext
        Loop
        XX = XX + 1
        I = I + 1
        If RS2.STATE = 1 Then RS2.Close
    Loop
    Label4.Caption = SUM
End If
GoTo TELOS:


ER:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"



TELOS:
If DB2.STATE = 1 Then DB2.Close
If RS2.STATE = 1 Then RS2.Close
End Sub

Private Sub Picture1_Click()
Call Shell(App.Path & "\EFARMOGES\calc.exe", vbNormalFocus)
End Sub

Private Sub Picture2_Click()
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
End Sub

Private Sub Text4_LostFocus()
Dim dd As Integer
Dim S As String

dd = Len(Text4.Text)
S = Text4.Text
'For I = 1 To dd
'    If Mid(S, I, 1) = "." Then
'        Mid(S, I, 1) = ","
'    End If
'Next I
Text4.Text = S

End Sub

Private Sub Text5_Change()

Dim S As String
Dim mhkos As Integer
mhkos = 0
mhkos = Len(Text5.Text)
If mhkos > 244 Then
    S = Text5.Text
    Text5.Text = Mid(S, 1, 244)
End If

End Sub


