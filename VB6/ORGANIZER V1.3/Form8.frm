VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form8 
   BackColor       =   &H80000013&
   Caption         =   "еныжкгсг тилокоциым"
   ClientHeight    =   10485
   ClientLeft      =   960
   ClientTop       =   360
   ClientWidth     =   11385
   LinkTopic       =   "Form8"
   ScaleHeight     =   10485
   ScaleWidth      =   11385
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   10200
      TabIndex        =   22
      Text            =   "Text6"
      Top             =   7440
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   9960
      TabIndex        =   21
      Text            =   "HELP_TEXT"
      Top             =   8040
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00E9C5AD&
      Caption         =   "епикоцг окым"
      Height          =   735
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   5160
      Width           =   1335
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Height          =   855
      Left            =   9120
      TabIndex        =   19
      Top             =   9360
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   1508
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
      Left            =   9240
      Top             =   8880
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
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
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   9960
      TabIndex        =   18
      Text            =   "HELP_TEXT"
      Top             =   8400
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00E9C5AD&
      Caption         =   "йахаяислос кистас"
      Height          =   735
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   7080
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   9120
      TabIndex        =   15
      Top             =   3480
      Width           =   735
   End
   Begin VB.ComboBox Combo4 
      Height          =   315
      Left            =   10080
      TabIndex        =   14
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00E9C5AD&
      Caption         =   "сглеяимг"
      Height          =   375
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00E9C5AD&
      Caption         =   "OK"
      Height          =   375
      Left            =   9480
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1200
      Width           =   495
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   10080
      TabIndex        =   9
      Text            =   "етос"
      Top             =   720
      Width           =   975
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   9000
      TabIndex        =   8
      Text            =   "лгмас"
      Top             =   720
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   7920
      TabIndex        =   7
      Text            =   "глеяа"
      Top             =   720
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   7800
      TabIndex        =   6
      Top             =   2280
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   7920
      TabIndex        =   5
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00E9C5AD&
      Caption         =   "епистяожг"
      Height          =   735
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   9600
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E9C5AD&
      Caption         =   "диацяажг тилокоциоу апо киста"
      Height          =   735
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6120
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E9C5AD&
      Caption         =   "еныжкгсг тилокоциым"
      Height          =   975
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   9000
      Width           =   1575
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   1200
      Top             =   8760
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
   Begin VB.ListBox List1 
      Height          =   8445
      Left            =   4680
      TabIndex        =   1
      Top             =   360
      Width           =   2895
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   8613
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   15187
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
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000013&
      Caption         =   "аяихлос тилокоциым пяос еныжкгсг"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7680
      TabIndex        =   16
      Top             =   3120
      Width           =   3615
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000013&
      Caption         =   "аяихлос епитацгс"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7800
      TabIndex        =   13
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000013&
      Caption         =   "глеяолгмиа еныжкгсгс"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7920
      TabIndex        =   12
      Top             =   360
      Width           =   2535
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Click()
Dim DAY_HM_EKSOF As String
Select Case Combo1.ListIndex
    Case 0
        DAY_HM_EKSOF = "1"
    Case 1
         DAY_HM_EKSOF = "2"
    Case 2
        DAY_HM_EKSOF = "3"
    Case 3
         DAY_HM_EKSOF = "4"
    Case 4
        DAY_HM_EKSOF = "5"
    Case 5
         DAY_HM_EKSOF = "6"
    Case 6
        DAY_HM_EKSOF = "7"
    Case 7
         DAY_HM_EKSOF = "8"
    Case 8
        DAY_HM_EKSOF = "9"
    Case 9
         DAY_HM_EKSOF = "10"
    Case 10
        DAY_HM_EKSOF = "11"
    Case 11
         DAY_HM_EKSOF = "12"
    Case 12
        DAY_HM_EKSOF = "13"
    Case 13
         DAY_HM_EKSOF = "14"
    Case 14
        DAY_HM_EKSOF = "15"
    Case 15
         DAY_HM_EKSOF = "16"
    Case 16
        DAY_HM_EKSOF = "17"
    Case 17
         DAY_HM_EKSOF = "18"
    Case 18
        DAY_HM_EKSOF = "19"
    Case 19
         DAY_HM_EKSOF = "20"
    Case 20
        DAY_HM_EKSOF = "21"
    Case 21
         DAY_HM_EKSOF = "22"
    Case 22
        DAY_HM_EKSOF = "23"
    Case 23
         DAY_HM_EKSOF = "24"
    Case 24
        DAY_HM_EKSOF = "25"
    Case 25
         DAY_HM_EKSOF = "26"
    Case 26
        DAY_HM_EKSOF = "27"
    Case 27
         DAY_HM_EKSOF = "28"
    Case 28
        DAY_HM_EKSOF = "29"
    Case 29
         DAY_HM_EKSOF = "30"
    Case 30
        DAY_HM_EKSOF = "31"
    End Select
End Sub

Private Sub Combo2_Change()
Dim MONTH_HM_EKSOF As String
Select Case Combo2.ListIndex
    Case 0
        MONTH_HM_EKSOF = "1"
    Case 1
         MONTH_HM_EKSOF = "2"
    Case 2
        MONTH_HM_EKSOF = "3"
    Case 3
         MONTH_HM_EKSOF = "4"
    Case 4
        MONTH_HM_EKSOF = "5"
    Case 5
         MONTH_HM_EKSOF = "6"
    Case 6
        MONTH_HM_EKSOF = "7"
    Case 7
         MONTH_HM_EKSOF = "8"
    Case 8
        MONTH_HM_EKSOF = "9"
    Case 9
         MONTH_HM_EKSOF = "10"
    Case 10
        MONTH_HM_EKSOF = "11"
    Case 11
         MONTH_HM_EKSOF = "12"
    End Select
End Sub

Private Sub Combo3_Change()
Dim ETOS_HM_EKSOF As String
Select Case Combo3.ListIndex
    Case 0
        ETOS_HM_EKSOF = "2005"
    Case 1
         ETOS_HM_EKSOF = "2006"
    Case 2
        ETOS_HM_EKSOF = "2007"
    Case 3
         ETOS_HM_EKSOF = "2008"
    Case 4
        ETOS_HM_EKSOF = "2009"
    Case 5
         ETOS_HM_EKSOF = "2010"
    Case 6
        ETOS_HM_EKSOF = "2011"
    Case 7
         ETOS_HM_EKSOF = "2012"
    Case 8
        ETOS_HM_EKSOF = "2013"
    Case 9
         ETOS_HM_EKSOF = "2014"
    Case 10
        ETOS_HM_EKSOF = "2015"
    Case 11
         ETOS_HM_EKSOF = "2016"
    Case 12
        ETOS_HM_EKSOF = "2017"
    Case 13
         ETOS_HM_EKSOF = "2018"
    Case 14
        ETOS_HM_EKSOF = "2019"
    Case 15
         ETOS_HM_EKSOF = "2020"
End Select
End Sub

Private Sub Combo4_Click()
Select Case Combo4.ListIndex
    Case 0
        Form8.Text2.Text = "летягта"
    Case 1
        Form8.Text2.Text = "епитацг"
    End Select
    Combo4.Text = ""
End Sub

Private Sub Command1_Click()
On Error GoTo ER:

Dim TIMOLOGIO As String
Dim STATE1, STATE2, STATE3, State, DATABASE_FILE, SS As String
Dim A, I, index As Integer
I = 0
index = 0

If RS1.State = 1 Then RS1.Close
If DB1.State = 1 Then DB1.Close
DB1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
       "Data Source=" & "\databases\ETAIRIES.mdb" & ";" & _
      "Persist Security Info=False"
DB1.Open App.Path & "\databases\ETAIRIES.mdb"
RS1.Open "[" & Form4.Text1.Text & "]", DB1, adOpenDynamic, adLockBatchOptimistic
'*************ELEGXOI******************************************
If List1.List(0) = "" Then GoTo LIST_1:
If Form8.Text1.Text = "" Then GoTo HMER_EKSO_1:
If IsDate(Form8.Text1.Text) = False Then GoTo HMER_EKSO_2:
If Form8.Text2.Text = "" Then GoTo typos_1:
'**********************************************************

' *********EKSOFLISI TIMOLOGION LISTAS*****************************
If List1.ListCount <> 0 Then ' AN YPARXOYN TIMOLOGIA GIA EKSOFLISI
    If MsgBox("хекете ма пяовыягсете се еныжкгсг тым епикецлемым тилокоциым", vbOKCancel, "") = vbOK Then
        A = List1.ListCount
        Text3.Text = A
        Do While I <= A - 1
            TIMOLOGIO = List1.List(I)
            State = "UPDATE " & UCase(Form4.Text1.Text) & _
                " SET глеяолгмиа_еныжкгсгс= " & "'" & Form8.Text1.Text & "'" & _
                " WHERE аяихлос_тилокоциоу='" & TIMOLOGIO & "'"

            STATE1 = " UPDATE " & UCase(Form4.Text1.Text) & _
                " SET еныжкгсг=" & "'1'" & _
                " WHERE аяихлос_тилокоциоу='" & TIMOLOGIO & "'"

            STATE2 = " UPDATE " & UCase(Form4.Text1.Text) & _
                " SET аяихлос_епитацгс = '" & Text2.Text & "'" & _
                " WHERE аяихлос_тилокоциоу='" & TIMOLOGIO & "'"
            DB1.Execute STATE2
            DB1.Execute STATE1
            DB1.Execute State
            I = I + 1
        Loop
        If Text2.Text = "летягта" Then
            I = 0
            Do While I <= A - 1
                TIMOLOGIO = List1.List(I)
                If RS1.BOF = RS1.EOF Then GoTo NIK:
                RS1.MoveFirst
NIK:
                Do While Not RS1.EOF
                    If RS1![аяихлос_тилокоциоу] = TIMOLOGIO Then
                        Text5.Text = RS1![посо]
                        STATE3 = " UPDATE " & UCase(Form4.Text1.Text) & _
                        " SET пистысг='" & Text5.Text & "'" & _
                        " WHERE аяихлос_тилокоциоу='" & TIMOLOGIO & "'"
                        DB1.Execute STATE3
                        RS1.MoveNext
                    Else
                        RS1.MoveNext
                    End If
                Loop
                I = I + 1
            Loop
        Else
        
        End If
            MsgBox ("г еныжкгсг тым тилокоциым поу епикенате окойкгяыхгйе"), , ""
    End If
    
    Form8.Text1.Text = ""
    Form8.Text2.Text = ""
    Combo1.Text = "глеяа"
    Combo2.Text = "лгмас"
    Combo3.Text = "етос"
    Combo4.Text = ""
    List1.Clear
    Text3.Text = List1.ListCount
    
    RS1.Fields.Refresh
    RS1.Close
    DB1.Close
    Dim DATABASE_FILE1 As String
    DATABASE_FILE1 = App.Path & "\databases\ETAIRIES.mdb"
    Adodc1.ConnectionString = _
    "PROVIDER=Microsoft.Jet.OLEDB.4.0;" & _
    "Data Source=" & DATABASE_FILE1 & ";"
    Adodc1.RecordSource = "SELECT аяихлос_тилокоциоу FROM " & UCase(Form4.Text1.Text) & _
        " WHERE еныжкгсг=0 AND тупос='тилок\пыкгсгс'" & _
        " ORDER BY глеяолгмиа_ейдосгс"
    Adodc1.Refresh
    DataGrid1.Refresh
    Set DataGrid1.DataSource = Adodc1
    Text6.Text = Adodc1.Recordset.RecordCount
    If Text6.Text <= 28 Then
        DataGrid1.Height = (1 + Text6.Text) * 298
    Else
        DataGrid1.Height = 8613
    End If
Else ' AN DEN YPARXOYN TIMOLOGIA GIA EKSOFLISI
    MsgBox ("дем евоум епикецеи тилокоциа пяос еныжкгсг"), , ""
End If
GoTo TELOS:
'*****************TELOS EKSOFLISIS***********************

'********DIAXEIRISI ELEGXON*******************************
LIST_1:
MsgBox ("дем евете епикенеи йамема тилокоцио"), vbCritical, "пяосовг!!"
index = 2
GoTo TELOS:

HMER_EKSO_1:
If index = 0 Then
    MsgBox ("дем дысате глеяолгмиа еныжкгсгс"), vbCritical, "пяосовг!!"
    index = 2
    GoTo TELOS:
Else
    GoTo TELOS:
End If

HMER_EKSO_2:
If index = 0 Then
    MsgBox ("дем дысате сыста тгм глеяолгмиа еныжкгсгс"), vbCritical, "пяосовг!!"
    index = 2
    GoTo TELOS:
Else
    GoTo TELOS:
End If

typos_1:
If index = 0 Then
    MsgBox ("дем дысате аяихло епитацгс"), vbCritical, "пяосовг!!"
    index = 2
    GoTo TELOS:
Else
    GoTo TELOS:
End If


ER:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
If RS1.State = 1 Then RS1.Close
If DB1.State = 1 Then DB1.Close
End Sub

Private Sub Command2_Click()
On Error GoTo ER:
If List1.ListIndex = True Then GoTo E:

Dim index As Integer
index = 0
Dim A As Integer


If List1.List(0) = "" Then
    GoTo NIK:
Else
        A = List1.ListIndex
        List1.RemoveItem (A)
        Text3.Text = List1.ListCount
        GoTo TELOS:
End If

NIK:
MsgBox ("г киста еимаи йемг"), , ""
index = 2

E:
If index = 0 Then
    MsgBox ("ха пяепеи пяыта ма епикенете то тилокоцио поу хекете ма диацяаьете апо тгм киста"), vbCritical, "пяосовг !!!"
    index = 2
Else
    GoTo TELOS:
End If


ER:
If index = 0 Then
    MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"
Else
    GoTo TELOS:
End If


TELOS:

End Sub

Private Sub Command3_Click()
On Error GoTo ER:

Form8.Hide
Unload Form8
GoTo TELOS:

ER:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
End Sub

Private Sub Command4_Click()
On Error GoTo ER:

Dim HM_EKSOF As String
Dim DATE_HM_EKSOF As Date
'***************** ELEGXOI **************************************
If IsNumeric(Combo1.Text) = False Then
    MsgBox ("дем дысате сыста глеяа"), vbCritical, "пяосовг !!!"
    GoTo TELOS:
End If
If IsNumeric(Combo2.Text) = False Then
    MsgBox ("дем дысате сыста лгма"), vbCritical, "пяосовг !!!"
    GoTo TELOS:
End If
If IsNumeric(Combo3.Text) = False Then
    MsgBox ("дем дысате сыста етос"), vbCritical, "пяосовг !!!"
    GoTo TELOS:
End If
If CInt(Combo1.Text) < 1 Or CInt(Combo1.Text) > 31 Then
    MsgBox ("о аяихлос поу дысате дем еимаи ецйуяг глеяа лгма"), vbCritical, "пяосовг !!!"
    GoTo TELOS:
End If
If CInt(Combo2.Text) < 1 Or CInt(Combo2.Text) > 12 Then
    MsgBox ("о аяихлос поу дысате дем еимаи ецйуяос лгмас етоус"), vbCritical, "пяосовг !!!"
    GoTo TELOS:
End If
If CInt(Combo3.Text) < 2005 Or CInt(Combo3.Text) > 2020 Then
    MsgBox ("то пяоцяалла упостгяифеи глеяолгмиес апо 2005 еыс 2020.паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"
    GoTo TELOS:
End If
'*************** LEITOYRGIA *******************************
HM_EKSOF = Combo1.Text & "/" & Combo2.Text & _
"/" & Combo3.Text

If IsDate(HM_EKSOF) = True Then
DATE_HM_EKSOF = CDate(HM_EKSOF)
Text1.Text = DATE_HM_EKSOF
Else
MsgBox ("дем дысате глеяолгмиа"), vbCritical, "пяосовг!!"
End If
GoTo TELOS:
'***********************************************************

ER:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
End Sub

Private Sub Command5_Click()
On Error GoTo ER:

Text1.Text = Date
GoTo TELOS:

ER:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
End Sub

Private Sub Command6_Click()
On Error GoTo ER:

List1.Clear
Text3.Text = List1.ListCount
GoTo TELOS:

ER:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
End Sub

Private Sub Command7_Click()
On Error GoTo ER:
List1.Clear

If RS1.State = 1 Then RS1.Close
If DB1.State = 1 Then DB1.Close
DB1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
       "Data Source=" & "\databases\ETAIRIES.mdb" & ";" & _
      "Persist Security Info=False"
DB1.Open App.Path & "\databases\ETAIRIES.mdb"
RS1.Open "[" & Form4.Text1.Text & "]", DB1, adOpenDynamic, adLockBatchOptimistic
If RS1.BOF = RS1.EOF Then GoTo NIK:
RS1.MoveNext
NIK:
Do While Not RS1.EOF
    If RS1![еныжкгсг] = 0 And RS1![тупос] = "тилок\пыкгсгс" Then
        List1.AddItem RS1![аяихлос_тилокоциоу]
        RS1.MoveNext
    Else
        RS1.MoveNext
    End If
Loop
Text3.Text = List1.ListCount
GoTo TELOS:

ER:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"


TELOS:
If RS1.State = 1 Then RS1.Close
If DB1.State = 1 Then DB1.Close
End Sub

Private Sub DataGrid1_DblClick()
Dim A As String
Dim B, C As Integer
C = 0
A = DataGrid1.Text
If List1.List(0) = "" Then
    List1.AddItem A
    Text3.Text = List1.ListCount
Else
    B = List1.ListCount
    For I = 0 To B - 1
        If A = List1.List(I) Then C = C + 1
    Next I
    If C = 0 Then
        List1.AddItem A
        Text3.Text = List1.ListCount
    Else
        MsgBox ("то тилокоцио : " & DataGrid1.Text & " еимаи гдг пеяаслемо стгм киста"), , "пяосовг"
    End If
End If
End Sub

Private Sub Form_Load()
On Error GoTo ER:
DataGrid1.DefColWidth = 3000
DataGrid1.Font.Size = 10
DataGrid1.HeadFont.Size = 10
DataGrid1.HeadFont.Bold = True
List1.Font.Size = 10
List1.Height = 8613
List1.Top = 120

'RITHMISI TIMON TON COMBO
Dim H, m, Y As Integer
H = 1
m = 1
Y = 2005

For I = 0 To 30
    Combo1.AddItem H + I
Next I

For I = 0 To 11
    Combo2.AddItem m + I
Next I

For I = 0 To 35
    Combo3.AddItem Y + I
Next I

Combo4.AddItem "летягта"
Combo4.AddItem "епитацг"

Dim State, STATECOUNT, DATABASE_FILE As String

State = "SELECT аяихлос_тилокоциоу FROM " & UCase(Form4.Text1.Text) & _
" WHERE еныжкгсг=0 AND тупос='тилок\пыкгсгс'" & _
" ORDER BY глеяолгмиа_ейдосгс"

STATECOUNT = "SELECT COUNT(аяихлос_тилокоциоу) FROM " & UCase(Form4.Text1.Text) & _
" WHERE еныжкгсг=0 AND тупос='тилок\пыкгсгс'"


DATABASE_FILE = App.Path & "\databases\ETAIRIES.mdb"

 Adodc2.ConnectionString = _
"PROVIDER=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & DATABASE_FILE & ";"
Adodc2.RecordSource = STATECOUNT
' Bind the ADODC to the DataGrid.
Set DataGrid2.DataSource = Adodc2
Text4.Text = DataGrid2.Text

If Text4.Text <= 28 Then
    DataGrid1.Height = (1 + Text4.Text) * 298
Else
    DataGrid1.Height = 8613
End If

Adodc1.ConnectionString = _
"PROVIDER=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & DATABASE_FILE & ";"
Adodc1.RecordSource = State
Set DataGrid1.DataSource = Adodc1

GoTo TELOS:

ER:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ER:

Form8.Hide
Unload Form8
GoTo TELOS:

ER:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
End Sub

Private Sub Text2_LostFocus()
Text2.Text = UCase(Text2.Text)
Text2.Text = Trim(Text2.Text)
End Sub
