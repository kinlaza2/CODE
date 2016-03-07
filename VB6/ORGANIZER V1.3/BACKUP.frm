VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form BACKUP 
   BackColor       =   &H80000013&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "амтицяажа асжакеиас"
   ClientHeight    =   10515
   ClientLeft      =   45
   ClientTop       =   450
   ClientWidth     =   15300
   ControlBox      =   0   'False
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   10515
   ScaleWidth      =   15300
   Begin VB.CommandButton Command7 
      BackColor       =   &H00F1C896&
      Caption         =   "диацяажг амтицяажым"
      Height          =   1335
      Left            =   10430
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7680
      Width           =   3495
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00F1C896&
      Caption         =   "амтийатастасеис ле амтицяажа"
      Height          =   1335
      Left            =   5950
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7680
      Width           =   3495
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00F1C896&
      Caption         =   "диавеияисг амтицяажым"
      Height          =   1335
      Left            =   1450
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7680
      Width           =   3495
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00E9C5AD&
      Caption         =   "епистяожг сто аявийо лемоу"
      Height          =   855
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   9480
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00E9C5AD&
      Caption         =   "дглиоуяциа амтицяажоу глеяокоциоу"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1350
      Left            =   6700
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1820
      Width           =   2000
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E9C5AD&
      Caption         =   "дглиоуяциа етгсиоу амтицяажоу  глеяокоциоу"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1350
      Left            =   11160
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1800
      Width           =   2000
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E9C5AD&
      Caption         =   "дглиоуяциа амтицяажоу тгкежымийоу йатакоцоу "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1350
      Left            =   2210
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1820
      Width           =   2000
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000013&
      Caption         =   "дглиоуяциа амтицяажым глеяокоциоу йаи тгкежымийоу йатакоцоу"
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
      Height          =   2655
      Left            =   960
      TabIndex        =   11
      Top             =   1080
      Width           =   13455
      Begin VB.Line Line4 
         Visible         =   0   'False
         X1              =   4485
         X2              =   4485
         Y1              =   120
         Y2              =   2640
      End
      Begin VB.Line Line19 
         Visible         =   0   'False
         X1              =   11205
         X2              =   11205
         Y1              =   120
         Y2              =   2600
      End
      Begin VB.Line Line18 
         Visible         =   0   'False
         X1              =   6727
         X2              =   6727
         Y1              =   20
         Y2              =   2550
      End
      Begin VB.Line Line17 
         Visible         =   0   'False
         X1              =   0
         X2              =   13440
         Y1              =   1380
         Y2              =   1380
      End
      Begin VB.Line Line11 
         Visible         =   0   'False
         X1              =   2242
         X2              =   2242
         Y1              =   120
         Y2              =   2640
      End
      Begin VB.Line Line5 
         Visible         =   0   'False
         X1              =   8970
         X2              =   8970
         Y1              =   120
         Y2              =   2640
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000013&
      Caption         =   "дглиоуяциа амтицяажым циа етаияиес йаи тилокоциа"
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
      Height          =   2655
      Left            =   960
      TabIndex        =   4
      Top             =   4560
      Width           =   13455
      Begin VB.Timer Timer2 
         Interval        =   100
         Left            =   4200
         Top             =   1080
      End
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   4200
         Top             =   600
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   495
         Left            =   4440
         TabIndex        =   12
         Top             =   1560
         Visible         =   0   'False
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   873
         _Version        =   393216
         Appearance      =   1
         Max             =   10
         Scrolling       =   1
      End
      Begin VB.CommandButton Command9 
         BackColor       =   &H00E9C5AD&
         Caption         =   "дглиоуяциа етгсиоу амтицяажоу етаияиым йаи тилокоциым"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1350
         Left            =   10200
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   720
         Width           =   2000
      End
      Begin VB.CommandButton Command8 
         BackColor       =   &H00E9C5AD&
         Caption         =   "дглиоуяциа амтицяажоу етаияиым йаи тилокоциым"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1350
         Left            =   1240
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   720
         Width           =   2000
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "паяайакы пеяилемете"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5040
         TabIndex        =   13
         Top             =   720
         Visible         =   0   'False
         Width           =   3855
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00C0C0C0&
         BorderColor     =   &H00000000&
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   1815
         Left            =   4080
         Top             =   480
         Visible         =   0   'False
         Width           =   5895
      End
      Begin VB.Line Line21 
         Visible         =   0   'False
         X1              =   11200
         X2              =   11200
         Y1              =   120
         Y2              =   2800
      End
      Begin VB.Line Line16 
         Visible         =   0   'False
         X1              =   0
         X2              =   14000
         Y1              =   1380
         Y2              =   1380
      End
      Begin VB.Line Line6 
         Visible         =   0   'False
         X1              =   4485
         X2              =   4485
         Y1              =   120
         Y2              =   2640
      End
      Begin VB.Line Line12 
         Visible         =   0   'False
         X1              =   2242
         X2              =   2242
         Y1              =   120
         Y2              =   2640
      End
      Begin VB.Line Line7 
         Visible         =   0   'False
         X1              =   8970
         X2              =   8970
         Y1              =   120
         Y2              =   2640
      End
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      Visible         =   0   'False
      X1              =   7695
      X2              =   7695
      Y1              =   45
      Y2              =   11005
   End
   Begin VB.Line Line22 
      Visible         =   0   'False
      X1              =   960
      X2              =   14700
      Y1              =   9120
      Y2              =   9120
   End
   Begin VB.Line Line20 
      Visible         =   0   'False
      X1              =   12170
      X2              =   12170
      Y1              =   3720
      Y2              =   9760
   End
   Begin VB.Line Line15 
      Visible         =   0   'False
      X1              =   360
      X2              =   1020
      Y1              =   5940
      Y2              =   5940
   End
   Begin VB.Line Line14 
      Visible         =   0   'False
      X1              =   960
      X2              =   480
      Y1              =   7200
      Y2              =   7200
   End
   Begin VB.Line Line13 
      Visible         =   0   'False
      X1              =   960
      X2              =   360
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Line Line10 
      BorderStyle     =   3  'Dot
      Visible         =   0   'False
      X1              =   3202
      X2              =   3202
      Y1              =   240
      Y2              =   9999
   End
   Begin VB.Line Line9 
      Visible         =   0   'False
      X1              =   14400
      X2              =   14400
      Y1              =   120
      Y2              =   9680
   End
   Begin VB.Line Line8 
      BorderStyle     =   3  'Dot
      Visible         =   0   'False
      X1              =   960
      X2              =   960
      Y1              =   360
      Y2              =   9040
   End
   Begin VB.Line Line3 
      BorderStyle     =   3  'Dot
      Visible         =   0   'False
      X1              =   9930
      X2              =   9930
      Y1              =   600
      Y2              =   9760
   End
   Begin VB.Line Line2 
      BorderStyle     =   3  'Dot
      Visible         =   0   'False
      X1              =   5445
      X2              =   5445
      Y1              =   480
      Y2              =   9120
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
      Caption         =   " амтицяажA  асжакеиас"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4600
      TabIndex        =   0
      Top             =   240
      Width           =   6135
   End
End
Attribute VB_Name = "BACKUP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim I As Integer
Private Sub Command1_Click()
On Error GoTo ER:
Dim STATEMENT, ONOMA, ONOMA1, APO, STO, STO2, CT As String
Dim MERA, MHNAS, ETOS
Dim C As Integer
C = 1
'******* SYNDESH ME BASH ********************************
Dim DBH As New ADODB.Connection
Dim RSH As New ADODB.Recordset
DBH.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
       "Data Source=" & "\databases\BACK_UPS\HELP_BACKUP.mdb" & ";" & _
      "Persist Security Info=False"
DBH.Open App.Path & "\databases\BACK_UPS\HELP_BACKUP.mdb"
RSH.Open "[BACKUP_THL]", DBH, adOpenDynamic, adLockBatchOptimistic
'*****DHMIOYRGIA ONOMATOS ARXEIOY KAI HMEROMHNIAS SAN STRING******
MERA = Day(Date)
MHNAS = Month(Date)
ETOS = Year(Date)
ONOMA = MERA & "_" & MHNAS & "_" & ETOS & "_THL"
ONOMA1 = MERA & "/" & MHNAS & "/" & ETOS
'***ELEGXOS AN YPARXEI ANTIGRAFO GIA THN SHMERINH HMEROMHNIA*******************************
If RSH.BOF = RSH.EOF Then GoTo NIK:
RSH.MoveFirst
NIK:
Do While Not RSH.EOF
    If ONOMA1 <> RSH![глеяолгмиа_дглиоуяциас_амтицяажоу] Then
        RSH.MoveNext
    Else
        C = C + 1
        RSH.MoveNext
    End If
Loop
'*******PROGRAMATISMOS********************

If C <> 1 Then
    CT = "евете гдг дглиоуяцгсг амтицяажо тгкежымийоу йатакоцоу циа тгм: " & ONOMA1 & ""
    MsgBox (CT), , "пяосовг"
Else
    APO = App.Path & "\DATABASES\telephone.mdb"
    STO = App.Path & "\DATABASES\BACK_UPS\BACKUP_THL\" & ONOMA & ".MDB"
    STO2 = "C:\WINDOWS\ORGANIZER_BACKUP\BACKUP_THL\" & ONOMA & ".MDB"
    FileCopy APO, STO
    FileCopy APO, STO2
    STATEMENT = "INSERT INTO BACKUP_THL(глеяолгмиа_дглиоуяциас_амтицяажоу,омола_аявеиоу,HMEROMHNIA)" & _
    "VALUES(" & _
    "'" & ONOMA1 & "'," & _
    "'" & ONOMA & "'," & _
    "#" & Date & "# )"
    DBH.Execute STATEMENT
 
    MsgBox ("г дглиоуяциа амтицяажоу тгкежымийоу йатакоцоу ециме ле епитувиа"), , ""
End If
GoTo TELOS:

ER:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
If DBH.STATE = 1 Then DBH.Close
If RSH.STATE = 1 Then RSH.Close
End Sub

Private Sub Command2_Click()
On Error GoTo ER:

Dim ETOS, ETOS_PLUS_1 As String
Dim TIMH_FLAG, TIMH_ETOS As Integer
Dim APO, STO, STO2 As String
Dim STATEMENT, STATEMENT1, STATEMENT_FLAG As String
Dim STATE1, STATE2 As String
Dim DILOSI1 As String
' H METABLHTH ETOS EINAI TO PROHGOYMENO ETOS
ETOS = CStr(Year(Date) - 1)
ETOS_PLUS_1 = CStr(Year(Date))
TIMH_ETOS = CInt(ETOS)
' DHMIOYRGIA SYNDESHS
Dim DBH As New ADODB.Connection
Dim RSH As New ADODB.Recordset
DBH.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
       "Data Source=" & "\databases\BACK_UPS\HELP_BACKUP.mdb" & ";" & _
      "Persist Security Info=False"
DBH.Open App.Path & "\databases\BACK_UPS\HELP_BACKUP.mdb"
RSH.Open "[BACKUP_YEAR_HMER]", DBH, adOpenDynamic, adLockBatchOptimistic
RSH.MoveFirst
' ANIXNEYSH TOY FLAG GIA TO ZHTOYMENO ETOS
Do While Not RSH.EOF
If RSH![ETOS] = TIMH_ETOS Then
    TIMH_FLAG = RSH![FLAG]
    RSH.MoveNext
Else
    RSH.MoveNext
End If
Loop
' AN HDH EXV ANTIGRAFO SXETIKO MYNHMA
If TIMH_FLAG = 1 Then
    DILOSI1 = "евете дглиоуяцгсг амтицяажо глеяокоциоу циа то етос:" & CStr(ETOS)
    MsgBox (DILOSI1), , "пяосовг"
Else
' AN DEN EXO TOTE DHMIOYRGO ANTIGRAFO
    APO = App.Path & "\DATABASES\HMEROLOGIO.mdb"
    STO = App.Path & "\DATABASES\BACK_UPS\BACKUP_HMER\BACKUP_HMER_ETOS\" & ETOS & "_HMER.MDB"
    FileCopy APO, STO
'*********************** DIAGRAFH PINAKON EKTOS ETOYS ********************************
    If RSH.STATE = 1 Then RSH.Close
    If DBH.STATE = 1 Then DBH.Close
    
    ' SYNDESH ME THN DHMIOYRGOYMENH BASH
    DBH.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
       "Data Source=" & "\databases\BACK_UPS\BACKUP_HMER\BACKUP_HMER_ETOS\" & ETOS & "_HMER.MDB" & ";" & _
      "Persist Security Info=False"
    DBH.Open App.Path & "\databases\BACK_UPS\BACKUP_HMER\BACKUP_HMER_ETOS\" & ETOS & "_HMER.MDB"
    RSH.Open "[ONOMATA_PINAKON]", DBH, adOpenDynamic, adLockBatchOptimistic
    If RSH.BOF = RSH.EOF Then GoTo NNN:
    RSH.MoveFirst
NNN:
    Do While Not RSH.EOF
    If Right(RSH![ONOMATA_PINAKON], 4) = ETOS Then
        RSH.MoveNext
    Else
        STATEMENT = " DROP TABLE " & RSH![ONOMATA_PINAKON]
        DBH.Execute STATEMENT
        RSH.MoveNext
    End If
    Loop
    '--------------------------------------------------------
    If RSH.BOF = RSH.EOF Then GoTo NNN1:
    RSH.MoveFirst
NNN1:
    Do While Not RSH.EOF
    If Right(RSH![ONOMATA_PINAKON], 4) = ETOS Then
        RSH.MoveNext
    Else
        STATEMENT1 = " DELETE FROM ONOMATA_PINAKON" & _
        " WHERE ONOMATA_PINAKON='" & RSH![ONOMATA_PINAKON] & "'"
        DBH.Execute STATEMENT1
        RSH.MoveNext
    End If
    Loop
    '-------------------------------------------------------------
    'KSANA ANOIGMA BOHTHITIKHS GIA TROPOIHSH FLAG
    If RSH.STATE = 1 Then RSH.Close
    If DBH.STATE = 1 Then DBH.Close
    DBH.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
       "Data Source=" & "\databases\BACK_UPS\HELP_BACKUP.mdb" & ";" & _
      "Persist Security Info=False"
    DBH.Open App.Path & "\databases\BACK_UPS\HELP_BACKUP.mdb"
    RSH.Open "[BACKUP_YEAR_HMER]", DBH, adOpenDynamic, adLockBatchOptimistic
    STATEMENT_FLAG = " UPDATE BACKUP_YEAR_HMER " & _
        " SET FLAG=1" & _
        " WHERE ETOS=" & TIMH_ETOS
    DBH.Execute STATEMENT_FLAG
    MsgBox ("г дглиоуяциа амтицяажоу глеяокоциоу циа то етос: " & ETOS & " ециме ле епитувиа"), , ""
    If RSH.STATE = 1 Then RSH.Close
    If DBH.STATE = 1 Then DBH.Close
    ' SYNDESH ME TREXON BASH HMEROLOGIO PROKEIMENOY NA DIAGRAPSO TIS
    ' EGRAFES POY MOLIS KRATHSA BACKUP
    DBH.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
    "Data Source=" & "\databases\HMEROLOGIO.mdb" & ";" & _
    "Persist Security Info=False"
    DBH.Open App.Path & "\databases\HMEROLOGIO.mdb"
    RSH.Open "[ONOMATA_PINAKON]", DBH, adOpenDynamic, adLockBatchOptimistic
    If RSH.BOF = RSH.EOF Then GoTo mm:
    RSH.MoveFirst
mm:
    Do While Not RSH.EOF
    If Right(RSH![ONOMATA_PINAKON], 4) = ETOS Then
        STATE1 = " DROP TABLE " & RSH![ONOMATA_PINAKON]
        DBH.Execute STATE1
        RSH.MoveNext
    Else
        RSH.MoveNext
    End If
    Loop
    STATE2 = " DELETE FROM ONOMATA_PINAKON " & _
    " WHERE ONOMATA_PINAKON LIKE '%" & ETOS & "'"
    DBH.Execute STATE2
    STO2 = "C:\WINDOWS\ORGANIZER_BACKUP\BACKUP_HMER\BACKUP_HMER_ETOS\" & ETOS & "_HMER.MDB"
    FileCopy STO, STO2
End If
GoTo TELOS:

ER:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
If RSH.STATE = 1 Then RSH.Close
If DBH.STATE = 1 Then DBH.Close
End Sub

Private Sub Command3_Click()
On Error GoTo ER:
Dim STATEMENT, ONOMA, ONOMA1, APO, STO, STO2, CT As String
Dim MERA, MHNAS, ETOS
Dim C As Integer
C = 1
'******* SYNDESH ME BASH ********************************
Dim DBH As New ADODB.Connection
Dim RSH As New ADODB.Recordset
DBH.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
       "Data Source=" & "\databases\BACK_UPS\HELP_BACKUP.mdb" & ";" & _
      "Persist Security Info=False"
DBH.Open App.Path & "\databases\BACK_UPS\HELP_BACKUP.mdb"
RSH.Open "[BACKUP_HMER]", DBH, adOpenDynamic, adLockBatchOptimistic
'*****DHMIOYRGIA ONOMATOS ARXEIOY KAI HMEROMHNIAS SAN STRING******
MERA = Day(Date)
MHNAS = Month(Date)
ETOS = Year(Date)
ONOMA = MERA & "_" & MHNAS & "_" & ETOS & "_HMER"
ONOMA1 = MERA & "/" & MHNAS & "/" & ETOS
'***ELEGXOS AN YPARXEI ANTIGRAFO GIA THN SHMERINH HMEROMHNIA*******************************
If RSH.BOF = RSH.EOF Then GoTo NIK:
RSH.MoveFirst
NIK:
Do While Not RSH.EOF
    If ONOMA1 <> RSH![глеяолгмиа_дглиоуяциас_амтицяажоу] Then
        RSH.MoveNext
    Else
        C = C + 1
        RSH.MoveNext
    End If
Loop
'*******PROGRAMATISMOS********************
If C <> 1 Then
    CT = "евете гдг дглиоуяцгсг амтицяажо глеяокоциоу циа тгм: " & ONOMA1 & ""
    MsgBox (CT), , "пяосовг"
Else
        'allagh toy date
    Dim date1
    If Day(Date) < 12 Then
        date1 = CDate(Day(Date) & " / " & Month(Date) & " / " & Year(Date))
    Else
        date1 = Date
    End If
    
    APO = App.Path & "\DATABASES\HMEROLOGIO.mdb"
    STO = App.Path & "\DATABASES\BACK_UPS\BACKUP_HMER\BACKUP_HMER\" & ONOMA & ".MDB"
    STO2 = "C:\WINDOWS\ORGANIZER_BACKUP\BACKUP_HMER\BACKUP_HMER\" & ONOMA & ".MDB"
    FileCopy APO, STO
    FileCopy APO, STO2
    STATEMENT = "INSERT INTO BACKUP_HMER(глеяолгмиа_дглиоуяциас_амтицяажоу,омола_аявеиоу,HMEROMHNIA)" & _
    "VALUES(" & _
    "'" & ONOMA1 & "'," & _
    "'" & ONOMA & "'," & _
    "'" & date1 & "' )"
    DBH.Execute STATEMENT
    
    MsgBox ("г дглиоуяциа амтицяажоу глеяокоциоу ециме ле епитувиа"), , ""
End If
GoTo TELOS:

ER:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
If DBH.STATE = 1 Then DBH.Close
If RSH.STATE = 1 Then RSH.Close
End Sub

Private Sub Command4_Click()
On Error GoTo ER:
BACKUP.Hide
Unload BACKUP
Form1.Enabled = True
GoTo TELOS:

ER:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
End Sub

Private Sub Command5_Click()
On Error GoTo ER:
LABEL_FOR_BACKUP = "диавеияисг амтицяажым"
Load Form9
Form9.Show
BACKUP.Enabled = False
GoTo TELOS:

ER:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
End Sub

Private Sub Command6_Click()
On Error GoTo ER:
LABEL_FOR_BACKUP = "амтийатастасг ле амтицяажа"
Load Form9
Form9.Show
BACKUP.Enabled = False
GoTo TELOS:

ER:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
End Sub

Private Sub Command7_Click()
On Error GoTo ER:
LABEL_FOR_BACKUP = "диацяажг амтицяажым"
Load Form9
Form9.Show
BACKUP.Enabled = False
GoTo TELOS:

ER:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
End Sub

Private Sub Command8_Click()
On Error GoTo ER:
Dim STATEMENT, ONOMA, ONOMA1, APO, STO, STO2, CT As String
Dim MERA, MHNAS, ETOS
Dim C As Integer
C = 1
'******* SYNDESH ME BASH ********************************
Dim DBH As New ADODB.Connection
Dim RSH As New ADODB.Recordset
DBH.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
       "Data Source=" & "\databases\BACK_UPS\HELP_BACKUP.mdb" & ";" & _
      "Persist Security Info=False"
DBH.Open App.Path & "\databases\BACK_UPS\HELP_BACKUP.mdb"
RSH.Open "[BACKUP_ETAIRIES]", DBH, adOpenDynamic, adLockBatchOptimistic
'*****DHMIOYRGIA ONOMATOS ARXEIOY KAI HMEROMHNIAS SAN STRING******
MERA = Day(Date)
MHNAS = Month(Date)
ETOS = Year(Date)
ONOMA = MERA & "_" & MHNAS & "_" & ETOS & "_ETAIRIES"
ONOMA1 = MERA & "/" & MHNAS & "/" & ETOS
'***ELEGXOS AN YPARXEI ANTIGRAFO GIA THN SHMERINH HMEROMHNIA*******************************
If RSH.BOF = RSH.EOF Then GoTo NIK:
RSH.MoveFirst
NIK:
Do While Not RSH.EOF
    If ONOMA1 <> RSH![глеяолгмиа_дглиоуяциас_амтицяажоу] Then
        RSH.MoveNext
    Else
        C = C + 1
        RSH.MoveNext
    End If
Loop
'*******PROGRAMATISMOS********************
If C <> 1 Then
    CT = "евете гдг дглиоуяцгсг амтицяажо етаияиым йаи тилокоциым циа тгм: " & ONOMA1 & ""
    MsgBox (CT), , "пяосовг"
Else
    'allagh toy date
    Dim date1
    If Day(Date) < 12 Then
        date1 = CDate(Day(Date) & " / " & Month(Date) & " / " & Year(Date))
    Else
        date1 = Date
    End If
    APO = App.Path & "\DATABASES\ETAIRIES.mdb"
    STO = App.Path & "\DATABASES\BACK_UPS\BACKUP_ETAIRIES\BACKUP_ETAIRIES\" & ONOMA & ".MDB"
    STO2 = "C:\WINDOWS\ORGANIZER_BACKUP\BACKUP_ETAIRIES\BACKUP_ETAIRIES\" & ONOMA & ".MDB"
    FileCopy APO, STO
    FileCopy APO, STO2
    
    '********** BACKUP FAKELON TXTS KAI ETAIRIES ******************************
    ' TXTS
    Dim FSO As New FileSystemObject
    Dim STO_PATH, APO_TXTS, STO_TXTS, STO2_TXTS As String
    STO_PATH = App.Path & "\databases\BACK_UPS\BACKUP_ETAIRIES\"
    
    APO_TXTS = App.Path & "\TXTS"
    STO_TXTS = STO_PATH & "BACKUP_ETAIRIES\TXTS\" & ONOMA & "_TXTS"
    STO2_TXTS = "C:\WINDOWS\ORGANIZER_BACKUP\BACKUP_ETAIRIES\" & _
    "BACKUP_ETAIRIES\TXTS\"
    FSO.CopyFolder APO_TXTS, STO_TXTS
    FSO.CopyFolder STO_TXTS, STO2_TXTS
    
    'ETAIRION
    Dim APO_ETAIRIES, STO_ETAIRIES, STO2_ETAIRIES As String
    APO_ETAIRIES = App.Path & "\ETAIRIES"
    
    STO_ETAIRIES = STO_PATH & "BACKUP_ETAIRIES\ETAIRIES\" & ONOMA
    
    STO2_ETAIRIES = "C:\WINDOWS\ORGANIZER_BACKUP\BACKUP_ETAIRIES\BACKUP_ETAIRIES\ETAIRIES\"
    
    FSO.CopyFolder APO_ETAIRIES, STO_ETAIRIES
    FSO.CopyFolder STO_ETAIRIES, STO2_ETAIRIES
    '***********************************************************
    
    STATEMENT = "INSERT INTO BACKUP_ETAIRIES(глеяолгмиа_дглиоуяциас_амтицяажоу,омола_аявеиоу,HMEROMHNIA)" & _
    "VALUES(" & _
    "'" & ONOMA1 & "'," & _
    "'" & ONOMA & "'," & _
    "'" & date1 & "')"
    DBH.Execute STATEMENT
    
    MsgBox ("г дглиоуяциа амтицяажоу етаияиым йаи тилокоциым ециме ле епитувиа"), , ""
End If
GoTo TELOS:

ER:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
If DBH.STATE = 1 Then DBH.Close
If RSH.STATE = 1 Then RSH.Close
End Sub

Private Sub Command9_Click()
On Error GoTo ER:

Dim APO_HMER, MEXRI_HMER As Date
Dim ETOS, ETOS_PLUS_1 As String
Dim TIMH_FLAG, TIMH_ETOS As Integer
Dim APO, STO, STO2 As String
Dim STATEMENT, STATEMENT1, STATEMENT_FLAG, STATEM, STATEM1 As String
Dim DILOSI1, SS1 As String
' H METABLHTH ETOS EINAI TO PROHGOYMENO ETOS
ETOS = CStr(Year(Date) - 1)
ETOS_PLUS_1 = CStr(Year(Date))
TIMH_ETOS = CInt(ETOS)
If MsgBox("хекете ма пяовыягсете се дглиоуяциа етгсиоу амтицяажоу циа то етос : " & ETOS & " ;", vbOKCancel, "пяосовг!") = vbOK Then
    If MsgBox("г диадийасиа аутг еимаи лг амтистяеьилг. еисте сицоуяои оти хекете ма пяовыягсете;", vbOKCancel, "пяосовг!") = vbOK Then
        
    Else
        GoTo TELOS:
    End If
Else
    GoTo TELOS:
End If

' DHMIOYRGIA SYNDESHS
Dim DBH As New ADODB.Connection
Dim RSH As New ADODB.Recordset
DBH.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & "\databases\BACK_UPS\HELP_BACKUP.mdb" & ";" & _
"Persist Security Info=False"
DBH.Open App.Path & "\databases\BACK_UPS\HELP_BACKUP.mdb"
RSH.Open "[BACKUP_YEAR_ETAIRIES]", DBH, adOpenDynamic, adLockBatchOptimistic
RSH.MoveFirst
' ANIXNEYSH TOY FLAG GIA TO ZHTOYMENO ETOS
Do While Not RSH.EOF
If RSH![ETOS] = TIMH_ETOS Then
    TIMH_FLAG = RSH![FLAG]
    RSH.MoveNext
Else
    RSH.MoveNext
End If
Loop
' AN HDH EXV ANTIGRAFO SXETIKO MYNHMA
If TIMH_FLAG = 1 Then
    DILOSI1 = "евете гдг дглиоуяцгсг амтицяажо етаияиым йаи тилокоциым циа то етос:" & CStr(ETOS)
    MsgBox (DILOSI1), , "пяосовг"
Else
'RITHMISEIS
'Shape1.Visible = True
'Label2.Visible = True
'ProgressBar1.Visible = True
BACKUP.MousePointer = 11
BACKUP.Command1.Enabled = False
BACKUP.Command2.Enabled = False
BACKUP.Command3.Enabled = False
BACKUP.Command4.Enabled = False
BACKUP.Command5.Enabled = False
BACKUP.Command6.Enabled = False
BACKUP.Command7.Enabled = False
BACKUP.Command8.Enabled = False
BACKUP.Command9.Enabled = False

' ******************************* AN DEN EXO TOTE DHMIOYRGO ANTIGRAFO ******************
'**************************************************************************************
' DHMIOYRGIA TOY PINAKA TZIROI_TEMP O OPOIOS PERIEXEI TOYS TZIROYS TOY ETOY POY KANO
'BACKUP P.X 2005
Dim APOD, MEXD As Date
APOD = CDate("1/1/" & ETOS)
MEXD = CDate("31/12/" & ETOS)
TZIROI_YP "DATABASES", "ETAIRIES", "TZIROI_TEMP", "ONOMATA_ETAIRION_ABCDEF", APOD, MEXD

'SYNENOSI TON PINAKON POY PERIEXOYN TOYS TZIROYS TOY ETOYS POY KANO BACKUP KAI TOY PROHGOYMENOY
'AYTOY P.X 2005 KAI 2004 TA APOTELESMATA PERIEXONTAI STON TZIROI_TEMP
JOIN_PIN "TZIROI", "TZIROI_TEMP", "DATABASES", "ETAIRIES"
'KANONOIKOPOIHSH TOY PINAKA POY PERIEXEI TOYS SYNOLIKOYS TZIROYS 2004+2005
KANON_PIN "TZIROI_TEMP", "DATABASES", "ETAIRIES"

' ANTIGRAFH ARXEIOY BASHS SE BACKUP STO DATABASES KAI STO WINDOWS
    APO = App.Path & "\DATABASES\ETAIRIES.mdb"
    STO = App.Path & "\DATABASES\BACK_UPS\BACKUP_ETAIRIES\BACKUP_ETAIRIES_ETOS\" & ETOS & "_ETAIRIES.MDB"
    FileCopy APO, STO
    
 '********** BACKUP FAKELON TXTS KAI ETAIRIES ******************************
    ' TXTS
    Dim FSO As New FileSystemObject
    Dim STO_PATH, APO_TXTS, STO_TXTS, STO2_TXTS As String
    STO_PATH = App.Path & "\databases\BACK_UPS\BACKUP_ETAIRIES\"
    
    APO_TXTS = App.Path & "\TXTS"
    STO_TXTS = STO_PATH & "BACKUP_ETAIRIES_ETOS\TXTS\" & "TXTS_" & ETOS
    STO2_TXTS = "C:\WINDOWS\ORGANIZER_BACKUP\BACKUP_ETAIRIES\" & _
    "BACKUP_ETAIRIES_ETOS\TXTS\"
    FSO.CopyFolder APO_TXTS, STO_TXTS
    FSO.CopyFolder STO_TXTS, STO2_TXTS
    
    'ETAIRION
    Dim APO_ETAIRIES, STO_ETAIRIES, STO2_ETAIRIES As String
    APO_ETAIRIES = App.Path & "\ETAIRIES"
    
    STO_ETAIRIES = STO_PATH & "BACKUP_ETAIRIES_ETOS\ETAIRIES\" & ETOS & "_ETAIRIES"
    
    STO2_ETAIRIES = "C:\WINDOWS\ORGANIZER_BACKUP\BACKUP_ETAIRIES\BACKUP_ETAIRIES_ETOS\ETAIRIES\"
   
    FSO.CopyFolder APO_ETAIRIES, STO_ETAIRIES
    FSO.CopyFolder STO_ETAIRIES, STO2_ETAIRIES
'***********************************************************
    
' DIAGRAFH EGRAFON EKTOS ETOYS POY KRATAO BACKUP, STHN BACKUP BASH
    If RSH.STATE = 1 Then RSH.Close
    If DBH.STATE = 1 Then DBH.Close
'****************** DIAGRAFH EGRAFON EKTOS ETOYS*******************************************
    ' SYNDESH ME THN DHMIOYRGOYMENH BASH
    APO_HMER = CDate("1/1/" & ETOS)
    MEXRI_HMER = CDate("31/12/" & ETOS)
    DBH.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
       "Data Source=" & "\databases\BACK_UPS\BACKUP_ETAIRIES\BACKUP_ETAIRIES_ETOS\" & ETOS & "_ETAIRIES.MDB" & ";" & _
      "Persist Security Info=False"
    DBH.Open App.Path & "\databases\BACK_UPS\BACKUP_ETAIRIES\BACKUP_ETAIRIES_ETOS\" & ETOS & "_ETAIRIES.MDB"
    RSH.Open "[ONOMATA_ETAIRION_ABCDEF]", DBH, adOpenDynamic, adLockBatchOptimistic
    If RSH.BOF = RSH.EOF Then GoTo NNN:
    RSH.MoveFirst
NNN:
    Do While Not RSH.EOF
        STATEM = " DELETE FROM " & RSH![омолата_етаияиым] & _
        " WHERE глеяолгмиа_ейдосгс <#" & APO_HMER & "#"
        DBH.Execute STATEM
        STATEM1 = " DELETE FROM " & RSH![омолата_етаияиым] & _
        " WHERE глеяолгмиа_ейдосгс >#" & MEXRI_HMER & "#"
        DBH.Execute STATEM1
        RSH.MoveNext
    Loop
' STO BACKUP EXO 2 PINAKES TOYS TZIROI KAI TZIROI_TEMP.O TEMP EINAI O 2004+2005 KAI DEN XREIAZETAI
    Dim JJJ As String
    JJJ = " DROP TABLE TZIROI_TEMP"
    DBH.Execute JJJ
    
'*****************************************************************
' DIAGRAFH EGRAFON ENTOS ETOYS POY KRATHSA BACKUP,STHN TREXON BASH
    If RSH.STATE = 1 Then RSH.Close
    If DBH.STATE = 1 Then DBH.Close
'*************************************************************
    ' SYNDESH ME THN TREXON BASH GIA DIAGRAFH EGRAFON POY KRATHSA BACKUP
    DBH.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
    "Data Source=" & "\databases\ETAIRIES.MDB" & ";" & _
    "Persist Security Info=False"
    DBH.Open App.Path & "\databases\ETAIRIES.MDB"
    RSH.Open "[ONOMATA_ETAIRION_ABCDEF]", DBH, adOpenDynamic, adLockBatchOptimistic
    If RSH.BOF = RSH.EOF Then GoTo mm:
    RSH.MoveFirst
mm:
    Do While Not RSH.EOF
        SS1 = " DELETE FROM " & RSH![омолата_етаияиым] & _
        " WHERE глеяолгмиа_ейдосгс >=#" & APO_HMER & "#" & _
        " AND глеяолгмиа_ейдосгс <=#" & MEXRI_HMER & "#"  ' PROSOXH !!!!!!!!!!!!!!!----
        DBH.Execute SS1
        RSH.MoveNext
    Loop

'STHN TREXON BASH EXO TON TZIROI POY ANTIPROSOPEYEI TO 2004 KAI TON TEMP 2004+2005
'O TZIROI DEN XREIAZETAI KAI SBHNETAI ENO O TZIROI_TEMP METONOMAZETAI SE TZIROI
    Dim JJJ1, JJJ2, JJJ3 As String
    JJJ1 = " DROP TABLE TZIROI"
    DBH.Execute JJJ1
    JJJ2 = " select * into TZIROI  from TZIROI_TEMP"
    JJJ3 = "DROP TABLE TZIROI_TEMP"
    DBH.Execute JJJ2
    DBH.Execute JJJ3
'*****************************************************************
   'KSANA ANOIGMA BOHTHITIKHS GIA TROPOIHSH FLAG
    If RSH.STATE = 1 Then RSH.Close
    If DBH.STATE = 1 Then DBH.Close
    DBH.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
       "Data Source=" & "\databases\BACK_UPS\HELP_BACKUP.mdb" & ";" & _
      "Persist Security Info=False"
    DBH.Open App.Path & "\databases\BACK_UPS\HELP_BACKUP.mdb"
    RSH.Open "[BACKUP_YEAR_ETAIRIES]", DBH, adOpenDynamic, adLockBatchOptimistic
    STATEMENT_FLAG = " UPDATE BACKUP_YEAR_ETAIRIES " & _
        " SET FLAG=1" & _
        " WHERE ETOS=" & TIMH_ETOS
    DBH.Execute STATEMENT_FLAG
'RITHMISEIS
    'Shape1.Visible = False
    'Label2.Visible = False
    'ProgressBar1.Visible = False
    BACKUP.MousePointer = 0
    BACKUP.Command1.Enabled = True
    BACKUP.Command2.Enabled = True
    BACKUP.Command3.Enabled = True
    BACKUP.Command4.Enabled = True
    BACKUP.Command5.Enabled = True
    BACKUP.Command6.Enabled = True
    BACKUP.Command7.Enabled = True
    BACKUP.Command8.Enabled = True
    BACKUP.Command9.Enabled = True
    MsgBox ("г дглиоуяциа амтицяажоу етаияиым йаи тилокоциым циа то етос: " & ETOS & " ециме ле епитувиа"), , ""
STO2 = "C:\WINDOWS\ORGANIZER_BACKUP\BACKUP_ETAIRIES\BACKUP_ETAIRIES_ETOS\" & ETOS & "_ETAIRIES.MDB"
FileCopy STO, STO2
End If
TOMH_PIN "ONOMATA_ETAIRION_ABCDEF", "TZIROI", "DATABASES", "ETAIRIES"

GoTo TELOS:

ER:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:

If RSH.STATE = 1 Then RSH.Close
If DBH.STATE = 1 Then DBH.Close
End Sub

Private Sub Form_Load()
On Error GoTo ER:
I = 0

Unload Form2
Unload Form3
Unload Form4
Unload Form5
Unload H
Unload HMEROLOGIO
Unload YPOL_TIMON
Unload YPOL_TIMON2
Unload PROXEIRO
Form1.Enabled = False
GoTo TELOS:

ER:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ER:
BACKUP.Hide
Unload BACKUP
Form1.Enabled = True
GoTo TELOS:

ER:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
End Sub

Private Sub Timer1_Timer()

I = I + 1
ProgressBar1.Value = I

If I = 10 Then
Timer1.Enabled = False
Timer2.Enabled = True
End If
End Sub

Private Sub Timer2_Timer()
I = I - 1
ProgressBar1.Value = I

If I = 0 Then
Timer1.Enabled = True
Timer2.Enabled = False
End If
End Sub
