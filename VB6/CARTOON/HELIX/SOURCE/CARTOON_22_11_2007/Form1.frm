VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00E0E0E0&
   Caption         =   "CARTOON"
   ClientHeight    =   13575
   ClientLeft      =   765
   ClientTop       =   555
   ClientWidth     =   17670
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   13575
   ScaleWidth      =   17670
   Begin VB.CommandButton Command9 
      Caption         =   "Command9"
      Height          =   735
      Left            =   3840
      TabIndex        =   62
      Top             =   10200
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox Text44 
      Height          =   615
      Left            =   2160
      TabIndex        =   61
      Text            =   "Text44"
      Top             =   9360
      Visible         =   0   'False
      Width           =   12975
   End
   Begin VB.TextBox Text43 
      Height          =   375
      Left            =   1680
      TabIndex        =   60
      Text            =   "Text43"
      Top             =   8040
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text9 
      Height          =   2895
      Left            =   2040
      TabIndex        =   59
      Text            =   "Text9"
      Top             =   6120
      Visible         =   0   'False
      Width           =   6375
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5760
      Top             =   10200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer4 
      Interval        =   60000
      Left            =   120
      Top             =   11040
   End
   Begin VB.TextBox Text30 
      Height          =   375
      Left            =   6360
      TabIndex        =   46
      Top             =   8880
      Width           =   11055
   End
   Begin VB.TextBox Text29 
      Height          =   285
      Left            =   120
      TabIndex        =   45
      Text            =   "Text29"
      Top             =   9360
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox Text28 
      Height          =   285
      Left            =   120
      TabIndex        =   44
      Text            =   "Text28"
      Top             =   9000
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox Text27 
      Height          =   375
      Left            =   120
      TabIndex        =   43
      Text            =   "Text27"
      Top             =   8520
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   120
      TabIndex        =   23
      Text            =   "Combo1"
      Top             =   4680
      Width           =   2175
   End
   Begin VB.TextBox Text10 
      Height          =   4215
      Left            =   6360
      ScrollBars      =   3  'Both
      TabIndex        =   22
      Top             =   9240
      Width           =   11055
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Command8"
      Height          =   495
      Left            =   1320
      TabIndex        =   21
      Top             =   10920
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Command7"
      Height          =   495
      Left            =   1080
      TabIndex        =   20
      Top             =   10320
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox Text8 
      Height          =   375
      Left            =   120
      TabIndex        =   19
      Text            =   "Text8"
      Top             =   8040
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Timer Timer2 
      Interval        =   60000
      Left            =   120
      Top             =   10560
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0C0&
      Caption         =   "CLEAR"
      Height          =   375
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   3600
      Width           =   975
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0C0&
      Caption         =   "DEFAULT"
      Height          =   375
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   3600
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "CLEAR"
      Height          =   375
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "DEFAULT"
      Height          =   375
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2520
      Width           =   975
   End
   Begin VB.TextBox Text7 
      Height          =   495
      Left            =   120
      TabIndex        =   14
      Top             =   3600
      Width           =   3255
   End
   Begin VB.TextBox Text6 
      Height          =   405
      Left            =   120
      TabIndex        =   11
      Top             =   2520
      Width           =   3255
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   1680
      Top             =   12120
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
      BackColor       =   &H00C0C0C0&
      Caption         =   "енодос"
      Height          =   855
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   11640
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   13560
      Top             =   120
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   15840
      TabIndex        =   7
      Text            =   "Text5"
      Top             =   120
      Width           =   1695
   End
   Begin VB.TextBox Text4 
      Height          =   525
      Left            =   1200
      TabIndex        =   4
      Top             =   1560
      Width           =   615
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   480
      TabIndex        =   3
      Top             =   1560
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "START"
      Height          =   735
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5400
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   525
      Left            =   1200
      TabIndex        =   1
      Top             =   600
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   600
      Width           =   615
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   8175
      Left            =   6360
      TabIndex        =   25
      Top             =   600
      Width           =   11055
      Begin VB.TextBox Text42 
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   360
         TabIndex        =   58
         Top             =   7440
         Width           =   10575
      End
      Begin VB.TextBox Text41 
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   360
         TabIndex        =   57
         Top             =   7080
         Width           =   10575
      End
      Begin VB.TextBox Text40 
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   360
         TabIndex        =   56
         Top             =   6720
         Width           =   10575
      End
      Begin VB.TextBox Text39 
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   360
         TabIndex        =   55
         Top             =   6360
         Width           =   10575
      End
      Begin VB.TextBox Text38 
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   360
         TabIndex        =   54
         Top             =   6000
         Width           =   10575
      End
      Begin VB.TextBox Text37 
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   360
         TabIndex        =   53
         Top             =   5640
         Width           =   10575
      End
      Begin VB.TextBox Text36 
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   360
         TabIndex        =   52
         Top             =   5280
         Width           =   10575
      End
      Begin VB.TextBox Text35 
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   360
         TabIndex        =   51
         Top             =   4920
         Width           =   10575
      End
      Begin VB.TextBox Text34 
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   360
         TabIndex        =   50
         Top             =   4560
         Width           =   10575
      End
      Begin VB.TextBox Text33 
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   360
         TabIndex        =   49
         Top             =   4200
         Width           =   10575
      End
      Begin VB.TextBox Text32 
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   360
         TabIndex        =   48
         Top             =   3840
         Width           =   10575
      End
      Begin VB.TextBox Text31 
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   360
         TabIndex        =   47
         Top             =   3480
         Width           =   10575
      End
      Begin VB.TextBox Text26 
         Height          =   375
         Left            =   7080
         TabIndex        =   41
         Top             =   2760
         Width           =   495
      End
      Begin VB.TextBox Text25 
         Height          =   375
         Left            =   7680
         TabIndex        =   40
         Top             =   2760
         Width           =   495
      End
      Begin VB.TextBox Text24 
         Height          =   285
         Left            =   360
         TabIndex        =   39
         Text            =   "то FTP амалеметаи ма текеиысеи стис"
         Top             =   2760
         Width           =   6375
      End
      Begin VB.Timer Timer3 
         Interval        =   500
         Left            =   7080
         Top             =   1800
      End
      Begin VB.TextBox Text23 
         Height          =   375
         Left            =   360
         TabIndex        =   38
         Text            =   "то ENCODING амалеметаи ма текеиысеи стис"
         Top             =   2280
         Width           =   6375
      End
      Begin VB.TextBox Text22 
         Height          =   375
         Left            =   7680
         TabIndex        =   37
         Top             =   2280
         Width           =   495
      End
      Begin VB.TextBox Text21 
         Height          =   375
         Left            =   7080
         TabIndex        =   36
         Top             =   2280
         Width           =   495
      End
      Begin VB.TextBox Text20 
         Height          =   375
         Left            =   360
         TabIndex        =   35
         Text            =   "Text20"
         Top             =   1800
         Width           =   6375
      End
      Begin VB.TextBox Text19 
         Height          =   375
         Left            =   7680
         TabIndex        =   34
         Top             =   1320
         Width           =   495
      End
      Begin VB.TextBox Text18 
         Height          =   375
         Left            =   7080
         TabIndex        =   33
         Top             =   1320
         Width           =   495
      End
      Begin VB.TextBox Text17 
         Height          =   375
         Left            =   360
         TabIndex        =   32
         Text            =   "то CAPTURE тоу AVI ха евеи текеиысеи стис"
         Top             =   1320
         Width           =   6375
      End
      Begin VB.TextBox Text16 
         Height          =   375
         Left            =   7680
         TabIndex        =   31
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox Text15 
         Height          =   405
         Left            =   7080
         TabIndex        =   30
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox Text14 
         Height          =   375
         Left            =   360
         TabIndex        =   29
         Text            =   "TO CAPTURE тоу AVI ха аявисеи стис"
         Top             =   840
         Width           =   6375
      End
      Begin VB.TextBox Text13 
         Height          =   405
         Left            =   7680
         TabIndex        =   28
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox Text12 
         Height          =   375
         Left            =   7080
         TabIndex        =   27
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox Text11 
         Height          =   375
         Left            =   360
         TabIndex        =   26
         Text            =   "екецвос оти г ежаялоцг тяевеи стис"
         Top             =   360
         Width           =   6375
      End
      Begin VB.Label Label8 
         BackColor       =   &H00E0E0E0&
         Caption         =   "г диадийасиа бяисйетаи се енекинг"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   360
         TabIndex        =   42
         Top             =   3120
         Width           =   7935
      End
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "INCLUDE RM FILES"
      Height          =   255
      Left            =   240
      TabIndex        =   24
      Top             =   4320
      Width           =   2175
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "PATH WHERE FILES ARE LOCATED"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   3240
      Width           =   3255
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "PATH WHERE AVI IS LOCATED"
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   2160
      Width           =   3135
   End
   Begin VB.Label Label4 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   3120
      TabIndex        =   10
      Top             =   120
      Width           =   10455
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   14160
      TabIndex        =   8
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "DURATION"
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "FROM"
      Height          =   255
      Left            =   600
      TabIndex        =   5
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Combo1_LostFocus()
If Combo1.Text <> "YES" Then Combo1.Text = "NO"
End Sub

Private Sub Command1_Click()
On Error GoTo er:
Dim STARTHOUR, STARTMINUTE, DURATIONHOUR, DURATIONMINUTE, ORA, LEPTA As Integer
Dim STATEMENT As String
Dim DURATIONENCODINGSUM, DURATIONENCODINGHOUR, DURATIONFTPHOUR, DURATIONFTPMINUTE As Integer
Dim DURATIONENCODINGMINUTE As Double
Dim S1, S2, S3, S4, S5, S6, S7, S8, S9, S10, S11, S12 As String
'****************** ELEGXOI ******************************************************
If Text1.Text = "" Then GoTo ER1
If Text2.Text = "" Then GoTo ER2
If Text3.Text = "" Then GoTo ER3
If Text4.Text = "" Then GoTo ER4
If IsNumeric(Text1.Text) <> True Then GoTo ER1
If IsNumeric(Text2.Text) <> True Then GoTo ER2
If IsNumeric(Text3.Text) <> True Then GoTo ER3
If IsNumeric(Text4.Text) <> True Then GoTo ER4
If Text1.Text < 0 Or Text1.Text > 23 Then GoTo ER1
If Text2.Text < 0 Or Text2.Text > 60 Then GoTo ER2
If Text3.Text < 0 Or Text3.Text > 23 Then GoTo ER3
If Text4.Text < 0 Or Text4.Text > 60 Then GoTo ER4
If Text3.Text = 0 And Text4.Text = 0 Then GoTo ER5

'*************** PROGRAMM ****************************

' EYRESH AN THA YPOLOGISO TA RM
If Combo1.Text <> "YES" Then Combo1.Text = "NO"
If Combo1.Text = "YES" Then
    FLAG_INCLUDE_RM = 1
Else
    FLAG_INCLUDE_RM = 0
End If

If Command1.Caption = "START" Then
    Frame1.Visible = True
    vapo_h = Text1.Text
    vapo_m = Text2.Text
    Text15.Text = vapo_h
    Text16.Text = vapo_m
    ' YPOLOGISMOS APO - 5
    If vapo_m - 5 >= 0 Then
        Vapo_5_h = vapo_h
        vapo_5_m = vapo_m - 5
    Else
        Vapo_5_h = vapo_h - 1
        vapo_5_m = 60 - (Abs(vapo_m - 5))
        If Vapo_5_h < 0 Then Vapo_5_h = 23
    End If
Text12.Text = Vapo_5_h
Text13.Text = vapo_5_m

  ' YPOLOGISMOS MEXRI
    STARTHOUR = CInt(Text1.Text)
    STARTMINUTE = CInt(Text2.Text)
    DURATIONHOUR = CInt(Text3.Text)
    DURATIONMINUTE = CInt(Text4.Text) + 2
    FINDHOUR CInt(Text1.Text), CInt(Text2.Text), CInt(Text3.Text), DURATIONMINUTE
    vmexri_h = TEMP_ORA
    vmexri_m = TEMP_MINUTE
    Text18.Text = vmexri_h
    Text19.Text = vmexri_m

' YPOLOGISA KAI BRHKA OTI GIA NA KANEI ENCODING 1m avi ARXEIO THELEI 0,4m. ARA THA EXO
'APO THN XR STIGMH MEXRI XREIAZETAI ((60*A + B)* 0,4)*4(MORFES ARXEION) + 10m ENA KENO ASFALEIAS.
    ' YPOLOGISMOS ENCODING
If FLAG_INCLUDE_RM = 1 Then
    DURATIONENCODINGSUM = CInt((((60 * DURATIONHOUR + DURATIONMINUTE) * 0.4) * 4) + 10)
    DURATIONENCODINGHOUR = Int(DURATIONENCODINGSUM / 60)
    DURATIONENCODINGMINUTE = DURATIONENCODINGSUM Mod 60
    FINDHOUR vmexri_h, vmexri_m, DURATIONENCODINGHOUR, DURATIONENCODINGMINUTE
    venco_h = TEMP_ORA
    venco_m = TEMP_MINUTE
    Text20.Text = "евете епикенеи ма дглиоуяцгхоум RM аявеиа"
    Text21.Text = venco_h
    Text22.Text = venco_m
Else
    DURATIONENCODINGSUM = CInt((((60 * DURATIONHOUR + DURATIONMINUTE) * 0.4) * 2) + 10)
    DURATIONENCODINGHOUR = Int(DURATIONENCODINGSUM / 60)
    DURATIONENCODINGMINUTE = DURATIONENCODINGSUM Mod 60
    FINDHOUR vmexri_h, vmexri_m, DURATIONENCODINGHOUR, DURATIONENCODINGMINUTE
    venco_h = TEMP_ORA
    venco_m = TEMP_MINUTE
    Text20.Text = "евете епикенеи ма лгм дглиоуяцгхоум RM аявеиа"
    Text21.Text = venco_h
    Text22.Text = venco_m
End If


'YPOLOGISMOS FTP DURATION
    ' karfoto
    DURATIONFTPHOUR = 0
    DURATIONFTPMINUTE = 1
    FINDHOUR venco_h, venco_m, DURATIONFTPHOUR, DURATIONFTPMINUTE
    vend_h = TEMP_ORA
    vend_m = TEMP_MINUTE
    Text26.Text = vend_h
    Text25.Text = vend_m
    AVI_PATH = Text6.Text
    FILES_PATH = Text7.Text

'******** INSERT SE BASH *******************************
    'S1 = "UPDATE CARTOON SET apo_5_h =" & Vapo_5_h & " WHERE ID=1"
    'S2 = "UPDATE CARTOON SET apo_5_m =" & vapo_5_m & " WHERE ID=1"
    'S3 = "UPDATE CARTOON SET apo_h =" & vapo_h & " WHERE ID=1"
    'S4 = "UPDATE CARTOON SET apo_m =" & vapo_m & " WHERE ID=1"
    'S5 = "UPDATE CARTOON SET mexri_h =" & vmexri_h & " WHERE ID=1"
    'S6 = "UPDATE CARTOON SET mexri_m = " & vmexri_m & " WHERE ID=1"
    'S7 = "UPDATE CARTOON SET enco_h =" & venco_h & " WHERE ID=1"
    'S8 = "UPDATE CARTOON SET enco_m =" & venco_m & " WHERE ID=1"
    'S9 = "UPDATE CARTOON SET end_h =" & vend_h & " WHERE ID=1"
    'S10 = "UPDATE CARTOON SET end_m=" & vend_m & " WHERE ID=1"
    'S11 = "UPDATE CARTOON SET PATHOF_AVI='" & Text6.Text & "' WHERE ID=1"
    'S12 = "UPDATE CARTOON SET PATHOF_FILES='" & Text7.Text & "' WHERE ID=1"
    'DB.Execute S1
    'DB.Execute S2
    'DB.Execute S3
    'DB.Execute S4
    'DB.Execute S5
    'DB.Execute S6
    'DB.Execute S7
    'DB.Execute S8
    'DB.Execute S9
    'DB.Execute S10
    'DB.Execute S11
    'DB.Execute S12
    Text1.Enabled = False
    Text2.Enabled = False
    Text3.Enabled = False
    Text4.Enabled = False
    Text6.Enabled = False
    Text7.Enabled = False
    Command3.Enabled = False
    Command4.Enabled = False
    Command5.Enabled = False
    Command6.Enabled = False
    Combo1.Enabled = False
    Timer2.Enabled = True
    Timer4.Enabled = True
    Command1.Caption = "STOP"
    Label4.Caption = "CARTOON NETWORK IS ON"
    Label4.ForeColor = &H0&
    Call Shell("C:\HELIX\CARTOON\MAILS\start.BAT", vbNormalFocus)
Else
    If MsgBox("пяойеите ма йкеисете то CARTOON NETWORK. еисте сицоуяои оти хекете ма пяовыягсете", vbOKCancel, "пяосовг !!!") = vbCancel Then
    GoTo TELOS
Else
    Text1.Enabled = True
    Text2.Enabled = True
    Text3.Enabled = True
    Text4.Enabled = True
    Text6.Enabled = True
    Text7.Enabled = True
    Combo1.Enabled = True
    Command3.Enabled = True
    Command4.Enabled = True
    Command5.Enabled = True
    Command6.Enabled = True
    Timer2.Enabled = False
    Timer4.Enabled = False
    Command1.Caption = "START"
    Label4.Caption = "CARTOON NETWORK IS NOT RUNNING      PLEASE ADD FROM AND DURATION  AND PRESS START     !!!!!!!!! "
    Label4.ForeColor = &HFF&
    Frame1.Visible = False
    ARXIKOPOIHSH2
    Command1.Caption = "START"
    'SEND MAIL THAT CARTOON STOP !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    Call Shell("C:\HELIX\CARTOON\MAILS\stop.BAT", vbNormalFocus)
End If
End If
GoTo TELOS:

'************ ANTIMETOPISIS ELEGXON ***************************
ER1:
MsgBox ("дем дысате сыста тгм ыяа емаянгс (педио ыяас)"), vbCritical, "пяосовг !!!"
Text1.Text = ""
GoTo TELOS:

ER2:
MsgBox ("дем дысате сыста тгм ыяа емаянгс (педио кептым)"), vbCritical, "пяосовг !!!"
Text2.Text = ""
GoTo TELOS:

ER3:
MsgBox ("дем дысате сыста тгм диаяйеиа (педио ыяас)"), vbCritical, "пяосовг !!!"
Text3.Text = ""
GoTo TELOS:

ER4:
MsgBox ("дем дысате сыста тгм диаяйеиа (педио кептым)"), vbCritical, "пяосовг !!!"
Text4.Text = ""
GoTo TELOS:

ER5:
MsgBox ("дем дысате сыста тгм диаяйеиа"), vbCritical, "пяосовг !!!"
Text3.Text = ""
Text4.Text = ""
GoTo TELOS:

er:
MsgBox ("йапоио амапамтево кахос елжамистгйе."), vbCritical, "пяосовг !!!"

TELOS:
End Sub



Private Sub Command11_Click()

End Sub

Private Sub Command12_Click()

End Sub

Private Sub Command13_Click()
 
End Sub

Private Sub Command14_Click()
   
End Sub

Private Sub Command15_Click()
 '*****************************   CHECK THAT 82 RM EXISTS *********************************************
            Text8.Text = ""
            Text8.Text = Dir(PATH82rm)
            If Text8.Text = "" Then
                EXIST_82RM_MAIL = "то 82Kbs RM аявеио евеи дглиоуяцгхеи.**********~~~"
                SIZE_82RM_MAIL = "то 82Kbs RM аявеио деивмеи ма лгм еимаи емтанеи. то лецехос тоу аявеиоу еимаи    отам то амалемолемо лецехос еимаи **********~~~"
                FLAG_82RM = 3
            Else
                EXIST_82RM_MAIL = "то 82Kbs RM аявеио евеи дглиоуяцгхеи."
                'CHECK THAT 82 RM SIZE IS CORRECT - BRHKA OTI GIA 1m AVI EXO  500<X<530
                Dim MY82RM_SIZE, MINSIZE_82RM, MAXSIZE_82RM As Double
                MINSIZE_82RM = 500 * ((60 * CInt(Text3.Text) + CInt(Text4.Text)))
                MAXSIZE_82RM = 530 * ((60 * CInt(Text3.Text) + CInt(Text4.Text)))
                MY25RM_SIZE = FileLen(PATH82rm)
                If (MY82RM_SIZE < MINSIZE_82RM) And (MY25RM_SIZE < MAXSIZE_82RM) Then
                    SIZE_82RM_MAIL = "то 82Kbs RM аявеио деивмеи ма еимаи емтанеи. то лецехос тоу аявеиоу еимаи    отам то амалемолемо лецехос еимаи  "
                    FLAG_82RM = 1
                Else
                     SIZE_82RM_MAIL = "то 82Kbs RM аявеио деивмеи ма лгм еимаи емтанеи. то лецехос тоу аявеиоу еимаи    отам то амалемолемо лецехос еимаи **********~~~"
                     FLAG_82RM = 2
                End If
            End If
End Sub



Private Sub Command10_Click()
    
End Sub

Private Sub Command2_Click()
On Error GoTo er:
If MsgBox("пяойеите ма йкеисете то CARTOON NETWORK. еисте сицоуяои оти хекете ма пяовыягсете", vbOKCancel, "пяосовг !!!") = vbCancel Then
    GoTo TELOS:
Else
    'SEND MAIL THAT CARTOON STOP !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    Call Shell("C:\HELIX\CARTOON\MAILS\down.BAT", vbNormalFocus)
    End
End If
GoTo TELOS:

er:
MsgBox ("йапоио амапамтево кахос елжамистгйе."), vbCritical, "пяосовг !!!"

TELOS:
End Sub

Private Sub Command3_Click()
Text6.Text = AVI_PATH_DEFAULT
End Sub

Private Sub Command4_Click()
Text6.Text = ""
End Sub

Private Sub Command5_Click()
Text7.Text = FILES_PATH_DEFAULT
End Sub

Private Sub Command6_Click()
Text7.Text = ""
End Sub

Private Sub Command7_Click()
Dim AVPATH, PATH263gp, PATH823gp, PATH25rm, PATH82rm As String
AVPATH = AVI_PATH & "\CAPTURE.AVI"
PATH263gp = FILES_PATH & "\26_2g.3gp"
PATH823gp = FILES_PATH & "\82_3g.3gp"
PATH25rm = FILES_PATH & "\25_2g.rm"
PATH82rm = FILES_PATH & "\82_3g.rm"
'MAIL INITIAL
If Hour(Time) = Vapo_5_h And Minute(Time) = vapo_5_m Then
    INITIAL_MESSAGE_MAIL_CAPTION = "CARTOON APPLICATION IS READY TO RUN"
    INITIAL_MESSAGE_MAIL = "CARTOON NETWORK CAPTURE WILL START AT" & Form1.Text1.Text & ":" & Form1.Text2.Text & "AND IT WILL CAPTURE FILM WITH DURATION " & Text3.Text & ":" & Text4.Text & ". THE WHOLE PROCUDURE WILL COMPLETED ON " & vend_h & ":" & vend_m
    'SEND MAIL
    SENDINITIALMAILNOW = 1
End If
' ENARKSH TOU SCRIPT
If Hour(Time) = vapo_h And Minute(Time) = vapo_m Then
    Call Shell("C:\CARTOON\START.bat", vbNormalFocus) 'IMPORTANT
    FLAGSTARTPROCEDURE = 1
    STARTPROCEDURE_MAIL = "г диадийасиа нейимгсе йамомийа~~~"
    Text31.Text = STARTPROCEDURE_MAIL
    'Wrap$ = Chr$(13) + Chr$(10)
    Text10.Text = Text10.Text & " " & STARTPROCEDURE_MAIL
    Label8.Visible = True
End If

If FLAGSTARTPROCEDURE = 1 Then
'****************************** AVI CHECK  **********************************8
        'CHECK THAT AVI EXISTS
    If Hour(Time) = vmexri_h And Minute(Time) = vmexri_m Then
        Text8.Text = ""
        Text8.Text = Dir(AVPATH)
        If Text8.Text = "" Then
            AVI_EXIST_MAIL = "то AVI аявеио дем евеи дглиоуяцгхеи.**********~~~"
            AVI_SIZE_MAIL = "**********~~~"
            FLAG_AVI = 3
        Else
            'CHECK THAT AVI SIZE IS CORRECT' me bash ta avi poy exo proekypse enas mesos oros 730 ~ 760 gia asfaleia that balo min 700MB kai'max 800mb      gia 1 minute
            AVI_EXIST_MAIL = "то AVI аявеио евеи дглиоуяцгхеи."
            Dim MYAVI_SIZE, MINSIZE_AVI, MAXSIZE_AVI As Double
            MINSIZE_AVI = 700000 * ((60 * CInt(Text3.Text) + CInt(Text4.Text)))
            MAXSIZE_AVI = 800000 * ((60 * CInt(Text3.Text) + CInt(Text4.Text)))
            MYAVI_SIZE = FileLen(AVPATH)
            If (MYAVI_SIZE > MINSIZE_AVI) And (MYAVI_SIZE < MAXSIZE_AVI) Then
                AVI_SIZE_MAIL = "то AVI аявеио деивмеи ма еимаи емтанеи. то лецехос тоу аявеиоу еимаи " & MYAVI_SIZE & _
                " отам то амалемолемо лецехос еимаи амалеса стис тилес " & MINSIZE_AVI & _
                " - " & MAXSIZE_AVI & " Kb"
                FLAG_AVI = 1
            Else
                AVI_SIZE_MAIL = "то AVI аявеио деивмеи ма MHN еимаи емтанеи. то лецехос тоу аявеиоу еимаи " & MYAVI_SIZE & _
                " отам то амалемолемо лецехос еимаи амалеса стис тилес " & MINSIZE_AVI & _
                " - " & MAXSIZE_AVI & " Kb **********~~~"
                FLAG_AVI = 2
            End If
            
        End If
        'Wrap$ = Chr$(13) + Chr$(10)
        Text10.Text = Text10.Text & " " & AVI_EXIST_MAIL & " " & AVI_SIZE_MAIL
        Text32.Text = AVI_EXIST_MAIL
        Text33.Text = AVI_SIZE_MAIL
    End If
    
    
'****************************  CHECK ENCODING FILES **************************************
    If Hour(Time) = venco_h And Minute(Time) = venco_m Then
        '************ 26 3GP CHECK ****************
        '26 3GP EXIST
        Text8.Text = ""
        Text8.Text = Dir(PATH263gp)
        If Text8.Text = "" Then
            EXIST_263GP_MAIL = "то 26Kbs 3GP аявеио дем евеи дглиоуяцгхеи.**********~~~"
            SIZE_263GP_MAIL = "**********~~~"
            FLAG_263GP = 3
        Else
            '3GP SIZR CORRECT   ' BRHKA OTI GIA 1m AVI EXO 235<X<265 KB
            EXIST_263GP_MAIL = "то 26Kbs 3GP аявеио евеи дглиоуяцгхеи."
            Dim MY263GP_SIZE, MINSIZE_263GP, MAXSIZE_263GP As Double
            MINSIZE_263GP = (235 / 1024) * ((60 * CInt(Text3.Text) + CInt(Text4.Text)))
            MAXSIZE_263GP = (265 / 1024) * ((60 * CInt(Text3.Text) + CInt(Text4.Text)))
            MY263GP_SIZE = (FileLen(PATH263gp) / 1024) / 1024
            If (MY263GP_SIZE > MINSIZE_263GP) And (MY263GP_SIZE < MAXSIZE_263GP) Then
                SIZE_263GP_MAIL = "то 26Kbs 3GP аявеио деивмеи ма еимаи емтанеи. то лецехос тоу аявеиоу еимаи " & _
                MY263GP_SIZE & " отам то амалемолемо лецехос еимаи амалеса стис тилес " & _
                MINSIZE_263GP & " - " & MAXSIZE_263GP & " MB"
                FLAG_263GP = 1
            Else
                SIZE_263GP_MAIL = "то 26Kbs 3GP аявеио деивмеи ма MHN еимаи емтанеи. то лецехос тоу аявеиоу еимаи " & MY263GP_SIZE & _
                " отам то амалемолемо лецехос еимаи амалеса стис тилес " & MINSIZE_263GP & _
                " - " & MAXSIZE_263GP & " MB**********~~~"
                FLAG_263GP = 2
            End If
        End If
        Wrap$ = Chr$(13) + Chr$(10)
        Text10.Text = Text10.Text & Wrap & EXIST_263GP_MAIL & Wrap & SIZE_263GP_MAIL
        Text34.Text = EXIST_263GP_MAIL
        Text35.Text = SIZE_263GP_MAIL
        '************ 82 3GP CHECK ****************
        '82 3GP EXIST
        Text27.Text = ""
        Text27.Text = Dir(PATH823gp)
        If Text27.Text = "" Then
            EXIST_823GP_MAIL = "то 82Kbs 3GP аявеио дем евеи дглиоуяцгхеи.**********~~~"
            SIZE_823GP_MAIL = "**********~~~"
            FLAG_823GP = 3
        Else
            EXIST_823GP_MAIL = "то 82Kbs 3GP аявеио евеи дглиоуяцгхеи."
            'CHECK THAT 82 3GP SIZE IS CORRECT - BRHKA OTI GIA 1m AVI EXO  540<X<570
            Dim MY823GP_SIZE, MINSIZE_823GP, MAXSIZE_823GP As Double
            MINSIZE_823GP = (540 / 1024) * ((60 * CInt(Text3.Text) + CInt(Text4.Text)))
            MAXSIZE_823GP = (570 / 1024) * ((60 * CInt(Text3.Text) + CInt(Text4.Text)))
            MY823GP_SIZE = (FileLen(PATH823gp) / 1024) / 1024
            If (MY823GP_SIZE > MINSIZE_823GP) And (MY823GP_SIZE < MAXSIZE_823GP) Then
                SIZE_823GP_MAIL = "то 82Kbs 3GP аявеио деивмеи ма еимаи емтанеи. то лецехос тоу аявеиоу еимаи " & MY823GP_SIZE & _
                " отам то амалемолемо лецехос еимаи амалеса стис тилес " & MINSIZE_823GP & _
                " - " & MAXSIZE_823GP & " MB"
                FLAG_823GP = 1
            Else
                SIZE_823GP_MAIL = "то 82Kbs 3GP аявеио деивмеи ма лгм еимаи емтанеи. то лецехос тоу аявеиоу еимаи " & MY823GP_SIZE & _
                " отам то амалемолемо лецехос еимаи амалеса стис тилес " & MINSIZE_823GP & _
                " - " & MAXSIZE_823GP & " MB**********~~~"
                FLAG_823GP = 2
            End If
        End If
        Wrap$ = Chr$(13) + Chr$(10)
        Text10.Text = Text10.Text & Wrap & EXIST_823GP_MAIL & Wrap & SIZE_823GP_MAIL
        Text36.Text = EXIST_823GP_MAIL
        Text37.Text = SIZE_823GP_MAIL
        ' ******** ELEGXOS AN THELO TA RM ************
        If FLAG_INCLUDE_RM = 1 Then
            '************ 25 RM CHECK ****************
                Text28.Text = ""
                Text28.Text = Dir(PATH25rm)
                If Text28.Text = "" Then
                    EXIST_25RM_MAIL = "то 25Kbs RM аявеио евеи дглиоуяцгхеи.**********~~~"
                    SIZE_25RM_MAIL = "**********~~~"
                    FLAG_25RM = 3
                Else
                    EXIST_25RM_MAIL = "то 25Kbs RM аявеио евеи дглиоуяцгхеи."
                    'CHECK THAT 25 RM SIZE IS CORRECT - BRHKA OTI GIA 1m AVI EXO  155<X<180
                    Dim MY25RM_SIZE, MINSIZE_25RM, MAXSIZE_25RM As Double
                    MINSIZE_25RM = (155 / 1024) * ((60 * CInt(Text3.Text) + CInt(Text4.Text)))
                    MAXSIZE_25RM = (180 / 1024) * ((60 * CInt(Text3.Text) + CInt(Text4.Text)))
                    MY25RM_SIZE = (FileLen(PATH25rm) / 1024) / 1024
                    If (MY25RM_SIZE > MINSIZE_25RM) And (MY25RM_SIZE < MAXSIZE_25RM) Then
                        SIZE_25RM_MAIL = "то 25Kbs RM аявеио деивмеи ма еимаи емтанеи. то лецехос тоу аявеиоу еимаи " & _
                        MY25RM_SIZE & " отам то амалемолемо лецехос еимаи амалеса стис тилес " & MINSIZE_25RM & _
                        " - " & MAXSIZE_25RM & " MB"
                        FLAG_25RM = 1
                    Else
                        SIZE_25RM_MAIL = "то 25Kbs RM аявеио деивмеи ма MHN еимаи емтанеи. то лецехос тоу аявеиоу еимаи " & _
                        MY25RM_SIZE & " отам то амалемолемо лецехос еимаи амалеса стис тилес " & MINSIZE_25RM & _
                        " - " & MAXSIZE_25RM & " MB**********~~~"
                        FLAG_25RM = 2
                    End If
                End If
                Wrap$ = Chr$(13) + Chr$(10)
                Text10.Text = Text10.Text & Wrap & EXIST_25RM_MAIL & Wrap & SIZE_25RM_MAIL
                Text38.Text = EXIST_25RM_MAIL
                Text39.Text = SIZE_25RM_MAIL
                '************ 82 RM CHECK ****************
                Text29.Text = ""
                Text29.Text = Dir(PATH82rm)
                If Text29.Text = "" Then
                    EXIST_82RM_MAIL = "то 82Kbs RM аявеио евеи дглиоуяцгхеи.**********~~~"
                    SIZE_82RM_MAIL = "**********~~~"
                    FLAG_82RM = 3
                Else
                    EXIST_82RM_MAIL = "то 82Kbs RM аявеио евеи дглиоуяцгхеи."
                    'CHECK THAT 82 RM SIZE IS CORRECT - BRHKA OTI GIA 1m AVI EXO  500<X<530
                    Dim MY82RM_SIZE, MINSIZE_82RM, MAXSIZE_82RM As Double
                    MINSIZE_82RM = (500 / 1024) * ((60 * CInt(Text3.Text) + CInt(Text4.Text)))
                    MAXSIZE_82RM = (530 / 1024) * ((60 * CInt(Text3.Text) + CInt(Text4.Text)))
                    MY82RM_SIZE = (FileLen(PATH82rm) / 1024) / 1024
                    If (MY82RM_SIZE > MINSIZE_82RM) And (MY25RM_SIZE < MAXSIZE_82RM) Then
                        SIZE_82RM_MAIL = "то 82Kbs RM аявеио деивмеи ма еимаи емтанеи. то лецехос тоу аявеиоу еимаи " & MY82RM_SIZE & _
                        " отам то амалемолемо лецехос еимаи  амалеса стис тилес " & MINSIZE_82RM & _
                        " - " & MAXSIZE_82RM & " MB"
                        FLAG_82RM = 1
                    Else
                        SIZE_82RM_MAIL = "то 82Kbs RM аявеио деивмеи ма MHN еимаи емтанеи. то лецехос тоу аявеиоу еимаи " & MY82RM_SIZE & _
                        " отам то амалемолемо лецехос еимаи  амалеса стис тилес " & MINSIZE_82RM & _
                        " - " & MAXSIZE_82RM & " MB**********~~~"
                        FLAG_82RM = FLAG_82RM = 2
                    End If
                End If
                Wrap$ = Chr$(13) + Chr$(10)
                Text10.Text = Text10.Text & Wrap & EXIST_82RM_MAIL & Wrap & SIZE_82RM_MAIL
                Text40.Text = EXIST_82RM_MAIL
                Text41.Text = SIZE_82RM_MAIL
            Else ' AN EXEI EPILEGEI NA MHN PAIZOYN TA RM
                EXIST_25RM_MAIL = "евете епикенеи то 25Kbs аявеио ма лгм дглиоуяцгхеи"
                SIZE_25RM_MAIL = "евете епикенеи то 25Kbs аявеио ма лгм дглиоуяцгхеи. йамемас екецвос циа то лецехос тоу"
                EXIST_82RM_MAIL = "евете епикенеи то 85Kbs аявеио ма лгм дглиоуяцгхеи"
                SIZE_28RM_MAIL = "евете епикенеи то 82Kbs аявеио ма лгм дглиоуяцгхеи. йамемас екецвос циа то лецехос тоу"
                FLAG_25RM = 1
                FLAG_82RM = 1
                Wrap$ = Chr$(13) + Chr$(10)
                Text10.Text = Text10.Text & Wrap & EXIST_25RM_MAIL & Wrap & SIZE_25RM_MAIL & _
                Wrap & EXIST_82RM_MAIL & Wrap & SIZE_28RM_MAIL
                Text38.Text = EXIST_25RM_MAIL
                Text39.Text = SIZE_25RM_MAIL
                Text40.Text = EXIST_82RM_MAIL
                Text41.Text = SIZE_28RM_MAIL
            End If
    End If

    'CHECK THAT FTP COMPLETED CORRECTLY (FIRST) AND DO ALL FINAL STEPS
    If Hour(Time) = vend_h And Minute(Time) = vend_m Then
        ' ALL NECESSARY STEPS REGARDING FTP
        Text42.Text = "то FTP ха пяепеи ма евеи окойкгяыхеи. текос диадийасиас"
        PREPARE_MAIL = 1
        Wrap$ = Chr$(13) + Chr$(10)
        Text10.Text = Text10.Text & Wrap & Text42.Text
    End If
    
End If 'TOY START PROCEDURE
TELOS:
End Sub



Private Sub Command8_Click()
' kanei ta ekshs. proton sent initial mail second check what caption toy put kai pote stelnei teliko mail.
'sto telos arxikopoiei oles tis metablhtes
' INITIAL
If SENDINITIALMAILNOW = 1 Then
    ' SEND TO INITIAL
    'INITIAL_MESSAGE_MAIL_CAPTION = "CARTOON APPLICATION IS READY TO RUN"
    'INITIAL_MESSAGE_MAIL =
    
    Text42.Text = "ok"
    SENDINITIALMAILNOW = 0
End If

'EMERGENCY
If FLAG_INCLUDE_RM = 0 Then
    If (FLAG_AVI = 3) Or (FLAG_263GP = 3) Or (FLAG_823GP = 3) Then
        MAIL_CAPTION = "CRITICAL ~~~~ то CARTOON дем доукеье сыста !!!!"
        Text30.Text = MAIL_CAPTION
        ARXIKOPOIHSH
        SENDMAILNOW = 1
        FLAGSTARTPROCEDURE = 0
    End If
 Else
    If (FLAG_AVI = 3) Or (FLAG_263GP = 3) Or (FLAG_823GP = 3) Or (FLAG_25RM = 3) Or (FLAG_82RM = 3) Then
        MAIL_CAPTION = "CRITICAL ~~~~ то CARTOON дем доукеье сыста !!!!"
        Text30.Text = MAIL_CAPTION
        ARXIKOPOIHSH
        SENDMAILNOW = 1
        FLAGSTARTPROCEDURE = 0
    End If
End If

'TELIKO MAIL
If PREPARE_MAIL = 1 Then
    If FLAG_INCLUDE_RM = 0 Then
        If (FLAG_AVI = 2) Or (FLAG_263GP = 2) Or (FLAG_823GP = 2) Then
            MAIL_CAPTION = "WARNING ~~~~ сто CARTOON йапоио(а) аявеио(а) жаиметаи ма лгм свглатистийе сыста."
            Text30.Text = MAIL_CAPTION
            ARXIKOPOIHSH
            SENDMAILNOW = 1
            FLAGSTARTPROCEDURE = 0
        Else
            MAIL_CAPTION = "SUCCESS ~~~~ то CARTOON доукеье сыста."
            Text30.Text = MAIL_CAPTION
            ARXIKOPOIHSH
            SENDMAILNOW = 1
            FLAGSTARTPROCEDURE = 0
        End If
    Else
        If (FLAG_AVI = 3) Or (FLAG_263GP = 3) Or (FLAG_823GP = 3) Or (FLAG_25RM = 3) Or (FLAG_82RM = 3) Then
            MAIL_CAPTION = "WARING ~~~~ сто CARTOON йапоио(а) аявеио(а) жаиметаи ма лгм свглатистийе сыста."
           Text30.Text = MAIL_CAPTION
            ARXIKOPOIHSH
            SENDMAILNOW = 1
            FLAGSTARTPROCEDURE = 0
        Else
            MAIL_CAPTION = "SUCCESS ~~~~ то CARTOON доукеье сыста."
            Text30.Text = MAIL_CAPTION
            ARXIKOPOIHSH
            SENDMAILNOW = 1
            FLAGSTARTPROCEDURE = 0
        End If
    End If
End If

If SENDMAILNOW = 1 Then
    Call Shell(App.Path & "\MAILS\mail.BAT", vbNormalFocus)
    'SEND TO MAIL ME CAPTION TEXT30 KAI BODY TEXT10
 '   ARXIKOPOIHSH
 '   SENDMAILNOW = 0
 '   Text10.Text = ""
'    Text30.Text = ""
End If

End Sub

Private Sub Command9_Click()

Call Shell(Text44.Text, vbNormalFocus)

End Sub

Private Sub Form_Load()
On Error GoTo er:
'**************************** ARXIKOPOIEISEIS ***********************************
ARXIKOPOIHSH   ' SYNARTHSH ARXIKOPOIHSHS
AVI_PATH_DEFAULT = "D:\HELIX\"

Dim MERA As String
If Day(Date) >= 1 And Day(Date) <= 9 Then
    MERA = "0" & Day(Date)
Else
    MERA = Day(Date)
End If
FILES_PATH_DEFAULT = "C:\HELIX\OUTPUT\" & MERA & "-" & Month(Date) & "-" & Year(Date)

Label4.Caption = "CARTOON NETWORK IS NOT RUNNING      PLEASE ADD FROM AND DURATION  AND PRESS START     !!!!!!!!! "
Label4.ForeColor = &HFF&
Frame1.Visible = False
Combo1.Text = "NO"
Combo1.AddItem "YES"
Combo1.AddItem "NO"
Text10.Text = ""
'DELETE_RECORDS = "DELETE * FROM CARTOON"
'INSERT_INITIAL_RECORDS = "INSERT INTO CARTOON (apo_5_h,apo_5_m,apo_h,apo_m,mexri_h,mexri_m,enco_h,enco_m,end_h,end_m,PATHOF_AVI,PATHOF_FILES,ID,PATHOF_AVI_DEFAULT,PATHOF_FILES_DEFAULT)" & _
    "VALUES ('','','','','','','','','','','','',1,'" & AVI_PATH_DEFAULT & "','" & FILES_PATH_DEFAULT & "')"
        
Dim STAT1, STAT2 As String
'STAT1 = "UPDATE CARTOON SET PATHOF_AVI_DEFAULT='" & AVI_PATH_DEFAULT & "' WHERE ID=1"
'STAT2 = "UPDATE CARTOON SET PATHOF_FILES_DEFAULT='" & FILES_PATH_DEFAULT & "' WHERE ID=1"
Label4.Enabled = True
Command1.Caption = "START"
'If RS.State = 1 Then RS.Close
'If DB.State = 1 Then DB.Close
'DB.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
'"Data Source=" & "\DB.mdb" & ";" & _
'"Persist Security Info=False"
'DB.Open App.Path & "\DB.mdb"
'RS.Open "[CARTOON]", DB, adOpenDynamic, adLockBatchOptimistic
'DB.Execute DELETE_RECORDS
'DB.Execute INSERT_INITIAL_RECORDS
'DB.Execute STAT1
'DB.Execute STAT2
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text6.Text = AVI_PATH_DEFAULT
Text7.Text = FILES_PATH_DEFAULT


'SEND MAIL THAT CARTOON OPENED !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call Shell("c:\HELIX\CARTOON\MAILS\UP.BAT", vbNormalFocus)

GoTo TELOS:

er:
MsgBox ("йапоио амапамтево кахос елжамистгйе."), vbCritical, "пяосовг !!!"

TELOS:
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo er:
If MsgBox("пяойеите ма йкеисете то CARTOON NETWORK. еисте сицоуяои оти хекете ма пяовыягсете", vbOKCancel, "пяосовг !!!") = vbCancel Then
    GoTo TELOS:
Else
    'SEND MAIL THAT CARTOON STOP !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    Call Shell("C:\HELIX\CARTOON\MAILS\down.BAT", vbNormalFocus)
    End
End If
GoTo TELOS:

er:
MsgBox ("йапоио амапамтево кахос елжамистгйе."), vbCritical, "пяосовг !!!"

TELOS:
End Sub

Private Sub Image1_Click()

End Sub

Private Sub Text1_LostFocus()
If Text1.Text = "0" Then Text1.Text = "00"
If Text1.Text = "1" Then Text1.Text = "01"
If Text1.Text = "2" Then Text1.Text = "02"
If Text1.Text = "3" Then Text1.Text = "03"
If Text1.Text = "4" Then Text1.Text = "04"
If Text1.Text = "5" Then Text1.Text = "05"
If Text1.Text = "6" Then Text1.Text = "06"
If Text1.Text = "7" Then Text1.Text = "07"
If Text1.Text = "8" Then Text1.Text = "08"
If Text1.Text = "9" Then Text1.Text = "09"
End Sub

Private Sub Text2_LostFocus()
If Text2.Text = "0" Then Text2.Text = "00"
If Text2.Text = "1" Then Text2.Text = "01"
If Text2.Text = "2" Then Text2.Text = "02"
If Text2.Text = "3" Then Text2.Text = "03"
If Text2.Text = "4" Then Text2.Text = "04"
If Text2.Text = "5" Then Text2.Text = "05"
If Text2.Text = "6" Then Text2.Text = "06"
If Text2.Text = "7" Then Text2.Text = "07"
If Text2.Text = "8" Then Text2.Text = "08"
If Text2.Text = "9" Then Text2.Text = "09"
End Sub

Private Sub Text3_LostFocus()
If Text3.Text = "0" Then Text3.Text = "00"
If Text3.Text = "1" Then Text3.Text = "01"
If Text3.Text = "2" Then Text3.Text = "02"
If Text3.Text = "3" Then Text3.Text = "03"
If Text3.Text = "4" Then Text3.Text = "04"
If Text3.Text = "5" Then Text3.Text = "05"
If Text3.Text = "6" Then Text3.Text = "06"
If Text3.Text = "7" Then Text3.Text = "07"
If Text3.Text = "8" Then Text3.Text = "08"
If Text3.Text = "9" Then Text3.Text = "09"
End Sub

Private Sub Text4_LostFocus()
If Text4.Text = "0" Then Text4.Text = "00"
If Text4.Text = "1" Then Text4.Text = "01"
If Text4.Text = "2" Then Text4.Text = "02"
If Text4.Text = "3" Then Text4.Text = "03"
If Text4.Text = "4" Then Text4.Text = "04"
If Text4.Text = "5" Then Text4.Text = "05"
If Text4.Text = "6" Then Text4.Text = "06"
If Text4.Text = "7" Then Text4.Text = "07"
If Text4.Text = "8" Then Text4.Text = "08"
If Text4.Text = "9" Then Text4.Text = "09"
End Sub

Private Sub Timer1_Timer()
Label3.Caption = Date
Text5.Text = Time
End Sub

Private Sub Timer2_Timer()
'On Error GoTo er:


'MAIL INITIAL
If Hour(Time) = Vapo_5_h And Minute(Time) = vapo_5_m Then
    INITIAL_MESSAGE_MAIL_CAPTION = "CARTOON APPLICATION IS READY TO RUN"
    INITIAL_MESSAGE_MAIL = "CARTOON NETWORK CAPTURE WILL START AT" & Form1.Text1.Text & ":" & Form1.Text2.Text & "AND IT WILL CAPTURE FILM WITH DURATION " & Text3.Text & ":" & Text4.Text & ". THE WHOLE PROCUDURE WILL COMPLETED ON " & vend_h & ":" & vend_m
    'SEND MAIL
    SENDINITIALMAILNOW = 1
End If
' ENARKSH TOU SCRIPT
If Hour(Time) = vapo_h And Minute(Time) = vapo_m Then
    If FLAG_INCLUDE_RM = 0 Then
        Call Shell("C:\HELIX\START_NORM.bat", vbNormalFocus) 'IMPORTANT
    Else
        Call Shell("C:\HELIX\START_RM.bat", vbNormalFocus) 'IMPORTANT
    End If
    FLAGSTARTPROCEDURE = 1
    STARTPROCEDURE_MAIL = "PROCEDURE START CORRECTLY~~~"
    Text31.Text = STARTPROCEDURE_MAIL
    'Wrap$ = Chr$(13) + Chr$(10)
    Text10.Text = Text10.Text & " " & STARTPROCEDURE_MAIL
    Label8.Visible = True
End If

If FLAGSTARTPROCEDURE = 1 Then
'****************************** AVI CHECK  **********************************8
        'CHECK THAT AVI EXISTS
    If Hour(Time) = vmexri_h And Minute(Time) = vmexri_m Then
        Dim AVPATH, MERA As String
        If Day(Date) >= 1 And Day(Date) <= 9 Then
            MERA = "0" & Day(Date)
        Else
            MERA = Day(Date)
        End If
        AVPATH = AVI_PATH & MERA & "-" & Month(Date) & "-" & Year(Date) & ".avi"
        Text8.Text = ""
        Text8.Text = Dir(AVPATH)
        Text43.Text = AVPATH
        If Text8.Text = "" Then
            AVI_EXIST_MAIL = "AVI FILE HAS NOT BEEN CREATED.**********~~~"
            AVI_SIZE_MAIL = "**********~~~"
            FLAG_AVI = 3
        Else
            'CHECK THAT AVI SIZE IS CORRECT' me bash ta avi poy exo proekypse enas mesos oros 730 ~ 760 gia asfaleia that balo min 700MB kai'max 800mb      gia 1 minute
            AVI_EXIST_MAIL = "AVI FILE EXISTS."
            Dim MYAVI_SIZE, MINSIZE_AVI, MAXSIZE_AVI As Double
            MINSIZE_AVI = 700000 * ((60 * CInt(Text3.Text) + CInt(Text4.Text)))
            MAXSIZE_AVI = 800000 * ((60 * CInt(Text3.Text) + CInt(Text4.Text)))
            MYAVI_SIZE = FileLen(AVPATH)
            If (MYAVI_SIZE > MINSIZE_AVI) And (MYAVI_SIZE < MAXSIZE_AVI) Then
                AVI_SIZE_MAIL = "AVI FILE SEEMS TO BE OK. SIZE OF THE FILE IS " & MYAVI_SIZE & _
                " WHEN EXPECTED SIZE IS BETWEEN " & MINSIZE_AVI & _
                " ~ " & MAXSIZE_AVI & " Kb"
                FLAG_AVI = 1
            Else
                AVI_SIZE_MAIL = "AVI FILE SEEMS NOT TO BE OK. SIZE OF THE FILE IS " & MYAVI_SIZE & _
                " WHEN EXPECTED SIZE IS BETWEEN " & MINSIZE_AVI & _
                " ~ " & MAXSIZE_AVI & " Kb **********~~~"
                FLAG_AVI = 2
            End If
            
        End If
        'Wrap$ = Chr$(13) + Chr$(10)
        Text10.Text = Text10.Text & " " & AVI_EXIST_MAIL & " " & AVI_SIZE_MAIL
        Text32.Text = AVI_EXIST_MAIL
        Text33.Text = AVI_SIZE_MAIL
    End If
    
    
'****************************  CHECK ENCODING FILES **************************************
    If Hour(Time) = venco_h And Minute(Time) = venco_m Then
        '************ 26 3GP CHECK ****************
        Dim PATH263gp, PATH823gp As String
        PATH263gp = FILES_PATH & "\26_2g.3gp"
        PATH823gp = FILES_PATH & "\82_3g.3gp"
        '26 3GP EXIST
        Text8.Text = ""
        Text8.Text = Dir(PATH263gp)
        If Text8.Text = "" Then
            EXIST_263GP_MAIL = "то 26Kbs 3GP HAS NOT BEEN CREATED.**********~~~"
            SIZE_263GP_MAIL = "**********~~~"
            FLAG_263GP = 3
        Else
            '3GP SIZR CORRECT   ' BRHKA OTI GIA 1m AVI EXO 235<X<265 KB
            EXIST_263GP_MAIL = "то 26Kbs 3GP HAS BEEN CREATED."
            Dim MY263GP_SIZE, MINSIZE_263GP, MAXSIZE_263GP As Double
            MINSIZE_263GP = (235 / 1024) * ((60 * CInt(Text3.Text) + CInt(Text4.Text)))
            MAXSIZE_263GP = (265 / 1024) * ((60 * CInt(Text3.Text) + CInt(Text4.Text)))
            MY263GP_SIZE = (FileLen(PATH263gp) / 1024) / 1024
            If (MY263GP_SIZE > MINSIZE_263GP) And (MY263GP_SIZE < MAXSIZE_263GP) Then
                SIZE_263GP_MAIL = "то 26Kbs 3GP SEEMS TO BE OK. THE SIZE OF THE FILE IS " & _
                MY263GP_SIZE & " WHEN EXPECTED SIZE IS BETWEEN " & _
                MINSIZE_263GP & " ~ " & MAXSIZE_263GP & " MB"
                FLAG_263GP = 1
            Else
                SIZE_263GP_MAIL = "то 26Kbs 3GP SEEMS NOT TO BE OK. THE SIZE OF THE FILE IS " & MY263GP_SIZE & _
                " WHEN EXPECTED SIZE IS BETWEEN " & MINSIZE_263GP & _
                " ~ " & MAXSIZE_263GP & " MB**********~~~"
                FLAG_263GP = 2
            End If
        End If
        'Wrap$ = Chr$(13) + Chr$(10)
        Text10.Text = Text10.Text & " " & EXIST_263GP_MAIL & " " & SIZE_263GP_MAIL
        Text34.Text = EXIST_263GP_MAIL
        Text35.Text = SIZE_263GP_MAIL
        '************ 82 3GP CHECK ****************
        '82 3GP EXIST
        Text27.Text = ""
        Text27.Text = Dir(PATH823gp)
        If Text27.Text = "" Then
            EXIST_823GP_MAIL = "то 82Kbs 3GP HAS NOT BEEN CREATED.**********~~~"
            SIZE_823GP_MAIL = "**********~~~"
            FLAG_823GP = 3
        Else
            EXIST_823GP_MAIL = "то 82Kbs 3GP HAS BEEN CREATED."
            'CHECK THAT 82 3GP SIZE IS CORRECT - BRHKA OTI GIA 1m AVI EXO  540<X<570
            Dim MY823GP_SIZE, MINSIZE_823GP, MAXSIZE_823GP As Double
            MINSIZE_823GP = (540 / 1024) * ((60 * CInt(Text3.Text) + CInt(Text4.Text)))
            MAXSIZE_823GP = (570 / 1024) * ((60 * CInt(Text3.Text) + CInt(Text4.Text)))
            MY823GP_SIZE = (FileLen(PATH823gp) / 1024) / 1024
            If (MY823GP_SIZE > MINSIZE_823GP) And (MY823GP_SIZE < MAXSIZE_823GP) Then
                SIZE_823GP_MAIL = "то 82Kbs 3GP SEEMS TO BE OK. THE SIZE OF THE FILE IS " & MY823GP_SIZE & _
                " WHEN EXPECTED SIZE IS BETWEEN " & MINSIZE_823GP & _
                " ~ " & MAXSIZE_823GP & " MB"
                FLAG_823GP = 1
            Else
                SIZE_823GP_MAIL = "то 82Kbs 3GP аявеио деивмеи ма лгм еимаи емтанеи. то лецехос тоу аявеиоу еимаи " & MY823GP_SIZE & _
                " WHEN EXPECTED SIZE IS BETWEEN " & MINSIZE_823GP & _
                " ~ " & MAXSIZE_823GP & " MB**********~~~"
                FLAG_823GP = 2
            End If
        End If
        'Wrap$ = Chr$(13) + Chr$(10)
        Text10.Text = Text10.Text & " " & EXIST_823GP_MAIL & " " & SIZE_823GP_MAIL
        Text36.Text = EXIST_823GP_MAIL
        Text37.Text = SIZE_823GP_MAIL
        ' ******** ELEGXOS AN THELO TA RM ************
        If FLAG_INCLUDE_RM = 1 Then
            '************ 25 RM CHECK ****************
                Dim PATH25rm, PATH82rm As String
                PATH25rm = FILES_PATH & "\25_2g.rm"
                PATH82rm = FILES_PATH & "\82_3g.rm"
                Text28.Text = ""
                Text28.Text = Dir(PATH25rm)
                If Text28.Text = "" Then
                    EXIST_25RM_MAIL = "то 25Kbs RM HAS NOT BEEN CREATED.**********~~~"
                    SIZE_25RM_MAIL = "**********~~~"
                    FLAG_25RM = 3
                Else
                    EXIST_25RM_MAIL = "то 25Kbs RM HAS BEEN CREATED."
                    'CHECK THAT 25 RM SIZE IS CORRECT - BRHKA OTI GIA 1m AVI EXO  155<X<180
                    Dim MY25RM_SIZE, MINSIZE_25RM, MAXSIZE_25RM As Double
                    MINSIZE_25RM = (155 / 1024) * ((60 * CInt(Text3.Text) + CInt(Text4.Text)))
                    MAXSIZE_25RM = (180 / 1024) * ((60 * CInt(Text3.Text) + CInt(Text4.Text)))
                    MY25RM_SIZE = (FileLen(PATH25rm) / 1024) / 1024
                    If (MY25RM_SIZE > MINSIZE_25RM) And (MY25RM_SIZE < MAXSIZE_25RM) Then
                        SIZE_25RM_MAIL = "то 25Kbs RM SEEMS TO BE OK. THE SIZE OF THE FILE IS " & _
                        MY25RM_SIZE & " WHEN EXPECTED SIZE IS BETWEEN " & MINSIZE_25RM & _
                        " ~ " & MAXSIZE_25RM & " MB"
                        FLAG_25RM = 1
                    Else
                        SIZE_25RM_MAIL = "то 25Kbs RM SEEMS NOT TO BE OK. THEE SIZE OF THE FILE IS " & _
                        MY25RM_SIZE & " WHEN EXPECTED SIZE IS BETWEEN " & MINSIZE_25RM & _
                        " ~ " & MAXSIZE_25RM & " MB**********~~~"
                        FLAG_25RM = 2
                    End If
                End If
                'Wrap$ = Chr$(13) + Chr$(10)
                Text10.Text = Text10.Text & " " & EXIST_25RM_MAIL & " " & SIZE_25RM_MAIL
                Text38.Text = EXIST_25RM_MAIL
                Text39.Text = SIZE_25RM_MAIL
                '************ 82 RM CHECK ****************
                Text29.Text = ""
                Text29.Text = Dir(PATH82rm)
                If Text29.Text = "" Then
                    EXIST_82RM_MAIL = "то 82Kbs RM HAS NOT BEEN CREATED.**********~~~"
                    SIZE_82RM_MAIL = "**********~~~"
                    FLAG_82RM = 3
                Else
                    EXIST_82RM_MAIL = "то 82Kbs RM HAS BEEN CREATED."
                    'CHECK THAT 82 RM SIZE IS CORRECT - BRHKA OTI GIA 1m AVI EXO  500<X<530
                    Dim MY82RM_SIZE, MINSIZE_82RM, MAXSIZE_82RM As Double
                    MINSIZE_82RM = (500 / 1024) * ((60 * CInt(Text3.Text) + CInt(Text4.Text)))
                    MAXSIZE_82RM = (530 / 1024) * ((60 * CInt(Text3.Text) + CInt(Text4.Text)))
                    MY82RM_SIZE = (FileLen(PATH82rm) / 1024) / 1024
                    If (MY82RM_SIZE > MINSIZE_82RM) And (MY25RM_SIZE < MAXSIZE_82RM) Then
                        SIZE_82RM_MAIL = "то 82Kbs RM FILE SEEMS TO BE OK. THE SIZE OF THE FILE IS " & MY82RM_SIZE & _
                        " WHEN EXPECTED SIZE IS BETWEEN" & MINSIZE_82RM & _
                        " ~ " & MAXSIZE_82RM & " MB"
                        FLAG_82RM = 1
                    Else
                        SIZE_82RM_MAIL = "то 82Kbs RM FILE SEEMS NOT TO BE OK. THE SIZE OF THE FILE IS " & MY82RM_SIZE & _
                        " WHEN EXPECTED SIZE IS BETWEEN " & MINSIZE_82RM & _
                        " ~ " & MAXSIZE_82RM & " MB**********~~~"
                        FLAG_82RM = FLAG_82RM = 2
                    End If
                End If
                'Wrap$ = Chr$(13) + Chr$(10)
                Text10.Text = Text10.Text & " " & EXIST_82RM_MAIL & " " & SIZE_82RM_MAIL
                Text40.Text = EXIST_82RM_MAIL
                Text41.Text = SIZE_82RM_MAIL
            Else ' AN EXEI EPILEGEI NA MHN PAIZOYN TA RM
                EXIST_25RM_MAIL = "YOU HAVE CHOOSEN 25Kbs NOT TO BE CREATED"
                SIZE_25RM_MAIL = "YOU HAVE CHOOSEN 25Kbs NOT TO BE CREATED. NO CHECK FOR THE SIZE"
                EXIST_82RM_MAIL = "YOU HAVE CHOOSEN 85Kbs NOT TO BE CREATED"
                SIZE_28RM_MAIL = "YOU HAVE CHOOSEN 82Kbs NOT TO BE CREATED. NO CHECK FOR THE SIZE"
                FLAG_25RM = 1
                FLAG_82RM = 1
                'Wrap$ = Chr$(13) + Chr$(10)
                Text10.Text = Text10.Text & " " & EXIST_25RM_MAIL & " " & SIZE_25RM_MAIL & _
                " " & EXIST_82RM_MAIL & " " & SIZE_28RM_MAIL
                Text38.Text = EXIST_25RM_MAIL
                Text39.Text = SIZE_25RM_MAIL
                Text40.Text = EXIST_82RM_MAIL
                Text41.Text = SIZE_28RM_MAIL
            End If
    End If

    'CHECK THAT FTP COMPLETED CORRECTLY (FIRST) AND DO ALL FINAL STEPS
    If Hour(Time) = vend_h And Minute(Time) = vend_m Then
        ' ALL NECESSARY STEPS REGARDING FTP
        Text42.Text = "FTP SHOULD BE COMPLETED. END OF PROCESS"
        PREPARE_MAIL = 1
        'Wrap$ = Chr$(13) + Chr$(10)
        Text10.Text = Text10.Text & " " & Text42.Text
    End If
    
End If 'TOY START PROCEDURE
GoTo TELOS:

er:
MsgBox ("йапоио амапамтево кахос елжамистгйе."), vbCritical, "пяосовг !!!"


TELOS:
End Sub

Private Sub Timer3_Timer()


If Text13.Text = "0" Then Text13.Text = "00"
If Text13.Text = "1" Then Text13.Text = "01"
If Text13.Text = "2" Then Text13.Text = "02"
If Text13.Text = "3" Then Text13.Text = "03"
If Text13.Text = "4" Then Text13.Text = "04"
If Text13.Text = "5" Then Text13.Text = "05"
If Text13.Text = "6" Then Text13.Text = "06"
If Text13.Text = "7" Then Text13.Text = "07"
If Text13.Text = "8" Then Text13.Text = "08"
If Text13.Text = "9" Then Text13.Text = "09"

If Text16.Text = "0" Then Text16.Text = "00"
If Text16.Text = "1" Then Text16.Text = "01"
If Text16.Text = "2" Then Text16.Text = "02"
If Text16.Text = "3" Then Text16.Text = "03"
If Text16.Text = "4" Then Text16.Text = "04"
If Text16.Text = "5" Then Text16.Text = "05"
If Text16.Text = "6" Then Text16.Text = "06"
If Text16.Text = "7" Then Text16.Text = "07"
If Text16.Text = "8" Then Text16.Text = "08"
If Text16.Text = "9" Then Text16.Text = "09"

If Text19.Text = "0" Then Text19.Text = "00"
If Text19.Text = "1" Then Text19.Text = "01"
If Text19.Text = "2" Then Text19.Text = "02"
If Text19.Text = "3" Then Text19.Text = "03"
If Text19.Text = "4" Then Text19.Text = "04"
If Text19.Text = "5" Then Text19.Text = "05"
If Text19.Text = "6" Then Text19.Text = "06"
If Text19.Text = "7" Then Text19.Text = "07"
If Text19.Text = "8" Then Text19.Text = "08"
If Text19.Text = "9" Then Text19.Text = "09"

If Text22.Text = "0" Then Text22.Text = "00"
If Text22.Text = "1" Then Text22.Text = "01"
If Text22.Text = "2" Then Text22.Text = "02"
If Text22.Text = "3" Then Text22.Text = "03"
If Text22.Text = "4" Then Text22.Text = "04"
If Text22.Text = "5" Then Text22.Text = "05"
If Text22.Text = "6" Then Text22.Text = "06"
If Text22.Text = "7" Then Text22.Text = "07"
If Text22.Text = "8" Then Text22.Text = "08"
If Text22.Text = "9" Then Text22.Text = "09"

If Text12.Text = "0" Then Text12.Text = "00"
If Text12.Text = "1" Then Text12.Text = "01"
If Text12.Text = "2" Then Text12.Text = "02"
If Text12.Text = "3" Then Text12.Text = "03"
If Text12.Text = "4" Then Text12.Text = "04"
If Text12.Text = "5" Then Text12.Text = "05"
If Text12.Text = "6" Then Text12.Text = "06"
If Text12.Text = "7" Then Text12.Text = "07"
If Text12.Text = "8" Then Text12.Text = "08"
If Text12.Text = "9" Then Text12.Text = "09"

If Text15.Text = "0" Then Text15.Text = "00"
If Text15.Text = "1" Then Text15.Text = "01"
If Text15.Text = "2" Then Text15.Text = "02"
If Text15.Text = "3" Then Text15.Text = "03"
If Text15.Text = "4" Then Text15.Text = "04"
If Text15.Text = "5" Then Text15.Text = "05"
If Text15.Text = "6" Then Text15.Text = "06"
If Text15.Text = "7" Then Text15.Text = "07"
If Text15.Text = "8" Then Text15.Text = "08"
If Text15.Text = "9" Then Text15.Text = "09"

If Text18.Text = "0" Then Text18.Text = "00"
If Text18.Text = "1" Then Text18.Text = "01"
If Text18.Text = "2" Then Text18.Text = "02"
If Text18.Text = "3" Then Text18.Text = "03"
If Text18.Text = "4" Then Text18.Text = "04"
If Text18.Text = "5" Then Text18.Text = "05"
If Text18.Text = "6" Then Text18.Text = "06"
If Text18.Text = "7" Then Text18.Text = "07"
If Text18.Text = "8" Then Text18.Text = "08"
If Text18.Text = "9" Then Text18.Text = "09"

If Text21.Text = "0" Then Text21.Text = "00"
If Text21.Text = "1" Then Text21.Text = "01"
If Text21.Text = "2" Then Text21.Text = "02"
If Text21.Text = "3" Then Text21.Text = "03"
If Text21.Text = "4" Then Text21.Text = "04"
If Text21.Text = "5" Then Text21.Text = "05"
If Text21.Text = "6" Then Text21.Text = "06"
If Text21.Text = "7" Then Text21.Text = "07"
If Text21.Text = "8" Then Text21.Text = "08"
If Text21.Text = "9" Then Text21.Text = "09"

If Text25.Text = "0" Then Text25.Text = "00"
If Text25.Text = "1" Then Text25.Text = "01"
If Text25.Text = "2" Then Text25.Text = "02"
If Text25.Text = "3" Then Text25.Text = "03"
If Text25.Text = "4" Then Text25.Text = "04"
If Text25.Text = "5" Then Text25.Text = "05"
If Text25.Text = "6" Then Text25.Text = "06"
If Text25.Text = "7" Then Text25.Text = "07"
If Text25.Text = "8" Then Text25.Text = "08"
If Text25.Text = "9" Then Text25.Text = "09"

If Text26.Text = "0" Then Text26.Text = "00"
If Text26.Text = "1" Then Text26.Text = "01"
If Text26.Text = "2" Then Text26.Text = "02"
If Text26.Text = "3" Then Text26.Text = "03"
If Text26.Text = "4" Then Text26.Text = "04"
If Text26.Text = "5" Then Text26.Text = "05"
If Text26.Text = "6" Then Text26.Text = "06"
If Text26.Text = "7" Then Text26.Text = "07"
If Text26.Text = "8" Then Text26.Text = "08"
If Text26.Text = "9" Then Text26.Text = "09"
End Sub

Private Sub Timer4_Timer()
'On Error GoTo er:
' kanei ta ekshs. proton sent initial mail second check what caption toy put kai pote stelnei teliko mail.
'sto telos arxikopoiei oles tis metablhtes
' INITIAL
If SENDINITIALMAILNOW = 1 Then
    Call Shell("C:HELIX\sendEmail-v155\sendEmail.exe -f CARTOON@velti.com -t nlazarou@velti.com -u " & INITIAL_MESSAGE_MAIL_CAPTION & " -m " & INITIAL_MESSAGE_MAIL & " -s hermes:25 ", vbNormalFocus)
    SENDINITIALMAILNOW = 0
End If

'EMERGENCY
If FLAG_INCLUDE_RM = 0 Then
    If (FLAG_AVI = 3) Or (FLAG_263GP = 3) Or (FLAG_823GP = 3) Then
        MAIL_CAPTION = "CRITICAL ~~~~ CARTOON WAS NOT WORKED AT ALL !!!!"
        Text30.Text = MAIL_CAPTION
        ARXIKOPOIHSH
        MAILSTATUS = 3
        SENDMAILNOW = 1
        FLAGSTARTPROCEDURE = 0
    End If
 Else
    If (FLAG_AVI = 3) Or (FLAG_263GP = 3) Or (FLAG_823GP = 3) Or (FLAG_25RM = 3) Or (FLAG_82RM = 3) Then
        MAIL_CAPTION = "CRITICAL ~~~~ CARTOON WAS NOT WORKED CORRECTLY !!!!"
        Text30.Text = MAIL_CAPTION
        ARXIKOPOIHSH
        MAILSTATUS = 3
        SENDMAILNOW = 1
        FLAGSTARTPROCEDURE = 0
    End If
End If

'TELIKO MAIL
If PREPARE_MAIL = 1 Then
    If FLAG_INCLUDE_RM = 0 Then
        If (FLAG_AVI = 2) Or (FLAG_263GP = 2) Or (FLAG_823GP = 2) Then
            MAIL_CAPTION = "WARNING ~~~~ IN CARTOON ONE OR SOME FILES SEEMS NOT TO BE CREATED COORECTLY."
            Text30.Text = MAIL_CAPTION
            ARXIKOPOIHSH
            MAILSTATUS = 2
            SENDMAILNOW = 1
            FLAGSTARTPROCEDURE = 0
        Else
            MAIL_CAPTION = "SUCCESS ~~~~ CARTOON WAS WORKED FINE."
            Text30.Text = MAIL_CAPTION
            ARXIKOPOIHSH
            MAILSTATUS = 2
            SENDMAILNOW = 1
            FLAGSTARTPROCEDURE = 0
        End If
    Else
        If (FLAG_AVI = 2) Or (FLAG_263GP = 2) Or (FLAG_823GP = 2) Or (FLAG_25RM = 2) Or (FLAG_82RM = 2) Then
            MAIL_CAPTION = "WARING ~~~~ IN CARTOON ONE OR SOME FILES SEEMS NOT TO BE CREATED COORECTLY."
           Text30.Text = MAIL_CAPTION
            ARXIKOPOIHSH
            MAILSTATUS = 2
            SENDMAILNOW = 1
            FLAGSTARTPROCEDURE = 0
        Else
            MAIL_CAPTION = "SUCCESS ~~~~ CARTOON WAS WORKED FINE."
            Text30.Text = MAIL_CAPTION
            ARXIKOPOIHSH
            MAILSTATUS = 2
            SENDMAILNOW = 1
            FLAGSTARTPROCEDURE = 0
        End If
    End If
End If

If SENDMAILNOW = 1 Then
    'SEND TO MAIL ME CAPTION TEXT30 KAI BODY TEXT10
    '******************** MAIL ************************
    If MAILSTATUS = 3 Then
        Call Shell("C:\HELIX\sendEmail-v155\sendEmail.exe -f CARTOON@velti.com -t nlazarou@velti.com -u " & Text30.Text & " -m " & Text10.Text & "  -s hermes:25 ", vbNormalFocus)
        ARXIKOPOIHSH
        SENDMAILNOW = 0
        MAILSTATUS = 0
    Else
        Dim HMER, MERA As String
        If Day(Date) >= 1 And Day(Date) <= 9 Then
            MERA = "0" & Day(Date)
    Else
            MERA = Day(Date)
    End If
    HMER = MERA & "-" & Month(Date) & "-" & Year(Date)
        Call Shell("C:\HELIX\sendEmail-v155\sendEmail.exe -f CARTOON@velti.com -t nlazarou@velti.com -u " & Text30.Text & " -m " & Text10.Text & " -a C:\HELIX\LOGS\FTP_LOGS\FTPLOG_" & HMER & ".TXT -s hermes:25 ", vbNormalFocus)
        Text44.Text = "C:\HELIX\sendEmail-v155\sendEmail.exe -f CARTOON@velti.com -t nlazarou@velti.com -u " & Text30.Text & " -m " & Text10.Text & " -a C:\HELIX\LOGS\FTP_LOGS\FTPLOG_" & HMER & ".TXT -s hermes:25 "
        ARXIKOPOIHSH
        SENDMAILNOW = 0
        MAILSTATUS = 0
    End If
    
    '****************** grapsimo logs ****************************************8
    Dim source, destination As String
    source = "C:\HELIX\log.txt"
    destination = "C:\HELIX\LOGS\log_" & Day(Date) & "-" & Month(Date) & "-" & Year(Date)
    FileCopy source, destination
    Text9.Text = ""
    Text9.Text = Text30.Text
    Wrap$ = Chr$(13) + Chr$(10)
    Text9.Text = Text9.Text & Wrap & Text10.Text
    CommonDialog1.FileName = "C:\HELIX\LOGS\log_" & Day(Date) & "-" & Month(Date) & "-" & Year(Date)
    Open CommonDialog1.FileName For Output As #1
    Print #1, Text9.Text
    Close #1
    Text10.Text = ""
    Text30.Text = ""
    Text9.Text = ""
End If
GoTo TELOS:

er:
MsgBox ("йапоио амапамтево кахос елжамистгйе."), vbCritical, "пяосовг !!!"

TELOS:
End Sub

