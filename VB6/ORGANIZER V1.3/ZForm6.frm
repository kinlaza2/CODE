VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form ZForm6 
   BackColor       =   &H80000013&
   Caption         =   "елжамисг тилокоциым  амтицяажоу"
   ClientHeight    =   10485
   ClientLeft      =   255
   ClientTop       =   345
   ClientWidth     =   14880
   LinkTopic       =   "Form6"
   ScaleHeight     =   10485
   ScaleWidth      =   14880
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   480
      Top             =   6840
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
   Begin MSDataGridLib.DataGrid DataGrid3 
      Height          =   495
      Left            =   240
      TabIndex        =   9
      Top             =   6240
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
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
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8280
      TabIndex        =   8
      Top             =   9360
      Width           =   1455
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Height          =   615
      Left            =   5640
      TabIndex        =   7
      Top             =   6240
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1085
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
      Height          =   330
      Left            =   3480
      Top             =   6240
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      TabIndex        =   6
      Top             =   8280
      Width           =   2020
   End
   Begin VB.CommandButton Command5 
      Caption         =   "ейтупысг"
      Height          =   855
      Left            =   13560
      TabIndex        =   5
      Top             =   9360
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "диацяажг еццяажгс"
      Height          =   375
      Left            =   10440
      TabIndex        =   4
      Top             =   10080
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.CommandButton Command3 
      Caption         =   "диояхысг еццяажгс"
      Height          =   375
      Left            =   10440
      TabIndex        =   3
      Top             =   9720
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.CommandButton Command2 
      Caption         =   "еисацыцг  меас еццяажгс"
      Height          =   375
      Left            =   10440
      TabIndex        =   2
      Top             =   9360
      Visible         =   0   'False
      Width           =   2895
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   120
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
      Height          =   855
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   9360
      Width           =   1335
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   7455
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   14415
      _ExtentX        =   25426
      _ExtentY        =   13150
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      WrapCellPointer =   -1  'True
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
   Begin VB.Label Label2 
      BackColor       =   &H80000013&
      Caption         =   "сумокийо посо :"
      Height          =   375
      Left            =   3000
      TabIndex        =   11
      Top             =   8280
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000013&
      Caption         =   "аяихлос тилокоциым ╧ епитацым :"
      Height          =   375
      Left            =   5400
      TabIndex        =   10
      Top             =   9360
      Width           =   2895
   End
End
Attribute VB_Name = "ZForm6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error GoTo er:

ZETAIRIES.Combo14.Text = "глеяа"
ZETAIRIES.Combo15.Text = "лгмас"
ZETAIRIES.Combo16.Text = "етос"
ZETAIRIES.Combo17.Text = "глеяа"
ZETAIRIES.Combo18.Text = "лгмас"
ZETAIRIES.Combo19.Text = "етос"
ZForm6.Hide
Unload ZForm6
GoTo TELOS:

er:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
End Sub

Private Sub Command2_Click()
If MsgBox("пяосовг!! OI летабокес стоивеиым ле тгм вягсг аутоу тоу йоулпиоу хекоум поку пяосовг диоти ам дем цимоум сыста лпояеи ма пяойакесоум кахои та опоиа дем лпояоум ма диыяхыхоум. ам дем еисте сицоуяои циа тгм емеяцеиа поу пяойеите ма ейтекесете паяайакы патгсте CANCEL", vbOKCancel, "пяосовг!!") = vbOK Then
If MsgBox("еисте сицоуяои;", vbOKCancel, "пяосовг!!") = vbOK Then

If Form6.Command2.Caption = "еисацыцг  меас еццяажгс" Then
Form6.Command3.Enabled = False
Form6.Command4.Enabled = False
Form6.Command2.Caption = "текос еисацыцгс меым еццяажым"
Form6.DataGrid1.AllowUpdate = True
Form6.DataGrid1.AllowAddNew = True
Else
Form6.Command3.Enabled = True
Form6.Command4.Enabled = True
Form6.Command2.Caption = "еисацыцг  меас еццяажгс"
Form6.DataGrid1.AllowUpdate = False
Form6.DataGrid1.AllowAddNew = False
End If

End If
End If
End Sub

Private Sub Command3_Click()
If MsgBox("пяосовг!! OI летабокес стоивеиым ле тгм вягсг аутоу тоу йоулпиоу хекоум поку пяосовг диоти ам дем цимоум сыста лпояеи ма пяойакесоум кахои та опоиа дем лпояоум ма диыяхыхоум. ам дем еисте сицоуяои циа тгм емеяцеиа поу пяойеите ма ейтекесете паяайакы патгсте CANCEL", vbOKCancel, "пяосовг!!") = vbOK Then
If MsgBox("еисте сицоуяои;", vbOKCancel, "пяосовг!!") = vbOK Then


If Form6.Command3.Caption = "диояхысг еццяажгс" Then
Form6.Command2.Enabled = False
Form6.Command4.Enabled = False
Form6.Command3.Caption = "текос диояхысгс еццяажгс"
Form6.DataGrid1.AllowUpdate = True
Else
Form6.Command2.Enabled = True
Form6.Command4.Enabled = True
Form6.Command3.Caption = "диояхысг еццяажгс"
Form6.DataGrid1.AllowUpdate = False
End If

End If
End If
End Sub

Private Sub Command4_Click()
If MsgBox("пяосовг!! OI летабокес стоивеиым ле тгм вягсг аутоу тоу йоулпиоу хекоум поку пяосовг диоти ам дем цимоум сыста лпояеи ма пяойакесоум кахои та опоиа дем лпояоум ма диыяхыхоум. ам дем еисте сицоуяои циа тгм емеяцеиа поу пяойеите ма ейтекесете паяайакы патгсте CANCEL", vbOKCancel, "пяосовг!!") = vbOK Then
If MsgBox("еисте сицоуяои;", vbOKCancel, "пяосовг!!") = vbOK Then


Dim STATEMENT As String
If Form6.Command4.Caption = "диацяажг еццяажгс" Then
Form6.Command2.Enabled = False
Form6.Command3.Enabled = False
Form6.Command4.Caption = "текос диацяажгс еццяажым"
Form6.DataGrid1.AllowUpdate = True
Form6.DataGrid1.AllowDelete = True
STATEMENT = " delete from " & Form4.Text1.Text & _
" WHERE аяихлос_тилокоциоу ='" & Form6.DataGrid1.Columns(0).Value & "'"
db1.Execute STATEMENT
Form6.Hide
Unload Form6
Load Form6
Form6.Show
Else
Form6.Command2.Enabled = True
Form6.Command3.Enabled = True
Form6.Command4.Caption = "еисацыцг  меас еццяажгс"
Form6.DataGrid1.AllowUpdate = False
Form6.DataGrid1.AllowDelete = False
End If

End If
End If
End Sub

Private Sub Command5_Click()
On Error GoTo er:

Command1.Visible = False
Command2.Visible = False
Command3.Visible = False
Command4.Visible = False
Command5.Visible = False
ZForm6.BackColor = &H80000005
ZForm6.PrintForm
Command1.Visible = True
Command2.Visible = True
Command3.Visible = True
Command4.Visible = True
Command5.Visible = True
ZForm6.BackColor = &H8000000F
GoTo TELOS:

er:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
End Sub

Private Sub Command6_Click()

End Sub

Private Sub DataGrid1_Click()

If PATIMA_DATAGRID_F6 = 0 Then
   ZETAIRIES.Text2.Text = ZForm6.DataGrid1.Columns(0).Text
Else
    ZETAIRIES.Text20.Text = ZForm6.DataGrid1.Columns(0).Text
End If
End Sub

Private Sub Form_Load()
On Error GoTo er:

DataGrid1.Font.Bold = True
If FLAG_FORM6 = 1 Then
    DataGrid1.DefColWidth = 1995
End If
If FLAG_FORM6 = 2 Then
    DataGrid1.DefColWidth = 2800
End If
If FLAG_FORM6 = 3 Then
    DataGrid1.DefColWidth = 2800
End If


Dim DATABASE_FILE As String
DATABASE_FILE = ZETAIRIES_DIADROMHS_BACKUP_DIAX
 Adodc1.ConnectionString = _
        "PROVIDER=Microsoft.Jet.OLEDB.4.0;" & _
        "Data Source=" & DATABASE_FILE & ";"
    Adodc1.RecordSource = STATE
    ' Bind the ADODC to the DataGrid.
    Set DataGrid1.DataSource = Adodc1

Adodc2.ConnectionString = _
"PROVIDER=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & DATABASE_FILE & ";"
Adodc2.RecordSource = SS
' Bind the ADODC to the DataGrid.
Set DataGrid2.DataSource = Adodc2
'rs.MoveFirst
Text1.Text = DataGrid2.Text

Adodc3.ConnectionString = _
"PROVIDER=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & DATABASE_FILE & ";"
Adodc3.RecordSource = FF
' Bind the ADODC to the DataGrid.
Set DataGrid3.DataSource = Adodc3
Text2.Text = DataGrid3.Text

If Text2.Text <= 33 Then
    DataGrid1.Height = (1 + Text2.Text) * 242.142857
    Text1.Top = (1 + Text2.Text) * 242.142857 + 280
    Label2.Top = (1 + Text2.Text) * 242.142857 + 280
Else
    DataGrid1.Height = 8475
    Text1.Top = 8755
    Label2.Top = 8755
End If
GoTo TELOS:

er:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo er:

ZETAIRIES.Combo14.Text = "глеяа"
ZETAIRIES.Combo15.Text = "лгмас"
ZETAIRIES.Combo16.Text = "етос"
ZETAIRIES.Combo17.Text = "глеяа"
ZETAIRIES.Combo18.Text = "лгмас"
ZETAIRIES.Combo19.Text = "етос"
ZForm6.Hide
Unload ZForm6
GoTo TELOS:

er:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
End Sub
