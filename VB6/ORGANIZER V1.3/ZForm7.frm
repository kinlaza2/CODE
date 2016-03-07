VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form ZForm7 
   BackColor       =   &H80000013&
   Caption         =   "йаятека амтицяажоу"
   ClientHeight    =   10590
   ClientLeft      =   105
   ClientTop       =   465
   ClientWidth     =   15225
   LinkTopic       =   "Form7"
   ScaleHeight     =   10590
   ScaleWidth      =   15225
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   4680
      TabIndex        =   16
      Text            =   "Text2"
      Top             =   480
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid DataGrid4 
      Height          =   615
      Left            =   2400
      TabIndex        =   15
      Top             =   120
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   375
      Left            =   0
      Top             =   480
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
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
      Caption         =   "Adodc4"
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
      Height          =   285
      Left            =   480
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   120
      Visible         =   0   'False
      Width           =   855
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   8655
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   14895
      _ExtentX        =   26273
      _ExtentY        =   15266
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   19
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
         Size            =   9.75
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
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E9C5AD&
      Caption         =   "ейтупысг"
      Height          =   735
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   9720
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E9C5AD&
      Caption         =   "епистяожг"
      Height          =   735
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   9720
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid DataGrid3 
      Height          =   855
      Left            =   4200
      TabIndex        =   5
      Top             =   7800
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
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
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   375
      Left            =   4200
      Top             =   8760
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
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
   Begin MSDataGridLib.DataGrid DataGrid2 
      Height          =   855
      Left            =   1320
      TabIndex        =   4
      Top             =   7680
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
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
      Height          =   375
      Left            =   1320
      Top             =   8640
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
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
      Height          =   330
      Left            =   4560
      Top             =   10080
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
   Begin VB.Line Line5 
      Visible         =   0   'False
      X1              =   7710
      X2              =   7710
      Y1              =   0
      Y2              =   11360
   End
   Begin VB.Line Line4 
      Visible         =   0   'False
      X1              =   14880
      X2              =   14880
      Y1              =   0
      Y2              =   11360
   End
   Begin VB.Line Line3 
      Visible         =   0   'False
      X1              =   10100
      X2              =   10100
      Y1              =   0
      Y2              =   11360
   End
   Begin VB.Line Line2 
      Visible         =   0   'False
      X1              =   12490
      X2              =   12490
      Y1              =   0
      Y2              =   11360
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      Visible         =   0   'False
      X1              =   7672
      X2              =   7672
      Y1              =   0
      Y2              =   11360
   End
   Begin VB.Label Label9 
      BackColor       =   &H80000013&
      Caption         =   "цемийа сумока"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      TabIndex        =   11
      Top             =   9600
      Width           =   1935
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
      Caption         =   "Label8"
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
      Left            =   12480
      TabIndex        =   10
      Top             =   9600
      Width           =   2415
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
      Caption         =   "Label7"
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
      Left            =   12480
      TabIndex        =   9
      Top             =   480
      Width           =   2415
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000013&
      Caption         =   "йаятека"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      TabIndex        =   8
      Top             =   0
      Width           =   6735
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
      Caption         =   "Label5"
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
      Left            =   10080
      TabIndex        =   7
      Top             =   9600
      Width           =   2415
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
      Caption         =   "Label4"
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
      TabIndex        =   6
      Top             =   9600
      Width           =   2415
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
      Caption         =   "Label3"
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
      Left            =   10080
      TabIndex        =   3
      Top             =   480
      Width           =   2415
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
      Caption         =   "Label2"
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
      TabIndex        =   2
      Top             =   480
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000013&
      Caption         =   "ей летажояас"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6000
      TabIndex        =   1
      Top             =   480
      Width           =   1695
   End
End
Attribute VB_Name = "ZForm7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error GoTo er:

ZETAIRIES.Combo21.Text = "глеяа"
ZETAIRIES.Combo22.Text = "лгмас"
ZETAIRIES.Combo23.Text = "етос"
ZETAIRIES.Combo24.Text = "глеяа"
ZETAIRIES.Combo25.Text = "лгмас"
ZETAIRIES.Combo26.Text = "етос"

ZForm7.Hide
Unload ZForm7
GoTo TELOS:

er:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
End Sub

Private Sub Command2_Click()
On Error GoTo er:

Command1.Visible = False
Command2.Visible = False
ZForm7.BackColor = &HFFFFFF
Label1.BackColor = &HFFFFFF
Label2.BackColor = &HFFFFFF
Label3.BackColor = &HFFFFFF
Label4.BackColor = &HFFFFFF
Label5.BackColor = &HFFFFFF
Label6.BackColor = &HFFFFFF
Label7.BackColor = &HFFFFFF
Label8.BackColor = &HFFFFFF
Label9.BackColor = &HFFFFFF
DataGrid1.BackColor = &HFFFFFF
DataGrid1.RecordSelectors = False

ZForm7.PrintForm
GoTo TELOS:

er:
Command1.Visible = True
Command2.Visible = True
Label1.BackColor = &H80000013
Label2.BackColor = &H80000013
Label3.BackColor = &H80000013
Label4.BackColor = &H80000013
Label5.BackColor = &H80000013
Label6.BackColor = &H80000013
Label7.BackColor = &H80000013
Label8.BackColor = &H80000013
Label9.BackColor = &H80000013
ZForm7.BackColor = &H80000013
DataGrid1.BackColor = &HFFFFFF
DataGrid1.RecordSelectors = True
MsgBox ("йапоио амапамтево кахос елжамистгйе.пихамом дем еимаи думатг г еуяесг ейтупытг, паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
'&H00FFFFFF&=LEYKO &H80000013&=galazio
Command1.Visible = True
Command2.Visible = True
Label1.BackColor = &H80000013
Label2.BackColor = &H80000013
Label3.BackColor = &H80000013
Label4.BackColor = &H80000013
Label5.BackColor = &H80000013
Label6.BackColor = &H80000013
Label7.BackColor = &H80000013
Label8.BackColor = &H80000013
Label9.BackColor = &H80000013
ZForm7.BackColor = &H80000013
DataGrid1.BackColor = &HFFFFFF
DataGrid1.RecordSelectors = True
End Sub

Private Sub DataGrid1_Click()
'On Error GoTo TELOS:

If DataGrid1.Columns(2).Text = "епитацг" Then
    Form4.Text20.Text = Form7.DataGrid1.Columns(1).Text
Else
    Form4.Text2.Text = Form7.DataGrid1.Columns(1).Text
End If
TELOS:
End Sub

Private Sub Form_Load()
On Error GoTo er:
Form1.Enabled = False

Label6.Caption = " йаятека етаияиас  : " & Form4.Text1.Text
DataGrid1.DefColWidth = 2400
DataGrid1.HeadFont.Bold = True
DataGrid1.HeadFont.Size = 8

If db1.STATE = 1 Then db1.Close
If rs1.STATE = 1 Then rs1.Close
db1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & ZETAIRIES_DIADROMHS_BACKUP_DIAX_1 & ";" & _
"Persist Security Info=False"
db1.Open ZETAIRIES_DIADROMHS_BACKUP_DIAX
rs1.Open "[" & ZETAIRIES.Text1.Text & "]", db1, adOpenDynamic, adLockBatchOptimistic


' ONOMAZO PERIODO TO XRONIKO DIASTHMA POY MOY ERXETAI
' APO THN FORM 4 DHLADH APO HAK KAI HMK

'********************************************************************* & _
YPOLOGISMOS XREOSEON PRIN APO HAK**************************************
Dim TEMP11, DATE11
Dim STATEX As String
Dim XREOSH As Double
Dim DATABASE_FILE As String
XREOSH = 0

TEMP11 = CDate("1/1/" & ETOSBACKUP)

If Day(TEMP11) < 12 Then
    DATE11 = CDate(Month(TEMP11) & "/" & Day(TEMP11) & "/" & Year(TEMP11))
Else
    DATE11 = TEMP11
End If

STATEX = "SELECT SUM(вяеысг) FROM " & ZETAIRIES.Text1.Text & _
" where (глеяолгмиа_ейдосгс between " & "#" & DATE11 & "#" & " and " & "#" & HAK & "#" & "-1" & ")"

DATABASE_FILE = ZETAIRIES_DIADROMHS_BACKUP_DIAX
Adodc2.ConnectionString = _
"PROVIDER=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & DATABASE_FILE & ";"
Adodc2.RecordSource = STATEX
Set DataGrid2.DataSource = Adodc2
If DataGrid2.Columns(0) = "" Then
    XREOSH = 0
Else
    XREOSH = DataGrid2.Columns(0)
End If
Adodc2.Refresh
DataGrid2.Refresh
'********************************************************* & _
YPOLOGISMOS PISTOSEON PRIN APO HAK**************************

Dim STATEP As String
Dim PISTOSI As Double
PISTOSI = 0
STATEP = "SELECT SUM(пистысг) FROM " & ZETAIRIES.Text1.Text & _
" where (глеяолгмиа_ейдосгс between " & "#" & DATE11 & "#" & " and " & "#" & HAK & "#" & "-1" & ")"

DATABASE_FILE = ZETAIRIES_DIADROMHS_BACKUP_DIAX
Adodc3.ConnectionString = _
"PROVIDER=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & DATABASE_FILE & ";"
Adodc3.RecordSource = STATEP
Set DataGrid3.DataSource = Adodc3
If DataGrid3.Columns(0) = "" Then
    PISTOSI = 0
Else
    PISTOSI = DataGrid3.Columns(0)
End If
Adodc3.Refresh
DataGrid3.Refresh

'******YPOLOGISMOS YPOLOIPOY PRIN APO HAK***************
Dim YPOLIPOPRIN As Double
YPOLIPOPRIN = XREOSH - PISTOSI


'*************************************************************** & _
YPOLOGISMOS YPOLOIPON MESA STHN PERIODO*************************
'*******************************************************
Dim SSS, SSS_UP As String
Dim XR, PI, YY, YYY, WE As Double
YY = 0
WE = 0

SSS = " select * into HELP_KARTELAS from " & ZETAIRIES.Text1.Text & " ORDER BY глеяолгмиа_ейдосгс,тупос DESC"
db1.Execute SSS

rsrs.Open "[HELP_KARTELAS]", db1, adOpenDynamic, adLockBatchOptimistic

If rsrs.BOF = rsrs.EOF Then GoTo NIK1:
rsrs.MoveFirst
NIK1:
Do While Not rsrs.EOF
    Text1.Text = rsrs![аяихлос_тилокоциоу]
    XR = rsrs![вяеысг]
    PI = rsrs![пистысг]
    YY = XR - PI
    WE = WE + YY
    
    SSS_UP = " UPDATE HELP_KARTELAS SET упокоипо='" & WE & "'" & _
    " WHERE аяихлос_тилокоциоу='" & Text1.Text & "'"
    db1.Execute SSS_UP
    rsrs.MoveNext
Loop


' ****YPOLOGISMOS TELIKHS XREOSHS,PISTOSIS,TELIKO YPOLOIPO******************

'YPOLOGISMOS XREOSEON PRIN APO HMK
Dim XREOSHT, PISTOSIT As Double
Dim STATEXT, STATEPT As String
XREOSHT = 0
PISTOSIT = 0
STATEXT = "SELECT SUM(вяеысг) FROM " & ZETAIRIES.Text1.Text & _
" where (глеяолгмиа_ейдосгс between " & "#" & DATE11 & "#" & " and " & "#" & HMK & "#" & ")"

Adodc2.Refresh
DataGrid2.Refresh
DATABASE_FILE = ZETAIRIES_DIADROMHS_BACKUP_DIAX
Adodc2.ConnectionString = _
"PROVIDER=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & DATABASE_FILE & ";"
Adodc2.RecordSource = STATEXT
Set DataGrid2.DataSource = Adodc2
If DataGrid2.Columns(0) = "" Then
    XREOSHT = 0
Else
    XREOSHT = DataGrid2.Columns(0)
End If
Adodc2.Refresh
DataGrid2.Refresh

'YPOLOGISMOS PISTOSEON PRIN APO HMK

STATEPT = "SELECT SUM(пистысг) FROM " & ZETAIRIES.Text1.Text & _
" where (глеяолгмиа_ейдосгс between " & "#" & DATE11 & "#" & " and " & "#" & HMK & "#" & ")"

Adodc3.Refresh
DataGrid3.Refresh
DATABASE_FILE = ZETAIRIES_DIADROMHS_BACKUP_DIAX
Adodc3.ConnectionString = _
"PROVIDER=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & DATABASE_FILE & ";"
Adodc3.RecordSource = STATEPT
Set DataGrid3.DataSource = Adodc3
If DataGrid3.Columns(0) = "" Then
    PISTOSIT = 0
Else
    PISTOSIT = DataGrid3.Columns(0)
End If
Adodc3.Refresh
DataGrid3.Refresh

'******YPOLOGISMOS YPOLOIPOY PRIN APO HAK***************
Dim YPOLIPOMETA As Double
YPOLIPOMETA = XREOSHT - PISTOSIT



'EMFANISH KARTELAS*****************************************
DATABASE_FILE = ZETAIRIES_DIADROMHS_BACKUP_DIAX
Adodc1.ConnectionString = _
"PROVIDER=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & DATABASE_FILE & ";"
Adodc1.RecordSource = STATE_KARTELAS
Set DataGrid1.DataSource = Adodc1

If DataGrid3.Columns(0) = "" Then
    PISTOSIT = 0
Else
    PISTOSIT = DataGrid3.Columns(0)
End If
If DataGrid2.Columns(0) = "" Then
    XREOSHT = 0
Else
    XREOSHT = DataGrid2.Columns(0)
End If

'******************* RITHMISI DATAGRID ***************************
Dim CCC As Integer
Adodc4.ConnectionString = _
"PROVIDER=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & DATABASE_FILE & ";"
Adodc4.RecordSource = "SELECT COUNT(аяихлос_тилокоциоу) FROM HELP_KARTELAS " & _
" where (глеяолгмиа_ейдосгс between " & "#" & HAK & "#" & " and " & "#" & HMK & "#" & ")"
Set DataGrid4.DataSource = Adodc4
CCC = CInt(DataGrid4.Text)
Text2.Text = CCC
If CCC <= 28 Then
        ZForm7.DataGrid1.Height = 255 + (CCC * 300)
        ZForm7.Label4.Top = ZForm7.DataGrid1.Height + 945
        ZForm7.Label5.Top = ZForm7.DataGrid1.Height + 945
        ZForm7.Label8.Top = ZForm7.DataGrid1.Height + 945
        ZForm7.Label9.Top = ZForm7.DataGrid1.Height + 945
    Else
        ZForm7.DataGrid1.Height = 8655
        ZForm7.Label4.Top = 9600
        ZForm7.Label5.Top = 9600
        ZForm7.Label8.Top = 9600
        ZForm7.Label9.Top = 9600
    End If
'*********************************************************************

YPOLIPOMETA = XREOSHT - PISTOSIT
Label2.Caption = XREOSH
Label3.Caption = PISTOSI
Label4.Caption = XREOSHT
Label5.Caption = PISTOSIT
Label7.Caption = YPOLIPOPRIN
Label8.Caption = YPOLIPOMETA
GoTo TELOS:

er:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"


TELOS:
If db1.STATE = 1 Then db1.Close
If rs1.STATE = 1 Then rs1.Close
If rsrs.STATE = 1 Then rsrs.Close
End Sub



Private Sub Form_Unload(Cancel As Integer)
On Error GoTo er:

If db1.STATE = 1 Then db1.Close
If rs1.STATE = 1 Then rs1.Close
db1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
       "Data Source=" & ZETAIRIES_DIADROMHS_BACKUP_DIAX_1 & ";" & _
      "Persist Security Info=False"
db1.Open ZETAIRIES_DIADROMHS_BACKUP_DIAX

Dim STATEMENT As String
STATEMENT = "drop table HELP_KARTELAS"
db1.Execute STATEMENT
Form1.Enabled = True

ZETAIRIES.Combo21.Text = "глеяа"
ZETAIRIES.Combo22.Text = "лгмас"
ZETAIRIES.Combo23.Text = "етос"
ZETAIRIES.Combo24.Text = "глеяа"
ZETAIRIES.Combo25.Text = "лгмас"
ZETAIRIES.Combo26.Text = "етос"

ZForm7.Hide
Unload ZForm7
GoTo TELOS:



er:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"


TELOS:
End Sub

