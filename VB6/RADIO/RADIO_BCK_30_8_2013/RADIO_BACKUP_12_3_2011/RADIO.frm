VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form RADIO 
   BackColor       =   &H80000013&
   Caption         =   "RADIO"
   ClientHeight    =   10485
   ClientLeft      =   8070
   ClientTop       =   450
   ClientWidth     =   5550
   FillColor       =   &H8000000A&
   ForeColor       =   &H8000000A&
   LinkTopic       =   "Form1"
   ScaleHeight     =   10485
   ScaleWidth      =   5550
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   2880
      TabIndex        =   9
      Text            =   "Text3"
      Top             =   9720
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00E9C5AD&
      Caption         =   "EXIT"
      Height          =   615
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3960
      Width           =   2415
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00E9C5AD&
      Caption         =   "CONFIGURE RADIO"
      Height          =   615
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3120
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2880
      TabIndex        =   6
      Text            =   "Boithitiko text gia na gyrisei URL"
      Top             =   6600
      Visible         =   0   'False
      Width           =   5055
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00E9C5AD&
      Caption         =   "STOP"
      Height          =   375
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   600
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E9C5AD&
      Caption         =   "HIDE"
      Height          =   615
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2280
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      Top             =   120
      Width           =   2415
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   10215
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   18018
      _Version        =   393216
      DefColWidth     =   150
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   2880
      Top             =   7080
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
      Caption         =   "PLAY"
      Height          =   375
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   600
      Width           =   855
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   975
      Left            =   2880
      TabIndex        =   0
      Top             =   1080
      Width           =   2415
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   4260
      _cy             =   1720
   End
End
Attribute VB_Name = "RADIO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error GoTo er:
Dim play_URL As String
Dim c As Integer
'  ##################### SYNDESH ME BASH GIA ANAKTHSH URL ###############################
Dim db As New ADODB.Connection
Dim RS As New ADODB.Recordset
db.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & "\DB.mdb" & ";" & _
"Persist Security Info=False"
db.Open App.Path & "\DB.mdb"
RS.Open "PIN", db, adOpenDynamic, adLockBatchOptimistic
c = 0
If RS.BOF = RS.EOF Then GoTo NIK:
RS.MoveFirst
NIK:
Do While Not RS.EOF
    If RS![Name] <> Trim(Text1.Text) Then
        RS.MoveNext
    Else
        c = c + 1
        Text2.Text = RS![URL]
        RS.MoveNext
    End If
Loop

If c <> 0 Then
    SOUN = Trim(Text2.Text)
Else
    SOUN = ""
    GoTo er1:
End If
WindowsMediaPlayer1.URL = SOUN
GoTo TELOS:

er1:
MsgBox ("TO URL DEN BRETHIKE"), vbCritical, "пяосовг !!!"
GoTo TELOS:

er:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"
GoTo TELOS:



TELOS:
End Sub

Private Sub Command2_Click()
RADIO.Hide
small.Show
End Sub

Private Sub Command3_Click()
WindowsMediaPlayer1.URL = ""
End Sub

Private Sub Command4_Click()
Load conf
RADIO.Hide
conf.Show
End Sub

Private Sub Command5_Click()
End
End Sub

Private Sub DataGrid1_Click()
Text1.Text = Trim(DataGrid1.Columns(0).Text)
End Sub

Private Sub DataGrid1_DblClick()
Text1.Text = Trim(DataGrid1.Columns(0).Text)
End Sub

Private Sub Form_Load()
Dim DATABASE_FILE As String
DATABASE_FILE = App.Path & "\DB.mdb"
Adodc1.ConnectionString = _
"PROVIDER=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & DATABASE_FILE & ";"
Adodc1.RecordSource = "SELECT NAME FROM PIN ORDER BY NAME"
' Bind the ADODC to the DataGrid.
Set DataGrid1.DataSource = Adodc1

DataGrid1.Font.Size = 10
DataGrid1.HeadFont.Size = 10
DataGrid1.HeadFont.Bold = True
DataGrid1.DefColWidth = 2250

Text3.Text = Adodc1.Recordset.RecordCount
    If Text3.Text <= 33 Then
        DataGrid1.Height = 300 + (CInt(Text3.Text) * 300)
    Else
        DataGrid1.Height = 10215
    End If
'DataGrid1.Height = 300 default

End Sub

