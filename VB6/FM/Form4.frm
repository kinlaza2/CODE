VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form4 
   Caption         =   "FM-RA"
   ClientHeight    =   5730
   ClientLeft      =   6060
   ClientTop       =   420
   ClientWidth     =   12990
   LinkTopic       =   "Form4"
   ScaleHeight     =   5730
   ScaleWidth      =   12990
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   11280
      Top             =   240
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   4095
      Left            =   360
      TabIndex        =   3
      Top             =   240
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   7223
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
   Begin VB.CommandButton Command2 
      Caption         =   "DELETE RECORD"
      Height          =   615
      Left            =   4680
      TabIndex        =   0
      Top             =   5040
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   4560
      Width           =   11415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   735
      Left            =   11520
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If RS1.State = 1 Then RS1.Close
If DB1.State = 1 Then DB1.Close
'Adodc1.Refresh
Form4.Hide
Unload Form4
End Sub

Private Sub Command2_Click()
'On Error GoTo ER:
If RS1.State = 1 Then RS1.Close
If DB1.State = 1 Then DB1.Close
DB1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & "\DB1.mdb" & ";" & _
      "Persist Security Info=False"
DB1.Open App.Path & "\DB1.mdb"
RS1.Open "[PIN]", DB1, adOpenDynamic, adLockBatchOptimistic

Dim STAT As String
STAT = "DELETE FROM PIN WHERE KEIMENO='" & Trim(Text1.Text) & "'"
DB1.Execute STAT
RS1.Fields.Refresh


If RS1.State = 1 Then RS1.Close
If DB1.State = 1 Then DB1.Close
DB1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & "\DB1.mdb" & ";" & _
      "Persist Security Info=False"
DB1.Open App.Path & "\DB1.mdb"
RS1.Open "[PIN]", DB1, adOpenDynamic, adLockBatchOptimistic
Adodc1.ConnectionString = _
"PROVIDER=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & DATABASE_FILE & ";"
Adodc1.RecordSource = STAT_EMFANISIS
' Bind the ADODC to the DataGrid.
Set DataGrid1.DataSource = Adodc1
Text1.Text = ""
Unload Form4
Load Form4
Form4.Show

GoTo TELOS:
er:
MsgBox ("йапоио амапамтево кахос елжамистгйе."), vbCritical, "пяосовг !!!"

TELOS:
End Sub

Private Sub DataGrid1_Click()
On Error GoTo TELOS:
TELOS:
End Sub

Private Sub DataGrid1_DblClick()
On Error GoTo TELOS:
Text1.Text = DataGrid1.Columns(0).Text
Form3.Text4.Text = DataGrid1.Columns(0).Text
TELOS:
End Sub

Private Sub Form_Load()
DataGrid1.Font.Size = 14
DataGrid1.DefColWidth = 10500
DataGrid1.HeadFont.Bold = True
DataGrid1.HeadFont.Size = 14

If RS1.State = 1 Then RS1.Close
If DB1.State = 1 Then DB1.Close
DB1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & "\DB1.mdb" & ";" & _
      "Persist Security Info=False"
DB1.Open App.Path & "\DB1.mdb"
RS1.Open "[PIN]", DB1, adOpenDynamic, adLockBatchOptimistic

Dim DATABASE_FILE As String
DATABASE_FILE = App.Path & "/DB1.MDB"
Adodc1.ConnectionString = _
"PROVIDER=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & DATABASE_FILE & ";"
Adodc1.RecordSource = STAT_EMFANISIS
' Bind the ADODC to the DataGrid.
Set DataGrid1.DataSource = Adodc1
'RS1.Fields.Refresh
'Adodc1.Refresh
'DataGrid1.Refresh
If RS1.State = 1 Then RS1.Close
If DB1.State = 1 Then DB1.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
If RS1.State = 1 Then RS1.Close
If DB1.State = 1 Then DB1.Close
'Adodc1.Refresh
Form4.Hide
Unload Form4
End Sub
