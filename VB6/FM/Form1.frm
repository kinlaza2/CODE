VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form1 
   Caption         =   "FM"
   ClientHeight    =   11535
   ClientLeft      =   5055
   ClientTop       =   420
   ClientWidth     =   14805
   LinkTopic       =   "Form1"
   ScaleHeight     =   11535
   ScaleWidth      =   14805
   Begin VB.CommandButton Command18 
      Caption         =   "0"
      Height          =   735
      Left            =   8880
      TabIndex        =   44
      Top             =   1080
      Width           =   255
   End
   Begin VB.CommandButton Command17 
      Caption         =   "9P"
      Height          =   495
      Left            =   11520
      TabIndex        =   43
      Top             =   1080
      Width           =   375
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   10800
      Top             =   6840
      Visible         =   0   'False
      Width           =   3615
      _ExtentX        =   6376
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
   Begin VB.CommandButton Command14 
      Caption         =   "8N"
      Height          =   255
      Left            =   10920
      TabIndex        =   42
      Top             =   1320
      Width           =   615
   End
   Begin VB.CommandButton Command6 
      Caption         =   "7N"
      Height          =   255
      Left            =   10320
      TabIndex        =   41
      Top             =   1320
      Width           =   615
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   615
      Left            =   11760
      Top             =   5880
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
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
   Begin MSDataGridLib.DataGrid DataGrid2 
      Height          =   855
      Left            =   12240
      TabIndex        =   40
      Top             =   2280
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
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
            LCID            =   2057
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
            LCID            =   2057
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3975
      Left            =   240
      TabIndex        =   39
      Top             =   1920
      Width           =   14295
      _ExtentX        =   25215
      _ExtentY        =   7011
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
            LCID            =   2057
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
            LCID            =   2057
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
   Begin VB.TextBox Text10 
      Height          =   375
      Left            =   10440
      TabIndex        =   34
      Top             =   600
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      Height          =   375
      Left            =   11760
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   33
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton Command16 
      Caption         =   "6L"
      Height          =   255
      Left            =   9720
      TabIndex        =   32
      Top             =   1319
      Width           =   615
   End
   Begin VB.CommandButton Command15 
      Caption         =   "5K"
      Height          =   255
      Left            =   9120
      TabIndex        =   31
      Top             =   1317
      Width           =   615
   End
   Begin VB.TextBox Text9 
      Height          =   285
      Left            =   5880
      TabIndex        =   30
      Text            =   "Text9"
      Top             =   1680
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Timer Timer4 
      Interval        =   10
      Left            =   5160
      Top             =   1560
   End
   Begin VB.CheckBox Check1 
      Caption         =   "SHOW ALL"
      Height          =   375
      Left            =   5280
      TabIndex        =   29
      Top             =   1320
      Width           =   855
   End
   Begin VB.TextBox Text8 
      Height          =   375
      Left            =   12240
      TabIndex        =   28
      Text            =   "Text8"
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Command13 
      Caption         =   "HMER"
      Height          =   375
      Left            =   13560
      TabIndex        =   27
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Timer Timer3 
      Interval        =   10
      Left            =   6120
      Top             =   720
   End
   Begin VB.Timer Timer2 
      Interval        =   10
      Left            =   5640
      Top             =   720
   End
   Begin VB.CommandButton Command12 
      Caption         =   "ALL"
      Height          =   255
      Left            =   9120
      TabIndex        =   18
      Top             =   1560
      Width           =   2775
   End
   Begin VB.CommandButton Command11 
      Caption         =   "4M"
      Height          =   255
      Left            =   10920
      TabIndex        =   17
      Top             =   1080
      Width           =   615
   End
   Begin VB.CommandButton Command10 
      Caption         =   "2S"
      Height          =   255
      Left            =   9720
      TabIndex        =   16
      Top             =   1080
      Width           =   615
   End
   Begin VB.CommandButton Command9 
      Caption         =   "3D"
      Height          =   255
      Left            =   10320
      TabIndex        =   15
      Top             =   1080
      Width           =   615
   End
   Begin VB.CommandButton Command8 
      Caption         =   "1S"
      Height          =   255
      Left            =   9120
      TabIndex        =   14
      Top             =   1080
      Width           =   615
   End
   Begin VB.CommandButton Command7 
      Caption         =   "елжамисг "
      Height          =   255
      Left            =   12480
      TabIndex        =   13
      Top             =   840
      Width           =   975
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   12240
      TabIndex        =   12
      Text            =   "Text7"
      Top             =   1560
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   12720
      TabIndex        =   11
      Text            =   "Text6"
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   11760
      TabIndex        =   10
      Text            =   "Text5"
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   12600
      TabIndex        =   9
      Text            =   "Text3"
      Top             =   480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Timer Timer1 
      Interval        =   60000
      Left            =   12960
      Top             =   1080
   End
   Begin VB.CommandButton Command5 
      Caption         =   "REMINDER"
      Height          =   495
      Left            =   13560
      TabIndex        =   8
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   9840
      TabIndex        =   7
      Text            =   "Text4"
      Top             =   12000
      Visible         =   0   'False
      Width           =   4815
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   10440
      TabIndex        =   6
      Text            =   "Combo1"
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   10215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   10215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "HIDE"
      Height          =   495
      Left            =   13680
      TabIndex        =   3
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "DELETE"
      Height          =   735
      Left            =   3600
      TabIndex        =   2
      Top             =   1080
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "UPDATE"
      Height          =   735
      Left            =   1920
      TabIndex        =   1
      Top             =   1080
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "INSERT"
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   1575
   End
   Begin VB.OLE OLE7 
      BackColor       =   &H00FFC0C0&
      Height          =   255
      Left            =   8400
      TabIndex        =   38
      Top             =   1620
      Width           =   255
   End
   Begin VB.OLE OLE6 
      BackColor       =   &H80000002&
      Height          =   255
      Left            =   7920
      TabIndex        =   37
      Top             =   1620
      Width           =   255
   End
   Begin VB.OLE OLE5 
      BackColor       =   &H0080FF80&
      DisplayType     =   1  'Icon
      Height          =   255
      Left            =   7440
      TabIndex        =   36
      Top             =   1620
      Width           =   255
   End
   Begin VB.Label Label5 
      Caption         =   "PRIORITY"
      Height          =   255
      Left            =   11040
      TabIndex        =   35
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "T"
      Height          =   255
      Left            =   7020
      TabIndex        =   26
      Top             =   1390
      Width           =   135
   End
   Begin VB.OLE OLE4 
      Class           =   "Package"
      Height          =   255
      Left            =   6960
      OleObjectBlob   =   "Form1.frx":0442
      OLETypeAllowed  =   0  'Linked
      SourceDoc       =   "C:\Users\ni\Desktop\MINE\FM\TASKS.txt"
      TabIndex        =   25
      Top             =   1080
      Width           =   255
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   6360
      Picture         =   "Form1.frx":0E5A
      Top             =   1200
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   6360
      Picture         =   "Form1.frx":129C
      Top             =   1200
      Width           =   480
   End
   Begin VB.Label Label3 
      Caption         =   "V"
      Height          =   255
      Left            =   8490
      TabIndex        =   24
      Top             =   1390
      Width           =   135
   End
   Begin VB.Label Label2 
      Caption         =   "W"
      Height          =   255
      Left            =   7980
      TabIndex        =   23
      Top             =   1390
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "E"
      Height          =   255
      Left            =   7520
      TabIndex        =   22
      Top             =   1390
      Width           =   135
   End
   Begin VB.OLE OLE3 
      Class           =   "Visio.Drawing.11"
      DisplayType     =   1  'Icon
      Height          =   255
      Left            =   8400
      OleObjectBlob   =   "Form1.frx":19DE
      SourceDoc       =   "C:\Users\ni\Desktop\MINE\FM\TEMP.vsd"
      TabIndex        =   21
      Top             =   1080
      Width           =   255
   End
   Begin VB.OLE OLE2 
      Class           =   "Word.Document.8"
      DisplayType     =   1  'Icon
      Height          =   255
      Left            =   7920
      OleObjectBlob   =   "Form1.frx":35F6
      SourceDoc       =   "C:\Users\ni\Desktop\MINE\FM\TEMP.doc"
      TabIndex        =   20
      Top             =   1080
      Width           =   255
   End
   Begin VB.OLE OLE1 
      Class           =   "Excel.Sheet.8"
      DisplayType     =   1  'Icon
      Height          =   255
      Left            =   7440
      OleObjectBlob   =   "Form1.frx":520E
      SourceDoc       =   "C:\Users\ni\Desktop\MINE\FM\TEMP.xls"
      TabIndex        =   19
      Top             =   1080
      Width           =   255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Adodc2_Click()

End Sub

Private Sub Combo1_LostFocus()
If Combo1.Text = "1S" Or Combo1.Text = "2S" Or Combo1.Text = "3D" Or Combo1.Text = "4M" Or Combo1.Text = "5K" Or Combo1.Text = "6L" Or Combo1.Text = "7N" Or Combo1.Text = "8N" Or Combo1.Text = "9PENDING" Or Combo1.Text = "0S" Then
Else
    Combo1.Text = "1S"
End If
End Sub

Private Sub Command1_Click()
'On Error GoTo TELOS:
Dim priorvalue As Integer

If Text10.Text = "" Then
    priorvalue = 40
Else
    If IsNumeric(Text10.Text) = False Then
        GoTo er1:
    Else
        priorvalue = Text10.Text
    End If
End If



If RS.State = 1 Then RS.Close
If DB.State = 1 Then DB.Close
DB.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & "\DB.mdb" & ";" & _
      "Persist Security Info=False"
DB.Open App.Path & "\DB.mdb"
RS.Open "[PIN]", DB, adOpenDynamic, adLockBatchOptimistic

Dim statement As String
statement = "INSERT INTO PIN(MAIN,COMMENT,SEVERITY,PRIORITY,DATE_STRING)" & _
    "VALUES (" & _
        "'" & UCase(Form1.Text1.Text) & "'," & _
        "'" & UCase(Form1.Text2.Text) & "', " & _
        "'" & Combo1.Text & "'," & _
        "'" & priorvalue & "'," & _
        "'" & Date & "')"
        
DB.Execute statement

RS.Fields.Refresh
RS.Close
DB.Close

If Text8.Text <> "ALL" Then
Text8.Text = Combo1.Text
End If

Select Case Check1.Value
    Case 0
            datagridsize (1)
            If Text8.Text = "ALL" Then
            INSERT_STAT = "SELECT MAIN,COMMENT FROM PIN ORDER BY SEVERITY,PRIORITY,DATE_STRING DESC"
            ElseIf Text8.Text = "1S" Then
            INSERT_STAT = "SELECT MAIN,COMMENT FROM PIN where SEVERITY='1S' ORDER BY SEVERITY,PRIORITY,DATE_STRING DESC"
            ElseIf Text8.Text = "2S" Then
            INSERT_STAT = "SELECT MAIN,COMMENT FROM PIN where SEVERITY='2S' ORDER BY SEVERITY,PRIORITY,DATE_STRING DESC"
            ElseIf Text8.Text = "3D" Then
            INSERT_STAT = "SELECT MAIN,COMMENT FROM PIN where SEVERITY='3D' ORDER BY SEVERITY,PRIORITY,DATE_STRING DESC"
            ElseIf Text8.Text = "4M" Then
            INSERT_STAT = "SELECT MAIN,COMMENT FROM PIN where SEVERITY='4M' ORDER BY SEVERITY,PRIORITY,DATE_STRING DESC"
            ElseIf Text8.Text = "5K" Then
            INSERT_STAT = "SELECT MAIN,COMMENT FROM PIN where SEVERITY='5K' ORDER BY SEVERITY,PRIORITY,DATE_STRING DESC"
            ElseIf Text8.Text = "6L" Then
            INSERT_STAT = "SELECT MAIN,COMMENT FROM PIN where SEVERITY='6L' ORDER BY SEVERITY,PRIORITY,DATE_STRING DESC"
            ElseIf Text8.Text = "7N" Then
            INSERT_STAT = "SELECT MAIN,COMMENT FROM PIN where SEVERITY='7N' ORDER BY SEVERITY,PRIORITY,DATE_STRING DESC"
            ElseIf Text8.Text = "8N" Then
            INSERT_STAT = "SELECT MAIN,COMMENT FROM PIN where SEVERITY='8N' ORDER BY SEVERITY,PRIORITY,DATE_STRING DESC"
            ElseIf Text8.Text = "9PENDING" Then
            INSERT_STAT = "SELECT MAIN,COMMENT FROM PIN where SEVERITY='9PENDING' ORDER BY SEVERITY,PRIORITY,DATE_STRING DESC"
            ElseIf Text8.Text = "0S" Then
            INSERT_STAT = "SELECT MAIN,COMMENT FROM PIN where SEVERITY='0S' ORDER BY SEVERITY,PRIORITY,DATE_STRING DESC"
            End If
    Case 1
            datagridsize (0)
            If Text8.Text = "ALL" Then
            INSERT_STAT = "SELECT MAIN,COMMENT,PRIORITY,SEVERITY FROM PIN ORDER BY SEVERITY,PRIORITY,DATE_STRING DESC"
            ElseIf Text8.Text = "1S" Then
            INSERT_STAT = "SELECT MAIN,COMMENT,PRIORITY,SEVERITY FROM PIN where SEVERITY='1S' ORDER BY SEVERITY,PRIORITY,DATE_STRING DESC"
            ElseIf Text8.Text = "2S" Then
            INSERT_STAT = "SELECT MAIN,COMMENT,PRIORITY,SEVERITY FROM PIN where SEVERITY='2S' ORDER BY SEVERITY,PRIORITY,DATE_STRING DESC"
            ElseIf Text8.Text = "3D" Then
            INSERT_STAT = "SELECT MAIN,COMMENT,PRIORITY,SEVERITY FROM PIN where SEVERITY='3D' ORDER BY SEVERITY,PRIORITY,DATE_STRING DESC"
            ElseIf Text8.Text = "4M" Then
            INSERT_STAT = "SELECT MAIN,COMMENT,PRIORITY,SEVERITY FROM PIN where SEVERITY='4M' ORDER BY SEVERITY,PRIORITY,DATE_STRING DESC"
            ElseIf Text8.Text = "5K" Then
            INSERT_STAT = "SELECT MAIN,COMMENT,PRIORITY,SEVERITY FROM PIN where SEVERITY='5K' ORDER BY SEVERITY,PRIORITY,DATE_STRING DESC"
            ElseIf Text8.Text = "6L" Then
            INSERT_STAT = "SELECT MAIN,COMMENT,PRIORITY,SEVERITY FROM PIN where SEVERITY='6L' ORDER BY SEVERITY,PRIORITY,DATE_STRING DESC"
            ElseIf Text8.Text = "7N" Then
            INSERT_STAT = "SELECT MAIN,COMMENT,PRIORITY,SEVERITY FROM PIN where SEVERITY='7N' ORDER BY SEVERITY,PRIORITY,DATE_STRING DESC"
            ElseIf Text8.Text = "8N" Then
            INSERT_STAT = "SELECT MAIN,COMMENT,PRIORITY,SEVERITY FROM PIN where SEVERITY='8N' ORDER BY SEVERITY,PRIORITY,DATE_STRING DESC"
            ElseIf Text8.Text = "9PENDING" Then
            INSERT_STAT = "SELECT MAIN,COMMENT,PRIORITY,SEVERITY FROM PIN where SEVERITY='9PENDING' ORDER BY SEVERITY,PRIORITY,DATE_STRING DESC"
            ElseIf Text8.Text = "0S" Then
            INSERT_STAT = "SELECT MAIN,COMMENT,PRIORITY,SEVERITY FROM PIN where SEVERITY='0S' ORDER BY SEVERITY,PRIORITY,DATE_STRING DESC"
            End If

    Case Else
End Select
    

Dim DATABASE_FILE As String
DATABASE_FILE = App.Path & "/DB.MDB"
Adodc1.ConnectionString = _
"PROVIDER=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & DATABASE_FILE & ";"
Adodc1.RecordSource = INSERT_STAT
' Bind the ADODC to the DataGrid.
DataGrid1.Refresh
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1

S2 = App.Path & "\TXTS\temp.TXT"
D2 = App.Path & "\TXTS\" & UCase(Trim(Text1.Text)) & ".TXT"
FileCopy S2, D2


S1 = App.Path & "\DOCUMENTS\TEMP.doc"
D1 = App.Path & "\DOCUMENTS\" & UCase(Trim(Text1.Text)) & ".doc"
FileCopy S1, D1

S2 = App.Path & "\DOCUMENTS\TEMP.xls"
D2 = App.Path & "\DOCUMENTS\" & UCase(Trim(Text1.Text)) & ".xls"
FileCopy S2, D2

S2 = App.Path & "\DOCUMENTS\TEMP.vsd"
D2 = App.Path & "\DOCUMENTS\" & UCase(Trim(Text1.Text)) & ".vsd"
FileCopy S2, D2

Text1.Text = ""
Text2.Text = ""
Text10.Text = ""
GoTo TELOS:

er1:
 MsgBox ("дем дысате сыста то педио PRIORITY"), vbCritical, "пяосовг !!!"
 GoTo TELOS:

TELOS:

End Sub

Private Sub Command10_Click()

Select Case Check1.Value
    Case 0
        datagridsize (1)
        statement_val = "SELECT MAIN,COMMENT FROM PIN where SEVERITY='2S' order by SEVERITY,PRIORITY,DATE_STRING DESC"
    Case 1
        datagridsize (0)
        statement_val = "SELECT MAIN,COMMENT,PRIORITY,SEVERITY FROM PIN where SEVERITY='2S' order by SEVERITY,PRIORITY,DATE_STRING DESC"
    Case Else
End Select

Dim DATABASE_FILE As String
DATABASE_FILE = App.Path & "/DB.MDB"
Adodc1.ConnectionString = _
"PROVIDER=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & DATABASE_FILE & ";"
Adodc1.RecordSource = statement_val
' Bind the ADODC to the DataGrid.

DataGrid1.Refresh
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1

Text8.Text = "2S"
Combo1.Text = "2S"
End Sub

Private Sub Command11_Click()

Select Case Check1.Value
    Case 0
        datagridsize (1)
        statement_val = "SELECT MAIN,COMMENT FROM PIN where SEVERITY='4M' order by SEVERITY,PRIORITY,DATE_STRING DESC"
    Case 1
        datagridsize (0)
        statement_val = "SELECT MAIN,COMMENT,PRIORITY,SEVERITY FROM PIN where SEVERITY='4M' order by SEVERITY,PRIORITY,DATE_STRING DESC"
    Case Else
End Select

Dim DATABASE_FILE As String
DATABASE_FILE = App.Path & "/DB.MDB"
Adodc1.ConnectionString = _
"PROVIDER=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & DATABASE_FILE & ";"
Adodc1.RecordSource = statement_val
' Bind the ADODC to the DataGrid.

DataGrid1.Refresh
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1

Text8.Text = "4M"
Combo1.Text = "4M"
End Sub

Private Sub Command12_Click()

Select Case Check1.Value
    Case 0
        datagridsize (1)
        statement_val = "SELECT MAIN,COMMENT FROM PIN order by SEVERITY,PRIORITY,DATE_STRING DESC"
    Case 1
        datagridsize (0)
        statement_val = "SELECT MAIN,COMMENT,PRIORITY,SEVERITY FROM PIN   order by SEVERITY,PRIORITY,DATE_STRING DESC"
    Case Else
End Select



Dim DATABASE_FILE As String
DATABASE_FILE = App.Path & "/DB.MDB"
Adodc1.ConnectionString = _
"PROVIDER=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & DATABASE_FILE & ";"
Adodc1.RecordSource = statement_val
' Bind the ADODC to the DataGrid.

DataGrid1.Refresh
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1

Text8.Text = "ALL"
Combo1.Text = "1S"
End Sub

Private Sub Command13_Click()
Load Form6
Form6.Show
Form1.Hide

End Sub





Private Sub Command14_Click()
Select Case Check1.Value
    Case 0
        datagridsize (1)
        statement_val = "SELECT MAIN,COMMENT FROM PIN where SEVERITY='8N' order by SEVERITY,PRIORITY,DATE_STRING DESC"
    Case 1
        datagridsize (0)
        statement_val = "SELECT MAIN,COMMENT,PRIORITY,SEVERITY FROM PIN where SEVERITY='8N' order by SEVERITY,PRIORITY,DATE_STRING DESC"
    Case Else
End Select



Dim DATABASE_FILE As String
DATABASE_FILE = App.Path & "/DB.MDB"
Adodc1.ConnectionString = _
"PROVIDER=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & DATABASE_FILE & ";"
Adodc1.RecordSource = statement_val
' Bind the ADODC to the DataGrid.

DataGrid1.Refresh
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1
Text8.Text = "8N"
Combo1.Text = "8N"
End Sub

Private Sub Command15_Click()
Select Case Check1.Value
    Case 0
        datagridsize (1)
        statement_val = "SELECT MAIN,COMMENT FROM PIN where SEVERITY='5K' order by SEVERITY,PRIORITY,PRIORITY,DATE_STRING DESC"
    Case 1
        datagridsize (0)
        statement_val = "SELECT MAIN,COMMENT,PRIORITY,SEVERITY FROM PIN where SEVERITY='5K' order by SEVERITY,PRIORITY,DATE_STRING DESC"
    Case Else
End Select

Dim DATABASE_FILE As String
DATABASE_FILE = App.Path & "/DB.MDB"
Adodc1.ConnectionString = _
"PROVIDER=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & DATABASE_FILE & ";"
Adodc1.RecordSource = statement_val
' Bind the ADODC to the DataGrid.

DataGrid1.Refresh
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1

Text8.Text = "5K"
Combo1.Text = "5K"
End Sub

Private Sub Command16_Click()
Select Case Check1.Value
    Case 0
        datagridsize (1)
        statement_val = "SELECT MAIN,COMMENT FROM PIN where SEVERITY='6L' order by SEVERITY,PRIORITY,DATE_STRING DESC"
    Case 1
        datagridsize (0)
        statement_val = "SELECT MAIN,COMMENT,PRIORITY,SEVERITY FROM PIN where SEVERITY='6L' order by SEVERITY,PRIORITY,DATE_STRING DESC"
    Case Else
End Select

Dim DATABASE_FILE As String
DATABASE_FILE = App.Path & "/DB.MDB"
Adodc1.ConnectionString = _
"PROVIDER=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & DATABASE_FILE & ";"
Adodc1.RecordSource = statement_val
' Bind the ADODC to the DataGrid.

DataGrid1.Refresh
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1

Text8.Text = "6L"
Combo1.Text = "6L"
End Sub

Private Sub Command17_Click()
Select Case Check1.Value
    Case 0
        datagridsize (1)
        statement_val = "SELECT MAIN,COMMENT FROM PIN where SEVERITY='9PENDING' order by SEVERITY,PRIORITY,DATE_STRING DESC"
    Case 1
        datagridsize (0)
        statement_val = "SELECT MAIN,COMMENT,PRIORITY,SEVERITY FROM PIN where SEVERITY='9PENDING' order by SEVERITY,PRIORITY,DATE_STRING DESC"
    Case Else
End Select



Dim DATABASE_FILE As String
DATABASE_FILE = App.Path & "/DB.MDB"
Adodc1.ConnectionString = _
"PROVIDER=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & DATABASE_FILE & ";"
Adodc1.RecordSource = statement_val
' Bind the ADODC to the DataGrid.

DataGrid1.Refresh
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1
Text8.Text = "9PENDING"
Combo1.Text = "9PENDING"
End Sub

Private Sub Command18_Click()
Select Case Check1.Value
    Case 0
        datagridsize (1)
        statement_val = "SELECT MAIN,COMMENT FROM PIN where SEVERITY='0S' order by SEVERITY,PRIORITY,DATE_STRING DESC"
    Case 1
        datagridsize (0)
        statement_val = "SELECT MAIN,COMMENT,PRIORITY,SEVERITY FROM PIN where SEVERITY='0S' order by SEVERITY,PRIORITY,DATE_STRING DESC"
    Case Else
End Select

Dim DATABASE_FILE As String
DATABASE_FILE = App.Path & "/DB.MDB"
Adodc1.ConnectionString = _
"PROVIDER=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & DATABASE_FILE & ";"
Adodc1.RecordSource = statement_val
' Bind the ADODC to the DataGrid.

DataGrid1.Refresh
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1


Text8.Text = "0S"
Combo1.Text = "0S"
End Sub

Private Sub Command2_Click()
On Error GoTo TELOS:

Dim priorvalue As Integer
If Text10.Text = "" Then
    priorvalue = 10
Else
    If IsNumeric(Text10.Text) = False Then
        GoTo er1:
    Else
        priorvalue = Text10.Text
    End If
End If

If RS.State = 1 Then RS.Close
If DB.State = 1 Then DB.Close
DB.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & "\DB.mdb" & ";" & _
      "Persist Security Info=False"
DB.Open App.Path & "\DB.mdb"
RS.Open "[PIN]", DB, adOpenDynamic, adLockBatchOptimistic

Dim S1, S2 As String
S1 = "DELETE FROM PIN WHERE MAIN='" & Text4.Text & "'"
DB.Execute S1



S2 = "INSERT INTO PIN(MAIN,COMMENT,SEVERITY,PRIORITY,DATE_STRING)" & _
    "VALUES (" & _
        "'" & UCase(Form1.Text1.Text) & "'," & _
        "'" & UCase(Form1.Text2.Text) & "', " & _
        "'" & Combo1.Text & "'," & _
        "'" & priorvalue & "'," & _
        "'" & Date & "')"
DB.Execute S2

txt_name = App.Path & "\TXTS\" & Text4.Text & ".TXT"
new_txt_name = App.Path & "\TXTS\" & UCase(Trim(Text1.Text)) & ".TXT"
If Dir(txt_name) <> "" Then
    'Kill txt_name
     Name txt_name As new_txt_name
Else
    SOURCE2 = App.Path & "\TXTS\temp.TXT"
    DESTIN2 = App.Path & "\TXTS\" & UCase(Trim(Text1.Text)) & ".TXT"
    FileCopy SOURCE2, DESTIN2
End If


doc_name = App.Path & "\DOCUMENTS\" & Text4.Text & ".doc"
new_doc_name = App.Path & "\DOCUMENTS\" & UCase(Trim(Text1.Text)) & ".doc"
If Dir(doc_name) <> "" Then
    'Kill txt_name
     Name doc_name As new_doc_name
Else
    SOURCE2 = App.Path & "\DOCUMENTS\temp.doc"
    DESTIN2 = App.Path & "\DOCUMENTS\" & UCase(Trim(Text1.Text)) & ".doc"
    FileCopy SOURCE2, DESTIN2
End If

xls_name = App.Path & "\DOCUMENTS\" & Text4.Text & ".xls"
new_xls_name = App.Path & "\DOCUMENTS\" & UCase(Trim(Text1.Text)) & ".xls"
If Dir(xls_name) <> "" Then
    'Kill txt_name
     Name xls_name As new_xls_name
Else
    SOURCE2 = App.Path & "\DOCUMENTS\temp.xls"
    DESTIN2 = App.Path & "\DOCUMENTS\" & UCase(Trim(Text1.Text)) & ".xls"
    FileCopy SOURCE2, DESTIN2
End If

vsd_name = App.Path & "\DOCUMENTS\" & Text4.Text & ".vsd"
new_vsd_name = App.Path & "\DOCUMENTS\" & UCase(Trim(Text1.Text)) & ".vsd"
If Dir(vsd_name) <> "" Then
    'Kill txt_name
     Name vsd_name As new_vsd_name
Else
    SOURCE2 = App.Path & "\DOCUMENTS\temp.vsd"
    DESTIN2 = App.Path & "\DOCUMENTS\" & UCase(Trim(Text1.Text)) & ".vsd"
    FileCopy SOURCE2, DESTIN2
End If



RS.Fields.Refresh
RS.Close
DB.Close


If Text8.Text <> "ALL" Then
Text8.Text = Combo1.Text
End If

Select Case Check1.Value
Case 0
            datagridsize (1)
            If Text8.Text = "ALL" Then
            UPDATE_STAT = "SELECT MAIN,COMMENT FROM PIN ORDER BY SEVERITY,PRIORITY,DATE_STRING DESC"
            ElseIf Text8.Text = "1S" Then
            UPDATE_STAT = "SELECT MAIN,COMMENT FROM PIN where SEVERITY='1S' ORDER BY SEVERITY,PRIORITY,DATE_STRING DESC"
            ElseIf Text8.Text = "2S" Then
            UPDATE_STAT = "SELECT MAIN,COMMENT FROM PIN where SEVERITY='2S' ORDER BY SEVERITY,PRIORITY,DATE_STRING DESC"
            ElseIf Text8.Text = "3D" Then
            UPDATE_STAT = "SELECT MAIN,COMMENT FROM PIN where SEVERITY='3D' ORDER BY SEVERITY,PRIORITY,DATE_STRING DESC"
            ElseIf Text8.Text = "4M" Then
            UPDATE_STAT = "SELECT MAIN,COMMENT FROM PIN where SEVERITY='4M' ORDER BY SEVERITY,PRIORITY,DATE_STRING DESC"
            ElseIf Text8.Text = "5K" Then
            UPDATE_STAT = "SELECT MAIN,COMMENT FROM PIN where SEVERITY='5K' ORDER BY SEVERITY,PRIORITY,DATE_STRING DESC"
            ElseIf Text8.Text = "6L" Then
            UPDATE_STAT = "SELECT MAIN,COMMENT FROM PIN where SEVERITY='6L' ORDER BY SEVERITY,PRIORITY,DATE_STRING DESC"
            ElseIf Text8.Text = "7N" Then
            UPDATE_STAT = "SELECT MAIN,COMMENT FROM PIN where SEVERITY='7N' ORDER BY SEVERITY,PRIORITY,DATE_STRING DESC"
            ElseIf Text8.Text = "8N" Then
            UPDATE_STAT = "SELECT MAIN,COMMENT FROM PIN where SEVERITY='8N' ORDER BY SEVERITY,PRIORITY,DATE_STRING DESC"
            ElseIf Text8.Text = "9PENDING" Then
            UPDATE_STAT = "SELECT MAIN,COMMENT FROM PIN where SEVERITY='9PENDING' ORDER BY SEVERITY,PRIORITY,DATE_STRING DESC"
            ElseIf Text8.Text = "0S" Then
            UPDATE_STAT = "SELECT MAIN,COMMENT FROM PIN where SEVERITY='0S' ORDER BY SEVERITY,PRIORITY,DATE_STRING DESC"
            End If
Case 1
            datagridsize (0)
            If Text8.Text = "ALL" Then
            UPDATE_STAT = "SELECT MAIN,COMMENT,PRIORITY,SEVERITY FROM PIN ORDER BY SEVERITY,PRIORITY,DATE_STRING DESC"
            ElseIf Text8.Text = "1S" Then
            UPDATE_STAT = "SELECT MAIN,COMMENT,PRIORITY,SEVERITY FROM PIN where SEVERITY='1S' ORDER BY SEVERITY,PRIORITY,DATE_STRING DESC"
            ElseIf Text8.Text = "2S" Then
            UPDATE_STAT = "SELECT MAIN,COMMENT,PRIORITY,SEVERITY FROM PIN where SEVERITY='2S' ORDER BY SEVERITY,PRIORITY,DATE_STRING DESC"
            ElseIf Text8.Text = "3D" Then
            UPDATE_STAT = "SELECT MAIN,COMMENT,PRIORITY,SEVERITY FROM PIN where SEVERITY='3D' ORDER BY SEVERITY,PRIORITY,DATE_STRING DESC"
            ElseIf Text8.Text = "4M" Then
            UPDATE_STAT = "SELECT MAIN,COMMENT,PRIORITY,SEVERITY FROM PIN where SEVERITY='4M' ORDER BY SEVERITY,PRIORITY,DATE_STRING DESC"
            ElseIf Text8.Text = "5K" Then
            UPDATE_STAT = "SELECT MAIN,COMMENT,PRIORITY,SEVERITY FROM PIN where SEVERITY='5K' ORDER BY SEVERITY,PRIORITY,DATE_STRING DESC"
            ElseIf Text8.Text = "6L" Then
            UPDATE_STAT = "SELECT MAIN,COMMENT,PRIORITY,SEVERITY  FROM PIN where SEVERITY='6L' ORDER BY SEVERITY,PRIORITY,DATE_STRING DESC"
            ElseIf Text8.Text = "7N" Then
            UPDATE_STAT = "SELECT MAIN,COMMENT,PRIORITY,SEVERITY  FROM PIN where SEVERITY='7N' ORDER BY SEVERITY,PRIORITY,DATE_STRING DESC"
            ElseIf Text8.Text = "8N" Then
            UPDATE_STAT = "SELECT MAIN,COMMENT,PRIORITY,SEVERITY  FROM PIN where SEVERITY='8N' ORDER BY SEVERITY,PRIORITY,DATE_STRING DESC"
            ElseIf Text8.Text = "9PENDING" Then
            UPDATE_STAT = "SELECT MAIN,COMMENT,PRIORITY,SEVERITY  FROM PIN where SEVERITY='9PENDING' ORDER BY SEVERITY,PRIORITY,DATE_STRING DESC"
            ElseIf Text8.Text = "0S" Then
            UPDATE_STAT = "SELECT MAIN,COMMENT,PRIORITY,SEVERITY  FROM PIN where SEVERITY='0S' ORDER BY SEVERITY,PRIORITY,DATE_STRING DESC"
            End If
Case Else
End Select

Dim DATABASE_FILE As String
DATABASE_FILE = App.Path & "/DB.MDB"
Adodc1.ConnectionString = _
"PROVIDER=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & DATABASE_FILE & ";"
Adodc1.RecordSource = UPDATE_STAT
' Bind the ADODC to the DataGrid.


DataGrid1.Refresh
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1


Text1.Text = ""
Text2.Text = ""
Text10.Text = ""
GoTo TELOS:

er1:
 MsgBox ("дем дысате сыста то педио PRIORITY"), vbCritical, "пяосовг !!!"
 GoTo TELOS:

TELOS:

End Sub

Private Sub Command3_Click()
'On Error GoTo TELOS:

If RS.State = 1 Then RS.Close
If DB.State = 1 Then DB.Close
DB.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & "\DB.mdb" & ";" & _
      "Persist Security Info=False"
DB.Open App.Path & "\DB.mdb"
RS.Open "[PIN]", DB, adOpenDynamic, adLockBatchOptimistic

Dim statement As String
statement = "DELETE FROM PIN WHERE MAIN='" & Text1.Text & "'"
DB.Execute statement

RS.Fields.Refresh
RS.Close
DB.Close

If Text8.Text <> "ALL" Then
Text8.Text = Combo1.Text
End If



Select Case Check1.Value
    Case 0
            datagridsize (1)
            If Text8.Text = "ALL" Then
            DELETE_STAT = "SELECT MAIN,COMMENT FROM PIN ORDER BY SEVERITY,PRIORITY,DATE_STRING DESC"
            ElseIf Text8.Text = "1S" Then
            DELETE_STAT = "SELECT MAIN,COMMENT FROM PIN where SEVERITY='1S' ORDER BY SEVERITY,PRIORITY,DATE_STRING DESC"
            ElseIf Text8.Text = "2S" Then
            DELETE_STAT = "SELECT MAIN,COMMENT FROM PIN where SEVERITY='2S' ORDER BY SEVERITY,PRIORITY,DATE_STRING DESC"
            ElseIf Text8.Text = "3D" Then
            DELETE_STAT = "SELECT MAIN,COMMENT FROM PIN where SEVERITY='3D' ORDER BY SEVERITY,PRIORITY,DATE_STRING DESC"
            ElseIf Text8.Text = "4M" Then
            DELETE_STAT = "SELECT MAIN,COMMENT FROM PIN where SEVERITY='4M' ORDER BY SEVERITY,PRIORITY,DATE_STRING DESC"
            ElseIf Text8.Text = "5K" Then
            DELETE_STAT = "SELECT MAIN,COMMENT FROM PIN where SEVERITY='5K' ORDER BY SEVERITY,PRIORITY,DATE_STRING DESC"
            ElseIf Text8.Text = "6L" Then
            DELETE_STAT = "SELECT MAIN,COMMENT FROM PIN where SEVERITY='6L' ORDER BY SEVERITY,PRIORITY,DATE_STRING DESC"
            ElseIf Text8.Text = "7N" Then
            DELETE_STAT = "SELECT MAIN,COMMENT FROM PIN where SEVERITY='7N' ORDER BY SEVERITY,PRIORITY,DATE_STRING DESC"
            ElseIf Text8.Text = "8N" Then
            DELETE_STAT = "SELECT MAIN,COMMENT FROM PIN where SEVERITY='8N' ORDER BY SEVERITY,PRIORITY,DATE_STRING DESC"
            ElseIf Text8.Text = "9PENDING" Then
            DELETE_STAT = "SELECT MAIN,COMMENT FROM PIN where SEVERITY='9PENDING' ORDER BY SEVERITY,PRIORITY,DATE_STRING DESC"
            ElseIf Text8.Text = "0S" Then
            DELETE_STAT = "SELECT MAIN,COMMENT FROM PIN where SEVERITY='0S' ORDER BY SEVERITY,PRIORITY,DATE_STRING DESC"
            End If
    Case 1
            datagridsize (0)
            If Text8.Text = "ALL" Then
            DELETE_STAT = "SELECT MAIN,COMMENT,PRIORITY,SEVERITY FROM PIN ORDER BY SEVERITY,PRIORITY,DATE_STRING DESC"
            ElseIf Text8.Text = "1S" Then
            DELETE_STAT = "SELECT MAIN,COMMENT,PRIORITY,SEVERITY FROM PIN where SEVERITY='1S' ORDER BY SEVERITY,PRIORITY,DATE_STRING DESC"
            ElseIf Text8.Text = "2S" Then
            DELETE_STAT = "SELECT MAIN,COMMENT,PRIORITY,SEVERITY FROM PIN where SEVERITY='2S' ORDER BY SEVERITY,PRIORITY,DATE_STRING DESC"
            ElseIf Text8.Text = "3D" Then
            DELETE_STAT = "SELECT MAIN,COMMENT,PRIORITY,SEVERITY FROM PIN where SEVERITY='3D' ORDER BY SEVERITY,PRIORITY,DATE_STRING DESC"
            ElseIf Text8.Text = "4M" Then
            DELETE_STAT = "SELECT MAIN,COMMENT,PRIORITY,SEVERITY FROM PIN where SEVERITY='4M' ORDER BY SEVERITY,PRIORITY,DATE_STRING DESC"
            ElseIf Text8.Text = "5K" Then
            DELETE_STAT = "SELECT MAIN,COMMENT,PRIORITY,SEVERITY FROM PIN where SEVERITY='5K' ORDER BY SEVERITY,PRIORITY,DATE_STRING DESC"
            ElseIf Text8.Text = "6L" Then
            DELETE_STAT = "SELECT MAIN,COMMENT,PRIORITY,SEVERITY FROM PIN where SEVERITY='6L' ORDER BY SEVERITY,PRIORITY,DATE_STRING DESC"
            ElseIf Text8.Text = "7N" Then
            DELETE_STAT = "SELECT MAIN,COMMENT,PRIORITY,SEVERITY FROM PIN where SEVERITY='7N' ORDER BY SEVERITY,PRIORITY,DATE_STRING DESC"
            ElseIf Text8.Text = "8N" Then
            DELETE_STAT = "SELECT MAIN,COMMENT,PRIORITY,SEVERITY FROM PIN where SEVERITY='8N' ORDER BY SEVERITY,PRIORITY,DATE_STRING DESC"
            ElseIf Text8.Text = "9PENDING" Then
            DELETE_STAT = "SELECT MAIN,COMMENT,PRIORITY,SEVERITY FROM PIN where SEVERITY='9PENDING' ORDER BY SEVERITY,PRIORITY,DATE_STRING DESC"
            ElseIf Text8.Text = "0S" Then
            DELETE_STAT = "SELECT MAIN,COMMENT,PRIORITY,SEVERITY FROM PIN where SEVERITY='0S' ORDER BY SEVERITY,PRIORITY,DATE_STRING DESC"
            End If
    Case Else
End Select



Dim DATABASE_FILE As String
DATABASE_FILE = App.Path & "/DB.MDB"
Adodc1.ConnectionString = _
"PROVIDER=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & DATABASE_FILE & ";"
Adodc1.RecordSource = DELETE_STAT
' Bind the ADODC to the DataGrid.


DataGrid1.Refresh
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1

txt_name = App.Path & "\TXTS\" & UCase(Trim(Text1.Text)) & ".TXT"
If Dir(txt_name) <> "" Then
    Kill txt_name
End If

doc_name = App.Path & "\DOCUMENTS\" & UCase(Trim(Text1.Text)) & ".doc"
If Dir(doc_name) <> "" Then
    Kill doc_name
End If
xls_name = App.Path & "\DOCUMENTS\" & UCase(Trim(Text1.Text)) & ".xls"
If Dir(xls_name) <> "" Then
    Kill xls_name
End If
vsd_name = App.Path & "\DOCUMENTS\" & UCase(Trim(Text1.Text)) & ".vsd"
If Dir(vsd_name) <> "" Then
    Kill vsd_name
End If

Text1.Text = ""
Text2.Text = ""
Text10.Text = ""




TELOS:
End Sub

Private Sub Command4_Click()
Form1.Hide
Load Form2
Form2.Show
End Sub

Private Sub Command5_Click()
Load Form3
Form3.Show
End Sub




Private Sub Command6_Click()
Select Case Check1.Value
    Case 0
        datagridsize (1)
        statement_val = "SELECT MAIN,COMMENT FROM PIN where SEVERITY='7N' order by SEVERITY,PRIORITY,DATE_STRING DESC"
    Case 1
        datagridsize (0)
        statement_val = "SELECT MAIN,COMMENT,PRIORITY,SEVERITY FROM PIN where SEVERITY='7N' order by SEVERITY,PRIORITY,DATE_STRING DESC"
    Case Else
End Select


Dim DATABASE_FILE As String
DATABASE_FILE = App.Path & "/DB.MDB"
Adodc1.ConnectionString = _
"PROVIDER=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & DATABASE_FILE & ";"
Adodc1.RecordSource = statement_val
' Bind the ADODC to the DataGrid.

DataGrid1.Refresh
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1
Text8.Text = "7N"
Combo1.Text = "7N"
End Sub

Private Sub Command7_Click()
STAT_EMFANISIS = "SELECT KEIMENO FROM PIN  where MERA=" & CInt(Day(Date)) & " AND MHNAS=" & CInt(Month(Date))
Load Form4
Form4.Show
End Sub

Private Sub Command8_Click()

Select Case Check1.Value
    Case 0
        datagridsize (1)
        statement_val = "SELECT MAIN,COMMENT FROM PIN where SEVERITY='1S' order by SEVERITY,PRIORITY,DATE_STRING DESC"
    Case 1
        datagridsize (0)
        statement_val = "SELECT MAIN,COMMENT,PRIORITY,SEVERITY FROM PIN where SEVERITY='1S' order by SEVERITY,PRIORITY,DATE_STRING DESC"
    Case Else
End Select

Dim DATABASE_FILE As String
DATABASE_FILE = App.Path & "/DB.MDB"
Adodc1.ConnectionString = _
"PROVIDER=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & DATABASE_FILE & ";"
Adodc1.RecordSource = statement_val
' Bind the ADODC to the DataGrid.

DataGrid1.Refresh
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1


Text8.Text = "1S"
Combo1.Text = "1S"

End Sub

Private Sub Command9_Click()


Select Case Check1.Value
    Case 0
        datagridsize (1)
        statement_val = "SELECT MAIN,COMMENT FROM PIN where SEVERITY='3D' order by SEVERITY,PRIORITY,DATE_STRING DESC"
    Case 1
        datagridsize (0)
        statement_val = "SELECT MAIN,COMMENT,PRIORITY,SEVERITY FROM PIN where SEVERITY='3D' order by SEVERITY,PRIORITY,DATE_STRING DESC"
    Case Else
End Select



Dim DATABASE_FILE As String
DATABASE_FILE = App.Path & "/DB.MDB"
Adodc1.ConnectionString = _
"PROVIDER=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & DATABASE_FILE & ";"
Adodc1.RecordSource = statement_val
' Bind the ADODC to the DataGrid.

DataGrid1.Refresh
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1
Text8.Text = "3D"
Combo1.Text = "3D"

End Sub

Private Sub DataGrid1_Click()
On Error GoTo TELOS:
Text1.Text = DataGrid1.Columns(0).Text
Text2.Text = DataGrid1.Columns(1).Text
Text4.Text = Text1.Text



Dim DATABASE_FILE As String
DATABASE_FILE = App.Path & "/DB.MDB"
Adodc2.ConnectionString = _
"PROVIDER=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & DATABASE_FILE & ";"
Adodc2.RecordSource = "SELECT SEVERITY,PRIORITY FROM PIN WHERE MAIN='" & Text1.Text & "' AND COMMENT ='" & Text2.Text & "';"
Set DataGrid2.DataSource = Adodc2
Combo1.Text = DataGrid2.Columns(0).Text
Text10.Text = DataGrid2.Columns(1).Text
DataGrid2.Refresh
Adodc2.Refresh
Set DataGrid2.DataSource = Adodc2

TELOS:

End Sub

Private Sub Form_Load()
On Error GoTo TELOS:
Unload Form3
Unload Form4
Text3.Text = Hour(Time)
Text5.Text = Minute(Time)
Text6.Text = Day(Date)
Text7.Text = Month(Date)
Text8.Text = "ALL"
Text9.Text = 0 'xreisimopoieitai gia na arxikopoihso thn katastash tou show all.

If RS.State = 1 Then RS.Close
If DB.State = 1 Then DB.Close
DB.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & "\DB.mdb" & ";" & _
      "Persist Security Info=False"
DB.Open App.Path & "\DB.mdb"
RS.Open "[PIN]", DB, adOpenDynamic, adLockBatchOptimistic

DataGrid1.Columns(0).Width = 0.1 * DataGrid1.Width
DataGrid1.Columns(1).Width = DataGrid1.Columns(0).Width


'DataGrid1.Font.Size = 13
'DataGrid1.DefColWidth = 6800
'DataGrid1.HeadFont.Bold = True
'DataGrid1.HeadFont.Size = 10
'DataGrid1.Font = "verdana"
'DataGrid1.Height = 9405

datagridsize (1)

'DataGrid1.Font.Size = 13
'DataGrid1.DefColWidth = 3400
'DataGrid1.HeadFont.Bold = True
'DataGrid1.HeadFont.Size = 10
'DataGrid1.Font = "verdana"
'DataGrid1.Height = 9405


'DataGrid1.Width = Form1.ScaleWidth
'DataGrid1.Columns(0).Width = 0.01 * DataGrid1.Width
'DataGrid1.Columns(1).Width = 1


Dim DATABASE_FILE As String
DATABASE_FILE = App.Path & "/DB.MDB"
Adodc1.ConnectionString = _
"PROVIDER=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & DATABASE_FILE & ";"
Adodc1.RecordSource = "SELECT MAIN,COMMENT FROM PIN ORDER BY SEVERITY,PRIORITY,DATE_STRING DESC"
' Bind the ADODC to the DataGrid.
Set DataGrid1.DataSource = Adodc1
'Text3.Text = Adodc1.Recordset.RecordCount
'If Text3.Text <= 33 Then
'    DataGrid1.Height = 327.059 + (CInt(Text3.Text) * 327.059)
'Else
'    DataGrid1.Height = 11120
'End If

Combo1.AddItem "0S"
Combo1.AddItem "1S"
Combo1.AddItem "2S"
Combo1.AddItem "3D"
Combo1.AddItem "4M"
Combo1.AddItem "5K"
Combo1.AddItem "6L"
Combo1.AddItem "7N"
Combo1.AddItem "8N"
Combo1.AddItem "9PENDING"
Combo1.Text = "1S"


TELOS:
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Image1_Click()
Load Form5
Form5.Show
End Sub

Private Sub Image1_DblClick()
Load Form5
Form5.Show


End Sub

Private Sub Image2_Click()
Load Form5
Form5.Show
End Sub

Private Sub Image2_DblClick()
Load Form5
Form5.Show
End Sub

Private Sub OLE5_Click()
OLE5.SourceDoc = App.Path & "\DOCUMENTS\" & UCase(Text1.Text) & ".xls"
OLE5.CreateLink App.Path & "\DOCUMENTS\" & UCase(Text1.Text) & ".xls"
'Dim name As String
'name = App.Path & "\DOCUMENTS\" & UCase(Text1.Text) & ".xls"
'OLE2.CreateEmbed App.Path & "\DOCUMENTS\" & UCase(Text1.Text) & ".xlsx"










End Sub

Private Sub OLE6_Click()
OLE6.SourceDoc = App.Path & "\DOCUMENTS\" & UCase(Text1.Text) & ".doc"
OLE6.CreateLink App.Path & "\DOCUMENTS\" & UCase(Text1.Text) & ".doc"
End Sub

Private Sub OLE7_Click()
OLE7.SourceDoc = App.Path & "\DOCUMENTS\" & UCase(Text1.Text) & ".vsd"
OLE7.CreateLink App.Path & "\DOCUMENTS\" & UCase(Text1.Text) & ".vsd"
End Sub

Private Sub Picture1_Click()
Text1.Text = ""
Text2.Text = ""
Text10.Text = ""
End Sub

Private Sub Timer1_Timer()
Text3.Text = Hour(Time)
Text5.Text = Minute(Time)
Text6.Text = Day(Date)
Text7.Text = Month(Date)

If CInt(Text3.Text) = 10 Then
    If CInt(Text5.Text) = 0 Then
        STAT_EMFANISIS = "SELECT KEIMENO FROM PIN WHERE F1=1 AND MERA=" & CInt(Text6.Text) & " AND MHNAS=" & CInt(Text7.Text)
        If RS1.State = 1 Then RS1.Close
        If DB1.State = 1 Then DB1.Close
        Load Form4
        Form4.Show
    End If
End If

If CInt(Text3.Text) = 12 Then
    If CInt(Text5.Text) = 0 Then
        STAT_EMFANISIS = "SELECT KEIMENO FROM PIN WHERE F2=1 AND MERA=" & CInt(Text6.Text) & " AND MHNAS=" & CInt(Text7.Text)
        If RS1.State = 1 Then RS1.Close
        If DB1.State = 1 Then DB1.Close
        Load Form4
        Form4.Show
    End If
End If

If CInt(Text3.Text) = 15 Then
    If CInt(Text5.Text) = 0 Then
        STAT_EMFANISIS = "SELECT KEIMENO FROM PIN WHERE F3=1 AND MERA=" & CInt(Text6.Text) & " AND MHNAS=" & CInt(Text7.Text)
        If RS1.State = 1 Then RS1.Close
        If DB1.State = 1 Then DB1.Close
        Load Form4
        Form4.Show
    End If
End If

If CInt(Text3.Text) = 18 Then
    If CInt(Text5.Text) = 0 Then
        STAT_EMFANISIS = "SELECT KEIMENO FROM PIN WHERE F4=1 AND MERA=" & CInt(Text6.Text) & " AND MHNAS=" & CInt(Text7.Text)
        If RS1.State = 1 Then RS1.Close
        If DB1.State = 1 Then DB1.Close
        Load Form4
        Form4.Show
    End If
End If

If CInt(Text3.Text) = 21 Then
    If CInt(Text5.Text) = 0 Then
        STAT_EMFANISIS = "SELECT KEIMENO FROM PIN WHERE F5=1 AND MERA=" & CInt(Text6.Text) & " AND MHNAS=" & CInt(Text7.Text)
        If RS1.State = 1 Then RS1.Close
        If DB1.State = 1 Then DB1.Close
        Load Form4
        Form4.Show
    End If
End If

End Sub

Private Sub Timer2_Timer()

Dim txt_name As String
txt_name = App.Path & "\TXTS\" & Text1.Text & ".TXT"
If Dir(txt_name) <> "" Then
    Timer3.Enabled = False
    If FileLen(txt_name) > 2 Then
        Image1.Visible = True
        Image2.Visible = False
    Else
        Image1.Visible = False
        Image2.Visible = True
    End If
Else
    Timer3.Enabled = True
    GoTo TELOS:
End If

xls_name = App.Path & "\DOCUMENTS\" & UCase(Trim(Text1.Text)) & ".xls"
doc_name = App.Path & "\DOCUMENTS\" & UCase(Trim(Text1.Text)) & ".doc"
vsd_name = App.Path & "\DOCUMENTS\" & UCase(Trim(Text1.Text)) & ".vsd"

If Dir(xls_name) <> "" Then
    OLE5.Visible = True
Else
    OLE5.Visible = False
End If

If Dir(doc_name) <> "" Then
    OLE6.Visible = True
Else
    OLE6.Visible = False
End If

If Dir(vsd_name) <> "" Then
    OLE7.Visible = True
Else
    OLE7.Visible = False
End If

TELOS:

End Sub

Private Sub Timer3_Timer()
        Image1.Visible = False
        Image2.Visible = False
        OLE5.Visible = False
        OLE6.Visible = False
        OLE7.Visible = False
End Sub

Private Sub Timer4_Timer()
On Error GoTo TELOS:

If Text9.Text <> Check1.Value Then
If RS.State = 1 Then RS.Close
If DB.State = 1 Then DB.Close
DB.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & "\DB.mdb" & ";" & _
      "Persist Security Info=False"
DB.Open App.Path & "\DB.mdb"
RS.Open "[PIN]", DB, adOpenDynamic, adLockBatchOptimistic
RS.Fields.Refresh
RS.Close
DB.Close

If Text8.Text <> "ALL" Then
Text8.Text = Combo1.Text
End If

Select Case Check1.Value
    Case 0
        datagridsize (1)
        If Text8.Text = "ALL" Then
        INSERT_STAT = "SELECT MAIN,COMMENT FROM PIN ORDER BY SEVERITY,PRIORITY,DATE_STRING DESC"
        ElseIf Text8.Text = "1S" Then
        INSERT_STAT = "SELECT MAIN,COMMENT FROM PIN where SEVERITY='1S' ORDER BY SEVERITY,PRIORITY,DATE_STRING DESC"
        ElseIf Text8.Text = "2S" Then
        INSERT_STAT = "SELECT MAIN,COMMENT FROM PIN where SEVERITY='2S' ORDER BY SEVERITY,PRIORITY,DATE_STRING DESC"
        ElseIf Text8.Text = "3D" Then
        INSERT_STAT = "SELECT MAIN,COMMENT FROM PIN where SEVERITY='3D' ORDER BY SEVERITY,PRIORITY,DATE_STRING DESC"
        ElseIf Text8.Text = "4M" Then
        INSERT_STAT = "SELECT MAIN,COMMENT FROM PIN where SEVERITY='4M' ORDER BY SEVERITY,PRIORITY,DATE_STRING DESC"
        ElseIf Text8.Text = "5K" Then
        INSERT_STAT = "SELECT MAIN,COMMENT FROM PIN where SEVERITY='5K' ORDER BY SEVERITY,PRIORITY,DATE_STRING DESC"
        ElseIf Text8.Text = "6L" Then
        INSERT_STAT = "SELECT MAIN,COMMENT FROM PIN where SEVERITY='6L' ORDER BY SEVERITY,PRIORITY,DATE_STRING DESC"
                ElseIf Text8.Text = "7N" Then
        INSERT_STAT = "SELECT MAIN,COMMENT FROM PIN where SEVERITY='7N' ORDER BY SEVERITY,PRIORITY,DATE_STRING DESC"

            ElseIf Text8.Text = "8N" Then
        INSERT_STAT = "SELECT MAIN,COMMENT FROM PIN where SEVERITY='8N' ORDER BY SEVERITY,PRIORITY,DATE_STRING DESC"

            ElseIf Text8.Text = "9PENDING" Then
        INSERT_STAT = "SELECT MAIN,COMMENT FROM PIN where SEVERITY='9PENDING' ORDER BY SEVERITY,PRIORITY,DATE_STRING DESC"
            ElseIf Text8.Text = "0S" Then
        INSERT_STAT = "SELECT MAIN,COMMENT FROM PIN where SEVERITY='0S' ORDER BY SEVERITY,PRIORITY,DATE_STRING DESC"
        End If
    Case 1
        datagridsize (0)
        If Text8.Text = "ALL" Then
        INSERT_STAT = "SELECT MAIN,COMMENT,PRIORITY,SEVERITY FROM PIN ORDER BY SEVERITY,PRIORITY,DATE_STRING DESC"
        ElseIf Text8.Text = "1S" Then
        INSERT_STAT = "SELECT MAIN,COMMENT,PRIORITY,SEVERITY FROM PIN where SEVERITY='1S' ORDER BY SEVERITY,PRIORITY,DATE_STRING DESC"
        ElseIf Text8.Text = "2S" Then
        INSERT_STAT = "SELECT MAIN,COMMENT,PRIORITY,SEVERITY FROM PIN where SEVERITY='2S' ORDER BY SEVERITY,PRIORITY,DATE_STRING DESC"
        ElseIf Text8.Text = "3D" Then
        INSERT_STAT = "SELECT MAIN,COMMENT,PRIORITY,SEVERITY FROM PIN where SEVERITY='3D' ORDER BY SEVERITY,PRIORITY,DATE_STRING DESC"
        ElseIf Text8.Text = "4M" Then
        INSERT_STAT = "SELECT MAIN,COMMENT,PRIORITY,SEVERITY FROM PIN where SEVERITY='4M' ORDER BY SEVERITY,PRIORITY,DATE_STRING DESC"
        ElseIf Text8.Text = "5K" Then
        INSERT_STAT = "SELECT MAIN,COMMENT,PRIORITY,SEVERITY FROM PIN where SEVERITY='5K' ORDER BY SEVERITY,PRIORITY,DATE_STRING DESC"
        ElseIf Text8.Text = "6L" Then
        INSERT_STAT = "SELECT MAIN,COMMENT,PRIORITY,SEVERITY FROM PIN where SEVERITY='6L' ORDER BY SEVERITY,PRIORITY,DATE_STRING DESC"
                ElseIf Text8.Text = "7N" Then
        INSERT_STAT = "SELECT MAIN,COMMENT,PRIORITY,SEVERITY FROM PIN where SEVERITY='7N' ORDER BY SEVERITY,PRIORITY,DATE_STRING DESC"
        ElseIf Text8.Text = "8N" Then
        INSERT_STAT = "SELECT MAIN,COMMENT,PRIORITY,SEVERITY FROM PIN where SEVERITY='8N' ORDER BY SEVERITY,PRIORITY,DATE_STRING DESC"
        ElseIf Text8.Text = "9PENDING" Then
        INSERT_STAT = "SELECT MAIN,COMMENT,PRIORITY,SEVERITY FROM PIN where SEVERITY='9PENDING' ORDER BY SEVERITY,PRIORITY,DATE_STRING DESC"
        ElseIf Text8.Text = "0S" Then
        INSERT_STAT = "SELECT MAIN,COMMENT,PRIORITY,SEVERITY FROM PIN where SEVERITY='0S' ORDER BY SEVERITY,PRIORITY,DATE_STRING DESC"
        
        End If
    Case Else
End Select



Dim DATABASE_FILE As String
DATABASE_FILE = App.Path & "/DB.MDB"
Adodc1.ConnectionString = _
"PROVIDER=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & DATABASE_FILE & ";"
Adodc1.RecordSource = INSERT_STAT
' Bind the ADODC to the DataGrid.


DataGrid1.Refresh
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1

Text9.Text = Check1.Value
End If

TELOS:
End Sub
