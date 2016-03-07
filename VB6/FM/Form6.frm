VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Form6 
   Caption         =   "FM-DIARY"
   ClientHeight    =   8850
   ClientLeft      =   4065
   ClientTop       =   450
   ClientWidth     =   15855
   LinkTopic       =   "Form6"
   ScaleHeight     =   8850
   ScaleWidth      =   15855
   Begin VB.TextBox Text5 
      Height          =   615
      Left            =   8880
      TabIndex        =   15
      Text            =   "Text5"
      Top             =   3360
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   9720
      TabIndex        =   14
      Text            =   "Text4"
      Top             =   4080
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FF8080&
      Caption         =   "EXIT"
      Height          =   855
      Left            =   10680
      TabIndex        =   11
      Top             =   600
      Width           =   1695
   End
   Begin MSComCtl2.MonthView MonthView2 
      Height          =   2370
      Left            =   12720
      TabIndex        =   10
      Top             =   3600
      Width           =   2805
      _ExtentX        =   4948
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   83361794
      TitleBackColor  =   -2147483635
      CurrentDate     =   40467
      MinDate         =   -53684
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   10560
      TabIndex        =   9
      Top             =   6120
      Width           =   5055
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FF8080&
      Caption         =   "SHOW"
      Height          =   495
      Left            =   13320
      TabIndex        =   8
      Top             =   3000
      Width           =   1455
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   855
      Left            =   10920
      Top             =   1200
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1508
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
      Height          =   735
      Left            =   10920
      Top             =   360
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   1296
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
      Height          =   4815
      Left            =   360
      TabIndex        =   7
      Top             =   3840
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   8493
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
            Type            =   1
            Format          =   "d/M/yyyy"
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
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FF8080&
      Caption         =   "DELETE"
      Height          =   735
      Left            =   13680
      TabIndex        =   6
      Top             =   7800
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FF8080&
      Caption         =   "FIND"
      Height          =   735
      Left            =   12120
      TabIndex        =   5
      Top             =   7800
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF8080&
      Caption         =   "INSERT"
      Height          =   735
      Left            =   10560
      TabIndex        =   4
      Top             =   7800
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   10560
      TabIndex        =   3
      Top             =   7080
      Width           =   5055
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3255
      Left            =   360
      TabIndex        =   2
      Top             =   360
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   5741
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
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   12720
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   2400
      Width           =   2775
   End
   Begin MSComCtl2.MonthView MonthView1 
      Height          =   2370
      Left            =   12720
      TabIndex        =   0
      Top             =   0
      Width           =   2805
      _ExtentX        =   4948
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   83361794
      TitleBackColor  =   -2147483635
      CurrentDate     =   40467
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   15240
      Picture         =   "Form6.frx":0000
      Top             =   7800
      Width           =   480
   End
   Begin VB.Label Label2 
      Caption         =   "COMMENT"
      Height          =   255
      Left            =   10560
      TabIndex        =   13
      Top             =   6720
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "HMEROMHNIA"
      Height          =   255
      Left            =   10560
      TabIndex        =   12
      Top             =   5880
      Width           =   1455
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'On Error GoTo er:
'ELEGXOI
If IsDate(Text3.Text) = False Then GoTo er1:
If Text3.Text = "" Then GoTo er2:
If Text2.Text = "" Then GoTo er3:



If RS.State = 1 Then RS.Close
If DB.State = 1 Then DB.Close
DB.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & "\HMER.mdb" & ";" & _
      "Persist Security Info=False"
DB.Open App.Path & "\HMER.mdb"
RS.Open "[PIN]", DB, adOpenDynamic, adLockBatchOptimistic
Dim c As Integer
c = 0
If RS.BOF = RS.EOF Then GoTo NIK:
RS.MoveFirst
NIK:
 '  PSAKSIMO EGRAFHS
    Do While Not RS.EOF
        If RS![MAIN] <> Text2.Text Or RS![date_string] <> CDate(Text3.Text) Then
            RS.MoveNext
        Else
                Text4.Text = RS![MAIN]
                Text5.Text = RS![date_string]
                c = c + 1
            RS.MoveNext
        End If
    Loop
  ' ANTIMETOPISI ANALOGA ME TO AN BRHKE H OXI EGRAFH
    If c <> 0 Then
       MsgBox ("г еццяажг поу фгтгхгйе упаявеи гдг"), vbCritical, "пяосовг!!"
        GoTo TELOS:
    End If










If RS.State = 1 Then RS.Close
If DB.State = 1 Then DB.Close
DB.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & "\HMER.mdb" & ";" & _
      "Persist Security Info=False"
DB.Open App.Path & "\HMER.mdb"
RS.Open "[PIN]", DB, adOpenDynamic, adLockBatchOptimistic
Dim statement As String
'statement = "INSERT INTO PIN (MAIN,DATE_STRING) VALUES (" & _
'        "'" & UCase(Trim(Text2.Text)) & "'," & _
'        "'" & UCase(Trim(Text3.Text)) & "'" & _
 '        ")"
statement = "INSERT INTO PIN (MAIN,DATE_STRING) VALUES (" & _
        "'" & UCase(Trim(Text2.Text)) & "'," & _
        " #" & Format(CDate(Text3.Text), "d / m / yyyy") & "#)"


DB.Execute statement

Text1.Text = Text3.Text
If RS.State = 1 Then RS.Close
If DB.State = 1 Then DB.Close
DB.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & "\HMER.mdb" & ";" & _
      "Persist Security Info=False"
DB.Open App.Path & "\HMER.mdb"
RS.Open "[PIN]", DB, adOpenDynamic, adLockBatchOptimistic
Dim DATABASE_FILE As String
DATABASE_FILE = App.Path & "/HMER.MDB"
Adodc1.ConnectionString = _
"PROVIDER=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & DATABASE_FILE & ";"
Adodc1.RecordSource = "SELECT * FROM PIN where DATE_STRING=#" & UCase(Trim(Text1.Text)) & "#  ORDER BY DATE_STRING DESC"
Adodc1.Refresh
DataGrid1.Refresh
Set DataGrid1.DataSource = Adodc1


Adodc2.ConnectionString = _
"PROVIDER=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & DATABASE_FILE & ";"
Adodc2.RecordSource = "SELECT * FROM PIN where DATE_STRING >= Date()  ORDER BY DATE_STRING ASC"
Adodc2.Refresh
DataGrid2.Refresh
Set DataGrid2.DataSource = Adodc2

GoTo TELOS:


'ANTIMETOPISI ERROR
er1:
MsgBox ("пкгйтяокоцгсате кахос тгм глеяолгмиа"), vbCritical, "пяосовг !!!"
GoTo TELOS:

er2:
MsgBox ("то педио глеяолгмиас еимаи йемо"), vbCritical, "пяосовг !!!"
GoTo TELOS:

er3:
MsgBox ("то педио COMMENT еимаи йемо"), vbCritical, "пяосовг !!!"
GoTo TELOS:


er:
MsgBox ("йапоио ацмысто сжакла елжамистгйе"), vbCritical, "пяосовг !!!"
GoTo TELOS:


TELOS:
End Sub

Private Sub Command2_Click()
'On Error GoTo er:

If IsDate(Text3.Text) = False Then GoTo er1:
If Text3.Text = "" Then GoTo er2:
If Text2.Text = "" Then GoTo er3:

If Command2.Caption = "FIND" Then  'FIND

If RS.State = 1 Then RS.Close
If DB.State = 1 Then DB.Close
DB.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & "\HMER.mdb" & ";" & _
      "Persist Security Info=False"
DB.Open App.Path & "\HMER.mdb"
RS.Open "[PIN]", DB, adOpenDynamic, adLockBatchOptimistic
Dim c As Integer
c = 0
If RS.BOF = RS.EOF Then GoTo NIK:
RS.MoveFirst
NIK:
 '  PSAKSIMO EGRAFHS
    Do While Not RS.EOF
        If RS![MAIN] <> Text2.Text Or RS![date_string] <> CDate(Text3.Text) Then
            RS.MoveNext
        Else
                Text4.Text = RS![MAIN]
                Text5.Text = RS![date_string]
                Text2.Text = RS![MAIN]
                Text3.Text = RS![date_string]
                c = c + 1
            RS.MoveNext
        End If
    Loop
  ' ANTIMETOPISI ANALOGA ME TO AN BRHKE H OXI EGRAFH
    If c = 0 Then
       MsgBox ("г еццяажг поу фгтгхгйе дем бяехгйе"), vbCritical, "пяосовг!!"
        Command2.Caption = "FIND"
        GoTo TELOS:
    Else
        Command2.Caption = "CHANGE"
    End If

Else   ' diorthosi
    
    If Text3.Text = Text4.Text And Text2.Text = Text5.Text Then
        MsgBox ("дем диояхысате типота"), vbCritical, "пяосовг !!!"
    Else
                 
            STATE1 = "UPDATE pin SET date_string=#" & Format(CDate(Text3.Text), "m/d/yyyy") & "# WHERE date_string=#" & Format(CDate(Text5.Text), "m/d/yyyy") & "# and main='" & _
            Text4.Text & "'"
            
            STATE2 = "UPDATE pin SET MAIN='" & UCase(Trim(Text2.Text)) & "' WHERE date_string=#" & Format(CDate(Text3.Text), "m/d/yyyy") & "# and main='" & _
            Text4.Text & "'"
            
            DB.Execute STATE1
            DB.Execute STATE2
            
           
            Command2.Caption = "FIND"
            Text2.Text = ""
            Text3.Text = ""
            Text4.Text = ""
            Text5.Text = ""
            
If RS.State = 1 Then RS.Close
If DB.State = 1 Then DB.Close
DB.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & "\HMER.mdb" & ";" & _
      "Persist Security Info=False"
DB.Open App.Path & "\HMER.mdb"
RS.Open "[PIN]", DB, adOpenDynamic, adLockBatchOptimistic
Dim DATABASE_FILE As String
DATABASE_FILE = App.Path & "/HMER.MDB"
Adodc1.ConnectionString = _
"PROVIDER=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & DATABASE_FILE & ";"
Adodc1.RecordSource = "SELECT * FROM PIN where DATE_STRING=#" & UCase(Trim(Text1.Text)) & "#  ORDER BY DATE_STRING DESC"
Adodc1.Refresh
DataGrid1.Refresh
Set DataGrid1.DataSource = Adodc1


Adodc2.ConnectionString = _
"PROVIDER=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & DATABASE_FILE & ";"
Adodc2.RecordSource = "SELECT * FROM PIN where DATE_STRING >= Date()  ORDER BY DATE_STRING ASC"
Adodc2.Refresh
DataGrid2.Refresh
Set DataGrid2.DataSource = Adodc2
            
        End If
End If
GoTo TELOS:


er1:
MsgBox ("дем дысате глеяолгмиа"), vbCritical, "пяосовг !!!"
GoTo TELOS:

er2:
MsgBox ("то педио апо еима йемо"), vbCritical, "пяосовг !!!"
GoTo TELOS:

er3:
MsgBox ("то педио левяи еима йемо"), vbCritical, "пяосовг !!!"
GoTo TELOS:

er4:
MsgBox ("дем дысате сыста глеяолгмиа"), vbCritical, "пяосовг !!!"
GoTo TELOS:

er:
MsgBox ("йапоио ацмысто сжакла елжамистгйе"), vbCritical, "пяосовг !!!"
GoTo TELOS:


TELOS:
End Sub

Private Sub Command3_Click()
On Error GoTo er:
If IsDate(Text3.Text) = False Then GoTo er1:
If Text3.Text = "" Then GoTo er2:
If Text2.Text = "" Then GoTo er3:



If RS.State = 1 Then RS.Close
If DB.State = 1 Then DB.Close
DB.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & "\HMER.mdb" & ";" & _
      "Persist Security Info=False"
DB.Open App.Path & "\HMER.mdb"
RS.Open "[PIN]", DB, adOpenDynamic, adLockBatchOptimistic
Dim c As Integer
c = 0
If RS.BOF = RS.EOF Then GoTo NIK:
RS.MoveFirst
NIK:
 '  PSAKSIMO EGRAFHS
    Do While Not RS.EOF
        If RS![MAIN] <> Text2.Text Or RS![date_string] <> CDate(Text3.Text) Then
            RS.MoveNext
        Else
                Text4.Text = RS![MAIN]
                Text5.Text = RS![date_string]
                c = c + 1
            RS.MoveNext
        End If
    Loop
  ' ANTIMETOPISI ANALOGA ME TO AN BRHKE H OXI EGRAFH
    If c = 0 Then
       MsgBox ("г еццяажг поу фгтгхгйе дем упаявеи"), vbCritical, "пяосовг!!"
        GoTo TELOS:
    End If




If RS.State = 1 Then RS.Close
If DB.State = 1 Then DB.Close
DB.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & "\HMER.mdb" & ";" & _
      "Persist Security Info=False"
DB.Open App.Path & "\HMER.mdb"
RS.Open "[PIN]", DB, adOpenDynamic, adLockBatchOptimistic
Dim statement As String
statement = "DELETE  * FROM PIN where MAIN='" & UCase(Trim(Text2.Text)) & "' AND date_string=#" & Format(CDate(Text3.Text), "m/d/yyyy") & "#"
DB.Execute statement

Text1.Text = Text3.Text
If RS.State = 1 Then RS.Close
If DB.State = 1 Then DB.Close
DB.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & "\HMER.mdb" & ";" & _
      "Persist Security Info=False"
DB.Open App.Path & "\HMER.mdb"
RS.Open "[PIN]", DB, adOpenDynamic, adLockBatchOptimistic
Dim DATABASE_FILE As String
DATABASE_FILE = App.Path & "/HMER.MDB"
Adodc1.ConnectionString = _
"PROVIDER=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & DATABASE_FILE & ";"
Adodc1.RecordSource = "SELECT * FROM PIN where DATE_STRING=#" & UCase(Trim(Text1.Text)) & "#  ORDER BY DATE_STRING DESC"
Adodc1.Refresh
DataGrid1.Refresh
Set DataGrid1.DataSource = Adodc1


Adodc2.ConnectionString = _
"PROVIDER=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & DATABASE_FILE & ";"
Adodc2.RecordSource = "SELECT * FROM PIN where DATE_STRING >= Date()  ORDER BY DATE_STRING ASC"
Adodc2.Refresh
DataGrid2.Refresh
Set DataGrid2.DataSource = Adodc2



GoTo TELOS:


'ANTIMETOPISI ERROR
er1:
MsgBox ("пкгйтяокоцгсате кахос тгм глеяолгмиа"), vbCritical, "пяосовг !!!"
GoTo TELOS:

er2:
MsgBox ("то педио глеяолгмиас еимаи йемо"), vbCritical, "пяосовг !!!"
GoTo TELOS:

er3:
MsgBox ("то педио COMMENT еимаи йемо"), vbCritical, "пяосовг !!!"
GoTo TELOS:


er:
MsgBox ("йапоио ацмысто сжакла елжамистгйе"), vbCritical, "пяосовг !!!"
GoTo TELOS:


TELOS:

End Sub

Private Sub Command4_Click()
On Error GoTo er:
If IsDate(Text1.Text) = False Then GoTo er1:
If Text2.Text = "" Then GoTo er2:



If RS.State = 1 Then RS.Close
Dim DATABASE_FILE As String
DATABASE_FILE = App.Path & "/HMER.MDB"
Adodc1.ConnectionString = _
"PROVIDER=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & DATABASE_FILE & ";"
Adodc1.RecordSource = "SELECT * FROM PIN where DATE_STRING=#" & Text1.Text & "#  ORDER BY DATE_STRING DESC"
Adodc1.Refresh
DataGrid1.Refresh
Set DataGrid1.DataSource = Adodc1
GoTo TELOS:


'ANTIMETOPISI ERROR
er1:
MsgBox ("пкгйтяокоцгсате кахос тгм глеяолгмиа"), vbCritical, "пяосовг !!!"
GoTo TELOS:

er2:
MsgBox ("то педио глеяолгмиас еимаи йемо"), vbCritical, "пяосовг !!!"
GoTo TELOS:

er:
MsgBox ("йапоио ацмысто сжакла елжамистгйе"), vbCritical, "пяосовг !!!"
GoTo TELOS:


TELOS:

End Sub

Private Sub Command5_Click()
Form1.Show
Form6.Hide
Unload Form6
End Sub

Private Sub DataGrid1_Click()
Text2.Text = DataGrid1.Columns(0).Text
Text3.Text = DataGrid1.Columns(1).Text
End Sub

Private Sub DataGrid2_Click()
Text2.Text = DataGrid2.Columns(0).Text
Text3.Text = CDate(DataGrid2.Columns(1).Text)
End Sub

Private Sub Form_Load()
On Error GoTo TELOS:
If RS.State = 1 Then RS.Close
If DB.State = 1 Then DB.Close
Text1.Text = Date

MonthView1.Value = Date
MonthView2.Value = Date

DB.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & "\HMER.mdb" & ";" & _
      "Persist Security Info=False"
DB.Open App.Path & "\HMER.mdb"
RS.Open "[PIN]", DB, adOpenDynamic, adLockBatchOptimistic
DataGrid1.Font.Size = 13
DataGrid1.DefColWidth = 6800
DataGrid1.HeadFont.Bold = True
DataGrid1.HeadFont.Size = 10
DataGrid1.Font = "verdana"

DataGrid1.Columns(0).Width = 1400
DataGrid1.Columns(1).Width = 14

Dim DATABASE_FILE As String
DATABASE_FILE = App.Path & "/HMER.MDB"
Adodc1.ConnectionString = _
"PROVIDER=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & DATABASE_FILE & ";"
Adodc1.RecordSource = "SELECT * FROM PIN where DATE_STRING=#" & Text1.Text & "#  ORDER BY DATE_STRING DESC"
' Bind the ADODC to the DataGrid.
Set DataGrid1.DataSource = Adodc1
'Text3.Text = Adodc1.Recordset.RecordCount
'If Text3.Text <= 33 Then
'    DataGrid1.Height = 327.059 + (CInt(Text3.Text) * 327.059)
'Else
'    DataGrid1.Height = 11120
'End If


DataGrid2.Font.Size = 13
DataGrid2.DefColWidth = 6800
DataGrid2.HeadFont.Bold = True
DataGrid2.HeadFont.Size = 10
DataGrid2.Font = "verdana"

DataGrid2.Columns(0).Width = 1400
DataGrid2.Columns(1).Width = 14

Adodc2.ConnectionString = _
"PROVIDER=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & DATABASE_FILE & ";"
Adodc2.RecordSource = "SELECT * FROM PIN where DATE_STRING >= Date() ORDER BY DATE_STRING ASC"
' Bind the ADODC to the DataGrid.
Set DataGrid2.DataSource = Adodc2


TELOS:
End Sub

Private Sub Form_Unload(Cancel As Integer)
Form1.Show
Form6.Hide
Unload Form6
End Sub

Private Sub Image1_Click()
Text3.Text = ""
Text2.Text = ""
Command2.Caption = "FIND"
End Sub

Private Sub MonthView1_DateClick(ByVal DateClicked As Date)
Text1.Text = MonthView1.Value
If RS.State = 1 Then RS.Close
Dim DATABASE_FILE As String
DATABASE_FILE = App.Path & "/HMER.MDB"
Adodc1.ConnectionString = _
"PROVIDER=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & DATABASE_FILE & ";"
Adodc1.RecordSource = "SELECT * FROM PIN where DATE_STRING=#" & Text1.Text & "#  ORDER BY DATE_STRING DESC"
Adodc1.Refresh
DataGrid1.Refresh
Set DataGrid1.DataSource = Adodc1
End Sub

Private Sub MonthView2_DateClick(ByVal DateClicked As Date)
Text3.Text = MonthView2.Value
End Sub

Private Sub Picture1_Click()

End Sub

