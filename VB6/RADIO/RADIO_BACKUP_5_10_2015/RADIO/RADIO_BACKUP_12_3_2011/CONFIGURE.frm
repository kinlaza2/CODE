VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form conf 
   BackColor       =   &H80000013&
   Caption         =   "CONFIGURE RADIO"
   ClientHeight    =   7620
   ClientLeft      =   6060
   ClientTop       =   450
   ClientWidth     =   11535
   LinkTopic       =   "Form1"
   ScaleHeight     =   7620
   ScaleWidth      =   11535
   Begin VB.CommandButton Command4 
      BackColor       =   &H00E9C5AD&
      Caption         =   "EXIT"
      Height          =   735
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2040
      Width           =   2055
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   7320
      TabIndex        =   8
      Text            =   "http://www.e-radio.gr/"
      Top             =   3600
      Width           =   1935
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   6360
      TabIndex        =   7
      Text            =   "HELP TEXT GIA UPDATE - NAME"
      Top             =   6360
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   4800
      TabIndex        =   6
      Top             =   7080
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00E9C5AD&
      Caption         =   "DELETE"
      Height          =   615
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1200
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E9C5AD&
      Caption         =   "SEARCH"
      Height          =   615
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1200
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E9C5AD&
      Caption         =   "INSERT"
      Height          =   615
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1200
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   5760
      TabIndex        =   2
      Top             =   720
      Width           =   5175
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   5760
      TabIndex        =   1
      Top             =   240
      Width           =   5175
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   720
      Top             =   7200
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   6975
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   12303
      _Version        =   393216
      DefColWidth     =   160
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
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
      Caption         =   "CHECK URL"
      Height          =   255
      Left            =   7320
      TabIndex        =   9
      Top             =   3000
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   11040
      Picture         =   "CONFIGURE.frx":0000
      Top             =   240
      Width           =   480
   End
End
Attribute VB_Name = "conf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error GoTo er:

Dim c As Integer
Dim UPDATE_RECORD1, UPDATE_RECORD2 As String
Dim db As New ADODB.Connection
Dim RS As New ADODB.Recordset
db.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & "\DB.mdb" & ";" & _
"Persist Security Info=False"
db.Open App.Path & "\DB.mdb"
RS.Open "PIN", db, adOpenDynamic, adLockBatchOptimistic

If Text1.Text = "" Then GoTo er1:


' ################ EYRESH AN YPARXRI HDH ##############################
c = 0
If RS.BOF = RS.EOF Then GoTo NIK:
RS.MoveFirst
NIK:
Do While Not RS.EOF
    If RS![name] <> Trim(Text1.Text) Then
        RS.MoveNext
    Else
        c = c + 1
        RS.MoveNext
    End If
Loop
'############################ INSERT AN DEN YPARXEI ALLIOS MYNHMA KAI META TELOS #####################
If c = 0 Then
    INSERT_RECORD = "INSERT INTO PIN  (NAME,URL) VALUES ('" & Trim(UCase(Text1.Text)) & "','" & Text2.Text & "')"
    db.Execute INSERT_RECORD
    Text1.Text = ""
    Text2.Text = ""
Else
    GoTo er2:
End If
RS.Close
db.Close
DataGrid1.Refresh
Adodc1.Refresh
GoTo TELOS:

er1:
MsgBox ("TO PEDIO NAME EINAI KENO"), vbCritical, "пяосовг !!!"
GoTo TELOS:

er2:
MsgBox ("TO NAME POU DOSATE YPARXEI HDH"), vbCritical, "пяосовг !!!"
GoTo TELOS:

er:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
End Sub

Private Sub Command2_Click()
'On Error GoTo er:

Dim c As Integer
Dim UPDATE_RECORD1, UPDATE_RECORD2 As String
Dim db As New ADODB.Connection
Dim RS As New ADODB.Recordset
db.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & "\DB.mdb" & ";" & _
"Persist Security Info=False"
db.Open App.Path & "\DB.mdb"
RS.Open "PIN", db, adOpenDynamic, adLockBatchOptimistic

If Text1.Text = "" Then GoTo er1:

If Command2.Caption = "SEARCH" Then
    ' ################ EYRESH AN YPARXRI HDH ##############################
    c = 0
    If RS.BOF = RS.EOF Then GoTo NIK:
        RS.MoveFirst
NIK:
        Do While Not RS.EOF
            If RS![name] <> Trim(UCase(Text1.Text)) Then
                RS.MoveNext
            Else
                c = c + 1
                Text4.Text = RS![name]
                Text1.Text = RS![name]
                Text2.Text = RS![URL]
                RS.MoveNext
            End If
        Loop
    If c <> 0 Then
        Command2.Caption = "UPDATE"
    Else
        GoTo er2:
    End If
Else
    c = 0
    If RS.BOF = RS.EOF Then GoTo NIK1:
        RS.MoveFirst
NIK1:
    Do While Not RS.EOF
            If RS![name] <> Trim(UCase(Text1.Text)) Then
                RS.MoveNext
            Else
                If Text4.Text = Text1.Text Then
                    c = 0
                Else
                    c = c + 1
                End If
                RS.MoveNext
            End If
        Loop
    If c = 0 Then
        UPDATE_RECORD1 = "UPDATE PIN SET URL='" & Trim(Text2.Text) & "' WHERE NAME='" & Trim(Text4.Text) & "'"
        UPDATE_RECORD2 = "UPDATE PIN SET NAME='" & Trim(UCase(Text1.Text)) & "' WHERE NAME='" & Trim(Text4.Text) & "'"
        db.Execute UPDATE_RECORD1
        db.Execute UPDATE_RECORD2
        RS.Close
        db.Close
        DataGrid1.Refresh
        Adodc1.Refresh
        Command2.Caption = "SEARCH"
    Else
        GoTo er3:
    End If
    
End If
GoTo TELOS:

er1:
MsgBox ("TO PEDIO NAME EINAI KENO"), vbCritical, "пяосовг !!!"
GoTo TELOS:

er2:
MsgBox ("TO NAME POU DOSATE DEN YPARXEI"), vbCritical, "пяосовг !!!"
GoTo TELOS:


er3:
MsgBox ("TO NAME POU PROSPATHITE NA DOSETE YPARXEI HDH"), vbCritical, "пяосовг !!!"
GoTo TELOS:

er:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
End Sub

Private Sub Command3_Click()
'On Error GoTo er:

Dim c As Integer
Dim DELETE_RECORD As String
Dim db As New ADODB.Connection
Dim RS As New ADODB.Recordset
db.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & "\DB.mdb" & ";" & _
"Persist Security Info=False"
db.Open App.Path & "\DB.mdb"
RS.Open "PIN", db, adOpenDynamic, adLockBatchOptimistic

If Text1.Text = "" Then GoTo er1:


' ################ EYRESH AN YPARXRI HDH ##############################
c = 0
If RS.BOF = RS.EOF Then GoTo NIK:
RS.MoveFirst
NIK:
Do While Not RS.EOF
    If RS![name] <> Trim(UCase(Text1.Text)) Then
        RS.MoveNext
    Else
        c = c + 1
        RS.MoveNext
    End If
Loop
'############################ INSERT AN DEN YPARXEI ALLIOS MYNHMA KAI META TELOS #####################
If c <> 0 Then
    DELETE_RECORD = "DELETE  FROM PIN WHERE NAME='" & Trim(UCase(Text1.Text)) & "'"
    db.Execute DELETE_RECORD
    Text1.Text = ""
    Text2.Text = ""
Else
    GoTo er2:
End If
RS.Close
db.Close
DataGrid1.Refresh
Adodc1.Refresh
GoTo TELOS:

er1:
MsgBox ("TO PEDIO NAME EINAI KENO"), vbCritical, "пяосовг !!!"
GoTo TELOS:

er2:
MsgBox ("TO NAME POU DOSATE DEN YPARXEI "), vbCritical, "пяосовг !!!"
GoTo TELOS:

er:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
End Sub

Private Sub Command4_Click()
conf.Hide
Load RADIO
RADIO.DataGrid1.Refresh
RADIO.Adodc1.Refresh
RADIO.Show
'Unload conf
RADIO.Text3.Text = Adodc1.Recordset.RecordCount
RADIO.DataGrid1.Font.Size = 10
RADIO.DataGrid1.HeadFont.Size = 10
RADIO.DataGrid1.HeadFont.Bold = True
RADIO.DataGrid1.DefColWidth = 2250

Text3.Text = RADIO.Adodc1.Recordset.RecordCount
    If RADIO.Text3.Text <= 33 Then
        RADIO.DataGrid1.Height = 300 + (CInt(Text3.Text) * 300)
    Else
        RADIO.DataGrid1.Height = 10215
    End If
End Sub

Private Sub DataGrid1_Click()
Text1.Text = Trim(DataGrid1.Columns(0).Text)
Text2.Text = Trim(DataGrid1.Columns(1).Text)
End Sub

Private Sub DataGrid1_DblClick()
Text1.Text = Trim(DataGrid1.Columns(0).Text)
Text2.Text = Trim(DataGrid1.Columns(1).Text)
End Sub

Private Sub Form_Load()

Dim DATABASE_FILE As String
DATABASE_FILE = App.Path & "\DB.mdb"
Adodc1.ConnectionString = _
"PROVIDER=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & DATABASE_FILE & ";"
Adodc1.RecordSource = "SELECT * FROM PIN ORDER BY NAME"
' Bind the ADODC to the DataGrid.
Set DataGrid1.DataSource = Adodc1
Text3.Text = Adodc1.Recordset.RecordCount


DataGrid1.Font.Size = 10
DataGrid1.HeadFont.Size = 10
DataGrid1.HeadFont.Bold = True
DataGrid1.DefColWidth = 2250

End Sub

Private Sub Form_Unload(Cancel As Integer)
conf.Hide
Load RADIO
RADIO.DataGrid1.Refresh
RADIO.Adodc1.Refresh
RADIO.Show
Unload conf
RADIO.Text3.Text = Adodc1.Recordset.RecordCount
RADIO.DataGrid1.Font.Size = 10
RADIO.DataGrid1.HeadFont.Size = 10
RADIO.DataGrid1.HeadFont.Bold = True
RADIO.DataGrid1.DefColWidth = 2250

Text3.Text = RADIO.Adodc1.Recordset.RecordCount
    If RADIO.Text3.Text <= 33 Then
        RADIO.DataGrid1.Height = 300 + (CInt(Text3.Text) * 300)
    Else
        RADIO.DataGrid1.Height = 10215
    End If
End Sub

Private Sub Image1_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Command2.Caption = "SEARCH"
End Sub

