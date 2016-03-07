VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form TZIROI_1 
   Caption         =   "Form10"
   ClientHeight    =   10485
   ClientLeft      =   60
   ClientTop       =   465
   ClientWidth     =   15270
   LinkTopic       =   "Form10"
   ScaleHeight     =   10485
   ScaleWidth      =   15270
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   3135
      Left            =   120
      TabIndex        =   3
      Top             =   4440
      Width           =   14655
      _ExtentX        =   25850
      _ExtentY        =   5530
      _Version        =   393216
      Rows            =   12
      Cols            =   14
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   240
      TabIndex        =   2
      Top             =   9720
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   1560
      Top             =   3840
      Width           =   3735
      _ExtentX        =   6588
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
      Height          =   2895
      Left            =   600
      TabIndex        =   1
      Top             =   840
      Width           =   12615
      _ExtentX        =   22251
      _ExtentY        =   5106
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
            LCID            =   1033
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
            LCID            =   1033
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
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9975
   End
   Begin VB.Line Line1 
      X1              =   7695
      X2              =   7695
      Y1              =   10
      Y2              =   10480
   End
End
Attribute VB_Name = "TZIROI_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
End
End Sub

Private Sub Form_Load()
On Error GoTo ER:
Dim DB As New ADODB.Connection
Dim RS As New ADODB.Recordset
Dim DATABASE_FILE, ON_BASHS As String
Dim C As Integer
C = 1


Label1.Caption = "тфияои етаияиас " & TZIROI.Text1.Text & " циа то етос : " & TZIROI.Combo1.Text
' ******* ELEGXOS AN TO ETOS EINAI TO TREXON H ANHKEI SE BACKUP *******
If TZIROI.Combo1.Text = Year(Date) Then
    DATABASE_FILE = "\databases\ETAIRIES.MDB\"
    ON_BASHS = "ETAIRIES.MDB"
Else
    DB.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
    "Data Source=" & "\databases\BACK_UPS\HELP_BACKUP.mdb" & ";" & _
    "Persist Security Info=False"
    DB.Open App.Path & "\databases\BACK_UPS\HELP_BACKUP.mdb"
    RS.Open [BACKUP_YEAR_ETAIRIES], DB, adOpenDynamic, adLockBatchOptimistic
    If RS.BOF = RS.EOF Then GoTo NIK:
    RS.MoveFirst
NIK:
    Do While Not RS.EOF
        If RS![ETOS] = CInt(TZIROI.Combo1.Text) Then
            If RS![FLAG] = 1 Then
                ' AN ONTOS YPARXEI TO ARXEIO
                C = 2
            End If
        End If
        RS.MoveNext
    Loop
End If
' AN DEN YPARXEI PINAKAS GIA TO EPILEGMENO ETOS LATHOS ELSE SYNDESI ME SOSTH BASH
If C = 2 Then
    DATABASE_FILE = "\databases\BACK_UPS\BACKUP_ETAIRIES\BACKUP_ETAIRIES_ETOS\"
    ON_BASHS = Trim(UCase(TZIROI.Combo1.Text)) & "_ETAIRIES"
Else
    GoTo ER1:
End If
DB.Close
RS.Close
'*********************************************************************












GoTo TELOS:


ER1:
MsgBox ("фгтгсате елжамисг стоивеиым циа етос циа то опоио дем упаявеи амтицяажо. паяайакы диояхысте"), vbCritical, "пяосовг !!!"
GoTo TELOS:

ER:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
End Sub

