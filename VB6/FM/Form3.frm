VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "FM-REMINDER"
   ClientHeight    =   7830
   ClientLeft      =   6060
   ClientTop       =   2415
   ClientWidth     =   12585
   LinkTopic       =   "Form3"
   ScaleHeight     =   7830
   ScaleWidth      =   12585
   Begin VB.CommandButton Command11 
      BackColor       =   &H80000009&
      Height          =   255
      Left            =   9480
      Picture         =   "Form3.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   1080
      Width           =   375
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H80000009&
      Height          =   255
      Left            =   10200
      Picture         =   "Form3.frx":0314
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   1080
      Width           =   375
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H80000009&
      Height          =   255
      Left            =   8760
      Picture         =   "Form3.frx":05FF
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   1080
      Width           =   375
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   8760
      TabIndex        =   28
      Text            =   "0"
      Top             =   720
      Width           =   1815
   End
   Begin VB.CommandButton Command8 
      Height          =   255
      Left            =   12000
      TabIndex        =   27
      Top             =   2520
      Width           =   375
   End
   Begin VB.CommandButton Command7 
      Height          =   315
      Left            =   7680
      TabIndex        =   26
      Top             =   1320
      Width           =   375
   End
   Begin VB.CheckBox Check5 
      Height          =   375
      Left            =   6720
      TabIndex        =   25
      Top             =   4200
      Width           =   375
   End
   Begin VB.CheckBox Check4 
      Height          =   255
      Left            =   6720
      TabIndex        =   24
      Top             =   3960
      Width           =   495
   End
   Begin VB.CommandButton Command6 
      Caption         =   "елжамисг сглеяимым"
      Height          =   735
      Left            =   3360
      TabIndex        =   21
      Top             =   6120
      Width           =   2175
   End
   Begin VB.CommandButton Command5 
      Height          =   255
      Left            =   7680
      TabIndex        =   20
      Top             =   4680
      Width           =   495
   End
   Begin VB.CommandButton Command4 
      Caption         =   "диацяажг окым"
      Height          =   735
      Left            =   8160
      TabIndex        =   19
      Top             =   6120
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      Caption         =   "елжамисг окым"
      Height          =   735
      Left            =   5760
      TabIndex        =   18
      Top             =   6120
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "енодос"
      Height          =   735
      Left            =   360
      TabIndex        =   17
      Top             =   6720
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "пяосхгйг"
      Height          =   735
      Left            =   5760
      TabIndex        =   16
      Top             =   5160
      Width           =   2175
   End
   Begin VB.CheckBox Check3 
      Height          =   255
      Left            =   6720
      TabIndex        =   10
      Top             =   3600
      Width           =   735
   End
   Begin VB.CheckBox Check2 
      Height          =   255
      Left            =   6720
      TabIndex        =   9
      Top             =   3240
      Width           =   495
   End
   Begin VB.CheckBox Check1 
      Height          =   495
      Left            =   6720
      TabIndex        =   8
      Top             =   2760
      Width           =   495
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   360
      TabIndex        =   7
      Top             =   1800
      Width           =   12135
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   6360
      TabIndex        =   5
      Top             =   720
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   4560
      TabIndex        =   4
      Top             =   720
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10800
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   720
      Width           =   1695
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   12000
      Top             =   1320
   End
   Begin VB.Label Label6 
      Caption         =   "09:00"
      Height          =   255
      Left            =   6000
      TabIndex        =   23
      Top             =   4320
      Width           =   495
   End
   Begin VB.Label Label5 
      Caption         =   "10:00"
      Height          =   255
      Left            =   6000
      TabIndex        =   22
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label Label14 
      Caption         =   "ока та паяапамы"
      Height          =   255
      Left            =   6000
      TabIndex        =   15
      Top             =   4680
      Width           =   1695
   End
   Begin VB.Label Label13 
      Caption         =   "06:00"
      Height          =   255
      Left            =   6000
      TabIndex        =   14
      Top             =   3960
      Width           =   615
   End
   Begin VB.Label Label12 
      Caption         =   "03:00"
      Height          =   255
      Left            =   6000
      TabIndex        =   13
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label Label11 
      Caption         =   "12:00"
      Height          =   255
      Left            =   6000
      TabIndex        =   12
      Top             =   3240
      Width           =   615
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Caption         =   "ыяа"
      Height          =   375
      Left            =   5160
      TabIndex        =   11
      Top             =   2400
      Width           =   2655
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "йеилемо "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      TabIndex        =   6
      Top             =   1320
      Width           =   3975
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "лгмас"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6720
      TabIndex        =   3
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "  леяа"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4800
      TabIndex        =   2
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10800
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'On Error GoTo ER:
' ********* ELEGXOS LATHON ************************
If Text2.Text = "" Or IsNumeric(Text2.Text) = False Then
    GoTo er1:
Else
    If CInt(Text2.Text) < 1 Or CInt(Text2.Text) > 31 Then
        GoTo er1:
    Else
    End If
End If
If Text3.Text = "" Or IsNumeric(Text3.Text) = False Then
    GoTo er2:
Else
    If CInt(Text3.Text) < 1 Or CInt(Text3.Text) > 12 Then
        GoTo er2:
    Else
    End If
End If
If IsNumeric(Combo1.Text) = False Then
    Combo1.Text = 0
Else
    If CInt(Combo1.Text) < 0 Or CInt(Combo1.Text) > 10 Then Combo1.Text = 0
End If
'*********************************************

If Check1.Value = 0 And Check2.Value = 0 And Check3.Value = 0 And Check4.Value = 0 And Check5.Value = 0 Then GoTo er3:
Dim F3, F2, F1 As Integer
If Check1.Value = 1 Then
    F1 = 1
Else
    F1 = 0
End If
If Check2.Value = 1 Then
    F2 = 1
Else
    F2 = 0
End If
If Check3.Value = 1 Then
    F3 = 1
Else
    F3 = 0
End If
If Check4.Value = 1 Then
    F4 = 1
Else
    F4 = 0
End If
If Check5.Value = 1 Then
    F5 = 1
Else
    F5 = 0
End If

Dim statement As String
If Combo1.Text = "0" Then
    If RS1.State = 1 Then RS1.Close
    If DB1.State = 1 Then DB1.Close
    DB1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
    "Data Source=" & "\DB1.mdb" & ";" & _
      "Persist Security Info=False"
    DB1.Open App.Path & "\DB1.mdb"
    RS1.Open "[PIN]", DB1, adOpenDynamic, adLockBatchOptimistic
    statement = "INSERT INTO PIN(KEIMENO,MERA,MHNAS,F1,F2,F3,F4,F5)" & _
        "VALUES (" & _
            "'" & Text4.Text & "'," & _
            "'" & CInt(Text2.Text) & "', " & _
            "'" & CInt(Text3.Text) & "'," & _
            "'" & F1 & "', " & _
            "'" & F2 & "', " & _
            "'" & F3 & "', " & _
            "'" & F4 & "', " & _
            "'" & F5 & "'" & ")"
    DB1.Execute statement
    
Else
    If RS1.State = 1 Then RS1.Close
    If DB1.State = 1 Then DB1.Close
    DB1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
    "Data Source=" & "\DB1.mdb" & ";" & _
      "Persist Security Info=False"
    DB1.Open App.Path & "\DB1.mdb"
    RS1.Open "[PIN]", DB1, adOpenDynamic, adLockBatchOptimistic
    statement = "INSERT INTO PIN(KEIMENO,MERA,MHNAS,F1,F2,F3,F4,F5)" & _
        "VALUES (" & _
            "'" & Text4.Text & "'," & _
            "'" & CInt(Text2.Text) & "', " & _
            "'" & CInt(Text3.Text) & "'," & _
            "'" & F1 & "', " & _
            "'" & F2 & "', " & _
            "'" & F3 & "', " & _
            "'" & F4 & "', " & _
            "'" & F5 & "'" & ")"
    DB1.Execute statement
       
    Dim MERAI, MHNASI, TEMPO As Integer
    MERAI = CInt(Text2.Text)
    MHNASI = CInt(Text3.Text)
    TEMPO = CInt(Combo1.Text)
    For I = 0 To TEMPO - 1
        If CInt(Text3.Text) = 4 Or CInt(Text3.Text) = 6 Or CInt(Text3.Text) = 9 Or CInt(Text3.Text) = 11 Then
                If MERAI + 1 <= 30 Then
                    MERAI = MERAI + 1
                Else
                    MERAI = MERAI + 1 - 30
                    MHNASI = MHNASI + 1
                    If MHNASI > 12 Then MHNASI = 1
                End If
        Else
            If CInt(Text3.Text) = 2 Then
                If MERAI + 1 <= 28 Then
                    MERAI = MERAI + 1
                Else
                    MERAI = MERAI + 1 - 28
                    MHNASI = MHNASI + 1
                    If MHNASI > 12 Then MHNASI = 1
                End If
            Else
                If MERAI + 1 <= 31 Then
                    MERAI = MERAI + 1
                Else
                    MERAI = MERAI + 1 - 31
                    MHNASI = MHNASI + 1
                    If MHNASI > 12 Then MHNASI = 1
                End If
            End If
        End If
        If RS1.State = 1 Then RS1.Close
        If DB1.State = 1 Then DB1.Close
        DB1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
            "Data Source=" & "\DB1.mdb" & ";" & _
            "Persist Security Info=False"
        DB1.Open App.Path & "\DB1.mdb"
        RS1.Open "[PIN]", DB1, adOpenDynamic, adLockBatchOptimistic
        statement = "INSERT INTO PIN(KEIMENO,MERA,MHNAS,F1,F2,F3,F4,F5)" & _
            "VALUES (" & _
                "'" & Text4.Text & "'," & _
                "'" & MERAI & "', " & _
                "'" & MHNASI & "'," & _
                "'" & F1 & "', " & _
                "'" & F2 & "', " & _
                "'" & F3 & "', " & _
                "'" & F4 & "', " & _
                "'" & F5 & "'" & ")"
        DB1.Execute statement

    Next I
    




End If

Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Check1.Value = 0
Check2.Value = 0
Check3.Value = 0
Check4.Value = 0
Check5.Value = 0
Combo1.Text = "0"
GoTo TELOS:

er1:
MsgBox ("кахос се леяа."), vbCritical, "пяосовг !!!"
Text2.Text = ""
GoTo TELOS:

er2:
MsgBox ("кахос се лгма."), vbCritical, "пяосовг !!!"
Text3.Text = ""
GoTo TELOS:

er3:
MsgBox ("пяепеи ма епикенете тоукавистом лиа ыяа."), vbCritical, "пяосовг !!!"
Text3.Text = ""
GoTo TELOS:

er:
MsgBox ("йапоио амапамтево кахос елжамистгйе."), vbCritical, "пяосовг !!!"
GoTo TELOS:

TELOS:

End Sub

Private Sub Command10_Click()
Combo1.Text = "10"
End Sub

Private Sub Command11_Click()
Combo1.Text = "5"
End Sub

Private Sub Command2_Click()
Form3.Hide
Unload Form3
End Sub

Private Sub Command3_Click()
STAT_EMFANISIS = "SELECT KEIMENO FROM PIN"
Load Form4
Form4.Show
If RS1.State = 1 Then RS1.Close
If DB1.State = 1 Then DB1.Close
End Sub

Private Sub Command4_Click()
Dim STAT As String
STAT = " DELETE * FROM PIN"

If MsgBox("хекете ма пяовыягсете се диацяажг окым;", vbOKCancel, "") = vbOK Then
DB1.Execute STAT
Else
End If
End Sub

Private Sub Command5_Click()
Check1.Value = 1
Check2.Value = 1
Check3.Value = 1
Check4.Value = 1
Check5.Value = 1
End Sub



Private Sub Command6_Click()
        STAT_EMFANISIS = "SELECT KEIMENO FROM PIN  where MERA=" & CInt(Day(Date)) & " AND MHNAS=" & CInt(Month(Date))
        If RS1.State = 1 Then RS1.Close
        If DB1.State = 1 Then DB1.Close
        Load Form4
        Form4.Show
End Sub

Private Sub Command7_Click()
Text2.Text = Day(Date)
Text3.Text = Month(Date)
End Sub

Private Sub Command8_Click()
Text4.Text = ""
End Sub

Private Sub Command9_Click()
Combo1.Text = "0"
End Sub

Private Sub Form_Load()
On Error GoTo er:
For I = 0 To 10
    Combo1.AddItem I
Next I
Combo1.Text = "0"
Label1.Caption = Date
Text1.Text = Time
If RS1.State = 1 Then RS1.Close
If DB1.State = 1 Then DB1.Close
DB1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & "\DB1.mdb" & ";" & _
      "Persist Security Info=False"
DB1.Open App.Path & "\DB1.mdb"
RS1.Open "[PIN]", DB1, adOpenDynamic, adLockBatchOptimistic
GoTo TELOS:

er:
MsgBox ("йапоио амапамтево кахос елжамистгйе."), vbCritical, "пяосовг !!!"

GoTo TELOS:

TELOS:
End Sub

Private Sub Form_Unload(Cancel As Integer)
Form3.Hide
Unload Form3
End Sub

Private Sub Timer1_Timer()
Label1.Caption = Date
Text1.Text = Time
End Sub

Private Sub Timer2_Timer()


End Sub
