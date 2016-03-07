VERSION 5.00
Begin VB.Form YPOL_TIMON2 
   BackColor       =   &H80000013&
   Caption         =   "аккацг тилым"
   ClientHeight    =   8205
   ClientLeft      =   2910
   ClientTop       =   1995
   ClientWidth     =   9135
   LinkTopic       =   "Form10"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8205
   ScaleWidth      =   9135
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E9C5AD&
      Caption         =   "епистяожг сто пяогцоулемо лемоу"
      Height          =   975
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   7080
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E9C5AD&
      Caption         =   "аккацг тилым"
      Height          =   855
      Left            =   3900
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   5640
      Width           =   1455
   End
   Begin VB.TextBox Text10 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   6600
      TabIndex        =   22
      Top             =   4680
      Width           =   1215
   End
   Begin VB.TextBox Text9 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   6600
      TabIndex        =   21
      Top             =   4080
      Width           =   1215
   End
   Begin VB.TextBox Text8 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   6600
      TabIndex        =   20
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   6600
      TabIndex        =   19
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   6600
      TabIndex        =   18
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2040
      TabIndex        =   11
      Top             =   4680
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2040
      TabIndex        =   10
      Top             =   4080
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2040
      TabIndex        =   9
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2040
      TabIndex        =   8
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2040
      TabIndex        =   7
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label13 
      BackColor       =   &H80000013&
      Caption         =   "телавиа  :"
      Height          =   255
      Left            =   5280
      TabIndex        =   17
      Top             =   4680
      Width           =   1335
   End
   Begin VB.Label Label12 
      BackColor       =   &H80000013&
      Caption         =   "пососто  :"
      Height          =   375
      Left            =   5280
      TabIndex        =   16
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Label Label11 
      BackColor       =   &H80000013&
      Caption         =   "ейптысг  :"
      Height          =   255
      Left            =   5280
      TabIndex        =   15
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Label Label10 
      BackColor       =   &H80000013&
      Caption         =   "ж.п.а_2"
      Height          =   255
      Left            =   5280
      TabIndex        =   14
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Label Label9 
      BackColor       =   &H80000013&
      Caption         =   "ж.п.а_1"
      Height          =   255
      Left            =   5280
      TabIndex        =   13
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
      Caption         =   "аккацг тилым:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4560
      TabIndex        =   12
      Top             =   1560
      Width           =   3975
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000013&
      Caption         =   "телавиа  :"
      Height          =   375
      Left            =   960
      TabIndex        =   6
      Top             =   4680
      Width           =   975
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000013&
      Caption         =   "пососто  :"
      Height          =   375
      Left            =   960
      TabIndex        =   5
      Top             =   4080
      Width           =   975
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000013&
      Caption         =   "ейптысг  :"
      Height          =   375
      Left            =   960
      TabIndex        =   4
      Top             =   3480
      Width           =   975
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000013&
      Caption         =   "ж.п.а_2  :"
      Height          =   375
      Left            =   960
      TabIndex        =   3
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000013&
      Caption         =   "ж.п.а_1  :"
      Height          =   255
      Left            =   960
      TabIndex        =   2
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000013&
      Caption         =   "упаявоусес тилес:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   840
      TabIndex        =   1
      Top             =   1560
      Width           =   3015
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
      Caption         =   $"YPOL_TIMON2.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8895
   End
End
Attribute VB_Name = "YPOL_TIMON2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error GoTo er:
Dim DBH As New ADODB.Connection
Dim RSH As New ADODB.Recordset
Dim STATE1, STATE2, STATE3, STATE4, STATE5 As String
Dim index As Integer
index = 2
'***********************elegxoi******************************
If Text6.Text = "" Then GoTo text6a:
If IsNumeric(Text6.Text) = False Then GoTo text6b:
If Text7.Text = "" Then GoTo text7a:
If IsNumeric(Text7.Text) = False Then GoTo text7b:
If Text8.Text = "" Then GoTo text8a:
If IsNumeric(Text8.Text) = False Then GoTo text8b:
If Text9.Text = "" Then GoTo text9a:
If IsNumeric(Text9.Text) = False Then GoTo text9b:
If Text10.Text = "" Then GoTo text10a:
If IsNumeric(Text10.Text) = False Then GoTo text10b:

' ************ йатавояисг меом тилом ****************************
If MsgBox("хекете ма пяовыягсете се амтийатастасг тым упаявоусым тилым", vbOKCancel, "пяосовг") = vbOK Then

If RSH.STATE = 1 Then RSH.Close
If DBH.STATE = 1 Then DBH.Close
DBH.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
       "Data Source=" & "\databases\HELP.mdb" & ";" & _
      "Persist Security Info=False"
DBH.Open App.Path & "\databases\HELP.mdb"
RSH.Open "[YPOL_TIMON]", DBH, adOpenDynamic, adLockBatchOptimistic
STATE1 = " UPDATE YPOL_TIMON SET TIMI=" & Text6.Text & " WHERE NUMBER=1"
STATE2 = " UPDATE YPOL_TIMON SET TIMI=" & Text7.Text & " WHERE NUMBER=2"
STATE3 = " UPDATE YPOL_TIMON SET TIMI=" & Text8.Text & " WHERE NUMBER=3"
STATE4 = " UPDATE YPOL_TIMON SET TIMI=" & Text9.Text & " WHERE NUMBER=4"
STATE5 = " UPDATE YPOL_TIMON SET TIMI=" & Text10.Text & " WHERE NUMBER=5"
DBH.Execute STATE1
DBH.Execute STATE2
DBH.Execute STATE3
DBH.Execute STATE4
DBH.Execute STATE5
YPOL_TIMON2.Hide
Unload YPOL_TIMON2
End If

GoTo TELOS:
'***** antimetopish elegxon *************************
text6a:
MsgBox ("дем дысате тгм тилг тоу ж.п.а_1"), vbCritical, "пяосовг!!"
index = 32

text6b:
If index = 2 Then
    MsgBox ("дем дысате сыста тгм тилг тоу ж.п.а_1"), vbCritical, "пяосовг!!"
    index = 32
Else
    GoTo TELOS:
End If

text7a:
If index = 2 Then
    MsgBox ("дем дысате тгм тилг тоу ж.п.а_2"), vbCritical, "пяосовг!!"
    index = 32
Else
    GoTo TELOS:
End If

text7b:
If index = 2 Then
    MsgBox ("дем дысате сыста тгм тилг тоу ж.п.а_2"), vbCritical, "пяосовг!!"
    index = 32
Else
    GoTo TELOS:
End If

text8a:
If index = 2 Then
    MsgBox ("дем дысате тгм тилг циа тгм ейптысг "), vbCritical, "пяосовг!!"
    index = 32
Else
    GoTo TELOS:
End If

text8b:
If index = 2 Then
    MsgBox ("дем дысате сыста тгм тилг циа тгм ейптысг"), vbCritical, "пяосовг!!"
    index = 32
Else
    GoTo TELOS:
End If

text9a:
If index = 2 Then
    MsgBox ("дем дысате тгм тилг циа то пососто"), vbCritical, "пяосовг!!"
    index = 32
Else
    GoTo TELOS:
End If

text9b:
If index = 2 Then
    MsgBox ("дем дысате сыста тгм тилг циа то пососто"), vbCritical, "пяосовг!!"
    index = 32
Else
    GoTo TELOS:
End If

text10a:
If index = 2 Then
    MsgBox ("дем дысате тгм тилг циа том аяихло телавиым"), vbCritical, "пяосовг!!"
    index = 32
Else
    GoTo TELOS:
End If

text10b:
If index = 2 Then
    MsgBox ("дем дысате сыста тгм тилг циа том аяихло телавиым"), vbCritical, "пяосовг!!"
    index = 32
Else
    GoTo TELOS:
End If

er:
If index = 2 Then
    MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"
    index = 32
Else
    GoTo TELOS:
End If

TELOS:
If DBH.STATE = 1 Then DBH.Close
If RSH.STATE = 1 Then RSH.Close
End Sub

Private Sub Command2_Click()
On Error GoTo er:

YPOL_TIMON.Enabled = True
Form1.Enabled = True
YPOL_TIMON2.Hide
Unload YPOL_TIMON2
GoTo TELOS:

er:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
End Sub

Private Sub Form_Load()
On Error GoTo er:

YPOL_TIMON.Enabled = False
Form1.Enabled = False


Dim DBH As New ADODB.Connection
Dim RSH As New ADODB.Recordset
If RSH.STATE = 1 Then RSH.Close
If DBH.STATE = 1 Then DBH.Close
DBH.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
       "Data Source=" & "\databases\HELP.mdb" & ";" & _
      "Persist Security Info=False"
DBH.Open App.Path & "\databases\HELP.mdb"
RSH.Open "[YPOL_TIMON]", DBH, adOpenDynamic, adLockBatchOptimistic
'********************* ANASIRSI TIMON APO BASH ***********************
RSH.MoveFirst
Do While Not RSH.EOF
    If RSH![Number] = 1 Then
        Text1.Text = RSH![TIMI]
        Text6.Text = RSH![TIMI]
        RSH.MoveNext
    End If
    If RSH![Number] = 2 Then
        Text2.Text = RSH![TIMI]
        Text7.Text = RSH![TIMI]
        RSH.MoveNext
    End If
    If RSH![Number] = 3 Then
        Text3.Text = RSH![TIMI]
        Text8.Text = RSH![TIMI]
        RSH.MoveNext
    End If
    If RSH![Number] = 4 Then
        Text4.Text = RSH![TIMI]
        Text9.Text = RSH![TIMI]
        RSH.MoveNext
    End If
    If RSH![Number] = 5 Then
        Text5.Text = RSH![TIMI]
        Text10.Text = RSH![TIMI]
        RSH.MoveNext
    End If
Loop
GoTo TELOS:

er:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
If RSH.STATE = 1 Then RSH.Close
If DBH.STATE = 1 Then DBH.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo er:

YPOL_TIMON.Enabled = True
Form1.Enabled = True
YPOL_TIMON2.Hide
Unload YPOL_TIMON2
GoTo TELOS:

er:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
End Sub
