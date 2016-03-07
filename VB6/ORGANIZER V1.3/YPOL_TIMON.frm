VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form YPOL_TIMON 
   BackColor       =   &H80000013&
   Caption         =   "тилес"
   ClientHeight    =   10485
   ClientLeft      =   105
   ClientTop       =   465
   ClientWidth     =   15180
   LinkTopic       =   "Form10"
   ScaleHeight     =   10485
   ScaleWidth      =   15180
   Begin VB.CommandButton Command4 
      BackColor       =   &H00E9C5AD&
      Caption         =   "аккацг тилым"
      Height          =   735
      Left            =   12480
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   8520
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00E9C5AD&
      Caption         =   "йахаяислос педиым"
      Height          =   615
      Left            =   9840
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E9C5AD&
      Caption         =   "упокоцислос"
      Height          =   615
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E9C5AD&
      Caption         =   "епистяожг сто аявийо лемоу"
      Height          =   855
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   9240
      Width           =   1695
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   6375
      Left            =   2400
      TabIndex        =   12
      Top             =   3480
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   11245
      _Version        =   393216
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   11400
      TabIndex        =   5
      Top             =   1080
      Width           =   2055
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   8640
      TabIndex        =   4
      Top             =   1080
      Width           =   2055
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   5880
      TabIndex        =   3
      Top             =   1080
      Width           =   2055
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   5880
      TabIndex        =   2
      Top             =   2400
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   3120
      TabIndex        =   1
      Top             =   1080
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   360
      TabIndex        =   0
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
      Caption         =   "тилес циа диажояетийа пососта йеядоус"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   14
      Top             =   3000
      Width           =   9255
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
      Caption         =   "текийг тилг"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5880
      TabIndex        =   11
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
      Caption         =   "пососто (%)"
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
      Left            =   11400
      TabIndex        =   10
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
      Caption         =   "жпа (%)"
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
      Left            =   8640
      TabIndex        =   9
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
      Caption         =   "ейптысг (%)"
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
      Left            =   5880
      TabIndex        =   8
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
      Caption         =   "аяихлос телавиым"
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
      Left            =   3120
      TabIndex        =   7
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
      Caption         =   "аявийг тилг"
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
      Left            =   360
      TabIndex        =   6
      Top             =   720
      Width           =   2055
   End
End
Attribute VB_Name = "YPOL_TIMON"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_LostFocus()
Dim dd As Integer
Dim S As String

dd = Len(Combo1.Text)
S = Combo1.Text
For i = 1 To dd
    If Mid(S, i, 1) = "." Then
        Mid(S, i, 1) = ","
    End If
Next i
Combo1.Text = S
End Sub

Private Sub Combo3_LostFocus()
Dim dd As Integer
Dim S As String

dd = Len(Combo3.Text)
S = Combo3.Text
For i = 1 To dd
    If Mid(S, i, 1) = "." Then
        Mid(S, i, 1) = ","
    End If
Next i
Combo3.Text = S
End Sub

Private Sub Command1_Click()
On Error GoTo er:
YPOL_TIMON.Hide
Unload YPOL_TIMON
GoTo TELOS:

er:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
End Sub

Private Sub Command2_Click()
On Error GoTo er:

Dim QWERTY As Integer

Dim FPA1, FPA2, index As Integer
Dim X, T, E, P, G, F, A, B, C  As Double
Dim G_1, G_2, G_3, G_4, G_5, G_6, G_7, G_8, G_9, G_10, G1, G2, G3, G4, G5, G6, G7, G8, G9, G10 As Double
index = 5
Dim DBHELP As New ADODB.Connection
Dim RSHELP As New ADODB.Recordset
'********** SYNDESH ME BASH ******************************************
If RSHELP.STATE = 1 Then RSHELP.Close
If DBHELP.STATE = 1 Then DBHELP.Close
DBHELP.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
       "Data Source=" & "\databases\HELP.mdb" & ";" & _
      "Persist Security Info=False"
DBHELP.Open App.Path & "\databases\HELP.mdb"
RSHELP.Open "[YPOL_TIMON]", DBHELP, adOpenDynamic, adLockBatchOptimistic
'*************** TIMES APO BASH SE METABLHTES *********************************
RSHELP.MoveFirst
For i = 1 To 5
If RSHELP![Number] = 1 Then FPA1 = RSHELP![TIMI]
If RSHELP![Number] = 2 Then FPA2 = RSHELP![TIMI]
RSHELP.MoveNext
Next i
If RSHELP.STATE = 1 Then RSHELP.Close
If DBHELP.STATE = 1 Then DBHELP.Close
' ************** ELEGXOS TIMON **************************************
If Text1.Text = "" Then GoTo AR_TIMI1:
If IsNumeric(Text1.Text) = False Then GoTo AR_TIMI2:

If Text2.Text = "" Then GoTo AR_TEMAXION1:
If IsNumeric(Text2.Text) = False Then GoTo AR_TEMAXION2:
QWERTY = CInt(Text2.Text)
Text2.Text = QWERTY

If Combo1.Text = "" Then GoTo EKPTOSH1:
If IsNumeric(Combo1.Text) = False Then GoTo EKPTOSH2:

If Combo2.Text = "" Then GoTo FPA_1:
If Combo2.Text = FPA1 Or Combo2.Text = FPA2 Then
Else
    GoTo FPA_2:
End If

If Combo3.Text = "" Then GoTo POSOSTO_1:
If IsNumeric(Combo3.Text) = False Then GoTo POSOSTO_2:
'************************************************************************

'****** ANATHESI TEXT KAI COMBO BOXES SE METABLHTES ***********************
X = Text1.Text
T = Text2.Text
E = Combo1.Text
F = Combo2.Text
P = Combo3.Text

'*************YPOLOGISMOS TIMON ***********************************
'GIA P
A = X - ((E / 100) * X)
B = A / T
C = B + (B * F / 100)
G = C + (C * P / 100)
Text3.Text = G
'GIA P-1
G_1 = C + (C * (P - 1) / 100)
'GIA P-2
G_2 = C + (C * (P - 2) / 100)
'GIA P-3
G_3 = C + (C * (P - 3) / 100)
'GIA P-4
G_4 = C + (C * (P - 4) / 100)
'GIA P-5
G_5 = C + (C * (P - 5) / 100)
'GIA P-6
G_6 = C + (C * (P - 6) / 100)
'GIA P-7
G_7 = C + (C * (P - 7) / 100)
'GIA P-8
G_8 = C + (C * (P - 8) / 100)
'GIA P-9
G_9 = C + (C * (P - 9) / 100)
'GIA P-10
G_10 = C + (C * (P - 10) / 100)
'GIA P+1
G1 = C + (C * (P + 1) / 100)
'GIA P+2
G2 = C + (C * (P + 2) / 100)
'GIA P+3
G3 = C + (C * (P + 3) / 100)
'GIA P+4
G4 = C + (C * (P + 4) / 100)
'GIA P+5
G5 = C + (C * (P + 5) / 100)
'GIA P+6
G6 = C + (C * (P + 6) / 100)
'GIA P+7
G7 = C + (C * (P + 7) / 100)
'GIA P+8
G8 = C + (C * (P + 8) / 100)
'GIA P+9
G9 = C + (C * (P + 9) / 100)
'GIA P+10
G10 = C + (C * (P + 10) / 100)

' ******* ANTOISTHXHSH TIMON STO FLEXGRID ********************
MSFlexGrid1.TextMatrix(1, 1) = P - 10
MSFlexGrid1.TextMatrix(2, 1) = P - 9
MSFlexGrid1.TextMatrix(3, 1) = P - 8
MSFlexGrid1.TextMatrix(4, 1) = P - 7
MSFlexGrid1.TextMatrix(5, 1) = P - 6
MSFlexGrid1.TextMatrix(6, 1) = P - 5
MSFlexGrid1.TextMatrix(7, 1) = P - 4
MSFlexGrid1.TextMatrix(8, 1) = P - 3
MSFlexGrid1.TextMatrix(9, 1) = P - 2
MSFlexGrid1.TextMatrix(10, 1) = P - 1
MSFlexGrid1.TextMatrix(11, 1) = P
MSFlexGrid1.TextMatrix(12, 1) = P + 1
MSFlexGrid1.TextMatrix(13, 1) = P + 2
MSFlexGrid1.TextMatrix(14, 1) = P + 3
MSFlexGrid1.TextMatrix(15, 1) = P + 4
MSFlexGrid1.TextMatrix(16, 1) = P + 5
MSFlexGrid1.TextMatrix(17, 1) = P + 6
MSFlexGrid1.TextMatrix(18, 1) = P + 7
MSFlexGrid1.TextMatrix(19, 1) = P + 8
MSFlexGrid1.TextMatrix(20, 1) = P + 9
MSFlexGrid1.TextMatrix(21, 1) = P + 10

MSFlexGrid1.TextMatrix(1, 2) = G_10
MSFlexGrid1.TextMatrix(2, 2) = G_9
MSFlexGrid1.TextMatrix(3, 2) = G_8
MSFlexGrid1.TextMatrix(4, 2) = G_7
MSFlexGrid1.TextMatrix(5, 2) = G_6
MSFlexGrid1.TextMatrix(6, 2) = G_5
MSFlexGrid1.TextMatrix(7, 2) = G_4
MSFlexGrid1.TextMatrix(8, 2) = G_3
MSFlexGrid1.TextMatrix(9, 2) = G_2
MSFlexGrid1.TextMatrix(10, 2) = G_1
MSFlexGrid1.TextMatrix(11, 2) = G
MSFlexGrid1.TextMatrix(12, 2) = G1
MSFlexGrid1.TextMatrix(13, 2) = G2
MSFlexGrid1.TextMatrix(14, 2) = G3
MSFlexGrid1.TextMatrix(15, 2) = G4
MSFlexGrid1.TextMatrix(16, 2) = G5
MSFlexGrid1.TextMatrix(17, 2) = G6
MSFlexGrid1.TextMatrix(18, 2) = G7
MSFlexGrid1.TextMatrix(19, 2) = G8
MSFlexGrid1.TextMatrix(20, 2) = G9
MSFlexGrid1.TextMatrix(21, 2) = G10
GoTo TELOS:

'*********** ANTIMETOPISH ELEGXON ********************************
AR_TIMI1:
MsgBox ("дем дысате аявийг тилг"), vbCritical, "пяосовг!!!"
index = 32

AR_TIMI2:
If index = 32 Then
    GoTo TELOS:
Else
    MsgBox ("дем дысате сыста тгм аявийг тилг"), vbCritical, "пяосовг!!!"
    index = 32
End If

AR_TEMAXION1:
If index = 32 Then
    GoTo TELOS:
Else
    MsgBox ("дем дысате аяихло телавиым"), vbCritical, "пяосовг!!!"
    index = 32
End If

AR_TEMAXION2:
If index = 32 Then
    GoTo TELOS:
Else
    MsgBox ("дем дысате сыста том аяихло телавиым"), vbCritical, "пяосовг!!!"
    index = 32
End If

EKPTOSH1:
If index = 32 Then
    GoTo TELOS:
Else
    MsgBox ("дем дысате тилг циа тгм ейптысг"), vbCritical, "пяосовг!!!"
    index = 32
End If

EKPTOSH2:
If index = 32 Then
    GoTo TELOS:
Else
    MsgBox ("дем дысате сыста тгм тилг циа тгм ейптысг"), vbCritical, "пяосовг!!!"
    index = 32
End If

FPA_1:
If index = 32 Then
    GoTo TELOS:
Else
    MsgBox ("дем дысате тилг циа то ж.п.а"), vbCritical, "пяосовг!!!"
    index = 32
End If

FPA_2:
If index = 32 Then
    GoTo TELOS:
Else
    MsgBox ("дем дысате сыста тгм тилг циа то ж.п.а"), vbCritical, "пяосовг!!!"
    index = 32
End If

POSOSTO_1:
If index = 32 Then
    GoTo TELOS:
Else
    MsgBox ("дем дысате пососто"), vbCritical, "пяосовг!!!"
    index = 32
End If

POSOSTO_2:
If index = 32 Then
    GoTo TELOS:
Else
    MsgBox ("дем дысате сыста то пососто"), vbCritical, "пяосовг!!!"
    index = 32
End If


er:
If index = 32 Then
    GoTo TELOS:
Else
    MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"
    index = 32
End If

TELOS:
End Sub

Private Sub Command3_Click()
On Error GoTo er:

Text1.Text = ""
Text2.Text = "1"
Text3.Text = ""
Combo1.Text = "0"
Combo2.Text = "жпа"
Combo3.Text = "20"
GoTo TELOS:

er:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
End Sub

Private Sub Command4_Click()
On Error GoTo er:

Load YPOL_TIMON2
YPOL_TIMON2.Show
GoTo TELOS:

er:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
End Sub

Private Sub Form_Load()
On Error GoTo er:

Dim FPA1, FPA2, E, P, T As Integer
Dim DBHELP As New ADODB.Connection
Dim RSHELP As New ADODB.Recordset
'********** SYNDESH ME BASH ******************************************
If RSHELP.STATE = 1 Then RSHELP.Close
If DBHELP.STATE = 1 Then DBHELP.Close
DBHELP.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
       "Data Source=" & "\databases\HELP.mdb" & ";" & _
      "Persist Security Info=False"
DBHELP.Open App.Path & "\databases\HELP.mdb"
RSHELP.Open "[YPOL_TIMON]", DBHELP, adOpenDynamic, adLockBatchOptimistic
'*************** TIMES APO BASH SE METABLHTES *********************************
RSHELP.MoveFirst
For i = 1 To 5
If RSHELP![Number] = 1 Then FPA1 = RSHELP![TIMI]
If RSHELP![Number] = 2 Then FPA2 = RSHELP![TIMI]
If RSHELP![Number] = 3 Then E = RSHELP![TIMI]
If RSHELP![Number] = 4 Then P = RSHELP![TIMI]
If RSHELP![Number] = 5 Then T = RSHELP![TIMI]
RSHELP.MoveNext
Next i
'***********ARXIKES TIMES *******************************************
Text2.Text = T
Combo1.Text = E
Combo2.Text = "жпа"
Combo3.Text = P
Combo2.AddItem FPA1
Combo2.AddItem FPA2
' ********** TIMES GIA COMBO SXETIKO ME EKPTOSH ***********************
Combo1.AddItem "2"
Combo1.AddItem "3"
Combo1.AddItem "4"
Combo1.AddItem "5"
Combo1.AddItem "6"
Combo1.AddItem "7"
Combo1.AddItem "8"
Combo1.AddItem "9"
Combo1.AddItem "10"
Combo1.AddItem "11"
Combo1.AddItem "12"
Combo1.AddItem "13"
Combo1.AddItem "14"
Combo1.AddItem "15"
Combo1.AddItem "16"
Combo1.AddItem "17"
Combo1.AddItem "18"
Combo1.AddItem "19"
Combo1.AddItem "20"
Combo1.AddItem "21"
Combo1.AddItem "22"
Combo1.AddItem "25"
Combo1.AddItem "30"
Combo1.AddItem "35"
Combo1.AddItem "40"
Combo1.AddItem "45"
Combo1.AddItem "50"
' ********** TIMES GIA COMBO SXETIKO ME POSOSTO ***********************
Combo3.AddItem "5"
Combo3.AddItem "10"
Combo3.AddItem "11"
Combo3.AddItem "12"
Combo3.AddItem "13"
Combo3.AddItem "14"
Combo3.AddItem "15"
Combo3.AddItem "16"
Combo3.AddItem "17"
Combo3.AddItem "18"
Combo3.AddItem "19"
Combo3.AddItem "20"
Combo3.AddItem "21"
Combo3.AddItem "22"
Combo3.AddItem "23"
Combo3.AddItem "24"
Combo3.AddItem "25"
Combo3.AddItem "26"
Combo3.AddItem "27"
Combo3.AddItem "28"
Combo3.AddItem "29"
Combo3.AddItem "30"
Combo3.AddItem "35"
Combo3.AddItem "40"
Combo3.AddItem "45"
Combo3.AddItem "50"
If RSHELP.STATE = 1 Then RSHELP.Close
If DBHELP.STATE = 1 Then DBHELP.Close

MSFlexGrid1.Rows = 22
MSFlexGrid1.Cols = 3
MSFlexGrid1.TextMatrix(0, 1) = "пососто"
MSFlexGrid1.TextMatrix(0, 2) = "тилг"
MSFlexGrid1.ColWidth(0) = 1
MSFlexGrid1.ColWidth(1) = 4570
MSFlexGrid1.ColWidth(2) = 4570
MSFlexGrid1.Font.Size = 10

MSFlexGrid1.Row = 0
MSFlexGrid1.Col = 0
MSFlexGrid1.RowSel = 21
MSFlexGrid1.ColSel = 2
MSFlexGrid1.FillStyle = flexFillRepeat
MSFlexGrid1.CellAlignment = flexAlignCenterCenter

MSFlexGrid1.Row = 11
MSFlexGrid1.Col = 1
MSFlexGrid1.RowSel = 11
MSFlexGrid1.ColSel = 2
MSFlexGrid1.FillStyle = flexFillRepeat
MSFlexGrid1.CellFontBold = True
MSFlexGrid1.CellForeColor = &HFF&
GoTo TELOS:

er:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo er:
YPOL_TIMON.Hide
Unload YPOL_TIMON
GoTo TELOS:

er:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
End Sub

Private Sub Text1_LostFocus()
Dim dd As Integer
Dim S As String

dd = Len(Text1.Text)
S = Text1.Text
For i = 1 To dd
    If Mid(S, i, 1) = "." Then
        Mid(S, i, 1) = ","
    End If
Next i
Text1.Text = S
End Sub

Private Sub Text2_LostFocus()
Dim dd As Integer
Dim S As String

dd = Len(Text2.Text)
S = Text2.Text
For i = 1 To dd
    If Mid(S, i, 1) = "." Then
        Mid(S, i, 1) = ","
    End If
Next i
Text2.Text = S
End Sub
