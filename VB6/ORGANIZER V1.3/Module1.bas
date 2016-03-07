Attribute VB_Name = "Module1"
Public db As New ADODB.Connection
Public rs As New ADODB.Recordset
Public total_rec_tele As Integer
Public db1 As New ADODB.Connection
Public rs1 As New ADODB.Recordset
Public rsrs As New ADODB.Recordset
Public EPIL_KARTELAS As Integer
Public EPIL_KARTELAS_GRAMMA As Integer
Public DAY_HM_EKDOSHS As String
Public MONTH_HM_EKDOSHS As String
Public ETOS_HM_EKDOSHS As String
Public DAY_HM_EKSOFLISIS As String
Public MONTH_HM_EKSOFLISIS As String
Public ETOS_HM_EKSOFLISIS As String
Public DAY_HM_EKSOFLISIS_2 As String
Public MONTH_HM_EKSOFLISIS_2 As String
Public ETOS_HM_EKSOFLISIS_2 As String
Public DAY_ANAZHTHSH_1 As String
Public MONTH_ANAZHTHSH_1 As String
Public ETOS_ANAZHTHSH_1 As String
Public DAY_ANAZHTHSH_2 As String
Public MONTH_ANAZHTHSH_2 As String
Public ETOS_ANAZHTHSH_2 As String
Public DAY_HM_APO_KARTELA As String
Public MONTH_HM_APO_KARTELA As String
Public ETOS_HM_APO_KARTELA As String
Public DAY_HM_MEXRI_KARTELA As String
Public MONTH_HM_MEXRI_KARTELA As String
Public ETOS_HM_MEXRI_KARTELA As String
Public FLAG_FORM6 As Integer

Sub ELEGXOSEYRESHSTHL(a)
Form2.Adodc1.Refresh
Dim DATABASE_FILE As String
DATABASE_FILE = App.Path & "/DATABASES/TELEPHONE.MDB"
Form2.Adodc1.ConnectionString = _
"PROVIDER=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & DATABASE_FILE & ";"
Form2.Adodc1.RecordSource = " SELECT * FROM TEL_PIN WHERE епихето LIKE '" & a & "%' ORDER BY епихето"
' Bind the ADODC to the DataGrid.
Set Form2.DataGrid1.DataSource = Form2.Adodc1
Form2.Text8.Text = Form2.Adodc1.Recordset.RecordCount
If Form2.Text8.Text <= 33 Then
    Form2.DataGrid1.Height = 327.059 + (CInt(Form2.Text8.Text) * 327.059)
Else
    Form2.DataGrid1.Height = 11120
End If
End Sub


Sub ANAZHTHSH_ETAIRION(ON_ETAIRIAS)
Form3.Adodc1.Refresh
Dim DATABASE_FILE As String
DATABASE_FILE = App.Path & "/DATABASES/ETAIRIES.MDB"
Form3.Adodc1.ConnectionString = _
"PROVIDER=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & DATABASE_FILE & ";"
 Form3.Adodc1.RecordSource = " SELECT * FROM ONOMATA_ETAIRION_ABCDEF " & _
"WHERE омолата_етаияиым LIKE '" & ON_ETAIRIAS & "%' ORDER BY омолата_етаияиым"
    ' Bind the ADODC to the DataGrid.
    Set Form3.DataGrid1.DataSource = Form3.Adodc1
Form3.Adodc1.Refresh

Form3.Adodc2.ConnectionString = _
"PROVIDER=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & DATABASE_FILE & ";"
 Form3.Adodc2.RecordSource = " SELECT COUNT(омолата_етаияиым) FROM ONOMATA_ETAIRION_ABCDEF " & _
"WHERE омолата_етаияиым LIKE '" & ON_ETAIRIAS & "%'"
    ' Bind the ADODC to the DataGrid.
    Set Form3.DataGrid2.DataSource = Form3.Adodc2
Form3.Adodc2.Refresh
Form3.Text4.Text = Form3.DataGrid2.Text
If Form3.Text4.Text <= 32 Then
    Form3.DataGrid1.Height = (Form3.Text4.Text * 300.46875) + 225
Else
    Form3.DataGrid1.Height = 9840
End If
Form3.Label3.Caption = "аяихлос сумеяцафолемым етаияиым : " & Form3.Text4.Text
End Sub

Sub ELEGXOSEYRESHSTHL_backup(a)
ZTHL.Adodc1.Refresh
Dim DATABASE_FILE As String
DATABASE_FILE = App.Path & "/DATABASES/BACK_UPS/BACKUP_THL/" & Form11THL.Text1.Text & ".mdb"
ZTHL.Adodc1.ConnectionString = _
"PROVIDER=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & DATABASE_FILE & ";"
ZTHL.Adodc1.RecordSource = " SELECT * FROM TEL_PIN WHERE епихето LIKE '" & a & "%' ORDER BY епихето"
' Bind the ADODC to the DataGrid.
Set ZTHL.DataGrid1.DataSource = ZTHL.Adodc1
End Sub
Sub ELX(s)
On Error GoTo er:
Dim L, LH, i, C, m As Integer
Dim PIN(51) As String
Dim PINHE(51) As String
Dim S_HELP, SH As String
Dim a As String
C = 0
m = 1
S_HELP = ""

For i = 0 To 50
PINHE(i) = ""
Next i

L = Len(s)
If L > 50 Then
    SH = Mid(s, 1, 50)
    LH = 50
Else
    SH = s
    LH = L
End If

For i = 1 To LH
a = Mid(SH, i, 1)
If ((Asc(a) >= 48) And (Asc(a) <= 57)) Or _
((Asc(a) >= 65) And (Asc(a) <= 90)) Or _
((Asc(a) >= 97) And (Asc(a) <= 122)) Or _
((Asc(a) >= 225) And (Asc(a) <= 249)) Or _
((Asc(a) >= 193) And (Asc(a) <= 217)) Or _
((Asc(a) = 32) Or (Asc(a) = 95)) Or _
((Asc(a) = 42) Or (Asc(a) = 35)) Then
    PIN(C) = a
    C = C + 1
Else

End If
Next i

i = 0

For i = 0 To C - 1
S_HELP = S_HELP & PIN(i)
Next i

Form1.Text1.Text = S_HELP
Form1.Text2.Text = C
GoTo TELOS:

er:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
End Sub

