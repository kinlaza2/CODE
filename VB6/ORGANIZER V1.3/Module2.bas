Attribute VB_Name = "Module2"
Public STATE1 As String
Public STATE2 As String
Public STATE3 As String
Public STATE4 As String
Public STATE5A As String
Public STATE5B As String
Public STATE6 As String
Public STATE7A As String
Public STATE7B As String
Public STATE8A As String
Public STATE8B As String
Public STATE9A As String
Public STATE9B As String
Public STATE As String
Public SS1 As String
Public SS2 As String
Public SS3 As String
Public SS4 As String
Public SS5A As String
Public SS5B As String
Public SS6 As String
Public SS7A As String
Public SS7B As String
Public SS8A As String
Public SS8B As String
Public SS9A As String
Public SS9B As String
Public SS As String

Public FF1 As String
Public FF2 As String
Public FF3 As String
Public FF4 As String
Public FF5A As String
Public FF5B As String
Public FF6 As String
Public FF7A As String
Public FF7B As String
Public FF8A As String
Public FF8B As String
Public FF9A As String
Public FF9B As String
Public FF As String

Public CT As Integer ' FORM4_TIMER2

Public STATE_KARTELAS As String
Public HAK
Public HMK
Public LABEL_FOR_BACKUP As String

Public PATIMA_DATAGRID_F6 As Integer

Public ZETAIRIES_BASHS_BACKUP_DIAX As String
Public ZETAIRIES_DIADROMHS_BACKUP_DIAX As String
Public ZETAIRIES_DIADROMHS_BACKUP_DIAX_1 As String
Public ETOSBACKUP As String
Public STROG As Double ' STROGILOPOIHSH SE KARTELA

Sub STROG_ARITH(ARITHMOS)
        Dim APOLITO, DEKADIKO, APOTELESMA As Double
        Dim AKERAIO, PROSIMO As Integer
        Dim temp As String
        
        If (ARITHMOS < 0.009 And ARITHMOS > -0.009) Then ARITHMOS = 0
        
        PROSIMO = Sgn(ARITHMOS)
        APOLITO = Abs(ARITHMOS)
        AKERAIO = Int(APOLITO)
        DEKADIKO = APOLITO - AKERAIO
        temp = DEKADIKO
        temp = Mid(temp, 1, 4)
        
        APOTELESMA = AKERAIO + temp
        If PROSIMO = 1 Then
            STROG = APOTELESMA
        Else
            STROG = -1 * APOTELESMA
        End If
End Sub


Sub TZIROI_YP(THESI_BASHS, ONOMA_BASHS, PIN_TZIROI, PIN_ONOM_ETAIR, AP_HM, MEX_HM)
Dim db As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset

Dim STATE, stat, STAT1, SSS, SSS1, SSS2, temp As String
Dim XREOSH, PISTOSI, YPOL As Double

XREOSH = 0
PISTOSI = 0

' FTIAKSIMO HMEROMHNION
If Day(AP_HM) < 12 Then
    APO_HMEROMHNIA = CDate(Month(AP_HM) & " / " & Day(AP_HM) & " / " & Year(AP_HM))
Else
    APO_HMEROMHNIA = AP_HM
End If

If Day(MEX_HM) < 12 Then
    MEXRI_HMER = CDate(Month(MEX_HM) & " / " & Day(MEX_HM) & " / " & Year(MEX_HM))
Else
    MEXRI_HMER = MEX_HM
End If

'P.X
'"Data Source=" & "\DATABASES\TZIROI.MDB" & ";"
'DB1.Open App.Path & "\DATABASES\TZIROI.MDB"

db.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & "\" & THESI_BASHS & "\" & ONOMA_BASHS & ".MDB ;" & _
"Persist Security Info=False"
db.Open App.Path & "\" & THESI_BASHS & "\" & ONOMA_BASHS & ".MDB"

'P.X DELETE * FROM TZIROI
STAT1 = " select * into " & PIN_TZIROI & " from TZIROI"
stat = "DELETE * FROM " & PIN_TZIROI
db.Execute STAT1
db.Execute stat

'ONOMA_PINAKA2=ONOMATA_ETAIRION_ABCDEF
rs.Open "[" & PIN_ONOM_ETAIR & "]", db, adOpenDynamic, adLockBatchOptimistic


If rs.BOF = rs.EOF Then GoTo NIK:
rs.MoveFirst
NIK:
Do While Not rs.EOF
    Form1.Text3.Text = rs![омолата_етаияиым]
    temp = "HELP_" & Form1.Text3.Text
    SSS = " select * into " & temp & " from " & Form1.Text3.Text
    db.Execute SSS
    
    SS1 = "DELETE  FROM " & temp & _
    " WHERE глеяолгмиа_ейдосгс <#" & APO_HMEROMHNIA & "#"
    SS2 = "DELETE  FROM " & temp & _
    " WHERE глеяолгмиа_ейдосгс >#" & MEXRI_HMER & "#"
    db.Execute SS1
    db.Execute SS2
    
    RS1.Open "[" & temp & "]", db, adOpenDynamic, adLockBatchOptimistic
    If RS1.BOF = RS1.EOF Then GoTo NIK1:
    RS1.MoveFirst
NIK1:
    Do While Not RS1.EOF
        XREOSH = XREOSH + RS1![вяеысг]
        PISTOSI = PISTOSI + RS1![пистысг]
        RS1.MoveNext
    Loop
          
        
        STROG_ARITH (XREOSH - PISTOSI)
        YPOL = STROG

           
    STATE = " INSERT INTO " & PIN_TZIROI & " (ETAIRIA,XREOSH,PISTOSI,YPOLOIPO)" & _
            " values (" & _
            "'" & UCase(Trim(Form1.Text3.Text)) & "'," & _
            "'" & XREOSH & "'," & _
            "'" & PISTOSI & "'," & _
            "'" & YPOL & "'" & _
            ")"
            db.Execute STATE
rs.MoveNext
XREOSH = 0
PISTOSI = 0
YPO = 0
RS1.Close
db.Execute "DROP TABLE " & temp
Loop
End Sub
Sub JOIN_PIN(PIN1, PIN2, THESI_BASHS, ONOMA_BASHS)
Dim DB1 As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset

Dim STATE, STATE1, STATE2, temp, SSS1, SSS2, SSS3 As String
Dim TEMP1, TEMP2 As Double
Dim C As Integer
C = 1

'P.X
'"Data Source=" & "\DATABASES\TZIROI.MDB" & ";"
'DB1.Open App.Path & "\DATABASES\TZIROI.MDB"

DB1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & "\" & THESI_BASHS & "\" & ONOMA_BASHS & ".MDB ;" & _
"Persist Security Info=False"
DB1.Open App.Path & "\" & THESI_BASHS & "\" & ONOMA_BASHS & ".MDB"


rs.Open "[" & PIN1 & "]", DB1, adOpenDynamic, adLockBatchOptimistic ' ASD_04
RS1.Open "[" & PIN2 & "]", DB1, adOpenDynamic, adLockBatchOptimistic 'QWE_05

 'SSS = " select * into " & TEMP & " from " & Form1.Text3.Text
SSS = " select * into TZIROI3_T  from " & PIN2
DB1.Execute SSS

If rs.BOF = rs.EOF Then GoTo NIK:
    rs.MoveFirst
NIK:

Do While Not rs.EOF
    If RS1.BOF = RS1.EOF Then GoTo NIK1:
        RS1.MoveFirst
NIK1:
        Do While Not RS1.EOF
            If rs![ETAIRIA] = RS1![ETAIRIA] Then
                temp = rs![ETAIRIA]
                TEMP1 = rs![XREOSH] + RS1![XREOSH]
                TEMP2 = rs![PISTOSI] + RS1![PISTOSI]
                STATE = " UPDATE TZIROI3_T" & _
                     " SET XREOSH=" & "'" & TEMP1 & "'" & _
                     " WHERE ETAIRIA=" & "'" & temp & "'"
                
                STATE1 = " UPDATE TZIROI3_T" & _
                     " SET PISTOSI=" & "'" & TEMP2 & "'" & _
                     " WHERE ETAIRIA=" & "'" & temp & "'"
                
                STATE2 = " UPDATE TZIROI3_T" & _
                     " SET YPOLOIPO=" & "'" & TEMP1 - TEMP2 & "'" & _
                     " WHERE ETAIRIA=" & "'" & temp & "'"
                DB1.Execute STATE
                DB1.Execute STATE1
                DB1.Execute STATE2
            End If
            RS1.MoveNext
        Loop
    rs.MoveNext
Loop

Form1.Text1.Text = SS
If rs.BOF = rs.EOF Then GoTo NIK3:
    rs.MoveFirst
NIK3:

Do While Not rs.EOF
    If RS1.BOF = RS1.EOF Then GoTo NIK4:
        RS1.MoveFirst
NIK4:
        Do While Not RS1.EOF
            If rs![ETAIRIA] = RS1![ETAIRIA] Then
                C = 2
            End If
            RS1.MoveNext
        Loop

        If C = 1 Then
                temp = rs![ETAIRIA]
                TEMP1 = rs![XREOSH]
                TEMP2 = rs![PISTOSI]
                STATE = "INSERT INTO TZIROI3_T (" & _
                        "ETAIRIA,XREOSH,PISTOSI,YPOLOIPO) " & _
                        "VALUES (" & _
                        "'" & temp & "'," & _
                        "'" & TEMP1 & "'," & _
                        "'" & TEMP2 & "'," & _
                        "'" & TEMP1 - TEMP2 & "'" & _
                        ")"
                DB1.Execute STATE
        End If
        C = 1
        rs.MoveNext
Loop
If RS1.STATE = 1 Then RS1.Close
If rs.STATE = 1 Then rs.Close

SSS1 = " DROP TABLE " & PIN2
SSS2 = "select * into " & PIN2 & " from TZIROI3_T"
SSS3 = "DROP TABLE TZIROI3_T"


DB1.Execute SSS1
DB1.Execute SSS2
DB1.Execute SSS3

If DB1.STATE = 1 Then DB1.Close
End Sub

Sub KANON_PIN(PIN, THESI_BASHS, ONOMA_BASHS)
Dim DB1 As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim X, P, X_NEW, P_NEW As Double
Dim STATE, STATE1, STATE2 As String

'P.X
'"Data Source=" & "\DATABASES\TZIROI.MDB" & ";"
'DB1.Open App.Path & "\DATABASES\TZIROI.MDB"

DB1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & "\" & THESI_BASHS & "\" & ONOMA_BASHS & ".MDB ;" & _
"Persist Security Info=False"
DB1.Open App.Path & "\" & THESI_BASHS & "\" & ONOMA_BASHS & ".MDB"
rs.Open "[" & PIN & "]", DB1, adOpenDynamic, adLockBatchOptimistic

If rs.BOF = rs.EOF Then GoTo NIK:
rs.MoveFirst
NIK:

Do While Not rs.EOF
    X = rs![XREOSH]
    P = rs![PISTOSI]
    If X > P Then
        X_NEW = X - P
        P_NEW = 0
    End If
    If X < P Then
        X_NEW = 0
        P_NEW = P - X
    End If
    If X = P Then
        X_NEW = 0
        P_NEW = 0
    End If
    STATE = " UPDATE " & PIN & _
    " SET XREOSH=" & "'" & X_NEW & "'" & _
    " WHERE ETAIRIA=" & "'" & rs![ETAIRIA] & "'"
    
    STATE1 = " UPDATE " & PIN & _
    " SET PISTOSI=" & "'" & P_NEW & "'" & _
    " WHERE ETAIRIA=" & "'" & rs![ETAIRIA] & "'"
    
    STATE2 = " UPDATE " & PIN & _
    " SET YPOLOIPO=" & "'" & X_NEW - P_NEW & "'" & _
    " WHERE ETAIRIA=" & "'" & rs![ETAIRIA] & "'"
    
    DB1.Execute STATE
    DB1.Execute STATE1
    DB1.Execute STATE2
    rs.MoveNext
Loop

End Sub

Sub TOMH_PIN(PIN1, PIN2, THESI_BASHS, ONOMA_BASHS) 'PIN O ABCEF KAI PIN2 O TZIROI_05
Dim DB1 As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
Dim STATE As String
Dim FLAG As Integer
FLAG = 0
DB1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & "\" & THESI_BASHS & "\" & ONOMA_BASHS & ".MDB ;" & _
"Persist Security Info=False"
DB1.Open App.Path & "\" & THESI_BASHS & "\" & ONOMA_BASHS & ".MDB"
rs.Open "[" & PIN2 & "]", DB1, adOpenDynamic, adLockBatchOptimistic 'PIN2=TZIROI
RS1.Open "[" & PIN1 & "]", DB1, adOpenDynamic, adLockBatchOptimistic 'PIN1=ABCDEF

If rs.BOF = rs.EOF Then GoTo NIK:
rs.MoveFirst
NIK:
Do While Not rs.EOF
    If RS1.BOF = RS1.EOF Then GoTo NIK2:
    RS1.MoveFirst
NIK2:
    Do While Not RS1.EOF
        If rs![ETAIRIA] = RS1![омолата_етаияиым] Then
            FLAG = 1
        End If
        RS1.MoveNext
    Loop
    If FLAG = 0 Then
        STATE = "DELETE FROM " & PIN2 & " WHERE ETAIRIA='" & rs![ETAIRIA] & "'"
        DB1.Execute STATE
    End If
    FLAG = 0
    rs.MoveNext
Loop
End Sub

Sub ANAF_MHN(THESI_BAS, ONOMA_BAS, THESI_BAS_GRAF_ANAF, ONOM_BAS_GRAF_ANAF, ON_ETAIR, ON_PIN_ANAF, XRONIA)
Dim db As New ADODB.Connection
Dim DB1 As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
Dim SIX, SIPO, SIPC, SIPG, SIE, UIX, UIPO, UIPC, UIPG, UIE As String
Dim SIX1, S2, S3, S4, S5, S6, S7, S8, S9, S10, s11, S12 As String

db.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & THESI_BAS & ONOMA_BAS & ";" & _
"Persist Security Info=False"
db.Open App.Path & THESI_BAS & ONOMA_BAS
rs.Open [ON_ETAIR], db, adOpenDynamic, adLockBatchOptimistic

DB1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & THESI_BAS_GRAF_ANAF & ONOM_BAS_GRAF_ANAF & ";" & _
"Persist Security Info=False"
DB1.Open App.Path & THESI_BAS_GRAF_ANAF & ONOM_BAS_GRAF_ANAF
RS1.Open "[" & ON_PIN_ANAF & "]", DB1, adOpenDynamic, adLockBatchOptimistic


SIX = "SELECT SUM(вяеысг) FROM " & ON_ETAIR & _
" WHERE глеяолгмиа_ейдосгс>=#1/1/XRONIA# AND глеяолгмиа_ейдосгс<=#31/1/XRONIA# "

UIX = "UPDATE " & ON_PIN_ANAF & " SET вяеысеис='" & db.Execute(SIX, , dbSQLPassThrough) & "' WHERE лгмас='иамоуаяиос'"
DB1.Execute UIX



End Sub
