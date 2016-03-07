Attribute VB_Name = "Module1"
'MAIL MESSAGES
Public MAIL_CAPTION As String
Public INITIAL_MESSAGE_MAIL As String  ' CAPTION TOY INITIAL MAIL
Public INITIAL_MESSAGE_MAIL_CAPTION As String ' BODY TOY INITIAL MAIL
Public STARTPROCEDURE_MAIL As String
Public AVI_EXIST_MAIL As String
Public AVI_SIZE_MAIL As String
Public EXIST_263GP_MAIL As String
Public SIZE_263GP_MAIL As String
Public EXIST_823GP_MAIL As String
Public SIZE_823GP_MAIL As String
Public EXIST_25RM_MAIL As String
Public SIZE_25RM_MAIL As String
Public EXIST_82RM_MAIL As String
Public SIZE_82RM_MAIL As String
Public FTPSTATUS As String
Public MAILSTATUS As Integer

'SQL QUIRIES
Public DELETE_RECORDS As String
Public INSERT_INITIAL_RECORDS As String
Public AVI_PATH_DEFAULT As String
Public FILES_PATH_DEFAULT As String
Public INSERT_RECORDS As String

'DATE VARIABLES
Public Vapo_5_h As Integer
Public vapo_5_m As Integer
Public vapo_h As Integer
Public vapo_m As Integer
Public vmexri_h As Integer
Public vmexri_m As Integer
Public venco_h As Integer
Public venco_m As Integer
Public vend_h As Integer
Public vend_m As Integer

'GENERAL VARIABLES
Public DB As New ADODB.Connection
Public RS As New ADODB.Recordset
Public TEMP_ORA, TEMP_MINUTE As Integer
Public AVI_PATH As String
Public FILES_PATH As String
Public FLAGSTARTPROCEDURE As Integer
Public FLAG_AVI As Integer
Public FLAG_263GP As Integer
Public FLAG_823GP As Integer
Public FLAG_25RM As Integer
Public FLAG_82RM As Integer
'Public FLAGFTP_READY As Integer
Public SENDMAILNOW As Integer
Public SENDINITIALMAILNOW As Integer
Public FLAG_INCLUDE_RM As Integer
Public START_APPLICATION_CAPTION_MAIL, START_APPLICATION_BODY_MAIL As String
Public PREPARE_MAIL As Integer


Sub ARXIKOPOIHSH2()
Vapo_5_h = -1
vapo_5_m = -1
vapo_h = -1
vapo_m = -1
vmexri_h = -1
vmexri_m = -1
venco_h = -1
venco_m = -1
vend_h = -1
vend_m = -1
'AVI_PATH = ""
'FILES_PATH = ""
Form1.Text6.Text = ""
Form1.Text7.Text = ""
Form1.Text1.Text = ""
Form1.Text2.Text = ""
Form1.Text3.Text = ""
Form1.Text4.Text = ""
End Sub
Sub ARXIKOPOIHSH()
'MAIL MESSAGES
MAILSTATUS = 0
MAIL_CAPTION = ""
INITIAL_MESSAGE_MAIL = ""
INITIAL_MESSAGE_MAIL_CAPTION = ""
STARTPROCEDURE_MAIL = ""
SENDMAILNOW = 0
SENDINITIALMAILNOW = 0
AVI_EXIST_MAIL = ""
AVI_SIZE_MAIL = ""
EXIST_263GP_MAIL = ""
SIZE_263GP_MAIL = ""
EXIST_823GP_MAIL = ""
SIZE_823GP_MAIL = ""
EXIST_25RM_MAIL = ""
SIZE_25RM_MAIL = ""
EXIST_82RM_MAIL = ""
SIZE_82RM_MAIL = ""
'Public FTPSTATUS
FTPSTATUS = ""
FLAG_AVI = 0
FLAG_263GP = 0
FLAG_823GP = 0
FLAG_25RM = 0
FLAG_82RM = 0
FLAGSTARTPROCEDURE = 0
FLAG_INCLUDE_RM = 0
PREPARE_MAIL = 0

Form1.Text31.Text = ""
Form1.Text32.Text = ""
Form1.Text33.Text = ""
Form1.Text34.Text = ""
Form1.Text35.Text = ""
Form1.Text36.Text = ""
Form1.Text37.Text = ""
Form1.Text38.Text = ""
Form1.Text39.Text = ""
Form1.Text40.Text = ""
Form1.Text41.Text = ""
Form1.Text42.Text = ""


End Sub
Sub FINDHOUR(ORA_BASHS, MINUTE_BASHS, HOUR_DURATION, MINUTE_DURATION)
    Dim ORA_F, LEPTA_F, STARTHOUR, STARTMINUTE, DURATIONHOUR, DURATIONMINUTE As Integer
    STARTHOUR = ORA_BASHS
    STARTMINUTE = MINUTE_BASHS
    DURATIONHOUR = HOUR_DURATION
    DURATIONMINUTE = MINUTE_DURATION
   If STARTMINUTE + DURATIONMINUTE < 60 Then
        If STARTHOUR + DURATIONHOUR >= 24 Then
            ORA_F = (STARTHOUR + DURATIONHOUR) - 24
            LEPTA_F = STARTMINUTE + DURATIONMINUTE
        Else
            ORA_F = STARTHOUR + DURATIONHOUR
            LEPTA_F = STARTMINUTE + DURATIONMINUTE
        End If
    Else
        If STARTHOUR + DURATIONHOUR + 1 >= 24 Then
            ORA_F = (STARTHOUR + DURATIONHOUR + 1) - 24
            LEPTA_F = (STARTMINUTE + DURATIONMINUTE) - 60
        Else
            ORA_F = STARTHOUR + DURATIONHOUR + 1
            LEPTA_F = (STARTMINUTE + DURATIONMINUTE) - 60
        End If
    End If
    TEMP_ORA = ORA_F
    TEMP_MINUTE = LEPTA_F
End Sub


