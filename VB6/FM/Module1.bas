Attribute VB_Name = "Module1"
Public DB As New ADODB.Connection
Public RS As New ADODB.Recordset
Public DB1 As New ADODB.Connection
Public RS1 As New ADODB.Recordset
Public DB2 As New ADODB.Connection
Public RS2 As New ADODB.Recordset
Public STAT_EMFANISIS As String


Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hwnd As Long, ByVal lpszOp As String, _
     ByVal lpszFile As String, ByVal lpszParams As String, _
     ByVal LpszDir As String, ByVal FsShowCmd As Long) _
    As Long

Sub datagridsize(ARITHMOS)
If ARITHMOS = 0 Then
    Form1.DataGrid1.Font.Size = 13
    Form1.DataGrid1.HeadFont.Bold = True
    Form1.DataGrid1.HeadFont.Size = 10
    Form1.DataGrid1.Font = "verdana"
    Form1.DataGrid1.Height = Form1.Height - 2600
    Form1.DataGrid1.Width = Form1.Width - 400
    Form1.DataGrid1.Columns(0).Width = 0.55 * Form1.DataGrid1.Width
    Form1.DataGrid1.Columns(1).Width = 0.2 * Form1.DataGrid1.Width
    Form1.DataGrid1.Columns(2).Width = 0.1 * Form1.DataGrid1.Width
    Form1.DataGrid1.Columns(3).Width = 0.1 * Form1.DataGrid1.Width
    
Else
    If ARITHMOS = 2 Then
        Form1.DataGrid1.Font.Size = 13
        'Form1.DataGrid1.DefColWidth = 6800
        Form1.DataGrid1.HeadFont.Bold = True
        Form1.DataGrid1.HeadFont.Size = 10
        Form1.DataGrid1.Font = "verdana"
        Form1.DataGrid1.Height = Form1.Height - 2600
        Form1.DataGrid1.Width = Form1.Width - 400
        Form1.DataGrid1.Columns(0).Width = 0.7 * Form1.DataGrid1.Width
        Form1.DataGrid1.Columns(1).Width = Form1.DataGrid1.Width - Form1.DataGrid1.Columns(0).Width - 600
        
        'Form1.DataGrid1.Columns(1).Width = 0.35 * Form1.DataGrid1.Width
    Else
        Form1.DataGrid1.Font.Size = 13
        'Form1.DataGrid1.DefColWidth = 6800
        Form1.DataGrid1.HeadFont.Bold = True
        Form1.DataGrid1.HeadFont.Size = 10
        Form1.DataGrid1.Font = "verdana"
        'Form1.DataGrid1.Left = Form1.Left - 100
        Form1.DataGrid1.Height = Form1.Height - 2600
        Form1.DataGrid1.Width = Form1.Width - 400
        Form1.DataGrid1.Columns(0).Width = 0.7 * Form1.DataGrid1.Width
        Form1.DataGrid1.Columns(1).Width = Form1.DataGrid1.Width - Form1.DataGrid1.Columns(0).Width - 600
    End If
End If




End Sub






    

