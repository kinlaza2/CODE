VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form PROXEIRO 
   BackColor       =   &H80000013&
   Caption         =   "пяовеияо"
   ClientHeight    =   10320
   ClientLeft      =   105
   ClientTop       =   645
   ClientWidth     =   15180
   LinkTopic       =   "Form1"
   ScaleHeight     =   10320
   ScaleWidth      =   15180
   Begin VB.CommandButton Command2 
      Caption         =   "ейтупысг"
      Height          =   615
      Left            =   13800
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00E9C5AD&
      Caption         =   "енодос"
      Height          =   615
      Left            =   13800
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00E9C5AD&
      Caption         =   "амоицла аккоу"
      Height          =   615
      Left            =   13800
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E9C5AD&
      Caption         =   "апохгйеусг"
      Height          =   615
      Left            =   13800
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   480
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   9840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "RTF"
      Filter          =   "Rich Text Format (*.RTF)|*.RTF|All Files (*.*)|*.*"
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   10095
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   13335
      _ExtentX        =   23521
      _ExtentY        =   17806
      _Version        =   393217
      ScrollBars      =   2
      TextRTF         =   $"RTFEdit1.frx":0000
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpenItem 
         Caption         =   "&Open..."
      End
      Begin VB.Menu mnuCloseItem 
         Caption         =   "&Close"
      End
      Begin VB.Menu mnuSaveAsItem 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu mnuPrintItem 
         Caption         =   "&Print"
      End
      Begin VB.Menu mnuExitItem 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuCutItem 
         Caption         =   "Cu&t"
      End
      Begin VB.Menu mnuCopyItem 
         Caption         =   "&Copy"
      End
      Begin VB.Menu mnuPasteItem 
         Caption         =   "&Paste"
      End
      Begin VB.Menu mnuFindItem 
         Caption         =   "&Find..."
      End
   End
   Begin VB.Menu mnuFormat 
      Caption         =   "F&ormat"
      Begin VB.Menu mnuAllcapsItem 
         Caption         =   "&All Caps"
      End
      Begin VB.Menu mnuFontItem 
         Caption         =   "&Font..."
      End
      Begin VB.Menu mnuBoldItem 
         Caption         =   "&Bold"
      End
      Begin VB.Menu mnuItalicItem 
         Caption         =   "&Italic"
      End
      Begin VB.Menu mnuUnderlineItem 
         Caption         =   "&Underline"
      End
   End
End
Attribute VB_Name = "PROXEIRO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Declare UnsavedChanges as a public Boolean (True/False)
'variable to track the current save state of the text.
'When the text is updated, the RichTextBox1_Change event
'procedure sets this variable to True.
Dim UnsavedChanges As Boolean
Dim MODE As Boolean


Private Sub Command1_Click()
On Error GoTo er:
RichTextBox1.SaveFile App.Path & "\KARTELES\PROXEIRO.RTF"
UnsavedChanges = False
GoTo TELOS:

er:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
End Sub

Private Sub Command2_Click()
On Error GoTo er:

RichTextBox1.SelPrint (Printer.hDC)
GoTo TELOS:

er:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
End Sub

Private Sub Command3_Click()
On Error GoTo er:

Dim Prompt As String
Dim Reply As Integer
If Command3.Caption = "амоицла аккоу" Then
    mnuSaveAsItem.Enabled = True
    mnuOpenItem.Enabled = True
    mnuCloseItem = True
    mnuExitItem = True
    Command1.Enabled = False
    Command2.Enabled = False
    Command4.Enabled = False
    'jump to error handler if the Cancel button is clicked
    CommonDialog1.CancelError = False
    If UnsavedChanges = True Then
        Prompt = "хекете ма апохгйеутоум ои аккацес поу йамате;"
        Reply = MsgBox(Prompt, vbYesNo, "")
        If Reply = vbYes Then
            RichTextBox1.SaveFile App.Path & "\KARTELES\PROXEIRO.RTF"
        End If
    End If
    RichTextBox1.Text = ""  'clear text box
    UnsavedChanges = False
    MODE = False
    Command3.Caption = "амоицла пяовеияоу"
Else
    mnuSaveAsItem.Enabled = False
    mnuOpenItem.Enabled = False
    mnuCloseItem = False
    mnuExitItem = False
    Command1.Enabled = True
    Command2.Enabled = True
    Command4.Enabled = True
    'jump to error handler if the Cancel button is clicked
    CommonDialog1.CancelError = False
    If UnsavedChanges = True Then
        Prompt = "хекете ма апохгйеутоум ои аккацес поу йамате;"
        Reply = MsgBox(Prompt, vbYesNo, "")
        If Reply = vbYes Then
            CommonDialog1.ShowSave
            RichTextBox1.SaveFile CommonDialog1.FileName, _
                rtfRTF
        End If
    End If
    RichTextBox1.Text = ""  'clear text box
    RichTextBox1.LoadFile App.Path & "\KARTELES\PROXEIRO.RTF"
    UnsavedChanges = False
    MODE = True
    Command3.Caption = "амоицла аккоу"
End If
GoTo TELOS:

er:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"


TELOS:
End Sub

Private Sub Command4_Click()
Dim Prompt As String
    Dim Reply As Integer
    'jump to error handler if the Cancel button is clicked
    
    On Error GoTo er:
    If UnsavedChanges = True Then
        Prompt = "хекете ма апохгйеутоум ои аккацес поу йамате;"
        Reply = MsgBox(Prompt, vbYesNo, "")
        If Reply = vbYes Then
            RichTextBox1.SaveFile App.Path & "\KARTELES\PROXEIRO.RTF"
        End If
    End If
    RichTextBox1.Text = ""  'clear text box
    UnsavedChanges = False
PROXEIRO.Hide
Unload PROXEIRO
GoTo TELOS:


er:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"


TELOS:
    Exit Sub
End Sub

Private Sub Form_Load()
mnuSaveAsItem.Enabled = False
mnuOpenItem.Enabled = False
mnuCloseItem = False
mnuExitItem = False
UnsavedChanges = False
MODE = True
    On Error GoTo er:
    RichTextBox1.LoadFile App.Path & "\KARTELES\PROXEIRO.RTF"
UnsavedChanges = False
GoTo TELOS:

er:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim Prompt As String
Dim Reply As Integer
'jump to error handler if the Cancel button is clicked
CommonDialog1.CancelError = False
On Error GoTo er:
If MODE = True Then
    If UnsavedChanges = True Then
        Prompt = "хекете ма апохгйеутоум ои аккацес поу йамате;"
        Reply = MsgBox(Prompt, vbYesNo, "")
        If Reply = vbYes Then
            RichTextBox1.SaveFile App.Path & "\KARTELES\PROXEIRO.RTF"
        End If
    End If
Else
    If UnsavedChanges = True Then
        Prompt = "хекете ма апохгйеутоум ои аккацес поу йамате;"
        Reply = MsgBox(Prompt, vbYesNo, "")
        If Reply = vbYes Then
            CommonDialog1.ShowSave
            RichTextBox1.SaveFile CommonDialog1.FileName, _
                rtfRTF
        End If
        UnsavedChanges = False
    End If
End If
RichTextBox1.Text = ""  'clear text box
UnsavedChanges = False
PROXEIRO.Hide
Unload PROXEIRO
GoTo TELOS:

er:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"


TELOS:
    Exit Sub
End Sub

Private Sub mnuBoldItem_Click()
    RichTextBox1.SelBold = Not RichTextBox1.SelBold
End Sub

Private Sub mnuCloseItem_Click()
    Dim Prompt As String
    Dim Reply As Integer
    'jump to error handler if the Cancel button is clicked
    CommonDialog1.CancelError = False
    On Error GoTo er:
    If UnsavedChanges = True Then
        Prompt = "хекете ма апохгйеутоум ои аккацес поу йамате;"
        Reply = MsgBox(Prompt, vbYesNo, "")
        If Reply = vbYes Then
            CommonDialog1.ShowSave
            RichTextBox1.SaveFile CommonDialog1.FileName, _
                rtfRTF
        End If
    End If
    RichTextBox1.Text = ""  'clear text box
    UnsavedChanges = False
    GoTo TELOS:
er:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
    Exit Sub
End Sub

Private Sub mnuCopyItem_Click()
    Clipboard.SetText RichTextBox1.SelRTF
End Sub

Private Sub mnuCutItem_Click()
    Clipboard.SetText RichTextBox1.SelRTF
    RichTextBox1.SelRTF = ""
End Sub

Private Sub mnuExitItem_Click()
    Dim Prompt As String
    Dim Reply As Integer
    CommonDialog1.CancelError = False
    On Error GoTo er:
    If UnsavedChanges = True Then
        Prompt = "хекете ма апохгйеутоум ои аккацес поу йамате;"
        Reply = MsgBox(Prompt, vbYesNo, "")
        If Reply = vbYes Then
            CommonDialog1.ShowSave
            RichTextBox1.SaveFile CommonDialog1.FileName, _
                rtfRTF
        End If
        UnsavedChanges = False
    End If
     'after file has been saved, quit program
PROXEIRO.Hide
Unload PROXEIRO
GoTo TELOS:

er:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
End Sub

Private Sub mnuFindItem_Click()
    Dim SearchStr As String  'text used for search
    Dim FoundPos As Integer  'location of found text
    SearchStr = InputBox("Enter search word", "Find")
    If SearchStr <> "" Then  'if search string not empty
        'find the first occurrence of the whole word
        FoundPos = RichTextBox1.Find(SearchStr, , , _
            rtfWholeWord)
        'if the word is found (if not -1)
        If FoundPos <> -1 Then
        'use Span method to select word (forward direction)
            RichTextBox1.Span " ", True, True
        Else
            MsgBox "то йеилемо поу дысате дем бяехгйе", , "еуяесг"
        End If
    End If
End Sub

Private Sub mnuFontItem_Click()
    'Force an error if the user clicks Cancel
    CommonDialog1.CancelError = True
    On Error GoTo Errhandler:
    'Set flags for special effects and all available fonts
    CommonDialog1.Flags = cdlCFEffects Or cdlCFBoth
    'Display font dialog box
    CommonDialog1.ShowFont
    'Set formatting properties with user selections:
    RichTextBox1.SelFontName = CommonDialog1.FontName
    RichTextBox1.SelFontSize = CommonDialog1.FontSize
    RichTextBox1.SelColor = CommonDialog1.Color
    RichTextBox1.SelBold = CommonDialog1.FontBold
    RichTextBox1.SelItalic = CommonDialog1.FontItalic
    RichTextBox1.SelUnderline = CommonDialog1.FontUnderline
    RichTextBox1.SelStrikeThru = CommonDialog1.FontStrikethru
Errhandler:
    'exit procedure if the user clicks Cancel
End Sub

Private Sub mnuItalicItem_Click()
    RichTextBox1.SelItalic = Not RichTextBox1.SelItalic
End Sub

Private Sub mnuAllcapsItem_Click()
    RichTextBox1.SelText = UCase(RichTextBox1.SelText)
End Sub

Private Sub mnuPrintItem_Click()
    'Prints the current document using the device
    'handle of the current printer
    RichTextBox1.SelPrint (Printer.hDC)
    GoTo TELOS:

er:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
End Sub

Private Sub mnuUnderlineItem_Click()
    RichTextBox1.SelUnderline = Not RichTextBox1.SelUnderline
End Sub

Private Sub mnuOpenItem_Click()
    CommonDialog1.CancelError = False
    
    On Error GoTo er:
    CommonDialog1.Flags = cdlOFNFileMustExist
    CommonDialog1.ShowOpen
    RichTextBox1.LoadFile CommonDialog1.FileName, rtfRTF
    UnsavedChanges = False
    GoTo TELOS:
    
er:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
End Sub

Private Sub mnuPasteItem_Click()
    RichTextBox1.SelRTF = Clipboard.GetText
End Sub

Private Sub mnuSaveAsItem_Click()
    CommonDialog1.CancelError = False
    On Error GoTo er:
    CommonDialog1.ShowSave
    'save specified file in RTF format
    RichTextBox1.SaveFile CommonDialog1.FileName, rtfRTF
    UnsavedChanges = False
    GoTo TELOS:
    
er:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
End Sub

Private Sub RichTextBox1_Change()
    'Set public variable UnsavedChanges to True each time
    'the text in the Rich textbox is modified.
    UnsavedChanges = True
    CommonDialog1.CancelError = True
End Sub
