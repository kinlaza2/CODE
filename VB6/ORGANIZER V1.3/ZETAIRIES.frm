VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form ZETAIRIES 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000013&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "тилокоциа  амтицяажоу"
   ClientHeight    =   10485
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15225
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   10485
   ScaleLeft       =   45
   ScaleMode       =   0  'User
   ScaleWidth      =   15225
   Begin VB.CommandButton Command30 
      BackColor       =   &H00E9C5AD&
      Caption         =   "омолата етаияиым"
      Height          =   855
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   111
      Top             =   120
      Width           =   1935
   End
   Begin VB.PictureBox Picture1 
      Height          =   375
      Left            =   6960
      Picture         =   "ZETAIRIES.frx":0000
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   110
      Top             =   3000
      Width           =   375
   End
   Begin VB.ComboBox Combo27 
      Height          =   315
      Left            =   4920
      TabIndex        =   106
      Text            =   "цемийо"
      Top             =   2280
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text25 
      Height          =   285
      Left            =   0
      TabIndex        =   105
      Top             =   360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text24 
      Height          =   405
      Left            =   1200
      TabIndex        =   104
      Text            =   "help_text2"
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text23 
      Height          =   375
      Left            =   0
      TabIndex        =   103
      Text            =   "HELP TEXT "
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command29 
      BackColor       =   &H00C00000&
      Caption         =   "еныжкгсг тилокоциым"
      Height          =   615
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   102
      Top             =   480
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton Command24 
      BackColor       =   &H00E9C5AD&
      Caption         =   "еуяесг"
      Height          =   615
      Left            =   3000
      MaskColor       =   &H80000010&
      Style           =   1  'Graphical
      TabIndex        =   97
      Top             =   5520
      Width           =   1575
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H80000013&
      Caption         =   "пкгяылес ле епитацг"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   120
      TabIndex        =   88
      Top             =   6600
      Width           =   7215
      Begin VB.CommandButton Command28 
         BackColor       =   &H00E9C5AD&
         Caption         =   "диацяажг"
         Height          =   615
         Left            =   5160
         Style           =   1  'Graphical
         TabIndex        =   101
         Top             =   2040
         Width           =   1455
      End
      Begin VB.CommandButton Command27 
         BackColor       =   &H00E9C5AD&
         Caption         =   "еуяесг"
         Height          =   615
         Left            =   2880
         Style           =   1  'Graphical
         TabIndex        =   100
         Top             =   2040
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   5160
         TabIndex        =   96
         Top             =   840
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   56164353
         CurrentDate     =   38388
      End
      Begin VB.TextBox Text21 
         Height          =   285
         Left            =   2520
         TabIndex        =   95
         Top             =   1320
         Width           =   2175
      End
      Begin VB.CommandButton Command23 
         BackColor       =   &H00E9C5AD&
         Caption         =   "еныжкгсг"
         Height          =   615
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   93
         Top             =   2040
         Width           =   1455
      End
      Begin VB.TextBox Text20 
         Height          =   285
         Left            =   2520
         TabIndex        =   92
         Top             =   360
         Width           =   2175
      End
      Begin VB.TextBox Text19 
         Height          =   285
         Left            =   2520
         TabIndex        =   91
         Top             =   840
         Width           =   2175
      End
      Begin VB.Label Label22 
         BackColor       =   &H80000013&
         Caption         =   "посо"
         Height          =   375
         Left            =   240
         TabIndex        =   94
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label21 
         BackColor       =   &H80000013&
         Caption         =   "аяихлос епитацгс"
         Height          =   255
         Left            =   240
         TabIndex        =   90
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label20 
         BackColor       =   &H80000013&
         Caption         =   "глеяолгмиа ейдосгс"
         Height          =   255
         Left            =   240
         TabIndex        =   89
         Top             =   840
         Width           =   1935
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H80000013&
      Caption         =   "амафгтгсг  йаи амахеыягсг йаятекас"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   7680
      TabIndex        =   72
      Top             =   960
      Width           =   7215
      Begin VB.ComboBox Combo26 
         Height          =   315
         Left            =   3720
         TabIndex        =   87
         Text            =   "етос"
         Top             =   2160
         Width           =   1215
      End
      Begin VB.ComboBox Combo25 
         Height          =   315
         Left            =   2160
         TabIndex        =   86
         Text            =   "лгмас"
         Top             =   2160
         Width           =   1215
      End
      Begin VB.ComboBox Combo24 
         Height          =   315
         Left            =   600
         TabIndex        =   85
         Text            =   "глеяа"
         Top             =   2160
         Width           =   1215
      End
      Begin VB.CommandButton Command22 
         BackColor       =   &H00E9C5AD&
         Caption         =   "сглеяимг"
         Height          =   375
         Left            =   5880
         Style           =   1  'Graphical
         TabIndex        =   84
         Top             =   2160
         Width           =   1215
      End
      Begin VB.CommandButton Command10 
         BackColor       =   &H00E9C5AD&
         Caption         =   "OK"
         Height          =   375
         Left            =   5280
         Style           =   1  'Graphical
         TabIndex        =   83
         Top             =   2160
         Width           =   495
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H00E9C5AD&
         Caption         =   "сглеяимг"
         Height          =   375
         Left            =   5880
         Style           =   1  'Graphical
         TabIndex        =   82
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00E9C5AD&
         Caption         =   "ой"
         Height          =   375
         Left            =   5280
         Style           =   1  'Graphical
         TabIndex        =   81
         Top             =   960
         Width           =   495
      End
      Begin VB.ComboBox Combo23 
         Height          =   315
         Left            =   3720
         TabIndex        =   80
         Text            =   "етос"
         Top             =   960
         Width           =   1215
      End
      Begin VB.ComboBox Combo22 
         Height          =   315
         Left            =   2040
         TabIndex        =   79
         Text            =   "лгмас"
         Top             =   960
         Width           =   1215
      End
      Begin VB.ComboBox Combo21 
         Height          =   315
         Left            =   480
         TabIndex        =   78
         Text            =   "глеяа"
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox Text18 
         Height          =   285
         Left            =   2880
         TabIndex        =   77
         Top             =   1560
         Width           =   1215
      End
      Begin VB.TextBox Text17 
         Height          =   285
         Left            =   2880
         TabIndex        =   76
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00E9C5AD&
         Caption         =   "елжамисг йаятекас"
         Height          =   855
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   73
         Top             =   3000
         Width           =   2175
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         X1              =   0
         X2              =   7200
         Y1              =   2640
         Y2              =   2640
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         X1              =   0
         X2              =   7200
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Label Label19 
         BackColor       =   &H80000013&
         Caption         =   "левяи"
         Height          =   375
         Left            =   2400
         TabIndex        =   75
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label18 
         BackColor       =   &H80000013&
         Caption         =   "апо"
         Height          =   255
         Left            =   2400
         TabIndex        =   74
         Top             =   480
         Width           =   735
      End
   End
   Begin VB.CommandButton Command18 
      BackColor       =   &H00E9C5AD&
      Caption         =   "сглеяимг"
      Height          =   375
      Left            =   13800
      Style           =   1  'Graphical
      TabIndex        =   65
      Top             =   6600
      Width           =   975
   End
   Begin VB.CommandButton Command17 
      BackColor       =   &H00E9C5AD&
      Caption         =   "ой"
      Height          =   375
      Left            =   13320
      Style           =   1  'Graphical
      TabIndex        =   64
      Top             =   7680
      Width           =   375
   End
   Begin VB.CommandButton Command16 
      Caption         =   "сглеяимг"
      Height          =   375
      Left            =   3840
      TabIndex        =   57
      Top             =   7320
      Width           =   975
   End
   Begin VB.CommandButton Command15 
      Caption         =   "OK"
      Height          =   375
      Left            =   3120
      TabIndex        =   56
      Top             =   7320
      Width           =   495
   End
   Begin VB.CommandButton Command14 
      BackColor       =   &H00E9C5AD&
      Caption         =   "сглеяимг"
      Height          =   375
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   52
      Top             =   4320
      Width           =   975
   End
   Begin VB.CommandButton Command13 
      BackColor       =   &H00E9C5AD&
      Caption         =   "OK"
      Height          =   375
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   51
      Top             =   4320
      Width           =   375
   End
   Begin VB.ComboBox Combo10 
      Height          =   315
      Left            =   4920
      TabIndex        =   50
      Text            =   "етос"
      Top             =   4320
      Width           =   855
   End
   Begin VB.ComboBox Combo9 
      Height          =   315
      Left            =   3960
      TabIndex        =   49
      Text            =   "лгма"
      Top             =   4320
      Width           =   855
   End
   Begin VB.ComboBox Combo8 
      Height          =   315
      Left            =   2880
      TabIndex        =   48
      Text            =   "глеяа"
      Top             =   4320
      Width           =   975
   End
   Begin VB.CommandButton Command12 
      BackColor       =   &H00E9C5AD&
      Caption         =   "сглеяимг"
      Height          =   375
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H00E9C5AD&
      Caption         =   "OK"
      Height          =   375
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   1680
      Width           =   375
   End
   Begin VB.ComboBox Combo7 
      Height          =   315
      Left            =   4920
      TabIndex        =   45
      Text            =   "етос"
      Top             =   1680
      Width           =   855
   End
   Begin VB.ComboBox Combo6 
      Height          =   315
      Left            =   3960
      TabIndex        =   44
      Text            =   "лгмас"
      Top             =   1680
      Width           =   855
   End
   Begin VB.ComboBox Combo5 
      Height          =   315
      Left            =   2880
      TabIndex        =   43
      Text            =   "глеяа"
      Top             =   1680
      Width           =   975
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   3600
      TabIndex        =   41
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00E9C5AD&
      Caption         =   "ой"
      Height          =   375
      Left            =   10920
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   120
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Caption         =   "еныжкгсг  тилокоциоу"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   2280
      TabIndex        =   32
      Top             =   7080
      Width           =   2655
      Begin VB.ComboBox Combo13 
         Height          =   315
         Left            =   4800
         TabIndex        =   55
         Text            =   "етос"
         Top             =   1080
         Width           =   735
      End
      Begin VB.ComboBox Combo12 
         Height          =   315
         Left            =   3720
         TabIndex        =   54
         Text            =   "лгмас"
         Top             =   1080
         Width           =   975
      End
      Begin VB.ComboBox Combo11 
         Height          =   315
         Left            =   2760
         TabIndex        =   53
         Text            =   "глеяа"
         Top             =   1080
         Width           =   855
      End
      Begin VB.ComboBox Combo4 
         Height          =   315
         Left            =   4680
         TabIndex        =   42
         Top             =   1920
         Width           =   855
      End
      Begin VB.CommandButton Command8 
         Caption         =   "еныжкгсг"
         Height          =   495
         Left            =   600
         TabIndex        =   39
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox Text15 
         Height          =   285
         Left            =   1800
         TabIndex        =   38
         Top             =   1920
         Width           =   2775
      End
      Begin VB.TextBox Text14 
         Height          =   285
         Left            =   1800
         TabIndex        =   37
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox Text13 
         Height          =   285
         Left            =   1800
         TabIndex        =   36
         Top             =   600
         Width           =   3735
      End
      Begin VB.Label Label16 
         Caption         =   "глеяолгмиа     еныжкгсгс"
         Height          =   375
         Left            =   120
         TabIndex        =   35
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label15 
         Caption         =   "аяихлос тилокоциоу"
         Height          =   375
         Left            =   120
         TabIndex        =   34
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label Label14 
         Caption         =   "аяихлос епитацгс"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   1920
         Width           =   2175
      End
   End
   Begin VB.TextBox Text12 
      Height          =   285
      Left            =   1800
      TabIndex        =   31
      Top             =   4920
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00E9C5AD&
      Caption         =   "епистяожг сто аявийо лемоу"
      Height          =   735
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   9600
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000013&
      Caption         =   "амафгтгсг  тилокоциым йаи епитацым"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4695
      Left            =   7680
      TabIndex        =   19
      Top             =   5640
      Width           =   7215
      Begin VB.ComboBox Combo28 
         Height          =   315
         Left            =   1080
         TabIndex        =   109
         Top             =   3240
         Width           =   975
      End
      Begin VB.TextBox Text26 
         Height          =   285
         Left            =   120
         TabIndex        =   108
         Text            =   "ока"
         Top             =   3240
         Width           =   855
      End
      Begin VB.CommandButton Command26 
         BackColor       =   &H00E9C5AD&
         Caption         =   "еуяесг епитацгс"
         Height          =   855
         Left            =   5880
         Style           =   1  'Graphical
         TabIndex        =   99
         Top             =   3600
         Width           =   1095
      End
      Begin VB.CommandButton Command25 
         BackColor       =   &H00E9C5AD&
         Caption         =   "еуяесг пистытийоу"
         Height          =   855
         Left            =   4560
         Style           =   1  'Graphical
         TabIndex        =   98
         Top             =   3600
         Width           =   1095
      End
      Begin VB.CommandButton Command21 
         BackColor       =   &H00E9C5AD&
         Caption         =   "йахаяислос окым тым педиым"
         Height          =   615
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   71
         Top             =   3960
         Width           =   1695
      End
      Begin VB.ComboBox Combo20 
         Height          =   315
         Left            =   1200
         TabIndex        =   70
         Top             =   2400
         Width           =   855
      End
      Begin VB.TextBox Text16 
         Height          =   285
         Left            =   120
         TabIndex        =   68
         Top             =   2400
         Width           =   855
      End
      Begin VB.CommandButton Command20 
         BackColor       =   &H00E9C5AD&
         Caption         =   "сглеяимг"
         Height          =   375
         Left            =   6120
         Style           =   1  'Graphical
         TabIndex        =   67
         Top             =   2040
         Width           =   975
      End
      Begin VB.CommandButton Command19 
         BackColor       =   &H00E9C5AD&
         Caption         =   "OK"
         Height          =   375
         Left            =   5640
         Style           =   1  'Graphical
         TabIndex        =   66
         Top             =   960
         Width           =   375
      End
      Begin VB.ComboBox Combo19 
         Height          =   315
         Left            =   6120
         TabIndex        =   63
         Text            =   "етос"
         Top             =   1560
         Width           =   975
      End
      Begin VB.ComboBox Combo18 
         Height          =   315
         Left            =   5040
         TabIndex        =   62
         Text            =   "лгмас"
         Top             =   1560
         Width           =   975
      End
      Begin VB.ComboBox Combo17 
         Height          =   315
         Left            =   3960
         TabIndex        =   61
         Text            =   "глеяа"
         Top             =   1560
         Width           =   975
      End
      Begin VB.ComboBox Combo16 
         Height          =   315
         Left            =   6120
         TabIndex        =   60
         Text            =   "етос"
         Top             =   480
         Width           =   975
      End
      Begin VB.ComboBox Combo15 
         Height          =   315
         Left            =   5040
         TabIndex        =   59
         Text            =   "лгмас"
         Top             =   480
         Width           =   975
      End
      Begin VB.ComboBox Combo14 
         Height          =   315
         Left            =   3960
         TabIndex        =   58
         Text            =   "глеяа"
         Top             =   480
         Width           =   975
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00E9C5AD&
         Caption         =   "еуяесг тилокоциоу"
         Height          =   855
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   3600
         Width           =   1095
      End
      Begin VB.TextBox Text11 
         Height          =   285
         Left            =   120
         TabIndex        =   28
         Top             =   1560
         Width           =   1935
      End
      Begin VB.TextBox Text10 
         Height          =   285
         Left            =   2640
         TabIndex        =   26
         Top             =   1560
         Width           =   1095
      End
      Begin VB.TextBox Text9 
         Height          =   285
         Left            =   2640
         TabIndex        =   24
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   120
         TabIndex        =   22
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label23 
         BackColor       =   &H80000013&
         Caption         =   "тупос пистытийоу"
         Height          =   255
         Left            =   120
         TabIndex        =   107
         Top             =   3000
         Width           =   1815
      End
      Begin VB.Label Label17 
         BackColor       =   &H80000013&
         Caption         =   "пкгяылема (маи ╧ ови)"
         Height          =   255
         Left            =   120
         TabIndex        =   69
         Top             =   2160
         Width           =   2055
      End
      Begin VB.Label Label12 
         BackColor       =   &H80000013&
         Caption         =   "посо"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label11 
         BackColor       =   &H80000013&
         Caption         =   "левяи"
         Height          =   255
         Left            =   2160
         TabIndex        =   25
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label Label10 
         BackColor       =   &H80000013&
         Caption         =   "апо"
         Height          =   255
         Left            =   2280
         TabIndex        =   23
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label9 
         BackColor       =   &H80000013&
         Caption         =   "аяих/ тилокоц."
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   480
         Width           =   1335
      End
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E9C5AD&
      Caption         =   "диацяажг тилокоциоу"
      Height          =   615
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   5520
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E9C5AD&
      Caption         =   "еццяажг тилокоциоу"
      Height          =   615
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   5520
      Width           =   1575
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   1800
      TabIndex        =   16
      Top             =   4320
      Width           =   975
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   3600
      TabIndex        =   15
      Top             =   3840
      Width           =   1215
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   1800
      TabIndex        =   14
      Top             =   3840
      Width           =   1575
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   1800
      TabIndex        =   13
      Top             =   1680
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   3600
      TabIndex        =   12
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   1800
      TabIndex        =   11
      Top             =   2280
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1800
      TabIndex        =   10
      Top             =   2760
      Width           =   3015
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1800
      TabIndex        =   9
      Top             =   1200
      Width           =   3015
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
      Height          =   405
      Left            =   7560
      TabIndex        =   2
      Top             =   120
      Width           =   3255
   End
   Begin VB.Line Line18 
      Visible         =   0   'False
      X1              =   4845
      X2              =   4845
      Y1              =   1080
      Y2              =   5280
   End
   Begin VB.Line Line17 
      Visible         =   0   'False
      X1              =   1725
      X2              =   1725
      Y1              =   1005
      Y2              =   5505
   End
   Begin VB.Line Line16 
      Visible         =   0   'False
      X1              =   100
      X2              =   7300
      Y1              =   4755
      Y2              =   4755
   End
   Begin VB.Line Line15 
      Visible         =   0   'False
      X1              =   100
      X2              =   7300
      Y1              =   4230
      Y2              =   4230
   End
   Begin VB.Line Line14 
      Visible         =   0   'False
      X1              =   100
      X2              =   7300
      Y1              =   3705
      Y2              =   3705
   End
   Begin VB.Line Line13 
      Visible         =   0   'False
      X1              =   100
      X2              =   7300
      Y1              =   3180
      Y2              =   3180
   End
   Begin VB.Line Line12 
      Visible         =   0   'False
      X1              =   100
      X2              =   7300
      Y1              =   2655
      Y2              =   2655
   End
   Begin VB.Line Line11 
      Visible         =   0   'False
      X1              =   100
      X2              =   7300
      Y1              =   2130
      Y2              =   2130
   End
   Begin VB.Line Line10 
      Visible         =   0   'False
      X1              =   100
      X2              =   7300
      Y1              =   1605
      Y2              =   1605
   End
   Begin VB.Line Line9 
      Visible         =   0   'False
      X1              =   165
      X2              =   7365
      Y1              =   5280
      Y2              =   5280
   End
   Begin VB.Line Line8 
      Visible         =   0   'False
      X1              =   7365
      X2              =   7365
      Y1              =   1080
      Y2              =   5280
   End
   Begin VB.Line Line7 
      Visible         =   0   'False
      X1              =   165
      X2              =   165
      Y1              =   1080
      Y2              =   5280
   End
   Begin VB.Line Line6 
      Visible         =   0   'False
      X1              =   165
      X2              =   7365
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Line Line5 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Visible         =   0   'False
      X1              =   45
      X2              =   7485
      Y1              =   6360
      Y2              =   6360
   End
   Begin VB.Line Line4 
      BorderWidth     =   3
      Visible         =   0   'False
      X1              =   7485
      X2              =   7485
      Y1              =   0
      Y2              =   10320
   End
   Begin VB.Line Line3 
      BorderWidth     =   3
      Visible         =   0   'False
      X1              =   7485
      X2              =   15045
      Y1              =   5490
      Y2              =   5490
   End
   Begin VB.Label Label13 
      BackColor       =   &H80000013&
      Caption         =   "аяихлос епитацгс"
      Height          =   255
      Left            =   240
      TabIndex        =   30
      Top             =   4920
      Width           =   1695
   End
   Begin VB.Label Label8 
      BackColor       =   &H80000013&
      Caption         =   "глеяолгмиа еныжкгсгс"
      Height          =   495
      Left            =   240
      TabIndex        =   8
      Top             =   4320
      Width           =   1455
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000013&
      Caption         =   "еныжкгсг"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   3720
      Width           =   1695
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000013&
      Caption         =   "еийома"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000013&
      Caption         =   "глеяолгмиа  ейдосгс"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000013&
      Caption         =   "тупос"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000013&
      Caption         =   "посо"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000013&
      Caption         =   "аяихлос тилокоциоу"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000013&
      Caption         =   "омола етаияиас"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5160
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "ZETAIRIES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Adodc1_WillMove(ByVal adReason As ADODB.EventReasonEnum, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)

End Sub

Private Sub Combo1_Click()
Select Case Combo1.ListIndex
    Case 0
        Form4.Text4.Text = "тилок\пыкгсгс"
        Form4.Text6.Text = ""
        Form4.Combo27.Visible = False
    Case 1
        Form4.Text4.Text = "пистытийо"
        Form4.Text6.Text = "ови"
        Form4.Combo27.Visible = True
    End Select
    Combo1.Text = ""
End Sub

Private Sub Combo10_Click()
Select Case Combo10.ListIndex
    Case 0
        ETOS_HM_EKSOFLISIS = "2005"
    Case 1
         ETOS_HM_EKSOFLISIS = "2006"
    Case 2
        ETOS_HM_EKSOFLISIS = "2007"
    Case 3
         ETOS_HM_EKSOFLISIS = "2008"
    Case 4
        ETOS_HM_EKSOFLISIS = "2009"
    Case 5
         ETOS_HM_EKSOFLISIS = "2010"
    Case 6
        ETOS_HM_EKSOFLISIS = "2011"
    Case 7
         ETOS_HM_EKSOFLISIS = "2012"
    Case 8
        ETOS_HM_EKSOFLISIS = "2013"
    Case 9
         ETOS_HM_EKSOFLISIS = "2014"
    Case 10
        ETOS_HM_EKSOFLISIS = "2015"
    Case 11
         ETOS_HM_EKSOFLISIS = "2016"
    Case 12
        ETOS_HM_EKSOFLISIS = "2017"
    Case 13
         ETOS_HM_EKSOFLISIS = "2018"
    Case 14
        ETOS_HM_EKSOFLISIS = "2019"
    Case 15
         ETOS_HM_EKSOFLISIS = "2020"
End Select
End Sub

Private Sub Combo11_Click()
Select Case Combo11.ListIndex
    Case 0
        DAY_HM_EKSOFLISIS_2 = "1"
    Case 1
         DAY_HM_EKSOFLISIS_2 = "2"
    Case 2
        DAY_HM_EKSOFLISIS_2 = "3"
    Case 3
         DAY_HM_EKSOFLISIS_2 = "4"
    Case 4
        DAY_HM_EKSOFLISIS_2 = "5"
    Case 5
         DAY_HM_EKSOFLISIS_2 = "6"
    Case 6
        DAY_HM_EKSOFLISIS_2 = "7"
    Case 7
         DAY_HM_EKSOFLISIS_2 = "8"
    Case 8
        DAY_HM_EKSOFLISIS_2 = "9"
    Case 9
         DAY_HM_EKSOFLISIS_2 = "10"
    Case 10
        DAY_HM_EKSOFLISIS_2 = "11"
    Case 11
         DAY_HM_EKSOFLISIS_2 = "12"
    Case 12
        DAY_HM_EKSOFLISIS_2 = "13"
    Case 13
         DAY_HM_EKSOFLISIS_2 = "14"
    Case 14
        DAY_HM_EKSOFLISIS_2 = "15"
    Case 15
         DAY_HM_EKSOFLISIS_2 = "16"
    Case 16
        DAY_HM_EKSOFLISIS_2 = "17"
    Case 17
         DAY_HM_EKSOFLISIS_2 = "18"
    Case 18
        DAY_HM_EKSOFLISIS_2 = "19"
    Case 19
         DAY_HM_EKSOFLISIS_2 = "20"
    Case 20
        DAY_HM_EKSOFLISIS_2 = "21"
    Case 21
         DAY_HM_EKSOFLISIS_2 = "22"
    Case 22
        DAY_HM_EKSOFLISIS_2 = "23"
    Case 23
         DAY_HM_EKSOFLISIS_2 = "24"
    Case 24
        DAY_HM_EKSOFLISIS_2 = "25"
    Case 25
         DAY_HM_EKSOFLISIS_2 = "26"
    Case 26
        DAY_HM_EKSOFLISIS_2 = "27"
    Case 27
         DAY_HM_EKSOFLISIS_2 = "28"
    Case 28
        DAY_HM_EKSOFLISIS_2 = "29"
    Case 29
         DAY_HM_EKSOFLISIS_2 = "30"
    Case 30
        DAY_HM_EKSOFLISIS_2 = "31"
    End Select
End Sub

Private Sub Combo12_Click()
Select Case Combo12.ListIndex
    Case 0
        MONTH_HM_EKSOFLISIS_2 = "1"
    Case 1
         MONTH_HM_EKSOFLISIS_2 = "2"
    Case 2
        MONTH_HM_EKSOFLISIS_2 = "3"
    Case 3
         MONTH_HM_EKSOFLISIS_2 = "4"
    Case 4
        MONTH_HM_EKSOFLISIS_2 = "5"
    Case 5
         MONTH_HM_EKSOFLISIS_2 = "6"
    Case 6
        MONTH_HM_EKSOFLISIS_2 = "7"
    Case 7
         MONTH_HM_EKSOFLISIS_2 = "8"
    Case 8
        MONTH_HM_EKSOFLISIS_2 = "9"
    Case 9
         MONTH_HM_EKSOFLISIS_2 = "10"
    Case 10
        MONTH_HM_EKSOFLISIS_2 = "11"
    Case 11
         MONTH_HM_EKSOFLISIS_2 = "12"
    End Select
End Sub

Private Sub Combo13_Click()
Select Case Combo13.ListIndex
    Case 0
        ETOS_HM_EKSOFLISIS_2 = "2005"
    Case 1
         ETOS_HM_EKSOFLISIS_2 = "2006"
    Case 2
        ETOS_HM_EKSOFLISIS_2 = "2007"
    Case 3
         ETOS_HM_EKSOFLISIS_2 = "2008"
    Case 4
        ETOS_HM_EKSOFLISIS_2 = "2009"
    Case 5
         ETOS_HM_EKSOFLISIS_2 = "2010"
    Case 6
        ETOS_HM_EKSOFLISIS_2 = "2011"
    Case 7
         ETOS_HM_EKSOFLISIS_2 = "2012"
    Case 8
        ETOS_HM_EKSOFLISIS_2 = "2013"
    Case 9
         ETOS_HM_EKSOFLISIS_2 = "2014"
    Case 10
        ETOS_HM_EKSOFLISIS_2 = "2015"
    Case 11
         ETOS_HM_EKSOFLISIS_2 = "2016"
    Case 12
        ETOS_HM_EKSOFLISIS_2 = "2017"
    Case 13
         ETOS_HM_EKSOFLISIS_2 = "2018"
    Case 14
        ETOS_HM_EKSOFLISIS_2 = "2019"
    Case 15
         ETOS_HM_EKSOFLISIS_2 = "2020"
End Select
End Sub

Private Sub Combo14_Click()
Select Case Combo14.ListIndex
    Case 0
        DAY_ANAZHTHSH_1 = "1"
    Case 1
         DAY_ANAZHTHSH_1 = "2"
    Case 2
        DAY_ANAZHTHSH_1 = "3"
    Case 3
         DAY_ANAZHTHSH_1 = "4"
    Case 4
        DAY_ANAZHTHSH_1 = "5"
    Case 5
         DAY_ANAZHTHSH_1 = "6"
    Case 6
        DAY_ANAZHTHSH_1 = "7"
    Case 7
         DAY_ANAZHTHSH_1 = "8"
    Case 8
        DAY_ANAZHTHSH_1 = "9"
    Case 9
         DAY_ANAZHTHSH_1 = "10"
    Case 10
        DAY_ANAZHTHSH_1 = "11"
    Case 11
         DAY_ANAZHTHSH_1 = "12"
    Case 12
        DAY_ANAZHTHSH_1 = "13"
    Case 13
         DAY_ANAZHTHSH_1 = "14"
    Case 14
        DAY_ANAZHTHSH_1 = "15"
    Case 15
         DAY_ANAZHTHSH_1 = "16"
    Case 16
        DAY_ANAZHTHSH_1 = "17"
    Case 17
         DAY_ANAZHTHSH_1 = "18"
    Case 18
        DAY_ANAZHTHSH_1 = "19"
    Case 19
         DAY_ANAZHTHSH_1 = "20"
    Case 20
        DAY_ANAZHTHSH_1 = "21"
    Case 21
         DAY_ANAZHTHSH_1 = "22"
    Case 22
        DAY_ANAZHTHSH_1 = "23"
    Case 23
         DAY_ANAZHTHSH_1 = "24"
    Case 24
        DAY_ANAZHTHSH_1 = "25"
    Case 25
         DAY_ANAZHTHSH_1 = "26"
    Case 26
        DAY_ANAZHTHSH_1 = "27"
    Case 27
         DAY_ANAZHTHSH_1 = "28"
    Case 28
        DAY_ANAZHTHSH_1 = "29"
    Case 29
         DAY_ANAZHTHSH_1 = "30"
    Case 30
        DAY_ANAZHTHSH_1 = "31"
    End Select
End Sub

Private Sub Combo15_Click()
Select Case Combo15.ListIndex
    Case 0
        MONTH_ANAZHTHSH_1 = "1"
    Case 1
         MONTH_ANAZHTHSH_1 = "2"
    Case 2
        MONTH_ANAZHTHSH_1 = "3"
    Case 3
         MONTH_ANAZHTHSH_1 = "4"
    Case 4
        MONTH_ANAZHTHSH_1 = "5"
    Case 5
         MONTH_ANAZHTHSH_1 = "6"
    Case 6
        MONTH_ANAZHTHSH_1 = "7"
    Case 7
         MONTH_ANAZHTHSH_1 = "8"
    Case 8
        MONTH_ANAZHTHSH_1 = "9"
    Case 9
         MONTH_ANAZHTHSH_1 = "10"
    Case 10
        MONTH_ANAZHTHSH_1 = "11"
    Case 11
         MONTH_ANAZHTHSH_1 = "12"
    End Select
End Sub

Private Sub Combo16_Click()
Select Case Combo16.ListIndex
    Case 0
        ETOS_ANAZHTHSH_1 = "2005"
    Case 1
         ETOS_ANAZHTHSH_1 = "2006"
    Case 2
        ETOS_ANAZHTHSH_1 = "2007"
    Case 3
         ETOS_ANAZHTHSH_1 = "2008"
    Case 4
        ETOS_ANAZHTHSH_1 = "2009"
    Case 5
         ETOS_ANAZHTHSH_1 = "2010"
    Case 6
        ETOS_ANAZHTHSH_1 = "2011"
    Case 7
         ETOS_ANAZHTHSH_1 = "2012"
    Case 8
        ETOS_ANAZHTHSH_1 = "2013"
    Case 9
         ETOS_ANAZHTHSH_1 = "2014"
    Case 10
        ETOS_ANAZHTHSH_1 = "2015"
    Case 11
         ETOS_ANAZHTHSH_1 = "2016"
    Case 12
        ETOS_ANAZHTHSH_1 = "2017"
    Case 13
         ETOS_ANAZHTHSH_1 = "2018"
    Case 14
        ETOS_ANAZHTHSH_1 = "2019"
    Case 15
         ETOS_ANAZHTHSH_1 = "2020"
End Select
End Sub




Private Sub Combo17_Click()
Select Case Combo17.ListIndex
    Case 0
        DAY_ANAZHTHSH_2 = "1"
    Case 1
         DAY_ANAZHTHSH_2 = "2"
    Case 2
        DAY_ANAZHTHSH_2 = "3"
    Case 3
         DAY_ANAZHTHSH_2 = "4"
    Case 4
        DAY_ANAZHTHSH_2 = "5"
    Case 5
         DAY_ANAZHTHSH_2 = "6"
    Case 6
        DAY_ANAZHTHSH_2 = "7"
    Case 7
         DAY_ANAZHTHSH_2 = "8"
    Case 8
        DAY_ANAZHTHSH_2 = "9"
    Case 9
         DAY_ANAZHTHSH_2 = "10"
    Case 10
        DAY_ANAZHTHSH_2 = "11"
    Case 11
         DAY_ANAZHTHSH_2 = "12"
    Case 12
        DAY_ANAZHTHSH_2 = "13"
    Case 13
         DAY_ANAZHTHSH_2 = "14"
    Case 14
        DAY_ANAZHTHSH_2 = "15"
    Case 15
         DAY_ANAZHTHSH_2 = "16"
    Case 16
        DAY_ANAZHTHSH_2 = "17"
    Case 17
         DAY_ANAZHTHSH_2 = "18"
    Case 18
        DAY_ANAZHTHSH_2 = "19"
    Case 19
         DAY_ANAZHTHSH_2 = "20"
    Case 20
        DAY_ANAZHTHSH_2 = "21"
    Case 21
         DAY_ANAZHTHSH_2 = "22"
    Case 22
        DAY_ANAZHTHSH_2 = "23"
    Case 23
         DAY_ANAZHTHSH_2 = "24"
    Case 24
        DAY_ANAZHTHSH_2 = "25"
    Case 25
         DAY_ANAZHTHSH_2 = "26"
    Case 26
        DAY_ANAZHTHSH_2 = "27"
    Case 27
         DAY_ANAZHTHSH_2 = "28"
    Case 28
        DAY_ANAZHTHSH_2 = "29"
    Case 29
         DAY_ANAZHTHSH_2 = "30"
    Case 30
        DAY_ANAZHTHSH_2 = "31"
    End Select
End Sub

Private Sub Combo18_Click()
Select Case Combo18.ListIndex
    Case 0
        MONTH_ANAZHTHSH_2 = "1"
    Case 1
         MONTH_ANAZHTHSH_2 = "2"
    Case 2
        MONTH_ANAZHTHSH_2 = "3"
    Case 3
         MONTH_ANAZHTHSH_2 = "4"
    Case 4
        MONTH_ANAZHTHSH_2 = "5"
    Case 5
         MONTH_ANAZHTHSH_2 = "6"
    Case 6
        MONTH_ANAZHTHSH_2 = "7"
    Case 7
         MONTH_ANAZHTHSH_2 = "8"
    Case 8
        MONTH_ANAZHTHSH_2 = "9"
    Case 9
         MONTH_ANAZHTHSH_2 = "10"
    Case 10
        MONTH_ANAZHTHSH_2 = "11"
    Case 11
         MONTH_ANAZHTHSH_2 = "12"
    End Select
End Sub

Private Sub Combo19_Click()
Select Case Combo19.ListIndex
    Case 0
        ETOS_ANAZHTHSH_2 = "2005"
    Case 1
         ETOS_ANAZHTHSH_2 = "2006"
    Case 2
        ETOS_ANAZHTHSH_2 = "2007"
    Case 3
         ETOS_ANAZHTHSH_2 = "2008"
    Case 4
        ETOS_ANAZHTHSH_2 = "2009"
    Case 5
         ETOS_ANAZHTHSH_2 = "2010"
    Case 6
        ETOS_ANAZHTHSH_2 = "2011"
    Case 7
         ETOS_ANAZHTHSH_2 = "2012"
    Case 8
        ETOS_ANAZHTHSH_2 = "2013"
    Case 9
         ETOS_ANAZHTHSH_2 = "2014"
    Case 10
        ETOS_ANAZHTHSH_2 = "2015"
    Case 11
         ETOS_ANAZHTHSH_2 = "2016"
    Case 12
        ETOS_ANAZHTHSH_2 = "2017"
    Case 13
         ETOS_ANAZHTHSH_2 = "2018"
    Case 14
        ETOS_ANAZHTHSH_2 = "2019"
    Case 15
         ETOS_ANAZHTHSH_2 = "2020"
End Select
End Sub

Private Sub Combo2_Click()
Select Case Combo2.ListIndex
    Case 0
        Form4.Text6.Text = "ови"
    Case 1
        Form4.Text6.Text = "маи"
    End Select
    Combo2.Text = ""
End Sub

Private Sub Combo20_Click()
Select Case Combo20.ListIndex
    Case 0
        Text16.Text = "маи"
    Case 1
        Text16.Text = "ови"
    End Select
End Sub

Private Sub Combo21_Click()
Select Case Combo21.ListIndex
    Case 0
        DAY_HM_APO_KARTELA = "1"
    Case 1
         DAY_HM_APO_KARTELA = "2"
    Case 2
        DAY_HM_APO_KARTELA = "3"
    Case 3
         DAY_HM_APO_KARTELA = "4"
    Case 4
        DAY_HM_APO_KARTELA = "5"
    Case 5
         DAY_HM_APO_KARTELA = "6"
    Case 6
        DAY_HM_APO_KARTELA = "7"
    Case 7
         DAY_HM_APO_KARTELA = "8"
    Case 8
        DAY_HM_APO_KARTELA = "9"
    Case 9
         DAY_HM_APO_KARTELA = "10"
    Case 10
        DAY_HM_APO_KARTELA = "11"
    Case 11
         DAY_HM_APO_KARTELA = "12"
    Case 12
        DAY_HM_APO_KARTELA = "13"
    Case 13
         DAY_HM_APO_KARTELA = "14"
    Case 14
        DAY_HM_APO_KARTELA = "15"
    Case 15
         DAY_HM_APO_KARTELA = "16"
    Case 16
        DAY_HM_APO_KARTELA = "17"
    Case 17
         DAY_HM_APO_KARTELA = "18"
    Case 18
        DAY_HM_APO_KARTELA = "19"
    Case 19
         DAY_HM_APO_KARTELA = "20"
    Case 20
        DAY_HM_APO_KARTELA = "21"
    Case 21
         DAY_HM_APO_KARTELA = "22"
    Case 22
        DAY_HM_APO_KARTELA = "23"
    Case 23
         DAY_HM_APO_KARTELA = "24"
    Case 24
        DAY_HM_APO_KARTELA = "25"
    Case 25
         DAY_HM_APO_KARTELA = "26"
    Case 26
        DAY_HM_APO_KARTELA = "27"
    Case 27
         DAY_HM_APO_KARTELA = "28"
    Case 28
        DAY_HM_APO_KARTELA = "29"
    Case 29
         DAY_HM_APO_KARTELA = "30"
    Case 30
        DAY_HM_APO_KARTELA = "31"
    End Select
End Sub

Private Sub Combo22_Click()
Select Case Combo22.ListIndex
    Case 0
        MONTH_HM_APO_KARTELA = "1"
    Case 1
         MONTH_HM_APO_KARTELA = "2"
    Case 2
        MONTH_HM_APO_KARTELA = "3"
    Case 3
         MONTH_HM_APO_KARTELA = "4"
    Case 4
        MONTH_HM_APO_KARTELA = "5"
    Case 5
         MONTH_HM_APO_KARTELA = "6"
    Case 6
        MONTH_HM_APO_KARTELA = "7"
    Case 7
         MONTH_HM_APO_KARTELA = "8"
    Case 8
        MONTH_HM_APO_KARTELA = "9"
    Case 9
         MONTH_HM_APO_KARTELA = "10"
    Case 10
        MONTH_HM_APO_KARTELA = "11"
    Case 11
         MONTH_HM_APO_KARTELA = "12"
    End Select
End Sub

Private Sub Combo23_Click()
Select Case Combo23.ListIndex
    Case 0
        ETOS_HM_APO_KARTELA = "2005"
    Case 1
         ETOS_HM_APO_KARTELA = "2006"
    Case 2
       ETOS_HM_APO_KARTELA = "2007"
    Case 3
         ETOS_HM_APO_KARTELA = "2008"
    Case 4
        ETOS_HM_APO_KARTELA = "2009"
    Case 5
         ETOS_HM_APO_KARTELA = "2010"
    Case 6
        ETOS_HM_APO_KARTELA = "2011"
    Case 7
         ETOS_HM_APO_KARTELA = "2012"
    Case 8
        ETOS_HM_APO_KARTELA = "2013"
    Case 9
         ETOS_HM_APO_KARTELA = "2014"
    Case 10
        ETOS_HM_APO_KARTELA = "2015"
    Case 11
         ETOS_HM_APO_KARTELA = "2016"
    Case 12
       ETOS_HM_APO_KARTELA = "2017"
    Case 13
         ETOS_HM_APO_KARTELA = "2018"
    Case 14
        ETOS_HM_APO_KARTELA = "2019"
    Case 15
         ETOS_HM_APO_KARTELA = "2020"
End Select
End Sub

Private Sub Combo24_Click()
Select Case Combo24.ListIndex
    Case 0
        DAY_HM_MEXRI_KARTELA = "1"
    Case 1
         DAY_HM_MEXRI_KARTELA = "2"
    Case 2
        DAY_HM_MEXRI_KARTELA = "3"
    Case 3
         DAY_HM_MEXRI_KARTELA = "4"
    Case 4
        DAY_HM_MEXRI_KARTELA = "5"
    Case 5
         DAY_HM_MEXRI_KARTELA = "6"
    Case 6
        DAY_HM_MEXRI_KARTELA = "7"
    Case 7
         DAY_HM_MEXRI_KARTELA = "8"
    Case 8
        DAY_HM_MEXRI_KARTELA = "9"
    Case 9
         DAY_HM_MEXRI_KARTELA = "10"
    Case 10
        DAY_HM_MEXRI_KARTELA = "11"
    Case 11
         DAY_HM_MEXRI_KARTELA = "12"
    Case 12
        DAY_HM_MEXRI_KARTELA = "13"
    Case 13
         DAY_HM_MEXRI_KARTELA = "14"
    Case 14
        DAY_HM_MEXRI_KARTELA = "15"
    Case 15
         DAY_HM_MEXRI_KARTELA = "16"
    Case 16
        DAY_HM_MEXRI_KARTELA = "17"
    Case 17
         DAY_HM_MEXRI_KARTELA = "18"
    Case 18
        DAY_HM_MEXRI_KARTELA = "19"
    Case 19
         DAY_HM_MEXRI_KARTELA = "20"
    Case 20
        DAY_HM_MEXRI_KARTELA = "21"
    Case 21
         DAY_HM_MEXRI_KARTELA = "22"
    Case 22
        DAY_HM_MEXRI_KARTELA = "23"
    Case 23
         DAY_HM_MEXRI_KARTELA = "24"
    Case 24
        DAY_HM_MEXRI_KARTELA = "25"
    Case 25
         DAY_HM_MEXRI_KARTELA = "26"
    Case 26
        DAY_HM_MEXRI_KARTELA = "27"
    Case 27
         DAY_HM_MEXRI_KARTELA = "28"
    Case 28
        DAY_HM_MEXRI_KARTELA = "29"
    Case 29
         DAY_HM_MEXRI_KARTELA = "30"
    Case 30
        DAY_HM_MEXRI_KARTELA = "31"
    End Select
End Sub

Private Sub Combo25_Click()
Select Case Combo25.ListIndex
    Case 0
        MONTH_HM_MEXRI_KARTELA = "1"
    Case 1
         MONTH_HM_MEXRI_KARTELA = "2"
    Case 2
        MONTH_HM_MEXRI_KARTELA = "3"
    Case 3
         MONTH_HM_MEXRI_KARTELA = "4"
    Case 4
        MONTH_HM_MEXRI_KARTELA = "5"
    Case 5
         MONTH_HM_MEXRI_KARTELA = "6"
    Case 6
        MONTH_HM_MEXRI_KARTELA = "7"
    Case 7
         MONTH_HM_MEXRI_KARTELA = "8"
    Case 8
        MONTH_HM_MEXRI_KARTELA = "9"
    Case 9
         MONTH_HM_MEXRI_KARTELA = "10"
    Case 10
        MONTH_HM_MEXRI_KARTELA = "11"
    Case 11
         MONTH_HM_MEXRI_KARTELA = "12"
    End Select
End Sub

Private Sub Combo26_Click()
Select Case Combo26.ListIndex
    Case 0
        ETOS_HM_MEXRI_KARTELA = "2005"
    Case 1
         ETOS_HM_MEXRI_KARTELA = "2006"
    Case 2
       ETOS_HM_MEXRI_KARTELA = "2007"
    Case 3
         ETOS_HM_MEXRI_KARTELA = "2008"
    Case 4
        ETOS_HM_MEXRI_KARTELA = "2009"
    Case 5
         ETOS_HM_MEXRI_KARTELA = "2010"
    Case 6
        ETOS_HM_MEXRI_KARTELA = "2011"
    Case 7
         ETOS_HM_MEXRI_KARTELA = "2012"
    Case 8
        ETOS_HM_MEXRI_KARTELA = "2013"
    Case 9
         ETOS_HM_MEXRI_KARTELA = "2014"
    Case 10
        ETOS_HM_MEXRI_KARTELA = "2015"
    Case 11
         ETOS_HM_MEXRI_KARTELA = "2016"
    Case 12
       ETOS_HM_MEXRI_KARTELA = "2017"
    Case 13
         ETOS_HM_MEXRI_KARTELA = "2018"
    Case 14
        ETOS_HM_MEXRI_KARTELA = "2019"
    Case 15
         ETOS_HM_MEXRI_KARTELA = "2020"
End Select
End Sub

Private Sub Combo28_Click()
Select Case Combo28.ListIndex
    Case 0
        Form4.Text26.Text = "ока"
    Case 1
        Form4.Text26.Text = "цемийо"
    Case 2
        Form4.Text26.Text = "COOP"
    End Select
    Combo28.Text = ""
End Sub

Private Sub Combo3_Click()
Select Case Combo3.ListIndex
    Case 0
        Form4.Text12.Text = "летягта"
    Case 1
        Form4.Text12.Text = "епитацг"
    End Select
    Combo3.Text = ""
End Sub

Private Sub Combo4_Click()
Select Case Combo4.ListIndex
    Case 0
        Form4.Text15.Text = "летягта"
    Case 1
        Form4.Text15.Text = "епитацг"
    End Select
    Combo4.Text = ""
End Sub

Private Sub Combo5_Click()
Select Case Combo5.ListIndex
    Case 0
        DAY_HM_EKDOSHS = "1"
    Case 1
         DAY_HM_EKDOSHS = "2"
    Case 2
        DAY_HM_EKDOSHS = "3"
    Case 3
         DAY_HM_EKDOSHS = "4"
    Case 4
        DAY_HM_EKDOSHS = "5"
    Case 5
         DAY_HM_EKDOSHS = "6"
    Case 6
        DAY_HM_EKDOSHS = "7"
    Case 7
         DAY_HM_EKDOSHS = "8"
    Case 8
        DAY_HM_EKDOSHS = "9"
    Case 9
         DAY_HM_EKDOSHS = "10"
    Case 10
        DAY_HM_EKDOSHS = "11"
    Case 11
         DAY_HM_EKDOSHS = "12"
    Case 12
        DAY_HM_EKDOSHS = "13"
    Case 13
         DAY_HM_EKDOSHS = "14"
    Case 14
        DAY_HM_EKDOSHS = "15"
    Case 15
         DAY_HM_EKDOSHS = "16"
    Case 16
        DAY_HM_EKDOSHS = "17"
    Case 17
         DAY_HM_EKDOSHS = "18"
    Case 18
        DAY_HM_EKDOSHS = "19"
    Case 19
         DAY_HM_EKDOSHS = "20"
    Case 20
        DAY_HM_EKDOSHS = "21"
    Case 21
         DAY_HM_EKDOSHS = "22"
    Case 22
        DAY_HM_EKDOSHS = "23"
    Case 23
         DAY_HM_EKDOSHS = "24"
    Case 24
        DAY_HM_EKDOSHS = "25"
    Case 25
         DAY_HM_EKDOSHS = "26"
    Case 26
        DAY_HM_EKDOSHS = "27"
    Case 27
         DAY_HM_EKDOSHS = "28"
    Case 28
        DAY_HM_EKDOSHS = "29"
    Case 29
         DAY_HM_EKDOSHS = "30"
    Case 30
        DAY_HM_EKDOSHS = "31"
    End Select
End Sub

Private Sub Combo6_Click()
Select Case Combo6.ListIndex
    Case 0
        MONTH_HM_EKDOSHS = "1"
    Case 1
         MONTH_HM_EKDOSHS = "2"
    Case 2
        MONTH_HM_EKDOSHS = "3"
    Case 3
         MONTH_HM_EKDOSHS = "4"
    Case 4
        MONTH_HM_EKDOSHS = "5"
    Case 5
         MONTH_HM_EKDOSHS = "6"
    Case 6
        MONTH_HM_EKDOSHS = "7"
    Case 7
         MONTH_HM_EKDOSHS = "8"
    Case 8
        MONTH_HM_EKDOSHS = "9"
    Case 9
         MONTH_HM_EKDOSHS = "10"
    Case 10
        MONTH_HM_EKDOSHS = "11"
    Case 11
         MONTH_HM_EKDOSHS = "12"
    End Select
End Sub

Private Sub Combo7_Click()
Select Case Combo7.ListIndex
    Case 0
        ETOS_HM_EKDOSHS = "2005"
    Case 1
         ETOS_HM_EKDOSHS = "2006"
    Case 2
        ETOS_HM_EKDOSHS = "2007"
    Case 3
         ETOS_HM_EKDOSHS = "2008"
    Case 4
        ETOS_HM_EKDOSHS = "2009"
    Case 5
         ETOS_HM_EKDOSHS = "2010"
    Case 6
        ETOS_HM_EKDOSHS = "2011"
    Case 7
         ETOS_HM_EKDOSHS = "2012"
    Case 8
        ETOS_HM_EKDOSHS = "2013"
    Case 9
         ETOS_HM_EKDOSHS = "2014"
    Case 10
        ETOS_HM_EKDOSHS = "2015"
    Case 11
         ETOS_HM_EKDOSHS = "2016"
    Case 12
        ETOS_HM_EKDOSHS = "2017"
    Case 13
         ETOS_HM_EKDOSHS = "2018"
    Case 14
        ETOS_HM_EKDOSHS = "2019"
    Case 15
         ETOS_HM_EKDOSHS = "2020"
End Select
End Sub

Private Sub Combo8_Click()
Select Case Combo8.ListIndex
    Case 0
        DAY_HM_EKSOFLISIS = "1"
    Case 1
         DAY_HM_EKSOFLISIS = "2"
    Case 2
        DAY_HM_EKSOFLISIS = "3"
    Case 3
         DAY_HM_EKSOFLISIS = "4"
    Case 4
        DAY_HM_EKSOFLISIS = "5"
    Case 5
         DAY_HM_EKSOFLISIS = "6"
    Case 6
        DAY_HM_EKSOFLISIS = "7"
    Case 7
         DAY_HM_EKSOFLISIS = "8"
    Case 8
        DAY_HM_EKSOFLISIS = "9"
    Case 9
         DAY_HM_EKSOFLISIS = "10"
    Case 10
        DAY_HM_EKSOFLISIS = "11"
    Case 11
         DAY_HM_EKSOFLISIS = "12"
    Case 12
        DAY_HM_EKSOFLISIS = "13"
    Case 13
         DAY_HM_EKSOFLISIS = "14"
    Case 14
        DAY_HM_EKSOFLISIS = "15"
    Case 15
         DAY_HM_EKSOFLISIS = "16"
    Case 16
        DAY_HM_EKSOFLISIS = "17"
    Case 17
         DAY_HM_EKSOFLISIS = "18"
    Case 18
        DAY_HM_EKSOFLISIS = "19"
    Case 19
         DAY_HM_EKSOFLISIS = "20"
    Case 20
        DAY_HM_EKSOFLISIS = "21"
    Case 21
         DAY_HM_EKSOFLISIS = "22"
    Case 22
        DAY_HM_EKSOFLISIS = "23"
    Case 23
         DAY_HM_EKSOFLISIS = "24"
    Case 24
        DAY_HM_EKSOFLISIS = "25"
    Case 25
         DAY_HM_EKSOFLISIS = "26"
    Case 26
        DAY_HM_EKSOFLISIS = "27"
    Case 27
         DAY_HM_EKSOFLISIS = "28"
    Case 28
        DAY_HM_EKSOFLISIS = "29"
    Case 29
         DAY_HM_EKSOFLISIS = "30"
    Case 30
        DAY_HM_EKSOFLISIS = "31"
    End Select
End Sub

Private Sub Combo9_Click()
Select Case Combo9.ListIndex
    Case 0
        MONTH_HM_EKSOFLISIS = "1"
    Case 1
         MONTH_HM_EKSOFLISIS = "2"
    Case 2
        MONTH_HM_EKSOFLISIS = "3"
    Case 3
         MONTH_HM_EKSOFLISIS = "4"
    Case 4
        MONTH_HM_EKSOFLISIS = "5"
    Case 5
         MONTH_HM_EKSOFLISIS = "6"
    Case 6
        MONTH_HM_EKSOFLISIS = "7"
    Case 7
         MONTH_HM_EKSOFLISIS = "8"
    Case 8
        MONTH_HM_EKSOFLISIS = "9"
    Case 9
         MONTH_HM_EKSOFLISIS = "10"
    Case 10
        MONTH_HM_EKSOFLISIS = "11"
    Case 11
         MONTH_HM_EKSOFLISIS = "12"
    End Select
End Sub

Private Sub Command1_Click()
Text1.Text = Text24.Text

' ELEGXOS AN PERNAO TO PROBLEPOMENO MHKOS PEDIOY GIA TEXT2,12 POY EINAI 20,30 ANTOISTIXA
' KAI APOKOPEI PERITOY MEROYS AN XREIAZETAI
Dim L2, L12 As Integer
L2 = Len(Text2.Text)
L12 = Len(Text12.Text)

If L2 > 20 Then
    Text2.Text = Mid(Text2.Text, 1, 20)
Else
    
End If

If L12 > 30 Then
    Text12.Text = Mid(Text12.Text, 1, 30)
Else
    
End If

On Error GoTo er:
If rs1.STATE = 1 Then rs1.Close
If db1.STATE = 1 Then db1.Close
db1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & ZETAIRIES_DIADROMHS_BACKUP_DIAX_1 & ";" & _
"Persist Security Info=False"
db1.Open ZETAIRIES_DIADROMHS_BACKUP_DIAX
rs1.Open "[" & ZETAIRIES.Text1.Text & "]", db1, adOpenDynamic, adLockBatchOptimistic

Dim XREOSH, PISTOSI, YPOLIPO As Double
Dim STATEMENT, half_statement, STP As String
Dim C, A, index
XREOSH = 0
PISTOSI = 0
YPOLIPO = 0
index = 1
C = 1
'***************************** ELEGXOI  ******************************
If rs1.BOF = rs1.EOF Then GoTo NIK:
rs1.MoveFirst
NIK:
Do While Not rs1.EOF
    If rs1![аяихлос_тилокоциоу] <> UCase(Text2.Text) Then
        rs1.MoveNext
    Else
        C = C + 1
        rs1.MoveNext
    End If
Loop
If Text2.Text = "" Then GoTo ar_tim1:
If C <> 1 Then GoTo ar_tim2:
If Text5.Text = "" Then GoTo ekdo_1:
If IsDate(Text5.Text) = False Then GoTo ekdo_2:
If Text4.Text = "" Then GoTo typos_1:
If (Text4.Text = "пистытийо" Or Text4.Text = "тилок\пыкгсгс") Then
    GoTo LAZA:
Else
    GoTo typos_2:
End If
LAZA:
If (Combo27.Text = "цемийо" Or Combo27.Text = "COOP") Then
    GoTo SYNEXEIA:
Else
    GoTo TYPOS_PIST_1:
End If

SYNEXEIA:
If Text3.Text = "" Then GoTo POSO_1:
If IsNumeric(Text3.Text) = False Then GoTo POSO_2:
If Text6.Text = "" Then GoTo eksoflisi_1:
If Text6.Text = "маи" Or Text6.Text = "ови" Then
GoTo nik_1:
Else
MsgBox ("дем пкгйтяокоцгсате сыста то педио еныжкгсг.пкгйтяокоцисте маи г ови"), vbCritical, "пяосовг"
GoTo TELOS:
End If

' PROGRAMATISMOS *********************************
nik_1:

' PERIPTOSH NAI*********************************
If Text6.Text = "маи" Then
 ' PROSTHETI ELEGXOI
 If Text7.Text = "" Then GoTo ime_ekso_1:
 If IsDate(Text7.Text) = False Then GoTo ime_ekso_2:
 If Text12.Text = "" Then GoTo EPIT_1:
 ' EGRAFH
 If Text12.Text = "летягта" Then
    XREOSH = Text3.Text
    PISTOSI = Text3.Text
 Else
    XREOSH = Text3.Text
    PISTOSI = 0
 End If
 YPOLIPO = 0
 STATEMENT = "INSERT INTO " & UCase(ZETAIRIES.Text1.Text) & " (" & _
    "аяихлос_тилокоциоу,тупос,глеяолгмиа_ейдосгс," & _
    "еныжкгсг,посо," & _
    "глеяолгмиа_еныжкгсгс,аяихлос_епитацгс," & _
    "вяеысг,пистысг,упокоипо)" & _
    "VALUES (" & _
        "'" & UCase(ZETAIRIES.Text2.Text) & "'," & _
        "'" & UCase(ZETAIRIES.Text4.Text) & "', " & _
        "'" & ZETAIRIES.Text5.Text & "'," & _
        "'1'," & _
        "'" & ZETAIRIES.Text3.Text & "'," & _
        "'" & ZETAIRIES.Text7.Text & "'," & _
        "'" & ZETAIRIES.Text12.Text & "'," & _
        "'" & XREOSH & "'," & _
        "'" & PISTOSI & "'," & _
        "'" & YPOLIPO & "'" & _
         ")"
 db1.Execute STATEMENT
 If ZETAIRIES.Text4.Text = "пистытийо" Then
    STP = " UPDATE " & UCase(ZETAIRIES.Text1.Text) & _
        " SET тупос_пистытийоу='" & UCase(Trim(Combo27.Text)) & _
        "' WHERE аяихлос_тилокоциоу='" & UCase(Text2.Text) & "'"
    db1.Execute STP
 End If
Else
' PERIPTOSH OXI ********************************
 If Text4.Text = "пистытийо" Then
    XREOSH = 0
    PISTOSI = Text3.Text
 Else
    XREOSH = Text3.Text
    PISTOSI = 0
 End If
 YPOLIPO = 0
 half_statement = "INSERT INTO " & UCase(ZETAIRIES.Text1.Text) & " (" & _
    "аяихлос_тилокоциоу,тупос,глеяолгмиа_ейдосгс," & _
    "еныжкгсг,посо," & _
    "аяихлос_епитацгс,вяеысг,пистысг,упокоипо)" & _
    "VALUES (" & _
        "'" & UCase(ZETAIRIES.Text2.Text) & "'," & _
        "'" & UCase(ZETAIRIES.Text4.Text) & "', " & _
        "'" & ZETAIRIES.Text5.Text & "'," & _
        "'0'," & _
        "'" & ZETAIRIES.Text3.Text & "'," & _
        "' '," & _
        "'" & XREOSH & "'," & _
        "'" & PISTOSI & "'," & _
        "'" & YPOLIPO & "'" & _
        ")"
 db1.Execute half_statement
 If ZETAIRIES4.Text4.Text = "пистытийо" Then
    STP = " UPDATE " & UCase(ZETAIRIES.Text1.Text) & _
        " SET тупос_пистытийоу='" & UCase(Trim(Combo27.Text)) & _
        "' WHERE аяихлос_тилокоциоу='" & UCase(Text2.Text) & "'"
    db1.Execute STP
 End If
End If
' KATHARISMOS PEDION*******************************
MsgBox ("г еццяажг окойкгяыхгйе"), , "OK"
ZETAIRIES.Text2.Text = ""
ZETAIRIES.Text4.Text = ""
ZETAIRIES.Text5.Text = ""
ZETAIRIES.Text3.Text = ""
ZETAIRIES.Text6.Text = ""
ZETAIRIES.Text7.Text = ""
ZETAIRIES.Text12.Text = ""
Combo5.Text = "глеяа"
Combo6.Text = "лгмас"
Combo7.Text = "етос"
Combo8.Text = "глеяа"
Combo9.Text = "лгмас"
Combo10.Text = "етос"
Combo27.Text = "цемийо"

GoTo TELOS:
' ANTIMETOPISI LATHON****************************
'****************************************************************
ar_tim1:
MsgBox ("дем дысате аяихло тилокоциоу"), vbCritical, "пяосовг !!!"
index = 32

ar_tim2:
If index = 32 Then
    GoTo TELOS:
Else
    MsgBox ("о аяихлос тилокоциоу поу дысате упаявеи гдг.паяайакы диояхысте"), vbCritical, "пяосовг !!!"
    index = 32
End If

POSO_1:
If index = 32 Then
    GoTo TELOS
Else
MsgBox ("дем дысате посо"), vbCritical, "пяосовг !!!"
index = 32
End If

POSO_2:
If index = 32 Then
    GoTo TELOS
Else
MsgBox ("дем дысате сыста то посо"), vbCritical, "пяосовг !!!"
index = 32
End If

typos_1:
If index = 32 Then
    GoTo TELOS
Else
MsgBox ("дем дысате тупо тилокоциоу"), vbCritical, "пяосовг !!!"
index = 32
End If

typos_2:
If index = 32 Then
    GoTo TELOS
Else
    MsgBox ("сто педио тупос ха пяепеи ма дысете лиа апо тис кенеис <<пистытийо>> ╧ <<тилок\пыкгсгс>>."), vbCritical, "пяосовг !!!"
    index = 32
End If

TYPOS_PIST_1:
If index = 32 Then
    GoTo TELOS
Else
    MsgBox ("ха пяепеи сам тупо пистытийоу ма дысете лиа апо тис кенеис <<цемийо>> ╧ <<COOP>>"), vbCritical, "пяосовг !!!"
    index = 32
End If

ekdo_1:
If index = 32 Then
    GoTo TELOS
Else
MsgBox ("дем дысате глеяолгмиа ейдосгс"), vbCritical, "пяосовг !!!"
index = 32
End If

ekdo_2:
If index = 32 Then
    GoTo TELOS
Else
MsgBox ("дем дысате сыста тгм глеяолгмиа ейдосгс"), vbCritical, "пяосовг !!!"
index = 32
End If

eksoflisi_1:
If index = 32 Then
    GoTo TELOS
Else
MsgBox ("дем ояисате ам еныжкгхеи.пкгйтяокоцгсте маи ╧ ови"), vbCritical, "пяосовг !!!"
index = 32
End If

ime_ekso_1:
If index = 32 Then
    GoTo TELOS
Else
MsgBox ("дем дысате глеяолгмиа еныжкгсгс"), vbCritical, "пяосовг !!!"
index = 32
End If

ime_ekso_2:
If index = 32 Then
    GoTo TELOS
Else
MsgBox ("дем дысате сыста тгм глеяолгмиа еныжкгсгс"), vbCritical, "пяосовг !!!"
index = 32
End If

EPIT_1:
If index = 32 Then
    GoTo TELOS
Else
MsgBox ("дем дысате аяихло епитацгс г акка стоивеиа пкгяылгс"), vbCritical, "пяосовг !!!"
index = 32
End If

er:
If index = 32 Then
    GoTo TELOS
Else
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"
index = 32
End If

TELOS:
If rs1.STATE = 1 Then rs1.Close
If db1.STATE = 1 Then db1.Close
End Sub

Private Sub Command10_Click()
On Error GoTo er:

Dim HM_MEXRI_KARTELA As String
Dim DATE_HM_MEXRI_KARTELA As Date
'***************** ELEGXOI **************************************
If IsNumeric(Combo24.Text) = False Then
    MsgBox ("дем дысате сыста глеяа"), vbCritical, "пяосовг !!!"
    GoTo TELOS:
End If
If IsNumeric(Combo25.Text) = False Then
    MsgBox ("дем дысате сыста лгма"), vbCritical, "пяосовг !!!"
    GoTo TELOS:
End If
If IsNumeric(Combo26.Text) = False Then
    MsgBox ("дем дысате сыста етос"), vbCritical, "пяосовг !!!"
    GoTo TELOS:
End If
If CInt(Combo24.Text) < 1 Or CInt(Combo24.Text) > 31 Then
    MsgBox ("о аяихлос поу дысате дем еимаи ецйуяг глеяа лгма"), vbCritical, "пяосовг !!!"
    GoTo TELOS:
End If
If CInt(Combo25.Text) < 1 Or CInt(Combo25.Text) > 12 Then
    MsgBox ("о аяихлос поу дысате дем еимаи ецйуяос лгмас етоус"), vbCritical, "пяосовг !!!"
    GoTo TELOS:
End If
If CInt(Combo26.Text) < 2005 Or CInt(Combo26.Text) > 2020 Then
    MsgBox ("то пяоцяалла упостгяифеи глеяолгмиес апо 2005 еыс 2020.паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"
    GoTo TELOS:
End If
'*************** LEITOYRGIA *******************************
HM_MEXRI_KARTELA = DAY_HM_MEXRI_KARTELA & "/" & MONTH_HM_MEXRI_KARTELA & _
"/" & ETOS_HM_MEXRI_KARTELA

If IsDate(HM_MEXRI_KARTELA) = True Then
DATE_HM_MEXRI_KARTELA = CDate(HM_MEXRI_KARTELA)
Text18.Text = DATE_HM_MEXRI_KARTELA
Else
MsgBox ("дем дысате глеяолгмиа"), vbCritical, "пяосовг!!"
End If
GoTo TELOS:
'***********************************************************
er:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
End Sub

Private Sub Command11_Click()
On Error GoTo TELOS:

Dim HM_EKDOSHS As String
Dim DATE_HM_EKDOSHS As Date
'***************** ELEGXOI **************************************
If IsNumeric(Combo5.Text) = False Then
    MsgBox ("дем дысате сыста глеяа"), vbCritical, "пяосовг !!!"
    GoTo TELOS:
End If
If IsNumeric(Combo6.Text) = False Then
    MsgBox ("дем дысате сыста лгма"), vbCritical, "пяосовг !!!"
    GoTo TELOS:
End If
If IsNumeric(Combo7.Text) = False Then
    MsgBox ("дем дысате сыста етос"), vbCritical, "пяосовг !!!"
    GoTo TELOS:
End If
If CInt(Combo5.Text) < 1 Or CInt(Combo5.Text) > 31 Then
    MsgBox ("о аяихлос поу дысате дем еимаи ецйуяг глеяа лгма"), vbCritical, "пяосовг !!!"
    GoTo TELOS:
End If
If CInt(Combo6.Text) < 1 Or CInt(Combo6.Text) > 12 Then
    MsgBox ("о аяихлос поу дысате дем еимаи ецйуяос лгмас етоус"), vbCritical, "пяосовг !!!"
    GoTo TELOS:
End If
If CInt(Combo7.Text) < 2005 Or CInt(Combo7.Text) > 2020 Then
    MsgBox ("то пяоцяалла упостгяифеи глеяолгмиес апо 2005 еыс 2020.паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"
    GoTo TELOS:
End If
'*************** LEITOYRGIA *******************************
HM_EKDOSHS = DAY_HM_EKDOSHS & "/" & MONTH_HM_EKDOSHS & _
"/" & ETOS_HM_EKDOSHS

If IsDate(HM_EKDOSHS) = True Then
    DATE_HM_EKDOSHS = CDate(HM_EKDOSHS)
    Text5.Text = DATE_HM_EKDOSHS
Else
    MsgBox ("дем дысате сыста глеяолгмиа"), vbCritical, "пяосовг!!"
End If
GoTo TELOS:
'***********************************************************

er:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
End Sub

Private Sub Command12_Click()
Text5.Text = Date
End Sub

Private Sub Command13_Click()
On Error GoTo TELOS:

Dim HM_EKSOFLISIS As String
Dim DATE_HM_EKSOFLISIS As Date
'***************** ELEGXOI **************************************
If IsNumeric(Combo8.Text) = False Then
    MsgBox ("дем дысате сыста глеяа"), vbCritical, "пяосовг !!!"
    GoTo TELOS:
End If
If IsNumeric(Combo9.Text) = False Then
    MsgBox ("дем дысате сыста лгма"), vbCritical, "пяосовг !!!"
    GoTo TELOS:
End If
If IsNumeric(Combo10.Text) = False Then
    MsgBox ("дем дысате сыста етос"), vbCritical, "пяосовг !!!"
    GoTo TELOS:
End If
If CInt(Combo8.Text) < 1 Or CInt(Combo8.Text) > 31 Then
    MsgBox ("о аяихлос поу дысате дем еимаи ецйуяг глеяа лгма"), vbCritical, "пяосовг !!!"
    GoTo TELOS:
End If
If CInt(Combo9.Text) < 1 Or CInt(Combo9.Text) > 12 Then
    MsgBox ("о аяихлос поу дысате дем еимаи ецйуяос лгмас етоус"), vbCritical, "пяосовг !!!"
    GoTo TELOS:
End If
If CInt(Combo10.Text) < 2005 Or CInt(Combo10.Text) > 2020 Then
    MsgBox ("то пяоцяалла упостгяифеи глеяолгмиес апо 2005 еыс 2020.паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"
    GoTo TELOS:
End If
'*************** LEITOYRGIA *******************************
HM_EKSOFLISIS = DAY_HM_EKSOFLISIS & "/" & MONTH_HM_EKSOFLISIS & _
"/" & ETOS_HM_EKSOFLISIS

If IsDate(HM_EKSOFLISIS) Then
    DATE_HM_EKSOFLISIS = CDate(HM_EKSOFLISIS)
    Text7.Text = DATE_HM_EKSOFLISIS
Else
    MsgBox ("дем дысате сыста глеяолгмиа"), vbCritical, "пяосовг!!"
End If
GoTo TELOS:
'***********************************************************

er:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
End Sub

Private Sub Command14_Click()
Text7.Text = Date
End Sub

Private Sub Command15_Click()
Dim HM_EKSOFLISIS_2 As String
Dim DATE_HM_EKSOFLISIS_2 As Date

HM_EKSOFLISIS_2 = DAY_HM_EKSOFLISIS_2 & "/" & MONTH_HM_EKSOFLISIS_2 & _
"/" & ETOS_HM_EKSOFLISIS_2

If IsDate(HM_EKSOFLISIS_2) = True Then
DATE_HM_EKSOFLISIS_2 = CDate(HM_EKSOFLISIS_2)
Text14.Text = DATE_HM_EKSOFLISIS_2
Else
MsgBox ("дем дысате глеяолгмиа"), vbCritical, "пяосовг!!"
End If
End Sub

Private Sub Command16_Click()
Text14.Text = Date

End Sub

Private Sub Command17_Click()
On Error GoTo er:

Dim HM_ANAZITISIS_2 As String
Dim DATE_HM_ANAZITISIS_2 As Date
'***************** ELEGXOI **************************************
If IsNumeric(Combo17.Text) = False Then
    MsgBox ("дем дысате сыста глеяа"), vbCritical, "пяосовг !!!"
    GoTo TELOS:
End If
If IsNumeric(Combo18.Text) = False Then
    MsgBox ("дем дысате сыста лгма"), vbCritical, "пяосовг !!!"
    GoTo TELOS:
End If
If IsNumeric(Combo19.Text) = False Then
    MsgBox ("дем дысате сыста етос"), vbCritical, "пяосовг !!!"
    GoTo TELOS:
End If
If CInt(Combo17.Text) < 1 Or CInt(Combo17.Text) > 31 Then
    MsgBox ("о аяихлос поу дысате дем еимаи ецйуяг глеяа лгма"), vbCritical, "пяосовг !!!"
    GoTo TELOS:
End If
If CInt(Combo18.Text) < 1 Or CInt(Combo18.Text) > 12 Then
    MsgBox ("о аяихлос поу дысате дем еимаи ецйуяос лгмас етоус"), vbCritical, "пяосовг !!!"
    GoTo TELOS:
End If
If CInt(Combo19.Text) < 2005 Or CInt(Combo19.Text) > 2020 Then
    MsgBox ("то пяоцяалла упостгяифеи глеяолгмиес апо 2005 еыс 2020.паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"
    GoTo TELOS:
End If
'*************** LEITOYRGIA *******************************
HM_ANAZITISIS_2 = DAY_ANAZHTHSH_2 & _
"/" & MONTH_ANAZHTHSH_2 & _
"/" & ETOS_ANAZHTHSH_2

If IsDate(HM_ANAZITISIS_2) Then
DATE_HM_ANAZITISIS_2 = CDate(HM_ANAZITISIS_2)
Text10.Text = DATE_HM_ANAZITISIS_2
Else
MsgBox ("дем дысате глеяолгмиа"), vbCritical, "пяосовг!!"
End If
GoTo TELOS:
'***********************************************************
er:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
End Sub

Private Sub Command18_Click()
On Error GoTo er:

Text9.Text = Date
GoTo TELOS:

er:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
End Sub

Private Sub Command19_Click()
On Error GoTo er:

Dim HM_ANAZITISIS_1 As String
Dim DATE_HM_ANAZITISIS_1 As Date
'***************** ELEGXOI **************************************
If IsNumeric(Combo14.Text) = False Then
    MsgBox ("дем дысате сыста глеяа"), vbCritical, "пяосовг !!!"
    GoTo TELOS:
End If
If IsNumeric(Combo15.Text) = False Then
    MsgBox ("дем дысате сыста лгма"), vbCritical, "пяосовг !!!"
    GoTo TELOS:
End If
If IsNumeric(Combo16.Text) = False Then
    MsgBox ("дем дысате сыста етос"), vbCritical, "пяосовг !!!"
    GoTo TELOS:
End If
If CInt(Combo14.Text) < 1 Or CInt(Combo14.Text) > 31 Then
    MsgBox ("о аяихлос поу дысате дем еимаи ецйуяг глеяа лгма"), vbCritical, "пяосовг !!!"
    GoTo TELOS:
End If
If CInt(Combo15.Text) < 1 Or CInt(Combo15.Text) > 12 Then
    MsgBox ("о аяихлос поу дысате дем еимаи ецйуяос лгмас етоус"), vbCritical, "пяосовг !!!"
    GoTo TELOS:
End If
If CInt(Combo16.Text) < 2005 Or CInt(Combo16.Text) > 2020 Then
    MsgBox ("то пяоцяалла упостгяифеи глеяолгмиес апо 2005 еыс 2020.паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"
    GoTo TELOS:
End If
'*************** LEITOYRGIA *******************************
HM_ANAZITISIS_1 = DAY_ANAZHTHSH_1 & _
"/" & MONTH_ANAZHTHSH_1 & _
"/" & ETOS_ANAZHTHSH_1

If IsDate(HM_ANAZITISIS_1) = True Then
DATE_HM_ANAZITISIS_1 = CDate(HM_ANAZITISIS_1)
Text9.Text = DATE_HM_ANAZITISIS_1
Else
MsgBox ("дем дысате глеяолгмиа"), vbCritical, "пяосовг!!"
End If
GoTo TELOS:
'***********************************************************
er:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
End Sub

Private Sub Command2_Click()
Text1.Text = Text24.Text
On Error GoTo er:
If rs1.STATE = 1 Then rs1.Close
If db1.STATE = 1 Then db1.Close
db1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & ZETAIRIES_DIADROMHS_BACKUP_DIAX_1 & ";" & _
"Persist Security Info=False"
db1.Open ZETAIRIES_DIADROMHS_BACKUP_DIAX
rs1.Open "[" & ZETAIRIES.Text1.Text & "]", db1, adOpenDynamic, adLockBatchOptimistic
Dim C As Integer
Dim STATEMENT As String
C = 1
If Text2.Text = "" Then GoTo KENO:
'**************** EYRESH AN YPARXEI TO SYGKEKRIMENO TIMOLOGIO *********************
If rs1.BOF = rs1.EOF Then GoTo NIK:
rs1.MoveFirst
NIK:
Do While Not rs1.EOF
    If rs1![аяихлос_тилокоциоу] <> UCase(Form4.Text2.Text) Then
        rs1.MoveNext
    Else
        If rs1![тупос] = "пистытийо" Or rs1![тупос] = "тилок\пыкгсгс" Then
            C = C + 1
        End If
        rs1.MoveNext
    End If
Loop
If C = 1 Then
    MsgBox ("дем бяехгйе тилокоцио ле том аяихло поу дысате"), vbCritical, "пяосовг !!!"
Else
    If MsgBox("еисте бебаиои оти хекете ма пяовыягсете стгм диацяажг тоу тилокоциоу;", vbOKCancel, "") = vbOK Then
        STATEMENT = " delete from " & UCase(Text1.Text) & _
        " where аяихлос_тилокоциоу= '" & UCase(Text2.Text) & "'"
        db1.Execute STATEMENT
        MsgBox ("то тилокоцио ле том аяихло поу дысате диацяажгйе"), , "OK"
    End If
    ZETAIRIES.Text2.Text = ""
    ZETAIRIES.Text4.Text = ""
    ZETAIRIES.Text5.Text = ""
    ZETAIRIES.Text3.Text = ""
    ZETAIRIES.Text6.Text = ""
    ZETAIRIES.Text7.Text = ""
    ZETAIRIES.Text12.Text = ""
    Combo5.Text = "глеяа"
    Combo6.Text = "лгмас"
    Combo7.Text = "етос"
    Combo8.Text = "глеяа"
    Combo9.Text = "лгмас"
    Combo10.Text = "етос"
End If
GoTo TELOS:

KENO:
MsgBox ("дем дысате йамема аяихло тилокоциоу"), vbCritical, "пяосовг !!!"
GoTo TELOS:

er:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
If rs1.STATE = 1 Then rs1.Close
If db1.STATE = 1 Then db1.Close
End Sub

Private Sub Command20_Click()
On Error GoTo er:

Text10.Text = Date
GoTo TELOS:

er:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
End Sub

Private Sub Command21_Click()
On Error GoTo er:

Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
Text11.Text = ""
Text16.Text = ""
Combo20.Text = ""
Combo28.Text = ""
Text26.Text = "ока"
GoTo TELOS:

er:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
End Sub

Private Sub Command22_Click()
On Error GoTo er:

Text18.Text = Date
GoTo TELOS:

er:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
End Sub

Private Sub Command23_Click()
On Error GoTo er:
Text1.Text = Text24.Text

' ELEGXOS AN PERNAO TO PROBLEPOMENO MHKOS PEDIOY GIA TEXT2,12 POY EINAI 20,30 ANTOISTIXA
' KAI APOKOPEI PERITOY MEROYS AN XREIAZETAI
Dim L20 As Integer
L20 = Len(Text20.Text)
If L20 > 20 Then
    Text20.Text = Mid(Text20.Text, 1, 20)
Else
    
End If


Dim index As Integer
Dim STATEMENT As String
Dim YPO As Double
Dim C As Integer
C = 1
index = 1
YPO = 0
If rs1.STATE = 1 Then rs1.Close
If db1.STATE = 1 Then db1.Close
db1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & ZETAIRIES_DIADROMHS_BACKUP_DIAX_1 & ";" & _
"Persist Security Info=False"
db1.Open ZETAIRIES_DIADROMHS_BACKUP_DIAX
rs1.Open "[" & ZETAIRIES.Text1.Text & "]", db1, adOpenDynamic, adLockBatchOptimistic
'EYRESH TOY AN YPARXEI HDH H EGRAFH
If rs1.BOF = rs1.EOF Then GoTo NIK:
rs1.MoveFirst
NIK:
Do While Not rs1.EOF
    If Text20.Text = rs1![аяихлос_тилокоциоу] Then
        C = C + 1
        rs1.MoveNext
    Else
        rs1.MoveNext
    End If
Loop
If C <> 1 Then
    MsgBox ("о аяихлос епитацгс поу дысате евеи намапеяастеи. паяайакы екецнте"), vbCritical, "пяосовг !!!"
Else
    ' ELEGXOS LATHON
    If Text20.Text = "" Then GoTo EPIT_1:
    If Text19.Text = "" Then GoTo HMER_1:
    If IsDate(Text19.Text) = False Then GoTo HMER_2:
    If Text21.Text = "" Then GoTo POSO_1:
    If IsNumeric(Text21.Text) = False Then GoTo POSO_2:

    ' PROGRAMATISMOS
    STATEMENT = "INSERT INTO " & ZETAIRIES.Text1.Text & _
    "(аяихлос_тилокоциоу,тупос,глеяолгмиа_ейдосгс," & _
    "еныжкгсг,посо,глеяолгмиа_еныжкгсгс,аяихлос_епитацгс," & _
    "вяеысг,пистысг,упокоипо)" & _
    " VALUES (" & _
    "'" & UCase(Text20.Text) & "'," & _
    "'епитацг'," & _
    "'" & UCase(ZETAIRIES.Text19.Text) & "'," & _
    "'0'," & _
    "'" & ZETAIRIES.Text21.Text & "'," & _
    "'" & ZETAIRIES.Text19.Text & "'," & _
    "'" & UCase(Text20.Text) & "'," & _
    "0," & _
    "'" & ZETAIRIES.Text21.Text & "'," & _
    "'" & YPO & "'" & _
    ")"
    db1.Execute STATEMENT
    MsgBox ("г епитацг йатавыягхгйе"), , "ой"
    Text19.Text = ""
    Text20.Text = ""
    Text21.Text = ""
End If
GoTo TELOS:
' ANTIMETOPISH LATHON
HMER_1:
MsgBox ("дем дысате глеяолгмиа"), vbCritical, "пяосовг !!!"
index = 32
GoTo TELOS:

HMER_2:
If index = 32 Then
GoTo TELOS:
Else
MsgBox ("дем дысате сыста тгм глеяылгмиа"), vbCritical, "пяосовг  !!!"
index = 32
End If

EPIT_1:
If index = 32 Then
GoTo TELOS:
Else
MsgBox ("дем дысате аяихло епитацгс"), vbCritical, "пяосовг  !!!"
index = 32
End If

POSO_1:
If index = 32 Then
GoTo TELOS:
Else
MsgBox ("дем дысате посо"), vbCritical, "пяосовг !!!"
index = 32
End If

POSO_2:
If index = 32 Then
GoTo TELOS:
Else
MsgBox ("дем дысате сыста то посо"), vbCritical, "пяосовг !!!"
index = 32
End If

er:
If index = 32 Then
GoTo TELOS:
Else
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"
End If

TELOS:
If rs1.STATE = 1 Then rs1.Close
If db1.STATE = 1 Then db1.Close
End Sub

Private Sub Command24_Click()
On Error GoTo er:
Text1.Text = Text24.Text
' ***************************** ORISMOI *****************************************
Dim STATEMENT, STP As String
Dim C As Integer
C = 1
If rs1.STATE = 1 Then rs1.Close
If db1.STATE = 1 Then db1.Close
db1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & ZETAIRIES_DIADROMHS_BACKUP_DIAX_1 & ";" & _
"Persist Security Info=False"
db1.Open ZETAIRIES_DIADROMHS_BACKUP_DIAX
rs1.Open "[" & ZETAIRIES.Text1.Text & "]", db1, adOpenDynamic, adLockBatchOptimistic

If Text2.Text = "" Then
    MsgBox ("дем дысате йамема аяихло тилокоциоу"), vbCritical, "пяосовг !!!"
    GoTo TELOS:
End If

If rs1.BOF = rs1.EOF Then GoTo NIK:
rs1.MoveFirst
NIK:

'***************************** KOYMPI E Y R E S H *****************************************************
'********************************************************************
If Command24.Caption = "еуяесг" Then
 '  PSAKSIMO EGRAFHS
    Do While Not rs1.EOF
        If rs1![аяихлос_тилокоциоу] <> UCase(Text2.Text) Then
            rs1.MoveNext
        Else
            If rs1![тупос] = "пистытийо" Or rs1![тупос] = "тилок\пыкгсгс" Then
                Text2.Text = rs1![аяихлос_тилокоциоу]
                Text5.Text = rs1![глеяолгмиа_ейдосгс]
                Text4.Text = rs1![тупос]
                Text3.Text = rs1![посо]
                Text6.Text = rs1![еныжкгсг]
                Text23.Text = rs1![аяихлос_тилокоциоу]
                If rs1![еныжкгсг] <> False Then
                    Text7.Text = rs1![глеяолгмиа_еныжкгсгс]
                    Text12.Text = rs1![аяихлос_епитацгс]
                Else
                    Text7.Text = ""
                    Text12.Text = ""
                End If
                C = C + 1
            End If
            rs1.MoveNext
        End If
    Loop
  ' ANTIMETOPISI ANALOGA ME TO AN BRHKE H OXI EGRAFH
    If C <> 1 Then
       If Text6.Text = False Then
            Text6.Text = "ови"
       Else
            Text6.Text = "маи"
       End If
       Command24.Caption = "диояхысг"
       Command1.Enabled = False
       Command2.Enabled = False
       Command29.Enabled = False
    Else
       MsgBox ("дем бяехгйе тилокоцио ле том аяихло поу дысате"), vbCritical, "пяосовг!!"
    End If
Else
'********************* KOYMPI D I O R T H O S I **********************************
'*********************************************************************
' ELEGXOS AN PERNAO TO PROBLEPOMENO MHKOS PEDIOY GIA TEXT2,12 POY EINAI 20,30 ANTOISTIXA
' KAI APOKOPEI PERITOY MEROYS AN XREIAZETAI

    
    
    Dim L2, L12 As Integer
    L2 = Len(Text2.Text)
    L12 = Len(Text12.Text)

    If L2 > 20 Then
        Text2.Text = Mid(Text2.Text, 1, 20)
    Else
    
    End If

    If L12 > 30 Then
        Text12.Text = Mid(Text12.Text, 1, 30)
    Else

    End If

  Dim C1 As Integer
  C1 = 0
  Dim XREOSH, PISTOSI, YPOLIPO As Double
  Dim S1, S2, S3, S4, S5, S6, S7, S8, S9, S10 As String
  XREOSH = 0
  PISTOSI = 0
  YPOLIPO = 0
  Dim SS1, SS2, SS3, SS4, SS5, SS6, SS7, SS8, SS9, SS10 As String
  ' ELEGXOI***********************************
  Dim index As Integer
  index = 1
  If Text2.Text = "" Then GoTo ar_tim1:
  If Text5.Text = "" Then GoTo ekdo_1:
  If IsDate(Text5.Text) = False Then GoTo ekdo_2:
  If Text4.Text = "" Then GoTo typos_1:
  If (Text4.Text = "пистытийо" Or Text4.Text = "тилок\пыкгсгс") Then
     GoTo LAZA:
  Else
     GoTo typos_2:
  End If
LAZA:
If (Combo27.Text = "цемийо" Or Combo27.Text = "COOP") Then
    GoTo SYNEXEIA:
Else
    GoTo TYPOS_PIST_1:
End If

SYNEXEIA:
  If Text3.Text = "" Then GoTo POSO_1:
  If IsNumeric(Text3.Text) = False Then GoTo POSO_2:
  If Text6.Text = "" Then GoTo eksoflisi_1:
  If Text6.Text = "маи" Or Text6.Text = "ови" Then
     GoTo nik_1:
  Else
     MsgBox ("дем пкгйтяокоцгсате сыста то педио еныжкгсг.пкгйтяокоцисте маи г ови,выяис ма ажгмете йема"), vbCritical, "пяосовг"
  GoTo TELOS:
  End If
  
nik_1:
  If rs1.BOF = rs1.EOF Then GoTo NIK111:
    rs1.MoveFirst
NIK111:
    Do While Not rs1.EOF
        If rs1![аяихлос_тилокоциоу] = Text2.Text Then
            C1 = C1 + 1
            rs1.MoveNext
        Else
            rs1.MoveNext
        End If
    Loop
    If C1 <> 0 Then
        If Text23.Text = Text2.Text Then
        
        Else
            MsgBox ("то моулеяо епитацгс поу дысате упаявеи гдг"), vbCritical, "пяосовг !!!"
            GoTo TELOS:
        End If
    Else
        
    End If

' PROGRAMATISMOS ***

    '********** PERIPTOSH еныжкгсг=NAI*********************************
  If Form4.Text6.Text = "маи" Then
        ' PROSTHETI ELEGXOI
        If Text7.Text = "" Then GoTo ime_ekso_1:
        If IsDate(Text7.Text) = False Then GoTo ime_ekso_2:
        If Text12.Text = "" Then GoTo EPIT_1:
    ' EGRAFH
    If MsgBox("хекете ма сумевисете се диояхысг тгс еццяажгс", vbOKCancel, "пяосовг") = vbOK Then
        If Text12.Text = "летягта" Then
            XREOSH = Text3.Text
            PISTOSI = Text3.Text
        Else
            XREOSH = Text3.Text
            PISTOSI = 0
        End If
        YPOLIPO = 0
        S1 = " UPDATE " & UCase(ZETAIRIES.Text1.Text) & _
        " SET тупос=" & "'" & UCase(ZETAIRIES.Text4.Text) & "'" & _
        " WHERE аяихлос_тилокоциоу=" & "'" & UCase(ZETAIRIES.Text2.Text) & "'"
 
        S2 = " UPDATE " & UCase(ZETAIRIES.Text1.Text) & _
        " SET глеяолгмиа_ейдосгс=" & "'" & ZETAIRIES.Text5.Text & "'" & _
        " WHERE аяихлос_тилокоциоу=" & "'" & UCase(ZETAIRIES.Text2.Text) & "'"
 
        S3 = " UPDATE " & UCase(ZETAIRIES.Text1.Text) & _
        " SET еныжкгсг=" & "'1'" & _
        " WHERE аяихлос_тилокоциоу=" & "'" & UCase(ZETAIRIES.Text2.Text) & "'"
 
        S4 = " UPDATE " & UCase(ZETAIRIES.Text1.Text) & _
        " SET посо=" & "'" & ZETAIRIES.Text3.Text & "'" & _
        " WHERE аяихлос_тилокоциоу=" & "'" & UCase(ZETAIRIES.Text2.Text) & "'"
 
        S5 = " UPDATE " & UCase(ZETAIRIES.Text1.Text) & _
        " SET глеяолгмиа_еныжкгсгс=" & "'" & ZETAIRIES.Text7.Text & "'" & _
        " WHERE аяихлос_тилокоциоу=" & "'" & UCase(ZETAIRIES.Text2.Text) & "'"

        S6 = " UPDATE " & UCase(ZETAIRIES.Text1.Text) & _
        " SET аяихлос_епитацгс=" & "'" & UCase(ZETAIRIES.Text12.Text) & "'" & _
        " WHERE аяихлос_тилокоциоу=" & "'" & UCase(ZETAIRIES.Text2.Text) & "'"

        S7 = " UPDATE " & UCase(ZETAIRIES.Text1.Text) & _
        " SET вяеысг=" & "'" & XREOSH & "'" & _
        " WHERE аяихлос_тилокоциоу=" & "'" & UCase(ZETAIRIES.Text2.Text) & "'"
    
        S8 = " UPDATE " & UCase(ZETAIRIES.Text1.Text) & _
        " SET пистысг=" & "'" & PISTOSI & "'" & _
        " WHERE аяихлос_тилокоциоу=" & "'" & UCase(ZETAIRIES.Text2.Text) & "'"

        S9 = " UPDATE " & UCase(ZETAIRIES.Text1.Text) & _
        " SET упокоипо=" & "'" & YPOLIPO & "'" & _
        " WHERE аяихлос_тилокоциоу=" & "'" & UCase(ZETAIRIES.Text2.Text) & "'"

        S10 = " UPDATE " & UCase(ZETAIRIES.Text1.Text) & _
        " SET аяихлос_тилокоциоу=" & "'" & UCase(ZETAIRIES.Text2.Text) & "'" & _
        " WHERE аяихлос_тилокоциоу=" & "'" & UCase(ZETAIRIES.Text23.Text) & "'"
        
        db1.Execute S1
        db1.Execute S2
        db1.Execute S3
        db1.Execute S4
        db1.Execute S5
        db1.Execute S6
        db1.Execute S7
        db1.Execute S8
        db1.Execute S9
        db1.Execute S10
        Command1.Enabled = True
        Command2.Enabled = True
        Command29.Enabled = True
        If Form4.Text4.Text = "пистытийо" Then
            STP = " UPDATE " & UCase(ZETAIRIES.Text1.Text) & _
            " SET тупос_пистытийоу='" & UCase(Trim(Combo27.Text)) & _
            "' WHERE аяихлос_тилокоциоу='" & UCase(Text2.Text) & "'"
        db1.Execute STP
        End If
        MsgBox ("г еццяажг диояхыхгйе"), , "OK"
        ZETAIRIES.Text2.Text = ""
        ZETAIRIES.Text4.Text = ""
        ZETAIRIES.Text5.Text = ""
        ZETAIRIES.Text3.Text = ""
        ZETAIRIES.Text6.Text = ""
        ZETAIRIES.Text7.Text = ""
        ZETAIRIES.Text12.Text = ""
        Combo5.Text = "глеяа"
        Combo6.Text = "лгмас"
        Combo7.Text = "етос"
        Combo8.Text = "глеяа"
        Combo9.Text = "лгмас"
        Combo10.Text = "етос"
        Combo27.Text = "цемийо"
        Command24.Caption = "еуяесг"
        Combo27.Text = "цемийо"
    
    Else
        ' KATHARISMOS PEDION*******************************
        Command1.Enabled = True
        Command2.Enabled = True
        Command29.Enabled = True
        ZETAIRIES.Text2.Text = ""
        ZETAIRIES.Text4.Text = ""
        ZETAIRIES.Text5.Text = ""
        ZETAIRIES.Text3.Text = ""
        ZETAIRIES.Text6.Text = ""
        ZETAIRIES.Text7.Text = ""
        ZETAIRIES.Text12.Text = ""
        Combo5.Text = "глеяа"
        Combo6.Text = "лгмас"
        Combo7.Text = "етос"
        Combo8.Text = "глеяа"
        Combo9.Text = "лгмас"
        Combo10.Text = "етос"
        Command24.Caption = "еуяесг"
        Combo27.Text = "цемийо"
    End If
  Else
    ' PERIPTOSH еныжкгсг=OXI ********************************
    If MsgBox("хекете ма сумевисете се диояхысг тгс еццяажгс", vbOKCancel, "пяосовг") = vbOK Then
        If Text4.Text = "пистытийо" Then
            XREOSH = 0
            PISTOSI = Text3.Text
        Else
            XREOSH = Text3.Text
            PISTOSI = 0
        End If
        YPOLIPO = 0
        SS1 = " UPDATE " & UCase(ZETAIRIES.Text1.Text) & _
        " SET тупос=" & "'" & UCase(ZETAIRIES.Text4.Text) & "'" & _
        " WHERE аяихлос_тилокоциоу=" & "'" & UCase(ZETAIRIES.Text2.Text) & "'"
 
        SS2 = " UPDATE " & UCase(ZETAIRIES.Text1.Text) & _
        " SET глеяолгмиа_ейдосгс=" & "'" & ZETAIRIES.Text5.Text & "'" & _
        " WHERE аяихлос_тилокоциоу=" & "'" & UCase(ZETAIRIES.Text2.Text) & "'"
 
        SS3 = " UPDATE " & UCase(ZETAIRIES.Text1.Text) & _
        " SET еныжкгсг=" & "'0'" & _
        " WHERE аяихлос_тилокоциоу=" & "'" & UCase(ZETAIRIES.Text2.Text) & "'"
 
        SS4 = " UPDATE " & UCase(ZETAIRIES.Text1.Text) & _
        " SET посо=" & "'" & ZETAIRIES.Text3.Text & "'" & _
        " WHERE аяихлос_тилокоциоу=" & "'" & UCase(ZETAIRIES.Text2.Text) & "'"

        SS10 = " UPDATE " & UCase(ZETAIRIES.Text1.Text) & _
        " SET глеяолгмиа_еныжкгсгс= NULL " & _
        " WHERE аяихлос_тилокоциоу=" & "'" & UCase(ZETAIRIES.Text2.Text) & "'"
  
        SS5 = " UPDATE " & UCase(ZETAIRIES.Text1.Text) & _
        " SET аяихлос_епитацгс=" & "' '" & _
        " WHERE аяихлос_тилокоциоу=" & "'" & UCase(ZETAIRIES.Text2.Text) & "'"
    
        SS6 = " UPDATE " & UCase(ZETAIRIES.Text1.Text) & _
        " SET вяеысг=" & "'" & XREOSH & "'" & _
        " WHERE аяихлос_тилокоциоу=" & "'" & UCase(ZETAIRIES.Text2.Text) & "'"
    
        SS7 = " UPDATE " & UCase(ZETAIRIES.Text1.Text) & _
        " SET пистысг=" & "'" & PISTOSI & "'" & _
        " WHERE аяихлос_тилокоциоу=" & "'" & UCase(ZETAIRIES.Text2.Text) & "'"
       
        SS8 = " UPDATE " & UCase(ZETAIRIES.Text1.Text) & _
        " SET упокоипо=" & "'" & YPOLIPO & "'" & _
        " WHERE аяихлос_тилокоциоу=" & "'" & UCase(ZETAIRIES.Text2.Text) & "'"
 
        SS9 = " UPDATE " & UCase(ZETAIRIES.Text1.Text) & _
        " SET аяихлос_тилокоциоу=" & "'" & UCase(ZETAIRIES.Text2.Text) & "'" & _
        " WHERE аяихлос_тилокоциоу=" & "'" & UCase(ZETAIRIES.Text23.Text) & "'"

        db1.Execute SS1
        db1.Execute SS2
        db1.Execute SS3
        db1.Execute SS4
        db1.Execute SS10
        db1.Execute SS5
        db1.Execute SS6
        db1.Execute SS7
        db1.Execute SS8
        db1.Execute SS9
        Command1.Enabled = True
        Command2.Enabled = True
        Command29.Enabled = True
        If Form4.Text4.Text = "пистытийо" Then
            STP = " UPDATE " & UCase(ZETAIRIES.Text1.Text) & _
            " SET тупос_пистытийоу='" & UCase(Trim(Combo27.Text)) & _
            "' WHERE аяихлос_тилокоциоу='" & UCase(Text2.Text) & "'"
        db1.Execute STP
        End If
        MsgBox ("г еццяажг диояхыхгйе"), , "OK"
    Else
        Command1.Enabled = True
        Command2.Enabled = True
        Command29.Enabled = True
        Command24.Caption = "еуяесг"
        ZETAIRIES.Text2.Text = ""
        ZETAIRIES.Text4.Text = ""
        ZETAIRIES.Text5.Text = ""
        ZETAIRIES.Text3.Text = ""
        ZETAIRIES.Text6.Text = ""
        ZETAIRIES.Text7.Text = ""
        ZETAIRIES.Text12.Text = ""
        Combo5.Text = "глеяа"
        Combo6.Text = "лгмас"
        Combo7.Text = "етос"
        Combo8.Text = "глеяа"
        Combo9.Text = "лгмас"
        Combo10.Text = "етос"
        Combo27.Text = "цемийо"
    End If
' KATHARISMOS PEDION*******************************
    ZETAIRIES.Text2.Text = ""
    ZETAIRIES.Text4.Text = ""
    ZETAIRIES.Text5.Text = ""
    ZETAIRIES.Text3.Text = ""
    ZETAIRIES.Text6.Text = ""
    ZETAIRIES.Text7.Text = ""
    ZETAIRIES.Text12.Text = ""
    Combo5.Text = "глеяа"
    Combo6.Text = "лгмас"
    Combo7.Text = "етос"
    Combo8.Text = "глеяа"
    Combo9.Text = "лгмас"
    Combo10.Text = "етос"
    Command24.Caption = "еуяесг"
    Combo27.Text = "цемийо"
  End If
End If
GoTo TELOS:

  
  ' ANTIMETOPISH LATHON
ar_tim1:
  MsgBox ("дем дысате аяихло тилокоциоу"), vbCritical, "пяосовг !!!"
  index = 32

ekdo_1:
  If index = 32 Then
    GoTo TELOS
  Else
    MsgBox ("дем дысате глеяолгмиа ейдосгс"), vbCritical, "пяосовг !!!"
    index = 32
  End If

ekdo_2:
  If index = 32 Then
    GoTo TELOS
  Else
    MsgBox ("дем дысате сыста тгм глеяолгмиа ейдосгс"), vbCritical, "пяосовг !!!"
    index = 32
  End If
  
typos_1:
  If index = 32 Then
    GoTo TELOS
  Else
    MsgBox ("дем дысате тупо тилокоциоу"), vbCritical, "пяосовг !!!"
    index = 32
  End If

typos_2:
  If index = 32 Then
    GoTo TELOS
  Else
    MsgBox ("сто педио тупос ха пяепеи ма дысете лиа апо тис кенеис <<пистытийо>> ╧ <<тилок\пыкгсгс>>."), vbCritical, "пяосовг !!!"
    index = 32
  End If
  
TYPOS_PIST_1:
If index = 32 Then
    GoTo TELOS
Else
    MsgBox ("ха пяепеи сам тупо пистытийоу ма дысете лиа апо тис кенеис <<цемийо>> ╧ <<COOP>>"), vbCritical, "пяосовг !!!"
    index = 32
End If
  
POSO_1:
  If index = 32 Then
    GoTo TELOS
  Else
    MsgBox ("дем дысате посо"), vbCritical, "пяосовг !!!"
    index = 32
  End If

POSO_2:
  If index = 32 Then
    GoTo TELOS
  Else
    MsgBox ("дем дысате сыста то посо"), vbCritical, "пяосовг !!!"
    index = 32
  End If


eksoflisi_1:
  If index = 32 Then
    GoTo TELOS
  Else
    MsgBox ("дем ояисате ам еныжкгхеи.пкгйтяокоцгсте маи г ови"), vbCritical, "пяосовг !!!"
    index = 32
  End If
  
ime_ekso_1:
  If index = 32 Then
    GoTo TELOS
  Else
    MsgBox ("дем дысате глеяолгмиа еныжкгсгс"), vbCritical, "пяосовг !!!"
    index = 32
  End If

ime_ekso_2:
  If index = 32 Then
    GoTo TELOS
  Else
    MsgBox ("дем дысате сыста тгм глеяолгмиа еныжкгсгс"), vbCritical, "пяосовг !!!"
    index = 32
  End If

EPIT_1:
  If index = 32 Then
    GoTo TELOS
  Else
    MsgBox ("дем дысате аяихло епитацгс г акка стоивеиа пкгяылгс"), vbCritical, "пяосовг !!!"
    index = 32
  End If
  
er:
If index = 32 Then
    GoTo TELOS
  Else
   MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"
    index = 32
  End If
  Command24.Caption = "еуяесг"
  
TELOS:
  If rs1.STATE = 1 Then rs1.Close
  If db1.STATE = 1 Then db1.Close
  End Sub

Private Sub Command25_Click()
Text1.Text = Text24.Text

On Error GoTo er:

If rs1.STATE = 1 Then rs1.Close
If db1.STATE = 1 Then db1.Close
db1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & ZETAIRIES_DIADROMHS_BACKUP_DIAX_1 & ";" & _
"Persist Security Info=False"
db1.Open ZETAIRIES_DIADROMHS_BACKUP_DIAX
rs1.Open "[" & ZETAIRIES.Text1.Text & "]", db1, adOpenDynamic, adLockBatchOptimistic

Dim ap, mexr, ap_temp, mexr_temp
Dim index As Integer
index = 1
'************** ELEGXOI LATHON ************************************************
If Text9.Text <> "" Then
    If IsDate(Text9.Text) = False Then GoTo imer_apo:
End If
If Text10.Text <> "" Then
    If IsDate(Text10.Text) = False Then GoTo imer_mexri:
End If
If (Text9.Text <> "") And (Text10.Text <> "") Then
    If ((CDate(Text9.Text)) > (CDate(Text10.Text))) Then GoTo imer:
End If
If Text10.Text <> "" Then
    If ((CDate(Text10.Text)) > Date) Then GoTo imer_2:
End If
If Text11.Text <> "" Then
    If IsNumeric(Text11.Text) = False Then GoTo POSO:
End If
If Text16.Text <> "" Then
    If ((Text16.Text = "маи") Or (Text16.Text = "ови")) Then

    Else
        GoTo EKSO:
    End If
End If
If (Text8.Text <> "") And ((Text9.Text <> "") Or (Text10.Text <> "") Or (Text11.Text <> "") Or (Text16.Text <> "")) Then GoTo lathos:
' LEITOYRGIA
' ANTOISTOIXHSH TIMHS STHN METABLHTH TYPOY DATE AP
If Text9.Text = "" Then
ap_temp = #1/1/2004#
Else
ap_temp = CDate(Text9.Text)
End If
' ANTOISTOIXHSH TIMHS STHN METABLHTH TYPOY DATE MEXR
If Text10.Text = "" Then
    mexr_temp = Date
Else
    mexr_temp = CDate(Text10.Text)
End If
' ftiaksimo telikon timon
If Day(ap_temp) < 12 Then
    ap = CDate(Month(ap_temp) & "/" & Day(ap_temp) & "/" & Year(ap_temp))
Else
    ap = ap_temp
End If
If Day(mexr_temp) < 12 Then
    mexr = CDate(Month(mexr_temp) & "/" & Day(mexr_temp) & "/" & Year(mexr_temp))
Else
    mexr = mexr_temp
End If

If Text26.Text = "ока" Or Text26.Text = "цемийо" Or Text26.Text = "COOP" Then
    
Else
    GoTo TYPOS_PISTOTIKOY:
End If
'**** ORISMOS EROTHMATON GIA KATHE PERIPTOSH TOY TYPOY PISTOTIKOY ***************************************
Select Case Text26.Text
Case "ока"
    '-----------------------------------1--------------------------------------------
    STATE1 = " select аяихлос_тилокоциоу,глеяолгмиа_ейдосгс,посо,тупос,тупос_пистытийоу from " & UCase(ZETAIRIES.Text1.Text) & _
    " WHERE тупос='пистытийо'" & _
    " order by глеяолгмиа_ейдосгс"
    SS1 = " select SUM(посо) from " & UCase(ZETAIRIES.Text1.Text) & _
    " WHERE тупос='пистытийо'"
    FF1 = " select COUNT(аяихлос_тилокоциоу) from " & UCase(ZETAIRIES.Text1.Text) & _
    " WHERE тупос='пистытийо'"
    '-----------------------------------2--------------------------------------------
    STATE2 = "select аяихлос_тилокоциоу,глеяолгмиа_ейдосгс,посо,тупос,тупос_пистытийоу from " & UCase(ZETAIRIES.Text1.Text) & _
    " where аяихлос_тилокоциоу ='" & UCase(ZETAIRIES.Text8.Text) & "'" & " AND тупос='пистытийо'" & _
    " order by глеяолгмиа_ейдосгс"
    SS2 = "select SUM(посо) from " & UCase(ZETAIRIES.Text1.Text) & _
    " where аяихлос_тилокоциоу ='" & UCase(ZETAIRIES.Text8.Text) & "'" & " AND тупос='пистытийо'"
    FF2 = "select COUNT(аяихлос_тилокоциоу) from " & UCase(ZETAIRIES.Text1.Text) & _
    " where аяихлос_тилокоциоу ='" & UCase(ZETAIRIES.Text8.Text) & "'" & " AND тупос='пистытийо'"
    '-----------------------------------3--------------------------------------------
    STATE3 = "select аяихлос_тилокоциоу,глеяолгмиа_ейдосгс,посо,тупос,тупос_пистытийоу from " & UCase(ZETAIRIES.Text1.Text) & _
    " where посо =" & ZETAIRIES.Text11.Text & " AND тупос='пистытийо'" & _
    " order by глеяолгмиа_ейдосгс"
    SS3 = "select SUM(посо) from " & UCase(ZETAIRIES.Text1.Text) & _
    " where посо =" & ZETAIRIES.Text11.Text & " AND тупос='пистытийо'"
    FF3 = "select COUNT(аяихлос_тилокоциоу) from " & UCase(ZETAIRIES.Text1.Text) & _
    " where посо =" & ZETAIRIES.Text11.Text & " AND тупос='пистытийо'"
    '-----------------------------------4--------------------------------------------
    STATE4 = "select аяихлос_тилокоциоу,глеяолгмиа_ейдосгс,посо,тупос,тупос_пистытийоу from " & UCase(ZETAIRIES.Text1.Text) & _
    " where (глеяолгмиа_ейдосгс between " & "#" & ap & "#" & " and " & "#" & mexr & "#" & ")" & " AND тупос='пистытийо'" & _
    " order by глеяолгмиа_ейдосгс"
    SS4 = "select SUM(посо) from " & UCase(ZETAIRIES.Text1.Text) & _
    " where (глеяолгмиа_ейдосгс between " & "#" & ap & "#" & " and " & "#" & mexr & "#" & ")" & " AND тупос='пистытийо'"
    FF4 = "select COUNT(аяихлос_тилокоциоу) from " & UCase(ZETAIRIES.Text1.Text) & _
    " where (глеяолгмиа_ейдосгс between " & "#" & ap & "#" & " and " & "#" & mexr & "#" & ")" & " AND тупос='пистытийо'"
    '-----------------------------------5A--------------------------------------------
    STATE5A = "select аяихлос_тилокоциоу,глеяолгмиа_ейдосгс,посо,тупос,тупос_пистытийоу from " & UCase(ZETAIRIES.Text1.Text) & _
    " where еныжкгсг=-1" & " AND тупос='пистытийо'" & _
    " order by глеяолгмиа_ейдосгс"
    SS5A = "select SUM(посо) from " & UCase(ZETAIRIES.Text1.Text) & _
    " where еныжкгсг=-1" & " AND тупос='пистытийо'"
    FF5A = "select COUNT(аяихлос_тилокоциоу) from " & UCase(ZETAIRIES.Text1.Text) & _
    " where еныжкгсг=-1" & " AND тупос='пистытийо'"
    '-----------------------------------5B--------------------------------------------
    STATE5B = "select аяихлос_тилокоциоу,глеяолгмиа_ейдосгс,посо,тупос,тупос_пистытийоу from " & UCase(ZETAIRIES.Text1.Text) & _
    " where еныжкгсг=0" & " AND тупос='пистытийо'" & _
    " order by глеяолгмиа_ейдосгс"
    SS5B = "select SUM(посо) from " & UCase(ZETAIRIES.Text1.Text) & _
    " where еныжкгсг=0" & " AND тупос='пистытийо'"
    FF5B = "select COUNT(аяихлос_тилокоциоу) from " & UCase(ZETAIRIES.Text1.Text) & _
    " where еныжкгсг=0" & " AND тупос='пистытийо'"
    '-----------------------------------6--------------------------------------------
    STATE6 = "select аяихлос_тилокоциоу,глеяолгмиа_ейдосгс,посо,тупос,тупос_пистытийоу from " & UCase(ZETAIRIES.Text1.Text) & _
    " where (глеяолгмиа_ейдосгс between " & "#" & ap & "#" & " and " & "#" & mexr & "#" & ") AND посо =" & ZETAIRIES.Text11.Text & " AND тупос='пистытийо'" & _
    " order by глеяолгмиа_ейдосгс"
    SS6 = "select SUM(посо) from " & UCase(ZETAIRIES.Text1.Text) & _
    " where (глеяолгмиа_ейдосгс between " & "#" & ap & "#" & " and " & "#" & mexr & "#" & ") AND посо =" & ZETAIRIES.Text11.Text & " AND тупос='пистытийо'"
    FF6 = "select COUNT(аяихлос_тилокоциоу) from " & UCase(ZETAIRIES.Text1.Text) & _
    " where (глеяолгмиа_ейдосгс between " & "#" & ap & "#" & " and " & "#" & mexr & "#" & ") AND посо =" & ZETAIRIES.Text11.Text & " AND тупос='пистытийо'"
    '-----------------------------------7A--------------------------------------------
    STATE7A = "select аяихлос_тилокоциоу,глеяолгмиа_ейдосгс,посо,тупос,тупос_пистытийоу from " & UCase(ZETAIRIES.Text1.Text) & _
    " where еныжкгсг=-1 AND посо=" & ZETAIRIES.Text11.Text & " AND тупос='пистытийо'" & _
    " order by глеяолгмиа_ейдосгс"
    SS7A = "select SUM(посо) from " & UCase(ZETAIRIES.Text1.Text) & _
    " where еныжкгсг=-1 AND посо=" & ZETAIRIES.Text11.Text & " AND тупос='пистытийо'"
    FF7A = "select COUNT(аяихлос_тилокоциоу) from " & UCase(ZETAIRIES.Text1.Text) & _
    " where еныжкгсг=-1 AND посо=" & ZETAIRIES.Text11.Text & " AND тупос='пистытийо'"
    '-----------------------------------7B--------------------------------------------
    STATE7B = "select аяихлос_тилокоциоу,глеяолгмиа_ейдосгс,посо,тупос,тупос_пистытийоу from " & UCase(ZETAIRIES.Text1.Text) & _
    " where еныжкгсг=0 AND посо=" & ZETAIRIES.Text11.Text & " AND тупос='пистытийо'" & _
    " order by глеяолгмиа_ейдосгс"
    SS7B = "select SUM(посо) from " & UCase(ZETAIRIES.Text1.Text) & _
    " where еныжкгсг=0 AND посо=" & ZETAIRIES.Text11.Text & " AND тупос='пистытийо'"
    FF7B = "select COUNT(аяихлос_тилокоциоу) from " & UCase(ZETAIRIES.Text1.Text) & _
    " where еныжкгсг=0 AND посо=" & ZETAIRIES.Text11.Text & " AND тупос='пистытийо'"
    '-----------------------------------8A--------------------------------------------
    STATE8A = "select аяихлос_тилокоциоу,глеяолгмиа_ейдосгс,посо,тупос,тупос_пистытийоу from " & UCase(ZETAIRIES.Text1.Text) & _
    " where (глеяолгмиа_ейдосгс between " & "#" & ap & "#" & " and " & "#" & mexr & "#" & ") AND еныжкгсг=-1 " & " AND тупос='пистытийо'" & _
    " order by глеяолгмиа_ейдосгс"
    SS8A = "select SUM(посо) from " & UCase(ZETAIRIES.Text1.Text) & _
    " where (глеяолгмиа_ейдосгс between " & "#" & ap & "#" & " and " & "#" & mexr & "#" & ") AND еныжкгсг=-1 " & " AND тупос='пистытийо'"
    FF8A = "select COUNT(аяихлос_тилокоциоу) from " & UCase(ZETAIRIES.Text1.Text) & _
    " where (глеяолгмиа_ейдосгс between " & "#" & ap & "#" & " and " & "#" & mexr & "#" & ") AND еныжкгсг=-1 " & " AND тупос='пистытийо'"
    '-----------------------------------8B--------------------------------------------
    STATE8B = "select аяихлос_тилокоциоу,глеяолгмиа_ейдосгс,посо,тупос,тупос_пистытийоу from " & UCase(ZETAIRIES.Text1.Text) & _
    " where (глеяолгмиа_ейдосгс between " & "#" & ap & "#" & " and " & "#" & mexr & "#" & ") AND еныжкгсг=0 " & " AND тупос='пистытийо'" & _
    " order by глеяолгмиа_ейдосгс"
    SS8B = "select SUM(посо) from " & UCase(ZETAIRIES.Text1.Text) & _
    " where (глеяолгмиа_ейдосгс between " & "#" & ap & "#" & " and " & "#" & mexr & "#" & ") AND еныжкгсг=0 " & " AND тупос='пистытийо'"
    FF8B = "select COUNT(аяихлос_тилокоциоу) from " & UCase(ZETAIRIES.Text1.Text) & _
    " where (глеяолгмиа_ейдосгс between " & "#" & ap & "#" & " and " & "#" & mexr & "#" & ") AND еныжкгсг=0 " & " AND тупос='пистытийо'"
    '-----------------------------------9A--------------------------------------------
    STATE9A = "select аяихлос_тилокоциоу,глеяолгмиа_ейдосгс,посо,тупос,тупос_пистытийоу from " & UCase(ZETAIRIES.Text1.Text) & _
    " where (глеяолгмиа_ейдосгс between " & "#" & ap & "#" & " and " & "#" & mexr & "#" & ") AND еныжкгсг=-1 AND посо =" & ZETAIRIES.Text11.Text & " AND тупос='пистытийо'" & _
    " order by глеяолгмиа_ейдосгс"
    SS9A = "select SUM(посо) from " & UCase(ZETAIRIES.Text1.Text) & _
    " where (глеяолгмиа_ейдосгс between " & "#" & ap & "#" & " and " & "#" & mexr & "#" & ") AND еныжкгсг=-1 AND посо =" & ZETAIRIES.Text11.Text & " AND тупос='пистытийо'"
    FF9A = "select COUNT(аяихлос_тилокоциоу) from " & UCase(ZETAIRIES.Text1.Text) & _
    " where (глеяолгмиа_ейдосгс between " & "#" & ap & "#" & " and " & "#" & mexr & "#" & ") AND еныжкгсг=-1 AND посо =" & ZETAIRIES.Text11.Text & " AND тупос='пистытийо'"
    '-----------------------------------9B--------------------------------------------
    STATE9B = "select аяихлос_тилокоциоу,глеяолгмиа_ейдосгс,посо,тупос,тупос_пистытийоу from " & UCase(ZETAIRIES.Text1.Text) & _
    " where (глеяолгмиа_ейдосгс between " & "#" & ap & "#" & " and " & "#" & mexr & "#" & ") AND еныжкгсг=0 AND посо =" & ZETAIRIES.Text11.Text & " AND тупос='пистытийо'" & _
    " order by глеяолгмиа_ейдосгс"
    SS9B = "select SUM(посо) from " & UCase(ZETAIRIES.Text1.Text) & _
    " where (глеяолгмиа_ейдосгс between " & "#" & ap & "#" & " and " & "#" & mexr & "#" & ") AND еныжкгсг=0 AND посо =" & ZETAIRIES.Text11.Text & " AND тупос='пистытийо'"
    FF9B = "select COUNT(аяихлос_тилокоциоу) from " & UCase(ZETAIRIES.Text1.Text) & _
    " where (глеяолгмиа_ейдосгс between " & "#" & ap & "#" & " and " & "#" & mexr & "#" & ") AND еныжкгсг=0 AND посо =" & ZETAIRIES.Text11.Text & " AND тупос='пистытийо'"
    ' DIAXORISMOS PERIPTOSEON KAI ANTOISTOIXHSH TIMHS STHN STATE
    '1) OLA KENA
    If ((Text8.Text = "") And (Text9.Text = "") And (Text10.Text = "") _
    And (Text11.Text = "") And (Text16.Text = "")) Then
        STATE = STATE1
        SS = SS1
        FF = FF1
    End If

    '2) MONO AR TIMOLOGIOY
    If Text8.Text <> "" Then
        STATE = STATE2
        SS = SS2
        FF = FF2
    End If

    '3) MONO POSO
    If Text11.Text <> "" Then
        STATE = STATE3
        SS = SS3
        FF = FF3
    End If

    '4) APO - MEXRI
    If ((Text9.Text <> "") Or (Text10.Text <> "")) Then
        STATE = STATE4
        SS = SS4
        FF = FF4
    End If
    
    '5A) PLHR 'H OXI
    If Text16.Text = "маи" Then
        STATE = STATE5A
        SS = SS5A
        FF = FF5A
    End If

    '5б) PLHR 'H OXI
    If Text16.Text = "ови" Then
        STATE = STATE5B
        SS = SS5B
        FF = FF5B
    End If

    '6) POSO & APO - MEXRI
    If ((Text11.Text <> "") And ((Text9.Text <> "") Or (Text10.Text <> ""))) Then
        STATE = STATE6
        SS = SS6
        FF = FF6
    End If

    '7A) POSO & "NAI" STO PLHROMENA 'H OXI
    If ((Text11.Text <> "") And (Text16.Text = "маи")) Then
        STATE = STATE7A
        SS = SS7A
        FF = FF7A
    End If

    '7B) POSO & "OXI" STO PLHROMENA 'H OXI
    If ((Text11.Text <> "") And (Text16.Text = "ови")) Then
        STATE = STATE7B
        SS = SS7B
        FF = FF7B
    End If

    '8A) APO MEXRI & PLHROMENA="NAI"
    If (((Text9.Text <> "") Or (Text10.Text <> "")) And (Text16.Text = "маи")) Then
        STATE = STATE8A
        SS = SS8A
        FF = FF8A
    End If

    '8B) APO MEXRI & PLHROMENA="OXI"
    If (((Text9.Text <> "") Or (Text10.Text <> "")) And (Text16.Text = "ови")) Then
        STATE = STATE8B
        SS = SS8B
        FF = FF8B
    End If

    '9A) KAI TA TRIA ME PLIROMENA NAI
    If ((Text11.Text <> "") And ((Text9.Text <> "") Or (Text10.Text <> "")) And (Text16.Text = "маи")) Then
        STATE = STATE9A
        SS = SS9A
        FF = FF9A
    End If

    '9B) KAI TA TRIA ME PLIROMENA OXI
    If ((Text11.Text <> "") And ((Text9.Text <> "") Or (Text10.Text <> "")) And (Text16.Text = "ови")) Then
        STATE = STATE9B
        SS = SS9B
        FF = FF9B
    End If

Case "цемийо"
    '-----------------------------------1--------------------------------------------
    STATE1 = " select аяихлос_тилокоциоу,глеяолгмиа_ейдосгс,посо,тупос,тупос_пистытийоу from " & UCase(ZETAIRIES.Text1.Text) & _
    " WHERE тупос='пистытийо' AND тупос_пистытийоу='цемийо'" & _
    " order by глеяолгмиа_ейдосгс"
    SS1 = " select SUM(посо) from " & UCase(ZETAIRIES.Text1.Text) & _
    " WHERE тупос='пистытийо' AND тупос_пистытийоу='цемийо'"
    FF1 = " select COUNT(аяихлос_тилокоциоу) from " & UCase(ZETAIRIES.Text1.Text) & _
    " WHERE тупос='пистытийо' AND тупос_пистытийоу='цемийо'"
    '-----------------------------------2--------------------------------------------
    STATE2 = "select аяихлос_тилокоциоу,глеяолгмиа_ейдосгс,посо,тупос,тупос_пистытийоу from " & UCase(ZETAIRIES.Text1.Text) & _
    " where аяихлос_тилокоциоу ='" & UCase(ZETAIRIES.Text8.Text) & "'" & " AND тупос='пистытийо' AND тупос_пистытийоу='цемийо'" & _
    " order by глеяолгмиа_ейдосгс"
    SS2 = "select SUM(посо) from " & UCase(ZETAIRIES.Text1.Text) & _
    " where аяихлос_тилокоциоу ='" & UCase(ZETAIRIES.Text8.Text) & "'" & " AND тупос='пистытийо' AND тупос_пистытийоу='цемийо'"
    FF2 = "select COUNT(аяихлос_тилокоциоу) from " & UCase(ZETAIRIES.Text1.Text) & _
    " where аяихлос_тилокоциоу ='" & UCase(ZETAIRIES.Text8.Text) & "'" & " AND тупос='пистытийо' AND тупос_пистытийоу='цемийо'"
    '-----------------------------------3--------------------------------------------
    STATE3 = "select аяихлос_тилокоциоу,глеяолгмиа_ейдосгс,посо,тупос,тупос_пистытийоу from " & UCase(ZETAIRIES.Text1.Text) & _
    " where посо =" & ZETAIRIES.Text11.Text & " AND тупос='пистытийо' AND тупос_пистытийоу='цемийо'" & _
    " order by глеяолгмиа_ейдосгс"
    SS3 = "select SUM(посо) from " & UCase(ZETAIRIES.Text1.Text) & _
    " where посо =" & ZETAIRIES.Text11.Text & " AND тупос='пистытийо' AND тупос_пистытийоу='цемийо'"
    FF3 = "select COUNT(аяихлос_тилокоциоу) from " & UCase(ZETAIRIES.Text1.Text) & _
    " where посо =" & ZETAIRIES.Text11.Text & " AND тупос='пистытийо' AND тупос_пистытийоу='цемийо'"
    '-----------------------------------4--------------------------------------------
    STATE4 = "select аяихлос_тилокоциоу,глеяолгмиа_ейдосгс,посо,тупос,тупос_пистытийоу from " & UCase(ZETAIRIES.Text1.Text) & _
    " where (глеяолгмиа_ейдосгс between " & "#" & ap & "#" & " and " & "#" & mexr & "#" & ")" & " AND тупос='пистытийо' AND тупос_пистытийоу='цемийо'" & _
    " order by глеяолгмиа_ейдосгс"
    SS4 = "select SUM(посо) from " & UCase(ZETAIRIES.Text1.Text) & _
    " where (глеяолгмиа_ейдосгс between " & "#" & ap & "#" & " and " & "#" & mexr & "#" & ")" & " AND тупос='пистытийо' AND тупос_пистытийоу='цемийо'"
    FF4 = "select COUNT(аяихлос_тилокоциоу) from " & UCase(ZETAIRIES.Text1.Text) & _
    " where (глеяолгмиа_ейдосгс between " & "#" & ap & "#" & " and " & "#" & mexr & "#" & ")" & " AND тупос='пистытийо' AND тупос_пистытийоу='цемийо'"
    '-----------------------------------5A--------------------------------------------
    STATE5A = "select аяихлос_тилокоциоу,глеяолгмиа_ейдосгс,посо,тупос,тупос_пистытийоу from " & UCase(ZETAIRIES.Text1.Text) & _
    " where еныжкгсг=-1" & " AND тупос='пистытийо' AND тупос_пистытийоу='цемийо'" & _
    " order by глеяолгмиа_ейдосгс"
    SS5A = "select SUM(посо) from " & UCase(ZETAIRIES.Text1.Text) & _
    " where еныжкгсг=-1" & " AND тупос='пистытийо' AND тупос_пистытийоу='цемийо'"
    FF5A = "select COUNT(аяихлос_тилокоциоу) from " & UCase(ZETAIRIES.Text1.Text) & _
    " where еныжкгсг=-1" & " AND тупос='пистытийо' AND тупос_пистытийоу='цемийо'"
    '-----------------------------------5B--------------------------------------------
    STATE5B = "select аяихлос_тилокоциоу,глеяолгмиа_ейдосгс,посо,тупос,тупос_пистытийоу from " & UCase(ZETAIRIES.Text1.Text) & _
    " where еныжкгсг=0" & " AND тупос='пистытийо' AND тупос_пистытийоу='цемийо'" & _
    " order by глеяолгмиа_ейдосгс"
    SS5B = "select SUM(посо) from " & UCase(ZETAIRIES.Text1.Text) & _
    " where еныжкгсг=0" & " AND тупос='пистытийо' AND тупос_пистытийоу='цемийо'"
    FF5B = "select COUNT(аяихлос_тилокоциоу) from " & UCase(ZETAIRIES.Text1.Text) & _
    " where еныжкгсг=0" & " AND тупос='пистытийо' AND тупос_пистытийоу='цемийо'"
    '-----------------------------------6--------------------------------------------
    STATE6 = "select аяихлос_тилокоциоу,глеяолгмиа_ейдосгс,посо,тупос,тупос_пистытийоу from " & UCase(ZETAIRIES.Text1.Text) & _
    " where (глеяолгмиа_ейдосгс between " & "#" & ap & "#" & " and " & "#" & mexr & "#" & ") AND посо =" & ZETAIRIES.Text11.Text & " AND тупос='пистытийо' AND тупос_пистытийоу='цемийо'" & _
    " order by глеяолгмиа_ейдосгс"
    SS6 = "select SUM(посо) from " & UCase(ZETAIRIES.Text1.Text) & _
    " where (глеяолгмиа_ейдосгс between " & "#" & ap & "#" & " and " & "#" & mexr & "#" & ") AND посо =" & ZETAIRIES.Text11.Text & " AND тупос='пистытийо' AND тупос_пистытийоу='цемийо'"
    FF6 = "select COUNT(аяихлос_тилокоциоу) from " & UCase(ZETAIRIES.Text1.Text) & _
    " where (глеяолгмиа_ейдосгс between " & "#" & ap & "#" & " and " & "#" & mexr & "#" & ") AND посо =" & ZETAIRIES.Text11.Text & " AND тупос='пистытийо' AND тупос_пистытийоу='цемийо'"
    '-----------------------------------7A--------------------------------------------
    STATE7A = "select аяихлос_тилокоциоу,глеяолгмиа_ейдосгс,посо,тупос,тупос_пистытийоу from " & UCase(ZETAIRIES.Text1.Text) & _
    " where еныжкгсг=-1 AND посо=" & ZETAIRIES.Text11.Text & " AND тупос='пистытийо' AND тупос_пистытийоу='цемийо'" & _
    " order by глеяолгмиа_ейдосгс"
    SS7A = "select SUM(посо) from " & UCase(ZETAIRIES.Text1.Text) & _
    " where еныжкгсг=-1 AND посо=" & ZETAIRIES.Text11.Text & " AND тупос='пистытийо' AND тупос_пистытийоу='цемийо'"
    FF7A = "select COUNT(аяихлос_тилокоциоу) from " & UCase(ZETAIRIES.Text1.Text) & _
    " where еныжкгсг=-1 AND посо=" & ZETAIRIES.Text11.Text & " AND тупос='пистытийо' AND тупос_пистытийоу='цемийо'"
    '-----------------------------------7B--------------------------------------------
    STATE7B = "select аяихлос_тилокоциоу,глеяолгмиа_ейдосгс,посо,тупос,тупос_пистытийоу from " & UCase(ZETAIRIES.Text1.Text) & _
    " where еныжкгсг=0 AND посо=" & ZETAIRIES.Text11.Text & " AND тупос='пистытийо' AND тупос_пистытийоу='цемийо'" & _
    " order by глеяолгмиа_ейдосгс"
    SS7B = "select SUM(посо) from " & UCase(ZETAIRIES.Text1.Text) & _
    " where еныжкгсг=0 AND посо=" & ZETAIRIES.Text11.Text & " AND тупос='пистытийо' AND тупос_пистытийоу='цемийо'"
    FF7B = "select COUNT(аяихлос_тилокоциоу) from " & UCase(ZETAIRIES.Text1.Text) & _
    " where еныжкгсг=0 AND посо=" & ZETAIRIES.Text11.Text & " AND тупос='пистытийо' AND тупос_пистытийоу='цемийо'"
    '-----------------------------------8A--------------------------------------------
    STATE8A = "select аяихлос_тилокоциоу,глеяолгмиа_ейдосгс,посо,тупос,тупос_пистытийоу from " & UCase(ZETAIRIES.Text1.Text) & _
    " where (глеяолгмиа_ейдосгс between " & "#" & ap & "#" & " and " & "#" & mexr & "#" & ") AND еныжкгсг=-1 " & " AND тупос='пистытийо' AND тупос_пистытийоу='цемийо'" & _
    " order by глеяолгмиа_ейдосгс"
    SS8A = "select SUM(посо) from " & UCase(ZETAIRIES.Text1.Text) & _
    " where (глеяолгмиа_ейдосгс between " & "#" & ap & "#" & " and " & "#" & mexr & "#" & ") AND еныжкгсг=-1 " & " AND тупос='пистытийо' AND тупос_пистытийоу='цемийо'"
    FF8A = "select COUNT(аяихлос_тилокоциоу) from " & UCase(ZETAIRIES.Text1.Text) & _
    " where (глеяолгмиа_ейдосгс between " & "#" & ap & "#" & " and " & "#" & mexr & "#" & ") AND еныжкгсг=-1 " & " AND тупос='пистытийо' AND тупос_пистытийоу='цемийо'"
    '-----------------------------------8B--------------------------------------------
    STATE8B = "select аяихлос_тилокоциоу,глеяолгмиа_ейдосгс,посо,тупос,тупос_пистытийоу from " & UCase(ZETAIRIES.Text1.Text) & _
    " where (глеяолгмиа_ейдосгс between " & "#" & ap & "#" & " and " & "#" & mexr & "#" & ") AND еныжкгсг=0 " & " AND тупос='пистытийо' AND тупос_пистытийоу='цемийо'" & _
    " order by глеяолгмиа_ейдосгс"
    SS8B = "select SUM(посо) from " & UCase(ZETAIRIES.Text1.Text) & _
    " where (глеяолгмиа_ейдосгс between " & "#" & ap & "#" & " and " & "#" & mexr & "#" & ") AND еныжкгсг=0 " & " AND тупос='пистытийо' AND тупос_пистытийоу='цемийо'"
    FF8B = "select COUNT(аяихлос_тилокоциоу) from " & UCase(ZETAIRIES.Text1.Text) & _
    " where (глеяолгмиа_ейдосгс between " & "#" & ap & "#" & " and " & "#" & mexr & "#" & ") AND еныжкгсг=0 " & " AND тупос='пистытийо' AND тупос_пистытийоу='цемийо'"
    '-----------------------------------9A--------------------------------------------
    STATE9A = "select аяихлос_тилокоциоу,глеяолгмиа_ейдосгс,посо,тупос,тупос_пистытийоу from " & UCase(ZETAIRIES.Text1.Text) & _
    " where (глеяолгмиа_ейдосгс between " & "#" & ap & "#" & " and " & "#" & mexr & "#" & ") AND еныжкгсг=-1 AND посо =" & ZETAIRIES.Text11.Text & " AND тупос='пистытийо' AND тупос_пистытийоу='цемийо'" & _
    " order by глеяолгмиа_ейдосгс"
    SS9A = "select SUM(посо) from " & UCase(ZETAIRIES.Text1.Text) & _
    " where (глеяолгмиа_ейдосгс between " & "#" & ap & "#" & " and " & "#" & mexr & "#" & ") AND еныжкгсг=-1 AND посо =" & ZETAIRIES.Text11.Text & " AND тупос='пистытийо' AND тупос_пистытийоу='цемийо'"
    FF9A = "select COUNT(аяихлос_тилокоциоу) from " & UCase(ZETAIRIES.Text1.Text) & _
    " where (глеяолгмиа_ейдосгс between " & "#" & ap & "#" & " and " & "#" & mexr & "#" & ") AND еныжкгсг=-1 AND посо =" & ZETAIRIES.Text11.Text & " AND тупос='пистытийо' AND тупос_пистытийоу='цемийо'"
    '-----------------------------------9B--------------------------------------------
    STATE9B = "select аяихлос_тилокоциоу,глеяолгмиа_ейдосгс,посо,тупос,тупос_пистытийоу from " & UCase(ZETAIRIES.Text1.Text) & _
    " where (глеяолгмиа_ейдосгс between " & "#" & ap & "#" & " and " & "#" & mexr & "#" & ") AND еныжкгсг=0 AND посо =" & ZETAIRIES.Text11.Text & " AND тупос='пистытийо' AND тупос_пистытийоу='цемийо'" & _
    " order by глеяолгмиа_ейдосгс"
    SS9B = "select SUM(посо) from " & UCase(ZETAIRIES.Text1.Text) & _
    " where (глеяолгмиа_ейдосгс between " & "#" & ap & "#" & " and " & "#" & mexr & "#" & ") AND еныжкгсг=0 AND посо =" & ZETAIRIES.Text11.Text & " AND тупос='пистытийо' AND тупос_пистытийоу='цемийо'"
    FF9B = "select COUNT(аяихлос_тилокоциоу) from " & UCase(ZETAIRIES.Text1.Text) & _
    " where (глеяолгмиа_ейдосгс between " & "#" & ap & "#" & " and " & "#" & mexr & "#" & ") AND еныжкгсг=0 AND посо =" & ZETAIRIES.Text11.Text & " AND тупос='пистытийо' AND тупос_пистытийоу='цемийо'"
    ' DIAXORISMOS PERIPTOSEON KAI ANTOISTOIXHSH TIMHS STHN STATE
    '1) OLA KENA
    If ((Text8.Text = "") And (Text9.Text = "") And (Text10.Text = "") _
    And (Text11.Text = "") And (Text16.Text = "")) Then
        STATE = STATE1
        SS = SS1
        FF = FF1
    End If

    '2) MONO AR TIMOLOGIOY
    If Text8.Text <> "" Then
        STATE = STATE2
        SS = SS2
        FF = FF2
    End If

    '3) MONO POSO
    If Text11.Text <> "" Then
        STATE = STATE3
        SS = SS3
        FF = FF3
    End If

    '4) APO - MEXRI
    If ((Text9.Text <> "") Or (Text10.Text <> "")) Then
        STATE = STATE4
        SS = SS4
        FF = FF4
    End If
    
    '5A) PLHR 'H OXI
    If Text16.Text = "маи" Then
        STATE = STATE5A
        SS = SS5A
        FF = FF5A
    End If

    '5б) PLHR 'H OXI
    If Text16.Text = "ови" Then
        STATE = STATE5B
        SS = SS5B
        FF = FF5B
    End If

    '6) POSO & APO - MEXRI
    If ((Text11.Text <> "") And ((Text9.Text <> "") Or (Text10.Text <> ""))) Then
        STATE = STATE6
        SS = SS6
        FF = FF6
    End If

    '7A) POSO & "NAI" STO PLHROMENA 'H OXI
    If ((Text11.Text <> "") And (Text16.Text = "маи")) Then
        STATE = STATE7A
        SS = SS7A
        FF = FF7A
    End If

    '7B) POSO & "OXI" STO PLHROMENA 'H OXI
    If ((Text11.Text <> "") And (Text16.Text = "ови")) Then
        STATE = STATE7B
        SS = SS7B
        FF = FF7B
    End If

    '8A) APO MEXRI & PLHROMENA="NAI"
    If (((Text9.Text <> "") Or (Text10.Text <> "")) And (Text16.Text = "маи")) Then
        STATE = STATE8A
        SS = SS8A
        FF = FF8A
    End If

    '8B) APO MEXRI & PLHROMENA="OXI"
    If (((Text9.Text <> "") Or (Text10.Text <> "")) And (Text16.Text = "ови")) Then
        STATE = STATE8B
        SS = SS8B
        FF = FF8B
    End If

    '9A) KAI TA TRIA ME PLIROMENA NAI
    If ((Text11.Text <> "") And ((Text9.Text <> "") Or (Text10.Text <> "")) And (Text16.Text = "маи")) Then
        STATE = STATE9A
        SS = SS9A
        FF = FF9A
    End If

    '9B) KAI TA TRIA ME PLIROMENA OXI
    If ((Text11.Text <> "") And ((Text9.Text <> "") Or (Text10.Text <> "")) And (Text16.Text = "ови")) Then
        STATE = STATE9B
        SS = SS9B
        FF = FF9B
    End If
    
Case "COOP"
    '-----------------------------------1--------------------------------------------
    STATE1 = " select аяихлос_тилокоциоу,глеяолгмиа_ейдосгс,посо,тупос,тупос_пистытийоу from " & UCase(ZETAIRIES.Text1.Text) & _
    " WHERE тупос='пистытийо' AND тупос_пистытийоу='COOP'" & _
    " order by глеяолгмиа_ейдосгс"
    SS1 = " select SUM(посо) from " & UCase(ZETAIRIES.Text1.Text) & _
    " WHERE тупос='пистытийо' AND тупос_пистытийоу='COOP'"
    FF1 = " select COUNT(аяихлос_тилокоциоу) from " & UCase(ZETAIRIES.Text1.Text) & _
    " WHERE тупос='пистытийо' AND тупос_пистытийоу='COOP'"
    '-----------------------------------2--------------------------------------------
    STATE2 = "select аяихлос_тилокоциоу,глеяолгмиа_ейдосгс,посо,тупос,тупос_пистытийоу from " & UCase(ZETAIRIES.Text1.Text) & _
    " where аяихлос_тилокоциоу ='" & ZETAIRIES.Text8.Text & "'" & " AND тупос='пистытийо' AND тупос_пистытийоу='COOP'" & _
    " order by глеяолгмиа_ейдосгс"
    SS2 = "select SUM(посо) from " & UCase(ZETAIRIES.Text1.Text) & _
    " where аяихлос_тилокоциоу ='" & ZETAIRIES.Text8.Text & "'" & " AND тупос='пистытийо' AND тупос_пистытийоу='COOP'"
    FF2 = "select COUNT(аяихлос_тилокоциоу) from " & UCase(ZETAIRIES.Text1.Text) & _
    " where аяихлос_тилокоциоу ='" & ZETAIRIES.Text8.Text & "'" & " AND тупос='пистытийо' AND тупос_пистытийоу='COOP'"
    '-----------------------------------3--------------------------------------------
    STATE3 = "select аяихлос_тилокоциоу,глеяолгмиа_ейдосгс,посо,тупос,тупос_пистытийоу from " & UCase(ZETAIRIES.Text1.Text) & _
    " where посо =" & ZETAIRIES.Text11.Text & " AND тупос='пистытийо' AND тупос_пистытийоу='COOP'" & _
    " order by глеяолгмиа_ейдосгс"
    SS3 = "select SUM(посо) from " & UCase(ZETAIRIES.Text1.Text) & _
    " where посо =" & ZETAIRIES.Text11.Text & " AND тупос='пистытийо' AND тупос_пистытийоу='COOP'"
    FF3 = "select COUNT(аяихлос_тилокоциоу) from " & UCase(ZETAIRIES.Text1.Text) & _
    " where посо =" & ZETAIRIES.Text11.Text & " AND тупос='пистытийо' AND тупос_пистытийоу='COOP'"
    '-----------------------------------4--------------------------------------------
    STATE4 = "select аяихлос_тилокоциоу,глеяолгмиа_ейдосгс,посо,тупос,тупос_пистытийоу from " & UCase(ZETAIRIES.Text1.Text) & _
    " where (глеяолгмиа_ейдосгс between " & "#" & ap & "#" & " and " & "#" & mexr & "#" & ")" & " AND тупос='пистытийо' AND тупос_пистытийоу='COOP'" & _
    " order by глеяолгмиа_ейдосгс"
    SS4 = "select SUM(посо) from " & UCase(ZETAIRIES.Text1.Text) & _
    " where (глеяолгмиа_ейдосгс between " & "#" & ap & "#" & " and " & "#" & mexr & "#" & ")" & " AND тупос='пистытийо' AND тупос_пистытийоу='COOP'"
    FF4 = "select COUNT(аяихлос_тилокоциоу) from " & UCase(ZETAIRIES.Text1.Text) & _
    " where (глеяолгмиа_ейдосгс between " & "#" & ap & "#" & " and " & "#" & mexr & "#" & ")" & " AND тупос='пистытийо' AND тупос_пистытийоу='COOP'"
    '-----------------------------------5A--------------------------------------------
    STATE5A = "select аяихлос_тилокоциоу,глеяолгмиа_ейдосгс,посо,тупос,тупос_пистытийоу from " & UCase(ZETAIRIES.Text1.Text) & _
    " where еныжкгсг=-1" & " AND тупос='пистытийо' AND тупос_пистытийоу='COOP'" & _
    " order by глеяолгмиа_ейдосгс"
    SS5A = "select SUM(посо) from " & UCase(ZETAIRIES.Text1.Text) & _
    " where еныжкгсг=-1" & " AND тупос='пистытийо' AND тупос_пистытийоу='COOP'"
    FF5A = "select COUNT(аяихлос_тилокоциоу) from " & UCase(ZETAIRIES.Text1.Text) & _
    " where еныжкгсг=-1" & " AND тупос='пистытийо' AND тупос_пистытийоу='COOP'"
    '-----------------------------------5B--------------------------------------------
    STATE5B = "select аяихлос_тилокоциоу,глеяолгмиа_ейдосгс,посо,тупос,тупос_пистытийоу from " & UCase(ZETAIRIES.Text1.Text) & _
    " where еныжкгсг=0" & " AND тупос='пистытийо' AND тупос_пистытийоу='COOP'" & _
    " order by глеяолгмиа_ейдосгс"
    SS5B = "select SUM(посо) from " & UCase(ZETAIRIES.Text1.Text) & _
    " where еныжкгсг=0" & " AND тупос='пистытийо' AND тупос_пистытийоу='COOP'"
    FF5B = "select COUNT(аяихлос_тилокоциоу) from " & UCase(ZETAIRIES.Text1.Text) & _
    " where еныжкгсг=0" & " AND тупос='пистытийо' AND тупос_пистытийоу='COOP'"
    '-----------------------------------6--------------------------------------------
    STATE6 = "select аяихлос_тилокоциоу,глеяолгмиа_ейдосгс,посо,тупос,тупос_пистытийоу from " & UCase(ZETAIRIES.Text1.Text) & _
    " where (глеяолгмиа_ейдосгс between " & "#" & ap & "#" & " and " & "#" & mexr & "#" & ") AND посо =" & ZETAIRIES.Text11.Text & " AND тупос='пистытийо' AND тупос_пистытийоу='COOP'" & _
    " order by глеяолгмиа_ейдосгс"
    SS6 = "select SUM(посо) from " & UCase(ZETAIRIES.Text1.Text) & _
    " where (глеяолгмиа_ейдосгс between " & "#" & ap & "#" & " and " & "#" & mexr & "#" & ") AND посо =" & ZETAIRIES.Text11.Text & " AND тупос='пистытийо' AND тупос_пистытийоу='COOP'"
    FF6 = "select COUNT(аяихлос_тилокоциоу) from " & UCase(ZETAIRIES.Text1.Text) & _
    " where (глеяолгмиа_ейдосгс between " & "#" & ap & "#" & " and " & "#" & mexr & "#" & ") AND посо =" & ZETAIRIES.Text11.Text & " AND тупос='пистытийо' AND тупос_пистытийоу='COOP'"
    '-----------------------------------7A--------------------------------------------
    STATE7A = "select аяихлос_тилокоциоу,глеяолгмиа_ейдосгс,посо,тупос,тупос_пистытийоу from " & UCase(ZETAIRIES.Text1.Text) & _
    " where еныжкгсг=-1 AND посо=" & ZETAIRIES.Text11.Text & " AND тупос='пистытийо' AND тупос_пистытийоу='COOP'" & _
    " order by глеяолгмиа_ейдосгс"
    SS7A = "select SUM(посо) from " & UCase(ZETAIRIES.Text1.Text) & _
    " where еныжкгсг=-1 AND посо=" & ZETAIRIES.Text11.Text & " AND тупос='пистытийо' AND тупос_пистытийоу='COOP'"
    FF7A = "select COUNT(аяихлос_тилокоциоу) from " & UCase(ZETAIRIES.Text1.Text) & _
    " where еныжкгсг=-1 AND посо=" & ZETAIRIES.Text11.Text & " AND тупос='пистытийо' AND тупос_пистытийоу='COOP'"
    '-----------------------------------7B--------------------------------------------
    STATE7B = "select аяихлос_тилокоциоу,глеяолгмиа_ейдосгс,посо,тупос,тупос_пистытийоу from " & UCase(ZETAIRIES.Text1.Text) & _
    " where еныжкгсг=0 AND посо=" & ZETAIRIES.Text11.Text & " AND тупос='пистытийо' AND тупос_пистытийоу='COOP'" & _
    " order by глеяолгмиа_ейдосгс"
    SS7B = "select SUM(посо) from " & UCase(ZETAIRIES.Text1.Text) & _
    " where еныжкгсг=0 AND посо=" & ZETAIRIES.Text11.Text & " AND тупос='пистытийо' AND тупос_пистытийоу='COOP'"
    FF7B = "select COUNT(аяихлос_тилокоциоу) from " & UCase(ZETAIRIES.Text1.Text) & _
    " where еныжкгсг=0 AND посо=" & ZETAIRIES.Text11.Text & " AND тупос='пистытийо' AND тупос_пистытийоу='COOP'"
    '-----------------------------------8A--------------------------------------------
    STATE8A = "select аяихлос_тилокоциоу,глеяолгмиа_ейдосгс,посо,тупос,тупос_пистытийоу from " & UCase(ZETAIRIES.Text1.Text) & _
    " where (глеяолгмиа_ейдосгс between " & "#" & ap & "#" & " and " & "#" & mexr & "#" & ") AND еныжкгсг=-1 " & " AND тупос='пистытийо' AND тупос_пистытийоу='COOP'" & _
    " order by глеяолгмиа_ейдосгс"
    SS8A = "select SUM(посо) from " & UCase(ZETAIRIES.Text1.Text) & _
    " where (глеяолгмиа_ейдосгс between " & "#" & ap & "#" & " and " & "#" & mexr & "#" & ") AND еныжкгсг=-1 " & " AND тупос='пистытийо' AND тупос_пистытийоу='COOP'"
    FF8A = "select COUNT(аяихлос_тилокоциоу) from " & UCase(ZETAIRIES.Text1.Text) & _
    " where (глеяолгмиа_ейдосгс between " & "#" & ap & "#" & " and " & "#" & mexr & "#" & ") AND еныжкгсг=-1 " & " AND тупос='пистытийо' AND тупос_пистытийоу='COOP'"
    '-----------------------------------8B--------------------------------------------
    STATE8B = "select аяихлос_тилокоциоу,глеяолгмиа_ейдосгс,посо,тупос,тупос_пистытийоу from " & UCase(ZETAIRIES.Text1.Text) & _
    " where (глеяолгмиа_ейдосгс between " & "#" & ap & "#" & " and " & "#" & mexr & "#" & ") AND еныжкгсг=0 " & " AND тупос='пистытийо' AND тупос_пистытийоу='COOP'" & _
    " order by глеяолгмиа_ейдосгс"
    SS8B = "select SUM(посо) from " & UCase(ZETAIRIES.Text1.Text) & _
    " where (глеяолгмиа_ейдосгс between " & "#" & ap & "#" & " and " & "#" & mexr & "#" & ") AND еныжкгсг=0 " & " AND тупос='пистытийо' AND тупос_пистытийоу='COOP'"
    FF8B = "select COUNT(аяихлос_тилокоциоу) from " & UCase(ZETAIRIES.Text1.Text) & _
    " where (глеяолгмиа_ейдосгс between " & "#" & ap & "#" & " and " & "#" & mexr & "#" & ") AND еныжкгсг=0 " & " AND тупос='пистытийо' AND тупос_пистытийоу='COOP'"
    '-----------------------------------9A--------------------------------------------
    STATE9A = "select аяихлос_тилокоциоу,глеяолгмиа_ейдосгс,посо,тупос,тупос_пистытийоу from " & UCase(ZETAIRIES.Text1.Text) & _
    " where (глеяолгмиа_ейдосгс between " & "#" & ap & "#" & " and " & "#" & mexr & "#" & ") AND еныжкгсг=-1 AND посо =" & ZETAIRIES.Text11.Text & " AND тупос='пистытийо' AND тупос_пистытийоу='COOP'" & _
    " order by глеяолгмиа_ейдосгс"
    SS9A = "select SUM(посо) from " & UCase(ZETAIRIES.Text1.Text) & _
    " where (глеяолгмиа_ейдосгс between " & "#" & ap & "#" & " and " & "#" & mexr & "#" & ") AND еныжкгсг=-1 AND посо =" & ZETAIRIES.Text11.Text & " AND тупос='пистытийо' AND тупос_пистытийоу='COOP'"
    FF9A = "select COUNT(аяихлос_тилокоциоу) from " & UCase(ZETAIRIES.Text1.Text) & _
    " where (глеяолгмиа_ейдосгс between " & "#" & ap & "#" & " and " & "#" & mexr & "#" & ") AND еныжкгсг=-1 AND посо =" & ZETAIRIES.Text11.Text & " AND тупос='пистытийо' AND тупос_пистытийоу='COOP'"
    '-----------------------------------9B--------------------------------------------
    STATE9B = "select аяихлос_тилокоциоу,глеяолгмиа_ейдосгс,посо,тупос,тупос_пистытийоу from " & UCase(ZETAIRIES.Text1.Text) & _
    " where (глеяолгмиа_ейдосгс between " & "#" & ap & "#" & " and " & "#" & mexr & "#" & ") AND еныжкгсг=0 AND посо =" & ZETAIRIES.Text11.Text & " AND тупос='пистытийо' AND тупос_пистытийоу='COOP'" & _
    " order by глеяолгмиа_ейдосгс"
    SS9B = "select SUM(посо) from " & UCase(ZETAIRIES.Text1.Text) & _
    " where (глеяолгмиа_ейдосгс between " & "#" & ap & "#" & " and " & "#" & mexr & "#" & ") AND еныжкгсг=0 AND посо =" & ZETAIRIES.Text11.Text & " AND тупос='пистытийо' AND тупос_пистытийоу='COOP'"
    FF9B = "select COUNT(аяихлос_тилокоциоу) from " & UCase(ZETAIRIES.Text1.Text) & _
    " where (глеяолгмиа_ейдосгс between " & "#" & ap & "#" & " and " & "#" & mexr & "#" & ") AND еныжкгсг=0 AND посо =" & ZETAIRIES.Text11.Text & " AND тупос='пистытийо' AND тупос_пистытийоу='COOP'"
    ' DIAXORISMOS PERIPTOSEON KAI ANTOISTOIXHSH TIMHS STHN STATE
    '1) OLA KENA
    If ((Text8.Text = "") And (Text9.Text = "") And (Text10.Text = "") _
    And (Text11.Text = "") And (Text16.Text = "")) Then
        STATE = STATE1
        SS = SS1
        FF = FF1
    End If

    '2) MONO AR TIMOLOGIOY
    If Text8.Text <> "" Then
        STATE = STATE2
        SS = SS2
        FF = FF2
    End If

    '3) MONO POSO
    If Text11.Text <> "" Then
        STATE = STATE3
        SS = SS3
        FF = FF3
    End If

    '4) APO - MEXRI
    If ((Text9.Text <> "") Or (Text10.Text <> "")) Then
        STATE = STATE4
        SS = SS4
        FF = FF4
    End If
    
    '5A) PLHR 'H OXI
    If Text16.Text = "маи" Then
        STATE = STATE5A
        SS = SS5A
        FF = FF5A
    End If

    '5б) PLHR 'H OXI
    If Text16.Text = "ови" Then
        STATE = STATE5B
        SS = SS5B
        FF = FF5B
    End If

    '6) POSO & APO - MEXRI
    If ((Text11.Text <> "") And ((Text9.Text <> "") Or (Text10.Text <> ""))) Then
        STATE = STATE6
        SS = SS6
        FF = FF6
    End If

    '7A) POSO & "NAI" STO PLHROMENA 'H OXI
    If ((Text11.Text <> "") And (Text16.Text = "маи")) Then
        STATE = STATE7A
        SS = SS7A
        FF = FF7A
    End If

    '7B) POSO & "OXI" STO PLHROMENA 'H OXI
    If ((Text11.Text <> "") And (Text16.Text = "ови")) Then
        STATE = STATE7B
        SS = SS7B
        FF = FF7B
    End If

    '8A) APO MEXRI & PLHROMENA="NAI"
    If (((Text9.Text <> "") Or (Text10.Text <> "")) And (Text16.Text = "маи")) Then
        STATE = STATE8A
        SS = SS8A
        FF = FF8A
    End If

    '8B) APO MEXRI & PLHROMENA="OXI"
    If (((Text9.Text <> "") Or (Text10.Text <> "")) And (Text16.Text = "ови")) Then
        STATE = STATE8B
        SS = SS8B
        FF = FF8B
    End If

    '9A) KAI TA TRIA ME PLIROMENA NAI
    If ((Text11.Text <> "") And ((Text9.Text <> "") Or (Text10.Text <> "")) And (Text16.Text = "маи")) Then
        STATE = STATE9A
        SS = SS9A
        FF = FF9A
    End If

    '9B) KAI TA TRIA ME PLIROMENA OXI
    If ((Text11.Text <> "") And ((Text9.Text <> "") Or (Text10.Text <> "")) And (Text16.Text = "ови")) Then
        STATE = STATE9B
        SS = SS9B
        FF = FF9B
    End If
End Select

FLAG_FORM6 = 2
PATIMA_DATAGRID_F6 = 0
Load ZForm6
ZForm6.Show
GoTo TELOS:
' ANTIMETOPISI LATHON
imer_apo:
MsgBox ("дем дысате сыста тгм глеяолгмиа сто пкаисио <<апо>>"), vbCritical, "пяосовг"
index = 32
GoTo TELOS:

imer_mexri:
If index = 32 Then
    GoTo TELOS:
Else
    MsgBox ("дем дысате сыста тгм глеяолгмиа сто пкаисио <<левяи>>"), vbCritical, "пяосовг"
    index = 32
    GoTo TELOS:
End If

imer:
If index = 32 Then
    GoTo TELOS:
Else
    MsgBox ("г глеяолгмиа <<апо>> еимаи лецакутеяг апо тгм <<левяи>>"), vbCritical, "пяосовг"
    index = 32
    GoTo TELOS:
End If

imer_2:
If index = 32 Then
    GoTo TELOS:
Else
    MsgBox ("г глеяолгмиа <<левяи>> еимаи лецакутеяг апо тгм сглеяимг"), vbCritical, "пяосовг"
    index = 32
    GoTo TELOS:
End If

POSO:
If index = 32 Then
    GoTo TELOS:
Else
    MsgBox ("дем дысате сыста то посо"), vbCritical, "пяосовг"
    index = 32
    GoTo TELOS:
End If

EKSO:
If index = 32 Then
    GoTo TELOS:
Else
    MsgBox ("сто педио <<пкгяылема>> пяепеи ам сулпкгяыметаи, ма сулпкгяыметаи ломо ле тис кенеис <<маи>> ╧ <<ови>>. паяайакы диояхысте"), vbCritical, "пяосовг"
    index = 32
    GoTo TELOS:
End If

lathos:
If index = 32 Then
    GoTo TELOS:
Else
    MsgBox ("ам супкгяысате то педио аяих/тилокоц. тоте лгм сулкгяыметаи йамема акко педио"), vbCritical, "пяосовг"
    Text9.Text = ""
    Text10.Text = ""
    Text11.Text = ""
    Text16.Text = ""
    index = 32
    GoTo TELOS:
End If

TYPOS_PISTOTIKOY:
If index = 32 Then
    GoTo TELOS:
Else
    MsgBox ("то педио тупос пистытийоу пяепеи ма сулпкгяыхеи ле лиа апо тис кенеис : ока ╧ цемийо ╧ COOP.паяайакы сулпкгяысте то педио ле лиа апо аутес тис кенеис. "), vbCritical, "пяосовг"
    index = 32
    GoTo TELOS:
End If

er:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
ZForm6.Text1.Left = 6210
ZForm6.Label2.Left = 4770
ZForm6.Text1.Width = 2800
If rs1.STATE = 1 Then rs1.Close
If db1.STATE = 1 Then db1.Close
End Sub

Private Sub Command26_Click()
Text1.Text = Text24.Text
On Error GoTo er:

If rs1.STATE = 1 Then rs1.Close
If db1.STATE = 1 Then db1.Close
db1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & ZETAIRIES_DIADROMHS_BACKUP_DIAX_1 & ";" & _
"Persist Security Info=False"
db1.Open ZETAIRIES_DIADROMHS_BACKUP_DIAX
rs1.Open "[" & ZETAIRIES.Text1.Text & "]", db1, adOpenDynamic, adLockBatchOptimistic

Dim ap, mexr, ap_temp, mexr_temp
Dim index As Integer
index = 1
'*************** ELEGXOS LATHON ***********************************
If Text9.Text <> "" Then
    If IsDate(Text9.Text) = False Then GoTo imer_apo:
End If
If Text10.Text <> "" Then
    If IsDate(Text10.Text) = False Then GoTo imer_mexri:
End If
If (Text9.Text <> "") And (Text10.Text <> "") Then
    If ((CDate(Text9.Text)) > (CDate(Text10.Text))) Then GoTo imer:
End If
If Text10.Text <> "" Then
    If ((CDate(Text10.Text)) > Date) Then GoTo imer_2:
End If
If Text11.Text <> "" Then
    If IsNumeric(Text11.Text) = False Then GoTo POSO:
End If
If Text16.Text <> "" Then
    If ((Text16.Text = "маи") Or (Text16.Text = "ови")) Then

    Else
        GoTo EKSO:
    End If
End If

If (Text8.Text <> "") And ((Text9.Text <> "") Or (Text10.Text <> "") Or (Text11.Text <> "") Or (Text16.Text <> "")) Then GoTo lathos:
' LEITOYRGIA
' ANTOISTOIXHSH TIMHS STHN METABLHTH TYPOY DATE AP
If Text9.Text = "" Then
ap_temp = #1/1/2004#
Else
ap_temp = CDate(Text9.Text)
End If
' ANTOISTOIXHSH TIMHS STHN METABLHTH TYPOY DATE MEXR
If Text10.Text = "" Then
mexr_temp = Date
Else
mexr_temp = CDate(Text10.Text)
End If
' ftiaksimo telikon timon
If Day(ap_temp) < 12 Then
    ap = CDate(Month(ap_temp) & "/" & Day(ap_temp) & "/" & Year(ap_temp))
Else
    ap = ap_temp
End If
If Day(mexr_temp) < 12 Then
    mexr = CDate(Month(mexr_temp) & "/" & Day(mexr_temp) & "/" & Year(mexr_temp))
Else
    mexr = mexr_temp
End If
'****************** ORISMOS EROTHMATON ******************************************

'---------------------------------------- 1---------------------------------------
STATE1 = " select аяихлос_епитацгс,глеяолгмиа_ейдосгс,посо,глеяолгмиа_еныжкгсгс,тупос from " & UCase(ZETAIRIES.Text1.Text) & _
" WHERE тупос='епитацг'" & _
" order by глеяолгмиа_ейдосгс"
SS1 = " select SUM(посо) from " & UCase(ZETAIRIES.Text1.Text) & _
" WHERE тупос='епитацг'"
FF1 = " select COUNT(аяихлос_тилокоциоу) from " & UCase(ZETAIRIES.Text1.Text) & _
" WHERE тупос='епитацг'"
'---------------------------------------- 2---------------------------------------
STATE2 = "select аяихлос_епитацгс,глеяолгмиа_ейдосгс,посо,глеяолгмиа_еныжкгсгс,тупос from " & UCase(ZETAIRIES.Text1.Text) & _
 " where аяихлос_тилокоциоу ='" & UCase(ZETAIRIES.Text8.Text) & "'" & " AND тупос='епитацг'" & _
 " order by глеяолгмиа_ейдосгс"
SS2 = "select SUM(посо) from " & UCase(ZETAIRIES.Text1.Text) & _
 " where аяихлос_тилокоциоу ='" & UCase(ZETAIRIES.Text8.Text) & "'" & " AND тупос='епитацг'"
FF2 = "select COUNT(аяихлос_тилокоциоу) from " & UCase(ZETAIRIES.Text1.Text) & _
 " where аяихлос_тилокоциоу ='" & UCase(ZETAIRIES.Text8.Text) & "'" & " AND тупос='епитацг'"
'---------------------------------------- 3---------------------------------------
STATE3 = "select аяихлос_епитацгс,глеяолгмиа_ейдосгс,посо,глеяолгмиа_еныжкгсгс,тупос from " & UCase(ZETAIRIES.Text1.Text) & _
 " where посо =" & ZETAIRIES.Text11.Text & " AND тупос='епитацг'" & _
 " order by глеяолгмиа_ейдосгс"
SS3 = "select SUM(посо) from " & UCase(ZETAIRIES.Text1.Text) & _
 " where посо =" & ZETAIRIES.Text11.Text & " AND тупос='епитацг'"
 FF3 = "select COUNT(аяихлос_тилокоциоу) from " & UCase(ZETAIRIES.Text1.Text) & _
 " where посо =" & ZETAIRIES.Text11.Text & " AND тупос='епитацг'"
 '---------------------------------------- 4---------------------------------------
STATE4 = "select аяихлос_епитацгс,глеяолгмиа_ейдосгс,посо,глеяолгмиа_еныжкгсгс,тупос from " & UCase(ZETAIRIES.Text1.Text) & _
" where (глеяолгмиа_ейдосгс between " & "#" & ap & "#" & " and " & "#" & mexr & "#" & ")" & " AND тупос='епитацг'" & _
" order by глеяолгмиа_ейдосгс"
SS4 = "select SUM(посо) from " & UCase(ZETAIRIES.Text1.Text) & _
" where (глеяолгмиа_ейдосгс between " & "#" & ap & "#" & " and " & "#" & mexr & "#" & ")" & " AND тупос='епитацг'"
FF4 = "select COUNT(аяихлос_тилокоциоу) from " & UCase(ZETAIRIES.Text1.Text) & _
" where (глеяолгмиа_ейдосгс between " & "#" & ap & "#" & " and " & "#" & mexr & "#" & ")" & " AND тупос='епитацг'"
'---------------------------------------- 5а---------------------------------------
STATE5A = "select аяихлос_епитацгс,глеяолгмиа_ейдосгс,посо,глеяолгмиа_еныжкгсгс,тупос from " & UCase(ZETAIRIES.Text1.Text) & _
 " where еныжкгсг=-1" & " AND тупос='епитацг'" & _
" order by глеяолгмиа_ейдосгс"
SS5A = "select SUM(посо) from " & UCase(ZETAIRIES.Text1.Text) & _
 " where еныжкгсг=-1" & " AND тупос='епитацг'"
FF5A = "select COUNT(аяихлос_тилокоциоу) from " & UCase(ZETAIRIES.Text1.Text) & _
 " where еныжкгсг=-1" & " AND тупос='епитацг'"
'---------------------------------------- 5б ---------------------------------------
STATE5B = "select аяихлос_епитацгс,глеяолгмиа_ейдосгс,посо,глеяолгмиа_еныжкгсгс,тупос from " & UCase(ZETAIRIES.Text1.Text) & _
 " where еныжкгсг=0" & " AND тупос='епитацг'" & _
" order by глеяолгмиа_ейдосгс"
SS5B = "select SUM(посо) from " & UCase(ZETAIRIES.Text1.Text) & _
 " where еныжкгсг=0" & " AND тупос='епитацг'"
FF5B = "select COUNT(аяихлос_тилокоциоу) from " & UCase(ZETAIRIES.Text1.Text) & _
 " where еныжкгсг=0" & " AND тупос='епитацг'"
'---------------------------------------- 6 ---------------------------------------
STATE6 = "select аяихлос_епитацгс,глеяолгмиа_ейдосгс,посо,глеяолгмиа_еныжкгсгс,тупос from " & UCase(ZETAIRIES.Text1.Text) & _
" where (глеяолгмиа_ейдосгс between " & "#" & ap & "#" & " and " & "#" & mexr & "#" & ") AND посо =" & ZETAIRIES.Text11.Text & " AND тупос='епитацг'" & _
" order by глеяолгмиа_ейдосгс"
SS6 = "select SUM(посо) from " & UCase(ZETAIRIES.Text1.Text) & _
" where (глеяолгмиа_ейдосгс between " & "#" & ap & "#" & " and " & "#" & mexr & "#" & ") AND посо =" & ZETAIRIES.Text11.Text & " AND тупос='епитацг'"
FF6 = "select COUNT(аяихлос_тилокоциоу) from " & UCase(ZETAIRIES.Text1.Text) & _
" where (глеяолгмиа_ейдосгс between " & "#" & ap & "#" & " and " & "#" & mexr & "#" & ") AND посо =" & ZETAIRIES.Text11.Text & " AND тупос='епитацг'"

'---------------------------------------- 7а ---------------------------------------
STATE7A = "select аяихлос_епитацгс,глеяолгмиа_ейдосгс,посо,глеяолгмиа_еныжкгсгс,тупос from " & UCase(ZETAIRIES.Text1.Text) & _
" where еныжкгсг=-1 AND посо=" & ZETAIRIES.Text11.Text & " AND тупос='епитацг'" & _
" order by глеяолгмиа_ейдосгс"
SS7A = "select SUM(посо) from " & UCase(ZETAIRIES.Text1.Text) & _
" where еныжкгсг=-1 AND посо=" & ZETAIRIES.Text11.Text & " AND тупос='епитацг'"
FF7A = "select COUNT(аяихлос_тилокоциоу) from " & UCase(ZETAIRIES.Text1.Text) & _
" where еныжкгсг=-1 AND посо=" & ZETAIRIES.Text11.Text & " AND тупос='епитацг'"
'---------------------------------------- 7б ---------------------------------------
STATE7B = "select аяихлос_епитацгс,глеяолгмиа_ейдосгс,посо,глеяолгмиа_еныжкгсгс,тупос from " & UCase(ZETAIRIES.Text1.Text) & _
" where еныжкгсг=0 AND посо=" & ZETAIRIES.Text11.Text & " AND тупос='епитацг'" & _
" order by глеяолгмиа_ейдосгс"
SS7B = "select SUM(посо) from " & UCase(ZETAIRIES.Text1.Text) & _
" where еныжкгсг=0 AND посо=" & ZETAIRIES.Text11.Text & " AND тупос='епитацг'"
FF7B = "select COUNT(аяихлос_тилокоциоу) from " & UCase(ZETAIRIES.Text1.Text) & _
" where еныжкгсг=0 AND посо=" & ZETAIRIES.Text11.Text & " AND тупос='епитацг'"
'---------------------------------------- 8а ---------------------------------------
STATE8A = "select аяихлос_епитацгс,глеяолгмиа_ейдосгс,посо,глеяолгмиа_еныжкгсгс,тупос from " & UCase(ZETAIRIES.Text1.Text) & _
" where (глеяолгмиа_ейдосгс between " & "#" & ap & "#" & " and " & "#" & mexr & "#" & ") AND еныжкгсг=-1 " & " AND тупос='епитацг'" & _
" order by глеяолгмиа_ейдосгс"
SS8A = "select SUM(посо) from " & UCase(ZETAIRIES.Text1.Text) & _
" where (глеяолгмиа_ейдосгс between " & "#" & ap & "#" & " and " & "#" & mexr & "#" & ") AND еныжкгсг=-1 " & " AND тупос='епитацг'"
FF8A = "select COUNT(аяихлос_тилокоциоу) from " & UCase(ZETAIRIES.Text1.Text) & _
" where (глеяолгмиа_ейдосгс between " & "#" & ap & "#" & " and " & "#" & mexr & "#" & ") AND еныжкгсг=-1 " & " AND тупос='епитацг'"
'---------------------------------------- 8б ---------------------------------------
STATE8B = "select аяихлос_епитацгс,глеяолгмиа_ейдосгс,посо,глеяолгмиа_еныжкгсгс,тупос from " & UCase(ZETAIRIES.Text1.Text) & _
" where (глеяолгмиа_ейдосгс between " & "#" & ap & "#" & " and " & "#" & mexr & "#" & ") AND еныжкгсг=0 " & " AND тупос='епитацг'" & _
" order by глеяолгмиа_ейдосгс"
SS8B = "select SUM(посо) from " & UCase(ZETAIRIES.Text1.Text) & _
" where (глеяолгмиа_ейдосгс between " & "#" & ap & "#" & " and " & "#" & mexr & "#" & ") AND еныжкгсг=0 " & " AND тупос='епитацг'"
FF8B = "select COUNT(аяихлос_тилокоциоу) from " & UCase(ZETAIRIES.Text1.Text) & _
" where (глеяолгмиа_ейдосгс between " & "#" & ap & "#" & " and " & "#" & mexr & "#" & ") AND еныжкгсг=0 " & " AND тупос='епитацг'"
'---------------------------------------- 9а ---------------------------------------
STATE9A = "select аяихлос_епитацгс,глеяолгмиа_ейдосгс,посо,глеяолгмиа_еныжкгсгс,тупос from " & UCase(ZETAIRIES.Text1.Text) & _
" where (глеяолгмиа_ейдосгс between " & "#" & ap & "#" & " and " & "#" & mexr & "#" & ") AND еныжкгсг=-1 AND посо =" & ZETAIRIES.Text11.Text & " AND тупос='епитацг'" & _
" order by глеяолгмиа_ейдосгс"
SS9A = "select SUM(посо) from " & UCase(ZETAIRIES.Text1.Text) & _
" where (глеяолгмиа_ейдосгс between " & "#" & ap & "#" & " and " & "#" & mexr & "#" & ") AND еныжкгсг=-1 AND посо =" & ZETAIRIES.Text11.Text & " AND тупос='епитацг'"
FF9A = "select COUNT(аяихлос_тилокоциоу) from " & UCase(ZETAIRIES.Text1.Text) & _
" where (глеяолгмиа_ейдосгс between " & "#" & ap & "#" & " and " & "#" & mexr & "#" & ") AND еныжкгсг=-1 AND посо =" & ZETAIRIES.Text11.Text & " AND тупос='епитацг'"
'---------------------------------------- 9б ---------------------------------------
STATE9B = "select аяихлос_епитацгс,глеяолгмиа_ейдосгс,посо,глеяолгмиа_еныжкгсгс,тупос from " & UCase(ZETAIRIES.Text1.Text) & _
" where (глеяолгмиа_ейдосгс between " & "#" & ap & "#" & " and " & "#" & mexr & "#" & ") AND еныжкгсг=0 AND посо =" & ZETAIRIES.Text11.Text & " AND тупос='епитацг'" & _
" order by глеяолгмиа_ейдосгс"
SS9B = "select SUM(посо) from " & UCase(ZETAIRIES.Text1.Text) & _
" where (глеяолгмиа_ейдосгс between " & "#" & ap & "#" & " and " & "#" & mexr & "#" & ") AND еныжкгсг=0 AND посо =" & ZETAIRIES.Text11.Text & " AND тупос='епитацг'"
FF9B = "select COUNT(аяихлос_тилокоциоу) from " & UCase(ZETAIRIES.Text1.Text) & _
" where (глеяолгмиа_ейдосгс between " & "#" & ap & "#" & " and " & "#" & mexr & "#" & ") AND еныжкгсг=0 AND посо =" & ZETAIRIES.Text11.Text & " AND тупос='епитацг'"


' DIAXORISMOS PERIPTOSEON KAI ANTOISTOIXHSH TIMHS STHN STATE
'1) OLA KENA
If ((Text8.Text = "") And (Text9.Text = "") And (Text10.Text = "") _
And (Text11.Text = "") And (Text16.Text = "")) Then
    STATE = STATE1
    SS = SS1
    FF = FF1
End If
    
'2) MONO AR TIMOLOGIOY
If Text8.Text <> "" Then
    STATE = STATE2
    SS = SS2
    FF = FF2
End If

'3) MONO POSO
If Text11.Text <> "" Then
    STATE = STATE3
    SS = SS3
    FF = FF3
End If

'4) APO - MEXRI
If ((Text9.Text <> "") Or (Text10.Text <> "")) Then
    STATE = STATE4
    SS = SS4
    FF = FF4
End If
    
'5A) PLHR 'H OXI
If Text16.Text = "маи" Then
    STATE = STATE5A
    SS = SS5A
    FF = FF5A
End If
    
'5б) PLHR 'H OXI
If Text16.Text = "ови" Then
    STATE = STATE5B
    SS = SS5B
    FF = FF5B
End If

'6) POSO & APO - MEXRI
If ((Text11.Text <> "") And ((Text9.Text <> "") Or (Text10.Text <> ""))) Then
    STATE = STATE6
    SS = SS6
    FF = FF6
End If

'7A) POSO & "NAI" STO PLHROMENA 'H OXI
If ((Text11.Text <> "") And (Text16.Text = "маи")) Then
    STATE = STATE7A
    SS = SS7A
    FF = FF7A
End If

'7B) POSO & "OXI" STO PLHROMENA 'H OXI
If ((Text11.Text <> "") And (Text16.Text = "ови")) Then
    STATE = STATE7B
    SS = SS7B
    FF = FF7B
End If

'8A) APO MEXRI & PLHROMENA="NAI"
If (((Text9.Text <> "") Or (Text10.Text <> "")) And (Text16.Text = "маи")) Then
    STATE = STATE8A
    SS = SS8A
    FF = FF8A
End If

'8B) APO MEXRI & PLHROMENA="OXI"
If (((Text9.Text <> "") Or (Text10.Text <> "")) And (Text16.Text = "ови")) Then
    STATE = STATE8B
    SS = SS8B
    FF = FF8B
End If

'9A) KAI TA TRIA ME PLIROMENA NAI
If ((Text11.Text <> "") And ((Text9.Text <> "") Or (Text10.Text <> "")) And (Text16.Text = "маи")) Then
    STATE = STATE9A
    SS = SS9A
    FF = FF9A
End If
    
'9B) KAI TA TRIA ME PLIROMENA OXI
If ((Text11.Text <> "") And ((Text9.Text <> "") Or (Text10.Text <> "")) And (Text16.Text = "ови")) Then
    STATE = STATE9B
    SS = SS9B
    FF = FF9B
End If

FLAG_FORM6 = 3
PATIMA_DATAGRID_F6 = 1
Load ZForm6
ZForm6.Show
GoTo TELOS:
' ANTIMETOPISI LATHON
imer_apo:
MsgBox ("дем дысате сыста тгм глеяолгмиа сто пкаисио <<апо>>"), vbCritical, "пяосовг"
index = 32
GoTo TELOS:

imer_mexri:
If index = 32 Then
    GoTo TELOS:
Else
    MsgBox ("дем дысате сыста тгм глеяолгмиа сто пкаисио <<левяи>>"), vbCritical, "пяосовг"
    index = 32
    GoTo TELOS:
End If

imer:
If index = 32 Then
    GoTo TELOS:
Else
    MsgBox ("г глеяолгмиа <<апо>> еимаи лецакутеяг апо тгм <<левяи>>"), vbCritical, "пяосовг"
    index = 32
    GoTo TELOS:
End If

imer_2:
If index = 32 Then
    GoTo TELOS:
Else
    MsgBox ("г глеяолгмиа <<левяи>> еимаи лецакутеяг апо тгм сглеяимг"), vbCritical, "пяосовг"
    index = 32
    GoTo TELOS:
End If

POSO:
If index = 32 Then
    GoTo TELOS:
Else
    MsgBox ("дем дысате сыста то посо"), vbCritical, "пяосовг"
    index = 32
    GoTo TELOS:
End If

EKSO:
If index = 32 Then
    GoTo TELOS:
Else
    MsgBox ("сто педио <<пкгяылема>> пяепеи ам сулпкгяыметаи, ма сулпкгяыметаи ломо ле тис кенеис <<маи>> ╧ <<ови>>. паяайакы диояхысте"), vbCritical, "пяосовг"
    index = 32
    GoTo TELOS:
End If

lathos:
If index = 32 Then
    GoTo TELOS:
Else
    MsgBox ("ам супкгяысате то педио аяих/тилокоц. тоте лгм сулкгяыметаи йамема акко педио"), vbCritical, "пяосовг"
    Text9.Text = ""
    Text10.Text = ""
    Text11.Text = ""
    Text16.Text = ""
    index = 32
    GoTo TELOS:
End If


er:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
ZForm6.Text1.Left = 6210
ZForm6.Text1.Width = 2800
ZForm6.Label2.Left = 4770
If rs1.STATE = 1 Then rs1.Close
If db1.STATE = 1 Then db1.Close
End Sub

Private Sub Command27_Click()
On Error GoTo er:
Text1.Text = Text24.Text

If rs1.STATE = 1 Then rs1.Close
If db1.STATE = 1 Then db1.Close
db1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & ZETAIRIES_DIADROMHS_BACKUP_DIAX_1 & ";" & _
"Persist Security Info=False"
db1.Open ZETAIRIES_DIADROMHS_BACKUP_DIAX
rs1.Open "[" & ZETAIRIES.Text1.Text & "]", db1, adOpenDynamic, adLockBatchOptimistic
Dim index, C, C1 As Integer
Dim STATEMENT1, STATEMENT2, STATEMENT3, STATEMENT4, STATEMENT5, STATEMENT6 As String
index = 1
C = 1
C1 = 0
If Command27.Caption = "еуяесг" Then
  If rs1.BOF = rs1.EOF Then GoTo NIK:
  rs1.MoveFirst
NIK:
  Do While Not rs1.EOF
    If rs1![аяихлос_тилокоциоу] <> Text20.Text Then
        rs1.MoveNext
    Else
        Text20.Text = rs1![аяихлос_тилокоциоу]
        Text19.Text = rs1![глеяолгмиа_ейдосгс]
        Text21.Text = rs1![посо]
        Text23.Text = rs1![аяихлос_тилокоциоу]
        rs1.MoveNext
        C = C + 1
   End If
  Loop
  If C <> 1 Then
    Command27.Caption = "диояхысг"
  Else
    MsgBox ("о аяихлос епитацгс поу дысате дем бяехгйе"), vbCritical, "пяосовг !!!"
  End If
Else  ' ***************************** DIORTHOSI ******************************
    ' ELEGXOS AN PERNAO TO PROBLEPOMENO MHKOS PEDIOY GIA TEXT20 POY EINAI 20
    ' KAI APOKOPEI PERITOY MEROYS AN XREIAZETAI
    Dim L20 As Integer
    L20 = Len(Text20.Text)
    If L20 > 20 Then
        Text2.Text = Mid(Text2.Text, 1, 20)
    Else
    
End If
    ' ELEGXOS LATHON
    If Text19.Text = "" Then GoTo HMER_1:
    If IsDate(Text19.Text) = False Then GoTo HMER_2:
    If Text20.Text = "" Then GoTo EPIT_1:
    If Text21.Text = "" Then GoTo POSO_1:
    If IsNumeric(Text21.Text) = False Then GoTo POSO_2:
    
    If rs1.BOF = rs1.EOF Then GoTo NIK1:
    rs1.MoveFirst
NIK1:
    Do While Not rs1.EOF
        If rs1![аяихлос_тилокоциоу] = Text20.Text Then
            C1 = C1 + 1
            rs1.MoveNext
        Else
            rs1.MoveNext
        End If
    Loop
    If C1 <> 0 Then
        If Text23.Text = Text20.Text Then
        
        Else
            MsgBox ("то моулеяо епитацгс поу дысате упаявеи гдг"), vbCritical, "пяосовг !!!"
            Command27.Caption = "еуяесг"
            Text20.Text = ""
            Text19.Text = ""
            Text21.Text = ""
            GoTo TELOS:
        End If
    Else
        
    End If
    
    ' PROGRAMATISMOS
    If MsgBox(("хекете ма пяовыягсете се диояхысг "), vbOKCancel, "") = vbOK Then
    
    STATEMENT1 = " UPDATE " & ZETAIRIES.Text1.Text & _
    " SET  аяихлос_тилокоциоу='" & Text20.Text & "'" & _
    " WHERE аяихлос_тилокоциоу ='" & Text23.Text & "'"

    STATEMENT2 = " UPDATE " & ZETAIRIES.Text1.Text & _
    " SET глеяолгмиа_ейдосгс ='" & Text19.Text & "'" & _
    " WHERE аяихлос_тилокоциоу ='" & Text23.Text & "'"

    STATEMENT3 = " UPDATE " & ZETAIRIES.Text1.Text & _
    " SET посо ='" & Text21.Text & "'" & _
    " WHERE аяихлос_тилокоциоу ='" & Text23.Text & "'"

    STATEMENT4 = " UPDATE " & ZETAIRIES.Text1.Text & _
    " SET пистысг ='" & Text21.Text & "'" & _
    " WHERE аяихлос_тилокоциоу ='" & Text23.Text & "'"
    
    STATEMENT5 = " UPDATE " & ZETAIRIES.Text1.Text & _
    " SET  аяихлос_епитацгс='" & Text20.Text & "'" & _
    " WHERE аяихлос_тилокоциоу ='" & Text23.Text & "'"
    
    STATEMENT6 = " UPDATE " & ZETAIRIES.Text1.Text & _
    " SET глеяолгмиа_еныжкгсгс ='" & Text19.Text & "'" & _
    " WHERE аяихлос_тилокоциоу ='" & Text23.Text & "'"
    
    db1.Execute STATEMENT6
    db1.Execute STATEMENT5
    db1.Execute STATEMENT3
    db1.Execute STATEMENT4
    db1.Execute STATEMENT2
    db1.Execute STATEMENT1
    
    Command27.Caption = "еуяесг"
    Text19.Text = ""
    Text20.Text = ""
    Text21.Text = ""
    MsgBox ("г диояхысг окойкгяыхгйе"), , ""
  Else
    Command27.Caption = "еуяесг"
    Text19.Text = ""
    Text20.Text = ""
    Text21.Text = ""
  End If
 End If
GoTo TELOS:
' ANTIMETOPISH LATHON
HMER_1:
MsgBox ("дем дысате глеяолгмиа"), vbCritical, "пяосовг !!!"
index = 32
GoTo TELOS:

HMER_2:
If index = 32 Then
GoTo TELOS:
Else
MsgBox ("дем дысате сыста тгм глеяылгмиа"), vbCritical, "пяосовг !!!"
index = 32
End If

EPIT_1:
If index = 32 Then
GoTo TELOS:
Else
MsgBox ("дем дысате аяихло епитацгс"), vbCritical, "пяосовг !!!"
index = 32
End If

POSO_1:
If index = 32 Then
GoTo TELOS:
Else
MsgBox ("дем дысате посо"), vbCritical, "пяосовг !!!"
index = 32
End If

POSO_2:
If index = 32 Then
GoTo TELOS:
Else
MsgBox ("дем дысате сыста то посо"), vbCritical, "пяосовг !!!"
index = 32
End If

er:
If index = 32 Then
GoTo TELOS:
Else
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"
End If

TELOS:
If rs1.STATE = 1 Then rs1.Close
If db1.STATE = 1 Then db1.Close
End Sub

Private Sub Command28_Click()
On Error GoTo er:
Text1.Text = Text24.Text
If rs1.STATE = 1 Then rs1.Close
If db1.STATE = 1 Then db1.Close
db1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & ZETAIRIES_DIADROMHS_BACKUP_DIAX_1 & ";" & _
"Persist Security Info=False"
db1.Open ZETAIRIES_DIADROMHS_BACKUP_DIAX
rs1.Open "[" & ZETAIRIES.Text1.Text & "]", db1, adOpenDynamic, adLockBatchOptimistic
Dim STATE As String
Dim C As Integer
C = 1
If rs1.BOF = rs1.EOF Then GoTo NIK:
rs1.MoveFirst
NIK:
  Do While Not rs1.EOF
    If rs1![аяихлос_тилокоциоу] <> UCase(Text20.Text) Then
        rs1.MoveNext
    Else
        C = C + 1
        rs1.MoveNext
    End If
  Loop
  If C <> 1 Then
   If MsgBox("хекете ма пяовыягсете се диацяажг тгс епитацгс", vbOKCancel, "пяосовг!!") = vbOK Then
    STATE = " DELETE FROM " & UCase(ZETAIRIES.Text1.Text) & _
    " WHERE аяихлос_тилокоциоу='" & UCase(Text20.Text) & "'"
    db1.Execute STATE
    MsgBox ("г диацяажг окойкгяыхгйе"), , "ой"
    End If
    Text19.Text = ""
    Text20.Text = ""
    Text21.Text = ""
  Else
    MsgBox ("дем бяехгйе епитацг ле то моулеяо поу дысате"), vbCritical, "пяосовг !!!"
    Text19.Text = ""
    Text20.Text = ""
    Text21.Text = ""
End If
GoTo TELOS:

er:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
If rs1.STATE = 1 Then rs1.Close
If db1.STATE = 1 Then db1.Close
End Sub

Private Sub Command29_Click()
Text1.Text = Text24.Text
On Error GoTo er:
Load ZForm8
ZForm8.Show
GoTo TELOS:

er:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
End Sub

Private Sub Command3_Click()
On Error GoTo TELOS:

If db1.STATE = 1 Then db1.Close
ZETAIRIES.Hide
Unload ZETAIRIES
Form11ETAIRIES.Enabled = True
Form11ETAIRIES.Show
GoTo TELOS:

er:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
End Sub

Private Sub Command30_Click()
On Error GoTo er:

Load ZForm3
ZForm3.Show
GoTo TELOS:

er:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
End Sub

Private Sub Command4_Click()
Text1.Text = Text24.Text
On Error GoTo er:

'If rs1.STATE = 1 Then rs1.Close
'If db1.STATE = 1 Then db1.Close
'db1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
'"Data Source=" & ZETAIRIES_DIADROMHS_BACKUP_DIAX_1 & ";" & _
'"Persist Security Info=False"
'db1.Open ZETAIRIES_DIADROMHS_BACKUP_DIAX
'rs1.Open "[" & ZETAIRIES.Text1.Text & "]", db1, adOpenDynamic, adLockBatchOptimistic

Dim ap, mexr, ap_temp, mexr_temp
Dim index As Integer
index = 1
'***************** ELEGXOS LATHON ****************************
If Text9.Text <> "" Then
    If IsDate(Text9.Text) = False Then GoTo imer_apo:
End If
If Text10.Text <> "" Then
    If IsDate(Text10.Text) = False Then GoTo imer_mexri:
End If
If (Text9.Text <> "") And (Text10.Text <> "") Then
    If ((CDate(Text9.Text)) > (CDate(Text10.Text))) Then GoTo imer:
End If
If Text10.Text <> "" Then
    If ((CDate(Text10.Text)) > Date) Then GoTo imer_2:
End If
If Text11.Text <> "" Then
    If IsNumeric(Text11.Text) = False Then GoTo POSO:
End If
If Text16.Text <> "" Then
    If ((Text16.Text = "маи") Or (Text16.Text = "ови")) Then

    Else
        GoTo EKSO:
    End If
End If
If (Text8.Text <> "") And ((Text9.Text <> "") Or (Text10.Text <> "") Or (Text11.Text <> "") Or (Text16.Text <> "")) Then GoTo lathos:
' LEITOYRGIA
' ANTOISTOIXHSH TIMHS STHN METABLHTH TYPOY DATE AP
If Text9.Text = "" Then
ap_temp = #1/1/2004#
Else
ap_temp = CDate(Text9.Text)
End If
' ANTOISTOIXHSH TIMHS STHN METABLHTH TYPOY DATE MEXR
If Text10.Text = "" Then
mexr_temp = Date
Else
mexr_temp = CDate(Text10.Text)
End If
' ftiaksimo telikon timon
If Day(ap_temp) < 12 Then
    ap = CDate(Month(ap_temp) & "/" & Day(ap_temp) & "/" & Year(ap_temp))
Else
    ap = ap_temp
End If
If Day(mexr_temp) < 12 Then
    mexr = CDate(Month(mexr_temp) & "/" & Day(mexr_temp) & "/" & Year(mexr_temp))
Else
    mexr = mexr_temp
End If

'*************** ORISMOS EROTHMATON *****************************

'--------------------------- 1 ---------------------------------
STATE1 = " select аяихлос_тилокоциоу,глеяолгмиа_ейдосгс,еныжкгсг,посо,глеяолгмиа_еныжкгсгс,аяихлос_епитацгс,тупос from " & UCase(ZETAIRIES.Text1.Text) & _
" WHERE тупос='тилок\пыкгсгс'" & _
" order by глеяолгмиа_ейдосгс"
SS1 = "SELECT SUM(посо) from " & UCase(ZETAIRIES.Text1.Text) & _
" WHERE тупос='тилок\пыкгсгс'"
FF1 = "SELECT COUNT(аяихлос_тилокоциоу) from " & UCase(ZETAIRIES.Text1.Text) & _
" WHERE тупос='тилок\пыкгсгс'"

'--------------------------- 2 ---------------------------------
STATE2 = "select аяихлос_тилокоциоу,глеяолгмиа_ейдосгс,еныжкгсг,посо,глеяолгмиа_еныжкгсгс,аяихлос_епитацгс,тупос from " & UCase(ZETAIRIES.Text1.Text) & _
 " where аяихлос_тилокоциоу ='" & UCase(ZETAIRIES.Text8.Text) & "'" & " AND тупос='тилок\пыкгсгс'" & _
 " order by глеяолгмиа_ейдосгс"
SS2 = "select SUM(посо) from " & UCase(ZETAIRIES.Text1.Text) & _
 " where аяихлос_тилокоциоу ='" & UCase(ZETAIRIES.Text8.Text) & "'" & " AND тупос='тилок\пыкгсгс'"
FF2 = "SELECT COUNT(аяихлос_тилокоциоу) from " & UCase(ZETAIRIES.Text1.Text) & _
 " where аяихлос_тилокоциоу ='" & UCase(ZETAIRIES.Text8.Text) & "'" & " AND тупос='тилок\пыкгсгс'"
 

'--------------------------- 3 ---------------------------------
STATE3 = "select аяихлос_тилокоциоу,глеяолгмиа_ейдосгс,еныжкгсг,посо,глеяолгмиа_еныжкгсгс,аяихлос_епитацгс,тупос from " & UCase(ZETAIRIES.Text1.Text) & _
 " where посо =" & ZETAIRIES.Text11.Text & " AND тупос='тилок\пыкгсгс'" & _
 " order by глеяолгмиа_ейдосгс"
SS3 = "select SUM(посо) from " & UCase(ZETAIRIES.Text1.Text) & _
 " where посо =" & ZETAIRIES.Text11.Text & " AND тупос='тилок\пыкгсгс'"
FF3 = "SELECT COUNT(аяихлос_тилокоциоу) from " & UCase(ZETAIRIES.Text1.Text) & _
 " where посо =" & ZETAIRIES.Text11.Text & " AND тупос='тилок\пыкгсгс'"
 
'--------------------------- 4 ---------------------------------
STATE4 = "select аяихлос_тилокоциоу,глеяолгмиа_ейдосгс,еныжкгсг,посо,глеяолгмиа_еныжкгсгс,аяихлос_епитацгс,тупос from " & UCase(ZETAIRIES.Text1.Text) & _
" where (глеяолгмиа_ейдосгс between " & "#" & ap & "#" & " and " & "#" & mexr & "#" & ")" & " AND тупос='тилок\пыкгсгс'" & _
" order by глеяолгмиа_ейдосгс"
SS4 = "select sum(посо) from " & UCase(ZETAIRIES.Text1.Text) & _
" where (глеяолгмиа_ейдосгс between " & "#" & ap & "#" & " and " & "#" & mexr & "#" & ")" & " AND тупос='тилок\пыкгсгс'"
FF4 = "SELECT COUNT(аяихлос_тилокоциоу) from " & UCase(ZETAIRIES.Text1.Text) & _
" where (глеяолгмиа_ейдосгс between " & "#" & ap & "#" & " and " & "#" & mexr & "#" & ")" & " AND тупос='тилок\пыкгсгс'"

'--------------------------- 5A ---------------------------------
STATE5A = "select аяихлос_тилокоциоу,глеяолгмиа_ейдосгс,еныжкгсг,посо,глеяолгмиа_еныжкгсгс,аяихлос_епитацгс,тупос from " & UCase(ZETAIRIES.Text1.Text) & _
 " where еныжкгсг=-1" & " AND тупос='тилок\пыкгсгс'" & _
" order by глеяолгмиа_ейдосгс"
SS5A = "select sum(посо) from " & UCase(ZETAIRIES.Text1.Text) & _
 " where еныжкгсг=-1" & " AND тупос='тилок\пыкгсгс'"
FF5A = "SELECT COUNT(аяихлос_тилокоциоу) from " & UCase(ZETAIRIES.Text1.Text) & _
 " where еныжкгсг=-1" & " AND тупос='тилок\пыкгсгс'"
'--------------------------- 5B ---------------------------------
STATE5B = "select аяихлос_тилокоциоу,глеяолгмиа_ейдосгс,еныжкгсг,посо,глеяолгмиа_еныжкгсгс,аяихлос_епитацгс,тупос from " & UCase(ZETAIRIES.Text1.Text) & _
 " where еныжкгсг=0" & " AND тупос='тилок\пыкгсгс'" & _
" order by глеяолгмиа_ейдосгс"
SS5B = "select sum(посо) from " & UCase(ZETAIRIES.Text1.Text) & _
 " where еныжкгсг=0" & " AND тупос='тилок\пыкгсгс'"
FF5B = "SELECT COUNT(аяихлос_тилокоциоу) from " & UCase(ZETAIRIES.Text1.Text) & _
 " where еныжкгсг=0" & " AND тупос='тилок\пыкгсгс'"
'--------------------------- 6 ---------------------------------
STATE6 = "select аяихлос_тилокоциоу,глеяолгмиа_ейдосгс,еныжкгсг,посо,глеяолгмиа_еныжкгсгс,аяихлос_епитацгс,тупос from " & UCase(ZETAIRIES.Text1.Text) & _
" where (глеяолгмиа_ейдосгс between " & "#" & ap & "#" & " and " & "#" & mexr & "#" & ") AND посо =" & ZETAIRIES.Text11.Text & " AND тупос='тилок\пыкгсгс'" & _
" order by глеяолгмиа_ейдосгс"
SS6 = "select sum(посо) from " & UCase(ZETAIRIES.Text1.Text) & _
" where (глеяолгмиа_ейдосгс between " & "#" & ap & "#" & " and " & "#" & mexr & "#" & ") AND посо =" & ZETAIRIES.Text11.Text & " AND тупос='тилок\пыкгсгс'"
FF6 = "SELECT COUNT(аяихлос_тилокоциоу) from " & UCase(ZETAIRIES.Text1.Text) & _
" where (глеяолгмиа_ейдосгс between " & "#" & ap & "#" & " and " & "#" & mexr & "#" & ") AND посо =" & ZETAIRIES.Text11.Text & " AND тупос='тилок\пыкгсгс'"

'--------------------------- 7A---------------------------------
STATE7A = "select аяихлос_тилокоциоу,глеяолгмиа_ейдосгс,еныжкгсг,посо,глеяолгмиа_еныжкгсгс,аяихлос_епитацгс,тупос from " & UCase(ZETAIRIES.Text1.Text) & _
" where еныжкгсг=-1 AND посо=" & ZETAIRIES.Text11.Text & " AND тупос='тилок\пыкгсгс'" & _
" order by глеяолгмиа_ейдосгс"
SS7A = "select sum(посо) from " & UCase(ZETAIRIES.Text1.Text) & _
" where еныжкгсг=-1 AND посо=" & ZETAIRIES.Text11.Text & " AND тупос='тилок\пыкгсгс'"
FF7A = "SELECT COUNT(аяихлос_тилокоциоу) from " & UCase(ZETAIRIES.Text1.Text) & _
" where еныжкгсг=-1 AND посо=" & ZETAIRIES.Text11.Text & " AND тупос='тилок\пыкгсгс'"


'--------------------------- 7B ---------------------------------
STATE7B = "select аяихлос_тилокоциоу,глеяолгмиа_ейдосгс,еныжкгсг,посо,глеяолгмиа_еныжкгсгс,аяихлос_епитацгс,тупос from " & UCase(ZETAIRIES.Text1.Text) & _
" where еныжкгсг=0 AND посо=" & ZETAIRIES.Text11.Text & " AND тупос='тилок\пыкгсгс'" & _
" order by глеяолгмиа_ейдосгс"
SS7B = "select sum(посо) from " & UCase(ZETAIRIES.Text1.Text) & _
" where еныжкгсг=0 AND посо=" & ZETAIRIES.Text11.Text & " AND тупос='тилок\пыкгсгс'"
FF7B = "SELECT COUNT(аяихлос_тилокоциоу) from " & UCase(ZETAIRIES.Text1.Text) & _
" where еныжкгсг=0 AND посо=" & ZETAIRIES.Text11.Text & " AND тупос='тилок\пыкгсгс'"


'--------------------------- 8A ---------------------------------
STATE8A = "select аяихлос_тилокоциоу,глеяолгмиа_ейдосгс,еныжкгсг,посо,глеяолгмиа_еныжкгсгс,аяихлос_епитацгс,тупос from " & UCase(ZETAIRIES.Text1.Text) & _
" where (глеяолгмиа_ейдосгс between " & "#" & ap & "#" & " and " & "#" & mexr & "#" & ") AND еныжкгсг=-1 " & " AND тупос='тилок\пыкгсгс'" & _
" order by глеяолгмиа_ейдосгс"
SS8A = "select sum(посо) from " & UCase(ZETAIRIES.Text1.Text) & _
" where (глеяолгмиа_ейдосгс between " & "#" & ap & "#" & " and " & "#" & mexr & "#" & ") AND еныжкгсг=-1 " & " AND тупос='тилок\пыкгсгс'"
FF8A = "SELECT COUNT(аяихлос_тилокоциоу) from " & UCase(ZETAIRIES.Text1.Text) & _
" where (глеяолгмиа_ейдосгс between " & "#" & ap & "#" & " and " & "#" & mexr & "#" & ") AND еныжкгсг=-1 " & " AND тупос='тилок\пыкгсгс'"

'--------------------------- 8B ---------------------------------
STATE8B = "select аяихлос_тилокоциоу,глеяолгмиа_ейдосгс,еныжкгсг,посо,глеяолгмиа_еныжкгсгс,аяихлос_епитацгс,тупос from " & UCase(ZETAIRIES.Text1.Text) & _
" where (глеяолгмиа_ейдосгс between " & "#" & ap & "#" & " and " & "#" & mexr & "#" & ") AND еныжкгсг=0 " & " AND тупос='тилок\пыкгсгс'" & _
" order by глеяолгмиа_ейдосгс"
SS8B = "select sum(посо) from " & UCase(ZETAIRIES.Text1.Text) & _
" where (глеяолгмиа_ейдосгс between " & "#" & ap & "#" & " and " & "#" & mexr & "#" & ") AND еныжкгсг=0 " & " AND тупос='тилок\пыкгсгс'"
FF8B = "SELECT COUNT(аяихлос_тилокоциоу) from " & UCase(ZETAIRIES.Text1.Text) & _
" where (глеяолгмиа_ейдосгс between " & "#" & ap & "#" & " and " & "#" & mexr & "#" & ") AND еныжкгсг=0 " & " AND тупос='тилок\пыкгсгс'"

'--------------------------- 9A---------------------------------
STATE9A = "select аяихлос_тилокоциоу,глеяолгмиа_ейдосгс,еныжкгсг,посо,глеяолгмиа_еныжкгсгс,аяихлос_епитацгс,тупос from " & UCase(ZETAIRIES.Text1.Text) & _
" where (глеяолгмиа_ейдосгс between " & "#" & ap & "#" & " and " & "#" & mexr & "#" & ") AND еныжкгсг=-1 AND посо =" & ZETAIRIES.Text11.Text & " AND тупос='тилок\пыкгсгс'" & _
" order by глеяолгмиа_ейдосгс"
SS9A = "select sum(посо) from " & UCase(ZETAIRIES.Text1.Text) & _
" where (глеяолгмиа_ейдосгс between " & "#" & ap & "#" & " and " & "#" & mexr & "#" & ") AND еныжкгсг=-1 AND посо =" & ZETAIRIES.Text11.Text & " AND тупос='тилок\пыкгсгс'"
FF9A = "SELECT COUNT(аяихлос_тилокоциоу) from " & UCase(ZETAIRIES.Text1.Text) & _
" where (глеяолгмиа_ейдосгс between " & "#" & ap & "#" & " and " & "#" & mexr & "#" & ") AND еныжкгсг=-1 AND посо =" & ZETAIRIES.Text11.Text & " AND тупос='тилок\пыкгсгс'"


'--------------------------- 9B ---------------------------------
STATE9B = "select аяихлос_тилокоциоу,глеяолгмиа_ейдосгс,еныжкгсг,посо,глеяолгмиа_еныжкгсгс,аяихлос_епитацгс,тупос from " & UCase(ZETAIRIES.Text1.Text) & _
" where (глеяолгмиа_ейдосгс between " & "#" & ap & "#" & " and " & "#" & mexr & "#" & ") AND еныжкгсг=0 AND посо =" & ZETAIRIES.Text11.Text & " AND тупос='тилок\пыкгсгс'" & _
" order by глеяолгмиа_ейдосгс"
SS9B = "select sum(посо) from " & UCase(ZETAIRIES.Text1.Text) & _
" where (глеяолгмиа_ейдосгс between " & "#" & ap & "#" & " and " & "#" & mexr & "#" & ") AND еныжкгсг=0 AND посо =" & ZETAIRIES.Text11.Text & " AND тупос='тилок\пыкгсгс'"
FF9B = "SELECT COUNT(аяихлос_тилокоциоу) from " & UCase(ZETAIRIES.Text1.Text) & _
" where (глеяолгмиа_ейдосгс between " & "#" & ap & "#" & " and " & "#" & mexr & "#" & ") AND еныжкгсг=0 AND посо =" & ZETAIRIES.Text11.Text & " AND тупос='тилок\пыкгсгс'"


' DIAXORISMOS PERIPTOSEON KAI ANTOISTOIXHSH TIMHS STHN STATE
'1) OLA KENA
If ((Text8.Text = "") And (Text9.Text = "") And (Text10.Text = "") _
And (Text11.Text = "") And (Text16.Text = "")) Then
    STATE = STATE1
    SS = SS1
    FF = FF1
End If
'2) MONO AR TIMOLOGIOY
If Text8.Text <> "" Then
    STATE = STATE2
    SS = SS2
    FF = FF2
End If
'3) MONO POSO
If Text11.Text <> "" Then
    STATE = STATE3
    SS = SS3
    FF = FF3
End If
'4) APO - MEXRI
If ((Text9.Text <> "") Or (Text10.Text <> "")) Then
    STATE = STATE4
    SS = SS4
    FF = FF4
End If
'5A) PLHR 'H OXI
If Text16.Text = "маи" Then
    STATE = STATE5A
    SS = SS5A
    FF = FF5A
End If
'5б) PLHR 'H OXI
If Text16.Text = "ови" Then
    STATE = STATE5B
    SS = SS5B
    FF = FF5B
End If
'6) POSO & APO - MEXRI
If ((Text11.Text <> "") And ((Text9.Text <> "") Or (Text10.Text <> ""))) _
Then
    STATE = STATE6
    SS = SS6
    FF = FF6
End If
'7A) POSO & "NAI" STO PLHROMENA 'H OXI
If ((Text11.Text <> "") And (Text16.Text = "маи")) Then
    STATE = STATE7A
    SS = SS7A
    FF = FF7A
End If
'7B) POSO & "OXI" STO PLHROMENA 'H OXI
If ((Text11.Text <> "") And (Text16.Text = "ови")) Then
    STATE = STATE7B
    SS = SS7B
    FF = FF7B
End If
'8A) APO MEXRI & PLHROMENA="NAI"
If (((Text9.Text <> "") Or (Text10.Text <> "")) And (Text16.Text = "маи")) Then
    STATE = STATE8A
    SS = SS8A
    FF = FF8A
End If
'8B) APO MEXRI & PLHROMENA="OXI"
If (((Text9.Text <> "") Or (Text10.Text <> "")) And (Text16.Text = "ови")) Then
    STATE = STATE8B
    SS = SS8B
    FF = FF8B
End If
'9A) KAI TA TRIA ME PLIROMENA NAI
If ((Text11.Text <> "") And ((Text9.Text <> "") Or (Text10.Text <> "")) And (Text16.Text = "маи")) Then
    STATE = STATE9A
    SS = SS9A
    FF = FF9A
End If
'9B) KAI TA TRIA ME PLIROMENA OXI
If ((Text11.Text <> "") And ((Text9.Text <> "") Or (Text10.Text <> "")) And (Text16.Text = "ови")) Then
    STATE = STATE9B
    SS = SS9B
    FF = FF9B
End If

FLAG_FORM6 = 1
PATIMA_DATAGRID_F6 = 0
Load ZForm6
ZForm6.Show
GoTo TELOS:

' ANTIMETOPISI LATHON
imer_apo:
MsgBox ("дем дысате сыста тгм глеяолгмиа сто пкаисио <<апо>>"), vbCritical, "пяосовг"
index = 32
GoTo TELOS:

imer_mexri:
If index = 32 Then
    GoTo TELOS:
Else
    MsgBox ("дем дысате сыста тгм глеяолгмиа сто пкаисио <<левяи>>"), vbCritical, "пяосовг"
    index = 32
    GoTo TELOS:
End If

imer:
If index = 32 Then
    GoTo TELOS:
Else
    MsgBox ("г глеяолгмиа <<апо>> еимаи лецакутеяг апо тгм <<левяи>>"), vbCritical, "пяосовг"
    index = 32
    GoTo TELOS:
End If

imer_2:
If index = 32 Then
    GoTo TELOS:
Else
    MsgBox ("г глеяолгмиа <<левяи>> еимаи лецакутеяг апо тгм сглеяимг"), vbCritical, "пяосовг"
    index = 32
    GoTo TELOS:
End If

POSO:
If index = 32 Then
    GoTo TELOS:
Else
    MsgBox ("дем дысате сыста то посо"), vbCritical, "пяосовг"
    index = 32
    GoTo TELOS:
End If

EKSO:
If index = 32 Then
    GoTo TELOS:
Else
    MsgBox ("сто педио <<пкгяылема>> пяепеи ам сулпкгяыметаи, ма сулпкгяыметаи ломо ле тис кенеис <<маи>> ╧ <<ови>>. паяайакы диояхысте"), vbCritical, "пяосовг"
    index = 32
    GoTo TELOS:
End If

lathos:
If index = 32 Then
    GoTo TELOS:
Else
    MsgBox ("ам супкгяысате то педио аяих/тилокоц. тоте лгм сулкгяыметаи йамема акко педио"), vbCritical, "пяосовг"
    Text9.Text = ""
    Text10.Text = ""
    Text11.Text = ""
    Text16.Text = ""
    index = 32
    GoTo TELOS:
End If

er:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
ZForm6.Text1.Left = 6605
ZForm6.Label2.Left = 5165
ZForm6.Text1.Width = 2020
'If rs1.STATE = 1 Then rs1.Close
'If db1.STATE = 1 Then db1.Close
End Sub

Private Sub Command5_Click()
Text1.Text = Text24.Text
On Error GoTo er:
Dim C As Integer
C = 0

If rs1.STATE = 1 Then rs1.Close
If db1.STATE = 1 Then db1.Close
db1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & ZETAIRIES_DIADROMHS_BACKUP_DIAX_1 & ";" & _
"Persist Security Info=False"
db1.Open ZETAIRIES_DIADROMHS_BACKUP_DIAX
rs1.Open "[" & ZETAIRIES.Text1.Text & "]", db1, adOpenDynamic, adLockBatchOptimistic

Dim k1, k2
Dim Y As String
Y = ETOSBACKUP
If Text17.Text = "" Then
    
Else
    If IsDate(Text17.Text) = False Then
        MsgBox ("дем дысате сыста тгм глеяолгмиа <<апо>>"), vbCritical, "пяосовг !!!"
        GoTo TELOS:
    End If
End If

If Text18.Text = "" Then
    
Else
    If IsDate(Text18.Text) = False Then
        MsgBox ("дем дысате сыста тгм глеяолгмиа <<левяи>>"), vbCritical, "пяосовг !!!"
        GoTo TELOS:
    End If
End If

'**RITHMISI HMEROMHNIAS APO(HAK) ***********************

If Text17.Text = "" Then
    k1 = CDate("1/1/" & Y)
Else
    k1 = CDate(Text17.Text)
End If
If Day(k1) < 12 Then
    HAK = CDate(Month(k1) & " / " & Day(k1) & " / " & Year(k1))
Else
    HAK = k1
End If

'**RITHMISI HMEROMHNIAS MEXRI(HMK) ***********************
If Text18.Text = "" Then
    k2 = CDate("31/12/" & Y)
Else
    k2 = CDate(Text18.Text)
End If
If Day(k2) < 12 Then
    HMK = CDate(Month(k2) & " / " & Day(k2) & " / " & Year(k2))
Else
    HMK = k2
End If

' ****KATARXHN DIAPISTONO AN H KARTELA EINAI KENH AN EXEI DHLADH PERASTH KATI********
If rs1.BOF = rs1.EOF Then GoTo NIK:
rs1.MoveNext
NIK:
Do While Not rs1.EOF
    C = C + 1
    rs1.MoveNext
Loop

'AN H KARTELA EINAI KENH
If C = 0 Then
    MsgBox ("дем упаявеи типота циа пяобокг"), vbCritical, "пяосовг !!!"
    GoTo TELOS:
Else ' ALLIOS
   STATE_KARTELAS = "SELECT глеяолгмиа_ейдосгс,аяихлос_тилокоциоу,тупос,вяеысг,пистысг,упокоипо FROM HELP_KARTELAS " & _
    " where (глеяолгмиа_ейдосгс between " & "#" & HAK & "#" & " and " & "#" & HMK & "#" & ")"
    '" order by глеяолгмиа_ейдосгс,тупос DESC"
    Load ZForm7
    ZForm7.Show
End If

GoTo TELOS:

er:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
End Sub

Private Sub Command6_Click()
On Error GoTo er:

Dim HM_APO_KARTELA As String
Dim DATE_HM_APO_KARTELA As Date
'***************** ELEGXOI **************************************
If IsNumeric(Combo21.Text) = False Then
    MsgBox ("дем дысате сыста глеяа"), vbCritical, "пяосовг !!!"
    GoTo TELOS:
End If
If IsNumeric(Combo22.Text) = False Then
    MsgBox ("дем дысате сыста лгма"), vbCritical, "пяосовг !!!"
    GoTo TELOS:
End If
If IsNumeric(Combo23.Text) = False Then
    MsgBox ("дем дысате сыста етос"), vbCritical, "пяосовг !!!"
    GoTo TELOS:
End If
If CInt(Combo21.Text) < 1 Or CInt(Combo21.Text) > 31 Then
    MsgBox ("о аяихлос поу дысате дем еимаи ецйуяг глеяа лгма"), vbCritical, "пяосовг !!!"
    GoTo TELOS:
End If
If CInt(Combo22.Text) < 1 Or CInt(Combo22.Text) > 12 Then
    MsgBox ("о аяихлос поу дысате дем еимаи ецйуяос лгмас етоус"), vbCritical, "пяосовг !!!"
    GoTo TELOS:
End If
If CInt(Combo23.Text) < 2005 Or CInt(Combo23.Text) > 2020 Then
    MsgBox ("то пяоцяалла упостгяифеи глеяолгмиес апо 2005 еыс 2020.паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"
    GoTo TELOS:
End If
'*************** LEITOYRGIA *******************************
HM_APO_KARTELA = DAY_HM_APO_KARTELA & "/" & MONTH_HM_APO_KARTELA & _
"/" & ETOS_HM_APO_KARTELA

If IsDate(HM_APO_KARTELA) = True Then
DATE_HM_APO_KARTELA = CDate(HM_APO_KARTELA)
Text17.Text = DATE_HM_APO_KARTELA
Else
MsgBox ("дем дысате глеяолгмиа"), vbCritical, "пяосовг!!"
End If
GoTo TELOS:
'***********************************************************
er:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
End Sub

Private Sub Command7_Click()
On Error GoTo er:

Text17.Text = Date
GoTo TELOS:

er:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
End Sub

Private Sub Command8_Click()
If rs1.STATE = 1 Then rs1.Close
If db1.STATE = 1 Then db1.Close
db1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
       "Data Source=" & "\databases\ETAIRIES.mdb" & ";" & _
      "Persist Security Info=False"
db1.Open App.Path & "\databases\ETAIRIES.mdb"
rs1.Open "[" & ZETAIRIES.Text1.Text & "]", db1, adOpenDynamic, adLockBatchOptimistic
Dim C As Integer
Dim STATEMENT, statement_B As String
Dim ar_timo As String
Dim imer_ekd As Date
Dim typos As String
Dim POSO, index As Integer
C = 1
index = 1
If Text13.Text = "" Then GoTo AR_EP:
If Text14.Text = "" Then GoTo ekdo_1:
If IsDate(Text14.Text) = False Then GoTo ekdo_2:
If Text15.Text = "" Then GoTo EPIT:

rs1.MoveFirst
Do While Not rs1.EOF
 If rs1![аяихлос_тилокоциоу] <> UCase(ZETAIRIES.Text13.Text) Then
    rs1.MoveNext
Else
    If rs1![еныжкгсг] = 0 Then
        ar_timo = rs1![аяихлос_тилокоциоу]
        imer_ekd = rs1![глеяолгмиа_ейдосгс]
        typos = rs1![тупос]
        POSO = rs1![посо]
        STATEMENT = " delete from " & UCase(Text1.Text) & _
        " where аяихлос_тилокоциоу='" & UCase(ar_timo) & "'"
        db1.Execute STATEMENT
        
        statement_B = "INSERT INTO " & UCase(ZETAIRIES.Text1.Text) & " (" & _
    "аяихлос_тилокоциоу,тупос,глеяолгмиа_ейдосгс," & _
    "еныжкгсг,посо," & _
    "глеяолгмиа_еныжкгсгс,аяихлос_епитацгс," & _
    "вяеысг,пистысг,упокоипо)" & _
    "VALUES (" & _
        "'" & UCase(ar_timo) & "'," & _
        "'" & UCase(typos) & "', " & _
        "'" & imer_ekd & "'," & _
        "'1'," & _
        "'" & POSO & "'," & _
        "'" & ZETAIRIES.Text14.Text & "'," & _
        "'" & ZETAIRIES.Text15.Text & "'," & _
        "'" & POSO & "'," & _
        "'0'," & _
        "'0'" & _
        ")"
        db1.Execute statement_B
        C = C + 1
        rs1.MoveNext
    Else
        MsgBox ("то тилокоцио еимаи гдг пкгяылемо")
        ZETAIRIES.Text13.Text = ""
        ZETAIRIES.Text14.Text = ""
        ZETAIRIES.Text15.Text = ""
        GoTo TELOS:
    End If
 End If
Loop
ZETAIRIES.Text13.Text = ""
ZETAIRIES.Text14.Text = ""
ZETAIRIES.Text15.Text = ""
ZETAIRIES.Combo11.Text = "глеяа"
ZETAIRIES.Combo12.Text = "лгмас"
ZETAIRIES.Combo13.Text = "етос"
If C = 1 Then
    MsgBox ("дем бяехгйе тилокоцио ле том аяихло поу дысате"), vbCritical, "пяосовг!!"
Else
    MsgBox ("то тилокоцио ле том аяихло поу дысате йатавыягхгйе ыс пкгяылемо"), , "пяосовг!!"
End If
GoTo TELOS:

AR_EP:
MsgBox ("дем дысате аяихло тилокоциоу"), vbCritical, "пяосовг"
index = 32

ekdo_1:
If index = 32 Then
    GoTo TELOS:
Else
    MsgBox ("дем дысате глеяолгмиа еныжкгсгс"), vbCritical, "пяосовг"
    index = 32
End If

ekdo_2:
If index = 32 Then
    GoTo TELOS:
Else
    MsgBox ("дем дысате сыста тгм глеяолгмиа еныжкгсгс"), vbCritical, "пяосовг"
    index = 32
End If

EPIT:
If index = 32 Then
    GoTo TELOS:
Else
    MsgBox ("дем дысате тяопо пкгяылгс"), vbCritical, "пяосовг"
    index = 32
End If

TELOS:
If rs1.STATE = 1 Then rs1.Close
If db1.STATE = 1 Then db1.Close
End Sub

Private Sub Command9_Click()
On Error GoTo er:
Dim C, index As Integer
C = 1
index = 1
If db1.STATE = 1 Then db1.Close
If rs1.STATE = 1 Then rs1.Close
db1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
       "Data Source=" & ZETAIRIES_DIADROMHS_BACKUP_DIAX_1 & ";" & _
      "Persist Security Info=False"
db1.Open ZETAIRIES_DIADROMHS_BACKUP_DIAX
rs1.Open "[ONOMATA_ETAIRION_ABCDEF]", db1, adOpenDynamic, adLockBatchOptimistic
If rs1.BOF = rs1.EOF Then GoTo NIK:
rs1.MoveFirst
NIK:
Do While Not rs1.EOF
    If rs1![омолата_етаияиым] = UCase(Text1.Text) Then
        C = C + 1
        rs1.MoveNext
    Else
        rs1.MoveNext
    End If
Loop
'**********************ELEGXOI**********************************
If Text1.Text = "" Then GoTo KENO:
If C = 1 Then GoTo ER_NAME:

Command1.Enabled = True
Command2.Enabled = True
Command4.Enabled = True
Command5.Enabled = True
Command6.Enabled = True
Command7.Enabled = True
Command8.Enabled = True
Command10.Enabled = True
Command23.Enabled = True
Command24.Enabled = True
Command24.Caption = "еуяесг"
Command26.Enabled = True
Command25.Enabled = True
Command27.Enabled = True
Command28.Enabled = True
Command29.Enabled = True
'Image1.Visible = True
Text1.Text = UCase(Text1.Text)
Text24.Text = Text1.Text

GoTo TELOS:

KENO:
MsgBox ("дем дысате йамема омола етаияиас"), vbCritical, "пяосовг !!!"
index = 32
GoTo TELOS:

ER_NAME:
If index = 32 Then
    GoTo TELOS:
Else
    MsgBox ("то омола етаияиас поу дысате дем упаявеи"), vbCritical, "пяосовг !!!"
    GoTo TELOS:
End If

er:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), , "пяосовг!!!"

TELOS:
If db1.STATE = 1 Then db1.Close
If rs1.STATE = 1 Then rs1.Close
End Sub

Private Sub DTPicker1_Change()
Text19.Text = DTPicker1.Value
End Sub

Private Sub DTPicker1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Text19.Text = "" Then Text19.Text = Date
End Sub

Private Sub Form_Load()
CT = 0
DTPicker1.Value = Date
Command24.Enabled = False
Command1.Enabled = False
Command2.Enabled = False
Command4.Enabled = False
Command8.Enabled = False
Command5.Enabled = False
Command23.Enabled = False
Command26.Enabled = False
Command25.Enabled = False
Command27.Enabled = False
Command28.Enabled = False
Command29.Enabled = False

Combo1.AddItem "тилок\пыкгсгс"
Combo1.AddItem "пистытийо"
Combo2.AddItem "ови"
Combo2.AddItem "маи"
Combo3.AddItem "летягта"
Combo3.AddItem "епитацг"
Combo4.AddItem "летягта"
Combo4.AddItem "епитацг"
'
Combo5.AddItem "1"
Combo5.AddItem "2"
Combo5.AddItem "3"
Combo5.AddItem "4"
Combo5.AddItem "5"
Combo5.AddItem "6"
Combo5.AddItem "7"
Combo5.AddItem "8"
Combo5.AddItem "9"
Combo5.AddItem "10"
Combo5.AddItem "11"
Combo5.AddItem "12"
Combo5.AddItem "13"
Combo5.AddItem "14"
Combo5.AddItem "15"
Combo5.AddItem "16"
Combo5.AddItem "17"
Combo5.AddItem "18"
Combo5.AddItem "19"
Combo5.AddItem "20"
Combo5.AddItem "21"
Combo5.AddItem "22"
Combo5.AddItem "23"
Combo5.AddItem "24"
Combo5.AddItem "25"
Combo5.AddItem "26"
Combo5.AddItem "27"
Combo5.AddItem "28"
Combo5.AddItem "29"
Combo5.AddItem "30"
Combo5.AddItem "31"

Combo6.AddItem "1"
Combo6.AddItem "2"
Combo6.AddItem "3"
Combo6.AddItem "4"
Combo6.AddItem "5"
Combo6.AddItem "6"
Combo6.AddItem "7"
Combo6.AddItem "8"
Combo6.AddItem "9"
Combo6.AddItem "10"
Combo6.AddItem "11"
Combo6.AddItem "12"

Combo7.AddItem "2005"
Combo7.AddItem "2006"
Combo7.AddItem "2007"
Combo7.AddItem "2008"
Combo7.AddItem "2009"
Combo7.AddItem "2010"
Combo7.AddItem "2011"
Combo7.AddItem "2012"
Combo7.AddItem "2013"
Combo7.AddItem "2014"
Combo7.AddItem "2015"
Combo7.AddItem "2016"
Combo7.AddItem "2017"
Combo7.AddItem "2018"
Combo7.AddItem "2019"
Combo7.AddItem "2020"

Combo8.AddItem "1"
Combo8.AddItem "2"
Combo8.AddItem "3"
Combo8.AddItem "4"
Combo8.AddItem "5"
Combo8.AddItem "6"
Combo8.AddItem "7"
Combo8.AddItem "8"
Combo8.AddItem "9"
Combo8.AddItem "10"
Combo8.AddItem "11"
Combo8.AddItem "12"
Combo8.AddItem "13"
Combo8.AddItem "14"
Combo8.AddItem "15"
Combo8.AddItem "16"
Combo8.AddItem "17"
Combo8.AddItem "18"
Combo8.AddItem "19"
Combo8.AddItem "20"
Combo8.AddItem "21"
Combo8.AddItem "22"
Combo8.AddItem "23"
Combo8.AddItem "24"
Combo8.AddItem "25"
Combo8.AddItem "26"
Combo8.AddItem "27"
Combo8.AddItem "28"
Combo8.AddItem "29"
Combo8.AddItem "30"
Combo8.AddItem "31"

Combo9.AddItem "1"
Combo9.AddItem "2"
Combo9.AddItem "3"
Combo9.AddItem "4"
Combo9.AddItem "5"
Combo9.AddItem "6"
Combo9.AddItem "7"
Combo9.AddItem "8"
Combo9.AddItem "9"
Combo9.AddItem "10"
Combo9.AddItem "11"
Combo9.AddItem "12"

Combo10.AddItem "2005"
Combo10.AddItem "2006"
Combo10.AddItem "2007"
Combo10.AddItem "2008"
Combo10.AddItem "2009"
Combo10.AddItem "2010"
Combo10.AddItem "2011"
Combo10.AddItem "2012"
Combo10.AddItem "2013"
Combo10.AddItem "2014"
Combo10.AddItem "2015"
Combo10.AddItem "2016"
Combo10.AddItem "2017"
Combo10.AddItem "2018"
Combo10.AddItem "2019"
Combo10.AddItem "2020"

Combo11.AddItem "1"
Combo11.AddItem "2"
Combo11.AddItem "3"
Combo11.AddItem "4"
Combo11.AddItem "5"
Combo11.AddItem "6"
Combo11.AddItem "7"
Combo11.AddItem "8"
Combo11.AddItem "9"
Combo11.AddItem "10"
Combo11.AddItem "11"
Combo11.AddItem "12"
Combo11.AddItem "13"
Combo11.AddItem "14"
Combo11.AddItem "15"
Combo11.AddItem "16"
Combo11.AddItem "17"
Combo11.AddItem "18"
Combo11.AddItem "19"
Combo11.AddItem "20"
Combo11.AddItem "21"
Combo11.AddItem "22"
Combo11.AddItem "23"
Combo11.AddItem "24"
Combo11.AddItem "25"
Combo11.AddItem "26"
Combo11.AddItem "27"
Combo11.AddItem "28"
Combo11.AddItem "29"
Combo11.AddItem "30"
Combo11.AddItem "31"

Combo12.AddItem "1"
Combo12.AddItem "2"
Combo12.AddItem "3"
Combo12.AddItem "4"
Combo12.AddItem "5"
Combo12.AddItem "6"
Combo12.AddItem "7"
Combo12.AddItem "8"
Combo12.AddItem "9"
Combo12.AddItem "10"
Combo12.AddItem "11"
Combo12.AddItem "12"

Combo13.AddItem "2005"
Combo13.AddItem "2006"
Combo13.AddItem "2007"
Combo13.AddItem "2008"
Combo13.AddItem "2009"
Combo13.AddItem "2010"
Combo13.AddItem "2011"
Combo13.AddItem "2012"
Combo13.AddItem "2013"
Combo13.AddItem "2014"
Combo13.AddItem "2015"
Combo13.AddItem "2016"
Combo13.AddItem "2017"
Combo13.AddItem "2018"
Combo13.AddItem "2019"
Combo13.AddItem "2020"

Combo14.AddItem "1"
Combo14.AddItem "2"
Combo14.AddItem "3"
Combo14.AddItem "4"
Combo14.AddItem "5"
Combo14.AddItem "6"
Combo14.AddItem "7"
Combo14.AddItem "8"
Combo14.AddItem "9"
Combo14.AddItem "10"
Combo14.AddItem "11"
Combo14.AddItem "12"
Combo14.AddItem "13"
Combo14.AddItem "14"
Combo14.AddItem "15"
Combo14.AddItem "16"
Combo14.AddItem "17"
Combo14.AddItem "18"
Combo14.AddItem "19"
Combo14.AddItem "20"
Combo14.AddItem "21"
Combo14.AddItem "22"
Combo14.AddItem "23"
Combo14.AddItem "24"
Combo14.AddItem "25"
Combo14.AddItem "26"
Combo14.AddItem "27"
Combo14.AddItem "28"
Combo14.AddItem "29"
Combo14.AddItem "30"
Combo14.AddItem "31"

Combo15.AddItem "1"
Combo15.AddItem "2"
Combo15.AddItem "3"
Combo15.AddItem "4"
Combo15.AddItem "5"
Combo15.AddItem "6"
Combo15.AddItem "7"
Combo15.AddItem "8"
Combo15.AddItem "9"
Combo15.AddItem "10"
Combo15.AddItem "11"
Combo15.AddItem "12"

Combo16.AddItem "2005"
Combo16.AddItem "2006"
Combo16.AddItem "2007"
Combo16.AddItem "2008"
Combo16.AddItem "2009"
Combo16.AddItem "2010"
Combo16.AddItem "2011"
Combo16.AddItem "2012"
Combo16.AddItem "2013"
Combo16.AddItem "2014"
Combo16.AddItem "2015"
Combo16.AddItem "2016"
Combo16.AddItem "2017"
Combo16.AddItem "2018"
Combo16.AddItem "2019"
Combo16.AddItem "2020"

Combo17.AddItem "1"
Combo17.AddItem "2"
Combo17.AddItem "3"
Combo17.AddItem "4"
Combo17.AddItem "5"
Combo17.AddItem "6"
Combo17.AddItem "7"
Combo17.AddItem "8"
Combo17.AddItem "9"
Combo17.AddItem "10"
Combo17.AddItem "11"
Combo17.AddItem "12"
Combo17.AddItem "13"
Combo17.AddItem "14"
Combo17.AddItem "15"
Combo17.AddItem "16"
Combo17.AddItem "17"
Combo17.AddItem "18"
Combo17.AddItem "19"
Combo17.AddItem "20"
Combo17.AddItem "21"
Combo17.AddItem "22"
Combo17.AddItem "23"
Combo17.AddItem "24"
Combo17.AddItem "25"
Combo17.AddItem "26"
Combo17.AddItem "27"
Combo17.AddItem "28"
Combo17.AddItem "29"
Combo17.AddItem "30"
Combo17.AddItem "31"

Combo18.AddItem "1"
Combo18.AddItem "2"
Combo18.AddItem "3"
Combo18.AddItem "4"
Combo18.AddItem "5"
Combo18.AddItem "6"
Combo18.AddItem "7"
Combo18.AddItem "8"
Combo18.AddItem "9"
Combo18.AddItem "10"
Combo18.AddItem "11"
Combo18.AddItem "12"

Combo19.AddItem "2005"
Combo19.AddItem "2006"
Combo19.AddItem "2007"
Combo19.AddItem "2008"
Combo19.AddItem "2009"
Combo19.AddItem "2010"
Combo19.AddItem "2011"
Combo19.AddItem "2012"
Combo19.AddItem "2013"
Combo19.AddItem "2014"
Combo19.AddItem "2015"
Combo19.AddItem "2016"
Combo19.AddItem "2017"
Combo19.AddItem "2018"
Combo19.AddItem "2019"
Combo19.AddItem "2020"

Combo20.AddItem "маи"
Combo20.AddItem "ови"

Combo21.AddItem "1"
Combo21.AddItem "2"
Combo21.AddItem "3"
Combo21.AddItem "4"
Combo21.AddItem "5"
Combo21.AddItem "6"
Combo21.AddItem "7"
Combo21.AddItem "8"
Combo21.AddItem "9"
Combo21.AddItem "10"
Combo21.AddItem "11"
Combo21.AddItem "12"
Combo21.AddItem "13"
Combo21.AddItem "14"
Combo21.AddItem "15"
Combo21.AddItem "16"
Combo21.AddItem "17"
Combo21.AddItem "18"
Combo21.AddItem "19"
Combo21.AddItem "20"
Combo21.AddItem "21"
Combo21.AddItem "22"
Combo21.AddItem "23"
Combo21.AddItem "24"
Combo21.AddItem "25"
Combo21.AddItem "26"
Combo21.AddItem "27"
Combo21.AddItem "28"
Combo21.AddItem "29"
Combo21.AddItem "30"
Combo21.AddItem "31"

Combo22.AddItem "1"
Combo22.AddItem "2"
Combo22.AddItem "3"
Combo22.AddItem "4"
Combo22.AddItem "5"
Combo22.AddItem "6"
Combo22.AddItem "7"
Combo22.AddItem "8"
Combo22.AddItem "9"
Combo22.AddItem "10"
Combo22.AddItem "11"
Combo22.AddItem "12"

Combo23.AddItem "2005"
Combo23.AddItem "2006"
Combo23.AddItem "2007"
Combo23.AddItem "2008"
Combo23.AddItem "2009"
Combo23.AddItem "2010"
Combo23.AddItem "2011"
Combo23.AddItem "2012"
Combo23.AddItem "2013"
Combo23.AddItem "2014"
Combo23.AddItem "2015"
Combo23.AddItem "2016"
Combo23.AddItem "2017"
Combo23.AddItem "2018"
Combo23.AddItem "2019"
Combo23.AddItem "2020"

Combo24.AddItem "1"
Combo24.AddItem "2"
Combo24.AddItem "3"
Combo24.AddItem "4"
Combo24.AddItem "5"
Combo24.AddItem "6"
Combo24.AddItem "7"
Combo24.AddItem "8"
Combo24.AddItem "9"
Combo24.AddItem "10"
Combo24.AddItem "11"
Combo24.AddItem "12"
Combo24.AddItem "13"
Combo24.AddItem "14"
Combo24.AddItem "15"
Combo24.AddItem "16"
Combo24.AddItem "17"
Combo24.AddItem "18"
Combo24.AddItem "19"
Combo24.AddItem "20"
Combo24.AddItem "21"
Combo24.AddItem "22"
Combo24.AddItem "23"
Combo24.AddItem "24"
Combo24.AddItem "25"
Combo24.AddItem "26"
Combo24.AddItem "27"
Combo24.AddItem "28"
Combo24.AddItem "29"
Combo24.AddItem "30"
Combo24.AddItem "31"

Combo25.AddItem "1"
Combo25.AddItem "2"
Combo25.AddItem "3"
Combo25.AddItem "4"
Combo25.AddItem "5"
Combo25.AddItem "6"
Combo25.AddItem "7"
Combo25.AddItem "8"
Combo25.AddItem "9"
Combo25.AddItem "10"
Combo25.AddItem "11"
Combo25.AddItem "12"

Combo26.AddItem "2005"
Combo26.AddItem "2006"
Combo26.AddItem "2007"
Combo26.AddItem "2008"
Combo26.AddItem "2009"
Combo26.AddItem "2010"
Combo26.AddItem "2011"
Combo26.AddItem "2012"
Combo26.AddItem "2013"
Combo26.AddItem "2014"
Combo26.AddItem "2015"
Combo26.AddItem "2016"
Combo26.AddItem "2017"
Combo26.AddItem "2018"
Combo26.AddItem "2019"
Combo26.AddItem "2020"

Combo27.AddItem "цемийо"
Combo27.AddItem "COOP"

Combo28.AddItem "ока"
Combo28.AddItem "цемийо"
Combo28.AddItem "COOP"

End Sub



Private Sub Form_Unload(Cancel As Integer)
On Error GoTo TELOS:

If db1.STATE = 1 Then db1.Close
ZETAIRIES.Hide
Unload ZETAIRIES
Form11ETAIRIES.Enabled = True
Form11ETAIRIES.Show
GoTo TELOS:

er:
MsgBox ("йапоио амапамтево кахос елжамистгйе. паяайакы епоийоимымгсте ле том упеухумо тгс ежаялоцгс"), vbCritical, "пяосовг !!!"

TELOS:
End Sub

Private Sub Picture1_Click()
Text2.Text = ""
Text5.Text = ""
Text4.Text = ""
Text3.Text = ""
Text6.Text = ""
Text7.Text = ""
Text12.Text = ""
Combo5.Text = "глеяа"
Combo8.Text = "глеяа"
Combo6.Text = "лгма"
Combo9.Text = "лгма"
Combo7.Text = "етос"
Combo10.Text = "етос"
Combo1.Text = ""
Combo3.Text = ""
Command24.Caption = "еуяесг"
Command1.Enabled = True
Command2.Enabled = True
Command29.Enabled = True
End Sub

Private Sub Text1_Change()
Text1.Text = Trim(Text1.Text)
End Sub

Private Sub Text10_Change()
Text10.Text = Trim(Text10.Text)
End Sub

Private Sub Text11_Change()
Text11.Text = Trim(Text11.Text)
End Sub

Private Sub Text11_LostFocus()
Dim dd As Integer
Dim S As String

dd = Len(Text11.Text)
S = Text11.Text
For i = 1 To dd
    If Mid(S, i, 1) = "," Then
        Mid(S, i, 1) = "."
    End If
Next i
Text11.Text = S
End Sub

Private Sub Text12_Change()
Text12.Text = Trim(Text12.Text)
End Sub

Private Sub Text16_Change()
Text16.Text = Trim(Text16.Text)
End Sub

Private Sub Text17_Change()
Text17.Text = Trim(Text17.Text)
End Sub

Private Sub Text18_Change()
Text18.Text = Trim(Text18.Text)
End Sub

Private Sub Text19_Change()
Text19.Text = Trim(Text19.Text)
End Sub

Private Sub Text2_Change()
Text2.Text = Trim(Text2.Text)
End Sub

Private Sub Text20_Change()
Text20.Text = Trim(Text20.Text)
End Sub

Private Sub Text21_Change()
Text21.Text = Trim(Text21.Text)

End Sub

Private Sub Text21_LostFocus()
Dim dd As Integer
Dim S As String

dd = Len(Text21.Text)
S = Text21.Text
For i = 1 To dd
    If Mid(S, i, 1) = "." Then
        Mid(S, i, 1) = ","
    End If
Next i
Text21.Text = S
End Sub

Private Sub Text23_Change()
Text23.Text = Trim(Text23.Text)
End Sub

Private Sub Text24_Change()
Text24.Text = Trim(Text24.Text)
End Sub

Private Sub Text25_Change()
Text25.Text = Trim(Text25.Text)
End Sub

Private Sub Text26_Change()
Text26.Text = Trim(Text26.Text)
End Sub

Private Sub Text3_Change()
Text3.Text = Trim(Text3.Text)
End Sub

Private Sub Text3_LostFocus()
Dim dd As Integer
Dim S As String

dd = Len(Text3.Text)
S = Text3.Text
For i = 1 To dd
    If Mid(S, i, 1) = "." Then
        Mid(S, i, 1) = ","
    End If
Next i
Text3.Text = S
End Sub

Private Sub Text4_Change()
Text4.Text = Trim(Text4.Text)
If Text4.Text = "пистытийо" Then Text6.Text = "ови"
If Text4.Text = "пистытийо" Then
    Combo27.Visible = True
Else
    Combo27.Visible = False
End If
End Sub

Private Sub Text5_Change()
Text5.Text = Trim(Text5.Text)
End Sub

Private Sub Text6_Change()
Text6.Text = Trim(Text6.Text)
End Sub

Private Sub Text7_Change()
Text7.Text = Trim(Text7.Text)
End Sub

Private Sub Text8_Change()
Text8.Text = Trim(Text8.Text)
End Sub

Private Sub Text9_Change()
Text9.Text = Trim(Text9.Text)
End Sub

Private Sub Timer1_Timer()
Dim txt_name As String
txt_name = App.Path & "\TXTS\" & Text24.Text & ".TXT"
If FileLen(txt_name) > 3 Then
    Image1.Visible = False
    Image6.Visible = True
    Timer2.Enabled = True
Else
    Image1.Visible = True
    Image6.Visible = False
    Timer2.Enabled = False
End If
End Sub

Private Sub Timer2_Timer()
If CT < 61 Then
Image6.Height = Image6.Height + 50
Image6.Width = Image6.Width + 50
CT = CT + 10
Else
Image6.Height = 360
Image6.Width = 360
CT = 0
End If
End Sub
