VERSION 5.00
Begin VB.Form small 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   75
   ClientLeft      =   23550
   ClientTop       =   45
   ClientWidth     =   150
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   75
   ScaleWidth      =   150
   ShowInTaskbar   =   0   'False
End
Attribute VB_Name = "small"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_DblClick()
small.Hide
Unload small
RADIO.Show
End Sub

