VERSION 5.00
Begin VB.Form small 
   BackColor       =   &H00FF00FF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   72
   ClientLeft      =   22560
   ClientTop       =   48
   ClientWidth     =   144
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   72
   ScaleWidth      =   144
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

