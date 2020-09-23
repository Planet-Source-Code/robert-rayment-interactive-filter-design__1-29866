VERSION 5.00
Begin VB.Form frmMsg 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form2"
   ClientHeight    =   1020
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   2130
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1020
   ScaleWidth      =   2130
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdYNO 
      Caption         =   "OK"
      Height          =   375
      Index           =   2
      Left            =   765
      TabIndex        =   2
      Top             =   315
      Width           =   615
   End
   Begin VB.CommandButton cmdYNO 
      Caption         =   "NO"
      Height          =   375
      Index           =   1
      Left            =   1320
      TabIndex        =   1
      Top             =   180
      Width           =   735
   End
   Begin VB.CommandButton cmdYNO 
      Caption         =   "YES"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   735
   End
End
Attribute VB_Name = "frmMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' frmMsg.frm
' Quick MesageBox to avoid BEEP
' But note needs:-
' Public ynYES, ynNO, ynOK

' Called using:-

' RRMsg Message$, ynYES, ynNO, ynOK
' where Message$ becomes frmMsg.Caption
' yn 1/0 shows or not cmd buttom

' Response yn = 1 for cmd button pressed


Private Sub cmdYNO_Click(Index As Integer)
' Public ynYES, ynNO, ynOK

Select Case Index
Case 0: ynYES = 1: ynNO = 0: ynOK = 0
Case 1: ynYES = 0: ynNO = 1: ynOK = 0
Case 2: ynYES = 0: ynNO = 0: ynOK = 1
End Select

Unload frmMsg

End Sub

Private Sub Form_Load()
Top = 6500
Left = 3500
End Sub
