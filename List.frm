VERSION 5.00
Begin VB.Form List 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   3825
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8685
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   8685
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame P 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3765
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   8625
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      Height          =   3795
      Left            =   15
      Top             =   15
      Width           =   8655
   End
End
Attribute VB_Name = "List"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub P_DblClick()
MM.playFullScreen
End Sub
