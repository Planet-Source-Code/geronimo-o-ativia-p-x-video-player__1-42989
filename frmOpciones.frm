VERSION 5.00
Begin VB.Form frmOpciones 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   1740
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2790
   LinkTopic       =   "Form1"
   Picture         =   "frmOpciones.frx":0000
   ScaleHeight     =   116
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   186
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label lblontop 
      BackStyle       =   0  'Transparent
      Caption         =   "       Mantener reproductor al frente"
      ForeColor       =   &H008080FF&
      Height          =   225
      Left            =   180
      TabIndex        =   4
      Top             =   1350
      Width           =   2565
   End
   Begin VB.Label lblacerca 
      BackStyle       =   0  'Transparent
      Caption         =   "       Acerca de ..."
      ForeColor       =   &H008080FF&
      Height          =   225
      Left            =   180
      TabIndex        =   3
      Top             =   1050
      Width           =   1935
   End
   Begin VB.Label lblfullscr 
      BackStyle       =   0  'Transparent
      Caption         =   "       Ver Pantalla Completa"
      ForeColor       =   &H008080FF&
      Height          =   225
      Left            =   180
      TabIndex        =   2
      Top             =   780
      Width           =   1935
   End
   Begin VB.Label lblvideohide 
      BackStyle       =   0  'Transparent
      Caption         =   "       Esconder Video"
      ForeColor       =   &H008080FF&
      Height          =   225
      Left            =   180
      TabIndex        =   1
      Top             =   480
      Width           =   1515
   End
   Begin VB.Label lblsonido 
      BackStyle       =   0  'Transparent
      Caption         =   "       Si"
      ForeColor       =   &H008080FF&
      Height          =   225
      Left            =   180
      TabIndex        =   0
      Top             =   210
      Width           =   555
   End
End
Attribute VB_Name = "frmOpciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim MM As New MovieModule

Private Sub lblacerca_Click()
If Acercade.Visible = False Then
Acercade.Show
Else
Acercade.Hide
End If
End Sub

Private Sub lblfullscr_Click()
MM.playFullScreen
End Sub

Private Sub lblontop_Click()
If lblontop.ForeColor = &HC0E0FF Then
Call noalfrente
lblontop.ForeColor = &H8080FF
Else
Call alfrente
lblontop.ForeColor = &HC0E0FF
End If
End Sub

Private Sub lblsonido_Click()
If lblsonido.Caption = "       Si" Then
MM.setAudioOff
lblsonido.Caption = "       No"
Else
MM.setAudioOn
lblsonido.Caption = "       Si"
End If
End Sub

Private Sub lblvideohide_Click()
    Dim File As String, OFLen As Double, Str As String
    File = App.Path & "\configuracion.pxv"
    OFLen = FileLen(File)
If lblvideohide.Caption = "       Esconder Video" Then
List.Hide
lblvideohide.Caption = "       Mostrar Video"
Else
List.Show
lblvideohide.Caption = "       Esconder Video"
End If
If List.Visible = True Then
WriteIni File, "General", "Lista", "Si-"
Else
WriteIni File, "General", "Lista", "No-"
End If
Call pegarlista
End Sub

