VERSION 5.00
Begin VB.Form Acercade 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   1725
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5730
   LinkTopic       =   "Form1"
   ScaleHeight     =   1725
   ScaleWidth      =   5730
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   180
      Top             =   90
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      ForeColor       =   &H000080FF&
      Height          =   1245
      Left            =   3210
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "Acercade.frx":0000
      Top             =   90
      Width           =   2445
   End
   Begin VB.Label lblhora 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Hora"
      ForeColor       =   &H0000C000&
      Height          =   195
      Left            =   4950
      TabIndex        =   3
      Top             =   1470
      Width           =   735
   End
   Begin VB.Label lblfecha 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha"
      ForeColor       =   &H0000C000&
      Height          =   195
      Left            =   3330
      TabIndex        =   2
      Top             =   1470
      Width           =   1395
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      Height          =   1665
      Left            =   30
      Top             =   30
      Width           =   5675
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Programado por Gerónimo Oñativia"
      ForeColor       =   &H000080FF&
      Height          =   195
      Left            =   150
      TabIndex        =   0
      Top             =   1450
      Width           =   2865
   End
   Begin VB.Image Image1 
      Height          =   1470
      Left            =   -90
      Picture         =   "Acercade.frx":00B3
      Top             =   90
      Width           =   3510
   End
End
Attribute VB_Name = "Acercade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Click()
Me.Hide
End Sub

Private Sub Form_Load()
    Dim File As String, OFLen As Double, Str As String, variab As String
    File = App.Path & "\configuracion.pxv"
    OFLen = FileLen(File)
    variab = ReadIni(File, "General", "Veces Abierto") - 1
    Text1.Text = "Esta es la " & variab & "° vez q usas este" & vbCrLf & "software. Gracias por bajarte la" & vbCrLf & "versión 1.0 del reproductor de" & vbCrLf & "video de Proyecto-X." & vbCrLf & "Ayudanos reportando los bugs a" & vbCrLf & "G_e_R_o@Hotmail.com"
End Sub
Private Sub Image1_Click()
Me.Hide
End Sub

Private Sub Label1_Click()
Me.Hide
End Sub

Private Sub Label2_Click()
Me.Hide
End Sub

Private Sub veces_Click()
Me.Hide
End Sub

Private Sub Timer1_Timer()
lblhora.Caption = Time$
lblfecha.Caption = Date$
End Sub
