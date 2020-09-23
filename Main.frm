VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   1695
   ClientLeft      =   2070
   ClientTop       =   825
   ClientWidth     =   5700
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   Picture         =   "Main.frx":058A
   ScaleHeight     =   1695
   ScaleWidth      =   5700
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   400
      Left            =   5640
      Top             =   1620
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      ForeColor       =   &H00C0C0C0&
      Height          =   225
      Left            =   4560
      Locked          =   -1  'True
      TabIndex        =   15
      Text            =   "x"
      Top             =   480
      Width           =   135
   End
   Begin VB.TextBox videosizew 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      ForeColor       =   &H00C0C0C0&
      Height          =   225
      Left            =   4710
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   480
      Width           =   585
   End
   Begin VB.TextBox videosizeh 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      ForeColor       =   &H00C0C0C0&
      Height          =   225
      Left            =   4170
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   480
      Width           =   405
   End
   Begin MSComDlg.CommonDialog C 
      Left            =   5760
      Top             =   1770
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.Slider Rate 
      Height          =   135
      Left            =   4440
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1020
      Width           =   645
      _ExtentX        =   1138
      _ExtentY        =   238
      _Version        =   393216
      Max             =   100
      SelStart        =   50
      TickStyle       =   3
      Value           =   50
   End
   Begin MSComctlLib.Slider Volume 
      Height          =   165
      Left            =   1950
      TabIndex        =   9
      Top             =   1020
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   291
      _Version        =   393216
      Max             =   100
      SelStart        =   85
      TickStyle       =   3
      Value           =   85
   End
   Begin MSComctlLib.Slider H 
      Height          =   135
      Left            =   1380
      TabIndex        =   10
      Top             =   750
      Width           =   3885
      _ExtentX        =   6853
      _ExtentY        =   238
      _Version        =   393216
      TickStyle       =   3
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1350
      TabIndex        =   13
      Top             =   720
      Width           =   3945
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Height          =   315
      Left            =   5370
      TabIndex        =   20
      Top             =   1350
      Width           =   315
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Proyecto-X Video Player V 1.0"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   1530
      TabIndex        =   19
      Top             =   60
      Width           =   3555
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0080C0FF&
      X1              =   4080
      X2              =   4080
      Y1              =   660
      Y2              =   450
   End
   Begin VB.Label lblvideoname 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   225
      Left            =   1530
      TabIndex        =   18
      Top             =   480
      Width           =   2505
   End
   Begin VB.Label lblversion 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   135
      Left            =   360
      TabIndex        =   17
      Top             =   570
      Width           =   345
   End
   Begin VB.Label btnoptions 
      BackStyle       =   0  'Transparent
      Height          =   225
      Left            =   3480
      TabIndex        =   16
      ToolTipText     =   "Mostrar Opciones"
      Top             =   1290
      Width           =   345
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Velocidad ::."
      Height          =   255
      Left            =   3510
      TabIndex        =   12
      Top             =   960
      Width           =   1605
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Vol ::."
      Height          =   255
      Left            =   1470
      TabIndex        =   11
      Top             =   960
      Width           =   1995
   End
   Begin VB.Label btnpause 
      BackStyle       =   0  'Transparent
      Height          =   225
      Left            =   1980
      TabIndex        =   6
      Top             =   1290
      Width           =   345
   End
   Begin VB.Label btnstop 
      BackStyle       =   0  'Transparent
      Height          =   225
      Left            =   1590
      TabIndex        =   5
      Top             =   1290
      Width           =   345
   End
   Begin VB.Label btnplay 
      BackStyle       =   0  'Transparent
      Height          =   225
      Left            =   1230
      TabIndex        =   4
      Top             =   1290
      Width           =   345
   End
   Begin VB.Label btnopen 
      BackStyle       =   0  'Transparent
      Height          =   225
      Left            =   3090
      TabIndex        =   3
      ToolTipText     =   "Abrir Video"
      Top             =   1290
      Width           =   345
   End
   Begin VB.Image Image1 
      Height          =   285
      Left            =   1200
      Picture         =   "Main.frx":2111
      Top             =   1260
      Width           =   2655
   End
   Begin VB.Label pxoverlogo 
      BackStyle       =   0  'Transparent
      Height          =   435
      Left            =   210
      TabIndex        =   2
      ToolTipText     =   "Proyecto-X Mp3 Player"
      Top             =   210
      Width           =   495
   End
   Begin VB.Label Minimizar 
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   5310
      TabIndex        =   1
      ToolTipText     =   "Minimizar"
      Top             =   210
      Width           =   195
   End
   Begin VB.Label Salir 
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   5490
      TabIndex        =   0
      ToolTipText     =   "Salir"
      Top             =   210
      Width           =   195
   End
   Begin VB.Menu mnuprincipal 
      Caption         =   "P-X Video Player"
      Visible         =   0   'False
      Begin VB.Menu mnuAbrirVideo 
         Caption         =   "Abrir Video"
      End
      Begin VB.Menu mnuReproducir 
         Caption         =   "Reproducir"
      End
      Begin VB.Menu mnuPausa 
         Caption         =   "Pausa"
      End
      Begin VB.Menu mnuParar 
         Caption         =   "Parar"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private WithEvents gSysTray As clsSysTray
Attribute gSysTray.VB_VarHelpID = -1
Private MouseDownForm
Private MouseDownFormX
Private MouseDownFormY
Private Type POINTAPI
    x As Long
    y As Long
End Type
Dim MM As New MovieModule

Private Sub btnopen_Click()
    Dim a As Long
    Dim b As Long
    C.DialogTitle = "Seleccione un archivo de video"
    C.Filter = "Avi Files (*.avi)|*.avi|Mpeg Files (*.mpeg)|*.mpeg|Mpg Files (*.mpg)|*.mpg|Mov Files (*.mov)|*.mov|All Files (*.*)|*.*"
    C.ShowOpen
    MM.Filename = C.Filename
    If C.Filename = "" Then GoTo nofile
    lblvideoname.Caption = C.FileTitle
    frmOpciones.lblvideohide.Caption = "       Esconder Video"
    H.Value = "0"
    MM.openMovieWindow List.P.hwnd, "child" 'this will open our movie in a child window
    C.Filename = ""
    MM.extractDefaultMovieSize a, b
    videosizeh.Text = CStr(a)
    videosizew.Text = CStr(b)
    Call ajustarvalores
nofile:
End Sub

Private Sub btnoptions_Click()
    Dim File As String, OFLen As Double, Str As String
    File = App.Path & "\configuracion.pxv"
    OFLen = FileLen(File)
If frmOpciones.Visible = True Then frmOpciones.Hide Else frmOpciones.Show
If frmOpciones.Visible = True Then
WriteIni File, "General", "Opciones", "Si-"
Else
WriteIni File, "General", "Opciones", "No-"
End If
Call pegarlista
End Sub

Private Sub btnpause_Click()
If MM.isMoviePlaying = True Then
MM.pauseMovie
Else
MM.resumeMovie
End If
End Sub

Private Sub btnplay_Click()
    On Error Resume Next
    MM.playMovie
    MM.setVolume Volume.Value * 10 ' set the new movie the selected volume
    MM.setSpeed Rate.Value * 20 'set the new movie to the selected speed
    H.Max = Val(MM.getLengthInSec) 'load the position bar with the max length
    MM.timeOut 0.5 'Give the mci device enough time to process

End Sub

Private Sub btnstop_Click()
    MM.stopMovie
    H.Value = "0"
    MM.setPositionTo H.Value
End Sub



Private Sub Form_Initialize()
If App.PrevInstance = True Then End
End Sub

Private Sub Form_Load()
    Set gSysTray = New clsSysTray
    gSysTray.LoadIcon Me.Icon, Me
    gSysTray.ToolTip = "P-X MP3 Player" & Chr(0)
    gSysTray.IconInSysTray
    Call pegarlista
    pausa = "no"
    Call shutup
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
MouseDownForm = 1
MouseDownFormX = x
MouseDownFormY = y
End Sub
Private Sub form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
MouseDownForm = 0
Call pegarlista
End Sub
Private Sub form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If MouseDownForm <> 1 Then
Exit Sub
End If
Dim z As POINTAPI
Call GetCursorPos(z)
frmMain.top = (z.y * 15) - MouseDownFormY
frmMain.left = (z.x * 15) - MouseDownFormX
Call pegarlista
End Sub

Private Sub gSysTray_LButtonDblClk()
Call alfrente
Call noalfrente
End Sub

Private Sub gSysTray_RButtonUP()
    On Error Resume Next
    frmMain.WindowState = vbNormal
    frmMain.Show
    PopupMenu Me.mnuprincipal
End Sub
Private Sub H_Click()
    MM.setPositionTo H.Value
End Sub

Private Sub Label4_Click()
Shell "start http://www.Proyecto-X.8m.com"
End Sub

Private Sub Minimizar_Click()
Me.Hide
List.Hide
frmOpciones.Hide
Acercade.Hide
End Sub

Private Sub mnuAbrirVideo_Click()
    Dim a As Long
    Dim b As Long
    C.DialogTitle = "Seleccione un archivo de video"
    C.Filter = "Avi Files (*.avi)|*.avi|Mpeg Files (*.mpeg)|*.mpeg|Mpg Files (*.mpg)|*.mpg|Mov Files (*.mov)|*.mov|All Files (*.*)|*.*"
    C.ShowOpen
    MM.Filename = C.Filename
    If C.Filename = "" Then GoTo nofiles
        frmOpciones.lblvideohide.Caption = "       Esconder Video"
    H.Value = "0"
    MM.openMovieWindow List.P.hwnd, "child" 'this will open our movie in a child window
    C.Filename = ""
    MM.extractDefaultMovieSize a, b
    videosizeh.Text = CStr(a)
    videosizew.Text = CStr(b)
    Call ajustarvalores
nofiles:
End Sub

Private Sub mnuParar_Click()
    MM.stopMovie
    H.Value = "0"
    MM.setPositionTo H.Value
End Sub

Private Sub mnuPausa_Click()
If MM.isMoviePlaying = True Then
MM.pauseMovie
Else
MM.resumeMovie
End If
End Sub

Private Sub mnuReproducir_Click()
    On Error Resume Next
    MM.playMovie
    MM.setVolume Volume.Value * 10 ' set the new movie the selected volume
    MM.setSpeed Rate.Value * 20 'set the new movie to the selected speed
    H.Max = Val(MM.getLengthInSec) 'load the position bar with the max length
    MM.timeOut 0.5 'Give the mci device enough time to process
End Sub
Private Sub Rate_Click()
    MM.setSpeed Rate.Value * 20
End Sub

Private Sub Rate_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    MM.setSpeed Rate.Value * 20
End Sub

Private Sub Salir_Click()
End
End Sub
Private Sub Salir_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Call shutdown
    gSysTray.RemoveFromSysTray
End Sub
Private Sub Volume_Click()
    MM.setVolume Volume.Value * 10
End Sub

Private Sub Volume_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    MM.setVolume Volume.Value * 10
End Sub

Private Sub Volume_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim volcalculo As String
Dim File As String, OFLen As Double, Str As String
File = App.Path & "\configuracion.pxv"
OFLen = FileLen(File)
volcalculo = Volume.Value * 10
WriteIni File, "General", "Volumen", volcalculo & "-"
End Sub
