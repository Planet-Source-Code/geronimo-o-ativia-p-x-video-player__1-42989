VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSysTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uId As Long
    uFlags As Long
    ucallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Private Declare Function Shell_NotifyIcon _
Lib "shell32" _
Alias "Shell_NotifyIconA" _
(ByVal dwMessage As Long, pnid As NOTIFYICONDATA) _
As Boolean

Private pct As NOTIFYICONDATA
Private WithEvents pctHook As PictureBox
Attribute pctHook.VB_VarHelpID = -1
Public Event LButtonDblClk()
Public Event RButtonUp()

Public Property Let ToolTip(ByVal vData As String)
    pct.szTip = vData
End Property

Public Sub LoadIcon(ByVal vData As Variant, ByRef vDatafrm As Form)
    Set pctHook = vDatafrm.Controls.Add("VB.PictureBox", "pctHook")
    Set pctHook.Picture = vDatafrm.Icon
    pct.hIcon = pctHook.Picture
    pct.hWnd = pctHook.hWnd
    Shell_NotifyIcon &H1, pct
End Sub

Private Sub Class_Initialize()
    pct.cbSize = Len(pct)
    pct.uId = 1&
    pct.uFlags = &H2 Or &H4 Or &H1
    pct.ucallbackMessage = &H200
    pct.szTip = Chr$(0)       'Default to no tooltip
End Sub

Private Sub pctHook_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Static Msg As Long
    Msg = X / Screen.TwipsPerPixelX
        If Msg = 515 Then
                RaiseEvent LButtonDblClk
        ElseIf Msg = 517 Then
                RaiseEvent RButtonUp
        End If
End Sub

Public Sub RemoveFromSysTray()
    Shell_NotifyIcon &H2, pct
End Sub

Public Sub IconInSysTray()
    Shell_NotifyIcon &H0, pct
End Sub



