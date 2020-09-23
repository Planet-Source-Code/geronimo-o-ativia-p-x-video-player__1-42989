Attribute VB_Name = "ponerorden"
Option Explicit
Private Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const HWND_TOPMOST = -1
Sub ajustarvalores()
Dim a, b, C, d As String
a = frmMain.videosizew.Text + 5
b = frmMain.videosizeh.Text + 5
List.P.Height = frmMain.videosizew.Text * 15
List.P.Width = frmMain.videosizeh.Text * 15
List.Height = a * 15
List.Width = b * 15
C = List.Height - 30
d = List.Width - 30
List.Shape1.Height = C
List.Shape1.Width = d
List.Shape1.top = 15
List.Shape1.left = 15
List.P.left = 2 * 15
List.P.top = 2 * 15
Call pegarlista
List.Show
End Sub

Sub pegarlista()
Dim cuentas, cuentas2, cuentas3, cuentas4, cuentas5, cuentas6 As String
If List.Width < frmMain.Width Then
List.top = frmMain.top + frmMain.Height
cuentas3 = (frmMain.Width - List.Width) / 2
cuentas4 = frmMain.left + cuentas3
List.left = cuentas4
frmOpciones.top = frmMain.top
cuentas5 = frmMain.left + frmMain.Width
frmOpciones.left = cuentas5
frmOpciones.Height = frmMain.Height
cuentas6 = frmMain.top + frmMain.Height + List.Height
Acercade.top = cuentas6
Acercade.left = frmMain.left + (frmMain.Width / 2)

Else
List.top = frmMain.top + frmMain.Height
cuentas = (List.Width - frmMain.Width) / 2
cuentas2 = frmMain.left - cuentas
List.left = cuentas2
frmOpciones.top = frmMain.top
cuentas5 = frmMain.left + frmMain.Width
frmOpciones.left = cuentas5
frmOpciones.Height = frmMain.Height
cuentas6 = frmMain.top + frmMain.Height + List.Height
Acercade.top = cuentas6
Acercade.left = frmMain.left
End If
End Sub
Sub shutdown()
    Dim File As String, OFLen As Double, Str As String
    File = App.Path & "\configuracion.pxv"
    OFLen = FileLen(File)
If List.Visible = True Then
WriteIni File, "General", "Lista", "Si-"
Else
WriteIni File, "General", "Lista", "No-"
End If
If frmOpciones.Visible = True Then
WriteIni File, "General", "Opciones", "Si-"
Else
WriteIni File, "General", "Opciones", "No-"
End If
End Sub
Sub shutup()
    Dim MM As New MovieModule
    Dim pausa, usad, usado As String
    Dim volcalc As Long, volcalc2 As String
    Dim File As String, OFLen As Double, Str As String
    File = App.Path & "\configuracion.pxv"
    OFLen = FileLen(File)

    ' Verificar Versión
    frmMain.lblversion.Caption = ReadIni(File, "General", "Versión")
    ' Comprobar si la lista y las opciones estaban
    ' visibles la ultima vez q se cerro el programa
    If ReadIni(File, "General", "Lista") = "Si" Then
    List.Show
    frmOpciones.lblvideohide.Caption = "       Esconder Video"
    End If
    If ReadIni(File, "General", "Lista") = "No" Then
    List.Hide
    frmOpciones.lblvideohide.Caption = "       Mostrar Video"
    End If
    If ReadIni(File, "General", "Opciones") = "Si" Then frmOpciones.Show
    If ReadIni(File, "General", "Opciones") = "No" Then frmOpciones.Hide
    ' Sumar 1 a las veces de uso
    usad = ReadIni(File, "General", "Veces Abierto")
    usado = usad + 1
    WriteIni File, "General", "Veces Abierto", usado & "-"
    ' Ajustar el volúmen
    volcalc = ReadIni(File, "General", "Volumen")
    volcalc2 = volcalc / 10
    MM.setVolume volcalc
    frmMain.Volume.Value = volcalc2
    Call pegarlista
End Sub
Sub noalfrente2()
    Dim MM As New MovieModule
    Dim pausa, usad, usado As String
    Dim volcalc As Long, volcalc2 As String
    Dim File As String, OFLen As Double, Str As String
    File = App.Path & "\configuracion.pxv"
    OFLen = FileLen(File)

    ' Verificar Versión
    frmMain.lblversion.Caption = ReadIni(File, "General", "Versión")
    ' Comprobar si la lista y las opciones estaban
    ' visibles la ultima vez q se cerro el programa
    If ReadIni(File, "General", "Lista") = "Si" Then
    List.Show
    frmOpciones.lblvideohide.Caption = "       Esconder Video"
    End If
    If ReadIni(File, "General", "Lista") = "No" Then
    List.Hide
    frmOpciones.lblvideohide.Caption = "       Mostrar Video"
    End If
    If ReadIni(File, "General", "Opciones") = "Si" Then frmOpciones.Show
    If ReadIni(File, "General", "Opciones") = "No" Then frmOpciones.Hide
    ' Ajustar el volúmen
    volcalc = ReadIni(File, "General", "Volumen")
    volcalc2 = volcalc / 10
    MM.setVolume volcalc
    frmMain.Volume.Value = volcalc2
    Call pegarlista
End Sub

Sub alfrente()
SetWindowPos frmMain.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
SetWindowPos frmOpciones.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
SetWindowPos List.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
SetWindowPos Acercade.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub
Sub noalfrente()
Unload frmMain
Load frmMain
frmMain.Show
Unload frmOpciones
Load frmOpciones
frmOpciones.Show
Unload List
Load List
Unload Acercade
Load Acercade
Call noalfrente2
End Sub
Public Function ReadIni(Filename As String, Section As String, Key As String) As String
Dim RetVal As String * 255, v As Long
v = GetPrivateProfileString(Section, Key, "", RetVal, 255, Filename)
ReadIni = left(RetVal, v - 1)
End Function

'reads ini section
Public Function ReadIniSection(Filename As String, Section As String) As String
Dim RetVal As String * 255, v As Long
v = GetPrivateProfileSection(Section, RetVal, 255, Filename)
ReadIniSection = left(RetVal, v - 1)
End Function

'writes ini
Public Sub WriteIni(Filename As String, Section As String, Key As String, Value As String)
WritePrivateProfileString Section, Key, Value, Filename
End Sub

'writes ini section
Public Sub WriteIniSection(Filename As String, Section As String, Value As String)
WritePrivateProfileSection Section, Value, Filename
End Sub

