Attribute VB_Name = "Module"
Option Explicit

        '
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function GetCapture Lib "user32" () As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Declare Function waveOutGetNumDevs Lib "winmm.dll" () As Long

Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Const SND_SYNC = &H0
Const SND_ASYNC = &H1
Const SND_NODEFAULT = &H2
Const SND_LOOP = &H8
Const SND_NOSTOP = &H10


Public Type POINTAPI
        X As Long
        Y As Long
End Type

Public Const SRCCOPY = &HCC0020

Global iTPPY As Long
Global iTPPX As Long
'======================='
Private Type OpenFileName
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Public Const OFN_READONLY = &H1
Public Const OFN_OVERWRITEPROMPT = &H2
Public Const OFN_HIDEREADONLY = &H4
Public Const OFN_NOCHANGEDIR = &H8
Public Const OFN_SHOWHELP = &H10
Public Const OFN_ENABLEHOOK = &H20
Public Const OFN_ENABLETEMPLATE = &H40
Public Const OFN_ENABLETEMPLATEHANDLE = &H80
Public Const OFN_NOVALIDATE = &H100
Public Const OFN_ALLOWMULTISELECT = &H200
Public Const OFN_EXTENSIONDIFFERENT = &H400
Public Const OFN_PATHMUSTEXIST = &H800
Public Const OFN_FILEMUSTEXIST = &H1000
Public Const OFN_CREATEPROMPT = &H2000
Public Const OFN_SHAREAWARE = &H4000
Public Const OFN_NOREADONLYRETURN = &H8000
Public Const OFN_NOTESTFILECREATE = &H10000
Public Const OFN_NONETWORKBUTTON = &H20000
Public Const OFN_NOLONGNAMES = &H40000                      '  force no long names for 4.x modules
Public Const OFN_EXPLORER = &H80000                         '  new look commdlg
Public Const OFN_NODEREFERENCELINKS = &H100000
Public Const OFN_LONGNAMES = &H200000                       '  force long names for 3.x modules

Public Const OFN_SHAREFALLTHROUGH = 2
Public Const OFN_SHARENOWARN = 1
Public Const OFN_SHAREWARN = 0

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OpenFileName) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OpenFileName) As Long
Public FileName As String
Public ExePath As String




Function SaveDialog(Form1 As Form, Filter As String, Title As String, InitDir As String) As String
 
 Dim ofn As OpenFileName
    Dim A As Long
    ofn.lStructSize = Len(ofn)
    ofn.hwndOwner = Form1.hwnd
    ofn.hInstance = App.hInstance
    If Right$(Filter, 1) <> "|" Then Filter = Filter + "|"
    For A = 1 To Len(Filter)
        If Mid$(Filter, A, 1) = "|" Then Mid$(Filter, A, 1) = Chr$(0)
    Next
    ofn.lpstrFilter = Filter
        ofn.lpstrFile = Space$(254)
        ofn.nMaxFile = 255
        ofn.lpstrFileTitle = Space$(254)
        ofn.nMaxFileTitle = 255
        ofn.lpstrInitialDir = InitDir
        ofn.lpstrTitle = Title
        ofn.flags = OFN_HIDEREADONLY Or OFN_OVERWRITEPROMPT Or OFN_CREATEPROMPT
        A = GetSaveFileName(ofn)

        If (A) Then
            SaveDialog = Trim$(ofn.lpstrFile)
        Else
            SaveDialog = ""
        End If

End Function


Function OpenDialog(Form1 As Form, Filter As String, Title As String, InitDir As String) As String
 
 Dim ofn As OpenFileName
    Dim A As Long
    ofn.lStructSize = Len(ofn)
    ofn.hwndOwner = Form1.hwnd
    ofn.hInstance = App.hInstance
    If Right$(Filter, 1) <> "|" Then Filter = Filter + "|"
    For A = 1 To Len(Filter)
        If Mid$(Filter, A, 1) = "|" Then Mid$(Filter, A, 1) = Chr$(0)
    Next
    ofn.lpstrFilter = Filter
        ofn.lpstrFile = Space$(254)
        ofn.nMaxFile = 255
        ofn.lpstrFileTitle = Space$(254)
        ofn.nMaxFileTitle = 255
        ofn.lpstrInitialDir = InitDir
        ofn.lpstrTitle = Title
        ofn.flags = OFN_HIDEREADONLY Or OFN_FILEMUSTEXIST
        A = GetOpenFileName(ofn)

        If (A) Then
            OpenDialog = Trim$(ofn.lpstrFile)
        Else
            OpenDialog = ""
        End If

End Function


Function DetectSnd()

Dim i As Long

i = waveOutGetNumDevs()

If i > 0 Then         ' There is at least one device.
      MsgBox "You got sound"
Else
      MsgBox "Cannot Play Wave Data"
End If

End Function



Sub Main()

Form1.Show

End Sub
