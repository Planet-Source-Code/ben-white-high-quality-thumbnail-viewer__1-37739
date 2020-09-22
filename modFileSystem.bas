Attribute VB_Name = "modFileSystem"
Option Explicit

Private Type OPENFILENAME
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

Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long

Public Function ShowOpen(hwnd As Long, Optional ByVal sInitialDir As String) As String
    Dim ofn As OPENFILENAME
    ofn.lStructSize = Len(ofn)
    ofn.hwndOwner = hwnd
    ofn.hInstance = App.hInstance
    ofn.lpstrFilter = "Image Files" + Chr$(0) + "*.jpg;*.bmp;*.gif;*.ico;*.cur;*.rle;*.wmf;*.emf" + Chr$(0)
    ofn.lpstrFile = Space$(254)
    ofn.nMaxFile = 255
    ofn.lpstrFileTitle = Space$(254)
    ofn.nMaxFileTitle = 255
    ofn.lpstrInitialDir = CurDir
    ofn.lpstrTitle = "Open File"
    ofn.flags = 4
    ofn.lpstrInitialDir = sInitialDir
    Dim lngRetVal As Long
    lngRetVal = GetOpenFileName(ofn)

    If (lngRetVal) Then
        ShowOpen = Replace(Trim$(ofn.lpstrFile), Chr(0), "")
    Else
        ShowOpen = ""
    End If
End Function
