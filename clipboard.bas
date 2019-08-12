Option Explicit

Private Declare PtrSafe Function IsClipboardFormatAvailable Lib "user32" (ByVal uFormat As Long) As Long
Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal Hwnd As Long) As Long
Private Declare PtrSafe Function GetClipboardData Lib "user32" (ByVal uFormat As Long) As Long
Private Declare PtrSafe Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare PtrSafe Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal strDest As Any, ByVal lpSource As Any, ByVal Length As Long)
Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As Long) As Long
Private Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
Private Declare PtrSafe Function SetClipboardData Lib "user32" (ByVal uFormat As Long, ByVal hData As Long) As Long
Private Declare PtrSafe Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long

Private Const GMEM_MOVABLE = &H2&
Private Const GMEM_DDESHARE = &H2000&
Private Const CF_TEXT = 1

'Error return codes from Clipboard2Text
Private Const CLIPBOARDFORMATNOTAVAILABLE = 1
Private Const CANNOTOPENCLIPBOARD = 2
Private Const CANNOTGETCLIPBOARDDATA = 3
Private Const CANNOTGLOBALLOCK = 4
Private Const CANNOTCLOSECLIPBOARD = 5
Private Const CANNOTGLOBALALLOC = 6
Private Const CANNOTEMPTYCLIPBOARD = 7
Private Const CANNOTSETCLIPBOARDDATA = 8
Private Const CANNOTGLOBALFREE = 9

'Function SetText
Function SetText(strText As String) As Variant
    Dim varRet As Variant
    Dim fSetClipboardData As Boolean
    Dim hMemory As Long
    Dim lpMemory As Long
    Dim lngSize As Long

    varRet = False
    fSetClipboardData = False

    ' Get the length, including one extra for a CHR$(0)
    ' at the end.
    lngSize = Len(strText) + 1
    hMemory = GlobalAlloc(GMEM_MOVABLE Or _
        GMEM_DDESHARE, lngSize)
    If Not CBool(hMemory) Then
        varRet = CVErr(CANNOTGLOBALALLOC)
        GoTo SetTextDone
    End If

    ' Lock the object into memory
    lpMemory = GlobalLock(hMemory)
    If Not CBool(lpMemory) Then
        varRet = CVErr(CANNOTGLOBALLOCK)
        GoTo SetTextGlobalFree
    End If

    ' Move the string into the memory we locked
    Call MoveMemory(lpMemory, strText, lngSize)

    ' Don't send clipboard locked memory.
    Call GlobalUnlock(hMemory)

    ' Open the clipboard
    If Not CBool(OpenClipboard(0&)) Then
        varRet = CVErr(CANNOTOPENCLIPBOARD)
        GoTo SetTextGlobalFree
    End If

    ' Remove the current contents of the clipboard
    If Not CBool(EmptyClipboard()) Then
        varRet = CVErr(CANNOTEMPTYCLIPBOARD)
        GoTo SetTextCloseClipboard
    End If

    ' Add our string to the clipboard as text
    If Not CBool(SetClipboardData(CF_TEXT, _
        hMemory)) Then
        varRet = CVErr(CANNOTSETCLIPBOARDDATA)
        GoTo SetTextCloseClipboard
    Else
        fSetClipboardData = True
    End If

SetTextCloseClipboard:
    ' Close the clipboard
    If Not CBool(CloseClipboard()) Then
        varRet = CVErr(CANNOTCLOSECLIPBOARD)
    End If

SetTextGlobalFree:
    If Not fSetClipboardData Then
        'If we have set the clipboard data, we no longer own
        ' the object--Windows does, so don't free it.
        If CBool(GlobalFree(hMemory)) Then
            varRet = CVErr(CANNOTGLOBALFREE)
        End If
    End If

SetTextDone:
    SetText = varRet
End Function

'Function GetText
Public Function GetText() As Variant
    Dim hMemory As Long
    Dim lpMemory As Long
    Dim strText As String
    Dim lngSize As Long
    Dim varRet As Variant

    varRet = ""

    ' Is there text on the clipboard? If not, error out.
    If Not CBool(IsClipboardFormatAvailable _
        (CF_TEXT)) Then
        varRet = CVErr(CLIPBOARDFORMATNOTAVAILABLE)
        GoTo GetTextDone
    End If

    ' Open the clipboard
    If Not CBool(OpenClipboard(0&)) Then
        varRet = CVErr(CANNOTOPENCLIPBOARD)
        GoTo GetTextDone
    End If

    ' Get the handle to the clipboard data
    hMemory = GetClipboardData(CF_TEXT)
    If Not CBool(hMemory) Then
        varRet = CVErr(CANNOTGETCLIPBOARDDATA)
        GoTo GetTextCloseClipboard
    End If

    ' Find out how big it is and allocate enough space
    ' in a string
    lngSize = GlobalSize(hMemory)
    strText = Space$(lngSize)

    ' Lock the handle so we can use it
    lpMemory = GlobalLock(hMemory)
    If Not CBool(lpMemory) Then
        varRet = CVErr(CANNOTGLOBALLOCK)
        GoTo GetTextCloseClipboard
    End If

    ' Move the information from the clipboard memory
    ' into our string
    Call MoveMemory(strText, lpMemory, lngSize)

    ' Truncate it at the first Null character because
    ' the value reported by lngSize is erroneously large
    strText = Left$(strText, InStr(1, strText, Chr$(0)) - 1)

    ' Free the lock
    Call GlobalUnlock(hMemory)

GetTextCloseClipboard:
    ' Close the clipboard
    If Not CBool(CloseClipboard()) Then
        varRet = CVErr(CANNOTCLOSECLIPBOARD)
    End If

GetTextDone:
    If Not IsError(varRet) Then
        GetText = strText
    Else
        GetText = varRet
    End If
End Function


'Sub copiar alguma célula para o clipboard
Sub CopyToClipBoard()
    Dim clip As Variant
    clip = Range("A1").Value
    SetText (clip)
End Sub


'Sub colar do clipboard para alguma célula
Sub PasteFromClipBoard()
   Dim clip As Variant
   clip = GetText()
   Range("A1").Value = clip
End Sub
