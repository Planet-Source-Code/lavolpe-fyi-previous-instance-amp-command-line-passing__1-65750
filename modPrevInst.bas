Attribute VB_Name = "modPrevInst"
Option Explicit

Private Declare Function EnumWindows Lib "user32.dll" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function SetProp Lib "user32.dll" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Private Declare Function GetProp Lib "user32.dll" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long

Private prevInstPropName As String  ' unique property value passed to IsPrevInstance
Private prevInstHwnd As Long        ' this will be the previous instance hWnd, if any
Private prevInstPropValue As Long   ' this will be the property value stored against the previous hWnd

Public Function IsPrevInstance(ByVal hWnd As Long, ByVal PropName As String, _
                        Optional ByRef propValue As Long = 1, _
                        Optional ByVal passCmdLineToPropValue As Boolean) As Long
    
    ' hWnd [in]: must be the hWnd of the instance being created (Me.hWnd)
    ' PropName [in]: unique property name, no other winodw should be using
    ' propValue [in/out]: the property value to be associated with this instance
    '            to pass command line parameters from other instances, suggest
    '            passing a textbox hWnd. This value is for your use, it must not be 0
    ' passCmdLineToPropValue [in]: since it is a common practice to pass the command line
    '   to the 1st instance when this is the 2nd instance, setting that parameter
    '   to True will allow this routine to pass the command line for you.
    
    ' Note that propValue is ByRef; therefore it can be changed by this function.
    ' Do not pass Read-Only values (i.e., do not pass Text1.hWnd as propValue)
    
    ' Return Values:
    ' If this is the first instance
    ' - function returns zero
    ' - The propValue is assigned to the PropName property
    ' - passCmdLineToPropValue is ignored
    
    ' If this is another instance...
    ' - function returns hWnd of the previous instance
    ' - propValue is the custom value set by the previous instance
    ' - if passCmdLineToPropValue, then function sends the Command$ variable
    '   to the value in propValue. Of course, this should be an hWnd of
    '   a textbox or other window that has a Text property.
    
    
    ' sanity checks
    If Trim$(PropName) = "" Then Exit Function
    If hWnd = 0 Then Exit Function
    ' should you want to add even more sanity checks, consider using
    ' the IsWindow API on passed hWnd, and also on the propValue just
    ' before using SendMessage below
    
    Const WM_SETTEXT As Long = &HC
    prevInstHwnd = 0    ' reset
    prevInstPropName = PropName ' set here so EnumWindowsProc can use it
    
    ' look for previous instance
    EnumWindows AddressOf EnumWindowsProc, hWnd
    prevInstPropName = vbNullString ' no longer need to waste extra memory
    
    If prevInstHwnd = 0 Then    ' no previous instance found
        ' ensure the property value passed is not zero
        If propValue = 0 Then propValue = 1
        ' now assign the property & value
        SetProp hWnd, PropName, propValue
    Else
        ' we do have a previous instance, set return values
        propValue = prevInstPropValue
        If passCmdLineToPropValue = True Then
            ' ensure a non-zero length string so the target's Change_Event occurs
            ' It is assumed the target will reset its text value to zero-length
            ' when done receiving this message
            SendMessage propValue, WM_SETTEXT, 0&, ByVal Command$ & " "
        End If
        IsPrevInstance = prevInstHwnd
    End If
End Function

Private Function EnumWindowsProc(ByVal hWnd As Long, ByVal lParam As Long) As Long
    
    ' the enumeration stops when the return value is zero or all windows have been enumerated
    prevInstPropValue = GetProp(hWnd, prevInstPropName)
    If prevInstPropValue = 0 Then
        EnumWindowsProc = 1 ' keep enumerating
    Else
        If hWnd = lParam Then
            ' safety check. Should this be called a 2nd time from the previous instance
            ' ensure we don't return true if this is the previous instance
            EnumWindowsProc = 1
        Else
            prevInstHwnd = hWnd
            ' stops enumerating cause we do not set its return value
        End If
    End If
End Function
