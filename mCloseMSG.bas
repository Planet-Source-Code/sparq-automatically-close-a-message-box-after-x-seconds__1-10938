Attribute VB_Name = "Module1"
Private Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long

'// Message we receive telling us to close the message box
Public Const NV_CLOSEMSGBOX As Long = &H5000&

Public Sub TimerProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTime As Long)
    
    'this is a callback function.  This means that windows "calls back" to this function
    'when it's time for the timer event to fire
    'first thing we do is kill the timer so that no other timer events will fire
    KillTimer hWnd, idEvent
    
    'select the type of manipulation that we want to perform
    Select Case idEvent
    Case NV_CLOSEMSGBOX '// we want to close this messagebox after 4 seconds
        Dim hMessageBox As Long
        
        'find the messagebox window
        'change the text to whatever the title of the message box is
        hMessageBox = FindWindow("#32770", "Self Closing Message Box")
        
        'if we found it make sure it has the keyboard focus and then send it an enter to dismiss it
        If hMessageBox Then
            Call SetForegroundWindow(hMessageBox)
            
            'this will result in the default option being chosen
            SendKeys "{enter}"
        End If
    End Select
End Sub
