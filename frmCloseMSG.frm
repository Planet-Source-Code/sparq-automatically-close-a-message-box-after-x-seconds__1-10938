VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdShowMsg 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1440
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long

Private Sub cmdShowMsg_Click()
    SetTimer hWnd, NV_CLOSEMSGBOX, 4000, AddressOf TimerProc
    If MsgBox("Watch this message box close itself after four seconds." & vbCrLf & _
              "The printer is out of paper." & vbCrLf & _
              "Retry or Cancel? (Example)", vbRetryCancel + vbDefaultButton1, "Self Closing Message Box") = vbRetry Then
        MsgBox "Retry!"
    Else
        MsgBox "Cancel"
    End If
  
End Sub
