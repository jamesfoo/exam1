Attribute VB_Name = "exam1_q2sub"
Option Explicit
Option Base 1

'This is the sub to answer question 2 of exam 1
Sub question2()

'define variables as belows:




'error handler
On Error GoTo errhandler

'sub sequences are belows:




Exit Sub

'error handler messages
errhandler:
    MsgBox "error in question 2 sub : " & Err.Description
    Stop
    
End Sub

