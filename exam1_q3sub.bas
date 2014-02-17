Attribute VB_Name = "exam1_q3sub"
Option Explicit
Option Base 1

'This is the sub to answer question 3 of exam 1
Sub question3()

'define variables as belows:




'error handler
On Error GoTo errhandler

'sub sequences are belows:




Exit Sub

'error handler messages
errhandler:
    MsgBox "error in question 3 sub : " & Err.Description
    Stop
    
End Sub

