Attribute VB_Name = "exam1_q4sub"
Option Explicit
Option Base 1

'This is the sub to answer question 4 of exam 1
Sub question4()

'define variables as belows:




'error handler
On Error GoTo errhandler

'sub sequences are belows:




Exit Sub

'error handler messages
errhandler:
    MsgBox "error in question 4 sub : " & Err.Description
    Stop
    
End Sub

