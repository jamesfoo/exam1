Attribute VB_Name = "exam1_q1sub"
Option Explicit
Option Base 1

'This is the sub to answer question 1 of exam 1
Sub question1()

'define variables as belows:
dim i as integer




'error handler
On Error GoTo errhandler

'sub sequences are belows:




Exit Sub

'error handler messages
errhandler:
    MsgBox "error in question 1 sub : " & Err.Description
    Stop
    
End Sub
