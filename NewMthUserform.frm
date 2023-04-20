VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} NewMthUserform 
   Caption         =   "Please select and press submit"
   ClientHeight    =   1080
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3390
   OleObjectBlob   =   "NewMthUserform.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "NewMthUserform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnSubmit_Click()

Set ws = ThisWorkbook.ActiveSheet

currentMth = ws.Range("B13")
currentYr = ws.Range("C13")
evalnewmth = Month(DateValue("01 " & newMth & " " & newYr))
evalcurrentMth = Month(DateValue("01 " & currentMth & " " & currentYr))

affirm = MsgBox("Press ok to proceed to CHANGE NEW MONTH! Press CANCEL to EXIT!", vbOKCancel)

    If affirm = vbCancel Then
        UserForm_Terminate
    
    'Debug.Print Month(DateValue("01 " & newMth & " " & newYr))
    'Debug.Print Month(DateValue("01 " & currentMth & " " & currentYr))
    ElseIf evalnewmth = evalcurrentMth Then
        MsgBox ("Current Month :" & UCase(currentMth) & ": and Selected Month :" & UCase(newMth) & ": is the SAME!")
        MsgBox ("Month and Year Selection Terminated!")
        UserForm_Terminate
        
    Else:
        Application.EnableEvents = False
        ws.Range("B13") = newMth
        ws.Range("C13") = newYr
        Me.Hide
        clearWorkspace
    
    End If
    
        
End Sub

Private Sub cbMonth_Change()

newMth = Me.cbMonth.Value
'Debug.Print newMth

End Sub

Private Sub cbYear_Change()

newYr = Me.cbYear.Value
'Debug.Print newYr

End Sub

Private Sub UserForm_Initialize()

Me.cbMonth.Clear
Me.cbYear.Clear

Set ws = ThisWorkbook.ActiveSheet

For x = 1 To 12
    
    Me.cbMonth.AddItem (MonthName(x))

Next x

thisYear = Year(Now())

Me.cbYear.AddItem (thisYear)
Me.cbYear.AddItem (thisYear + 1)

currentMth = ws.Range("B13")
currentYr = ws.Range("C13")

Me.cbMonth.Text = currentMth
Me.cbYear.Text = currentYr

End Sub

Private Sub UserForm_Terminate()

    Me.Hide
    Unload Me

End Sub
