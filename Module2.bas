Attribute VB_Name = "Module2"
Global EICC1, EICC2 As Integer 'EICC is a global variable to indicate EICC violation. 1 is violate. 0 is non violation
Global PlanningMode As Boolean
Global newMth, currentMth As String
Global newYr, currentYr As Integer

Sub reset_disable_events()
Dim i As VbMsgBoxResult

i = MsgBox("Enable Events? (Y) to Enable. (N) to Disable.", vbYesNo)
If i = vbYes Then
    Application.EnableEvents = True
Else: Application.EnableEvents = False
End If

End Sub

'Sub Check_ColorIndex()
'MsgBox Selection.DisplayFormat.Interior.ColorIndex
'End Sub
'
'Sub Check_fontColorIndex()
'MsgBox Selection.DisplayFormat.Interior.ColorIndex
'End Sub
'
'Sub chk_cf_ColorIndex()
'MsgBox Selection.DisplayFormat.Interior.ColorIndex
'End Sub
'
'Function ChkFill(CellRng As Range) As Double
'
'    On Error GoTo errhndlr
'    ChkFill = CellRng.DisplayFormat.Interior.ColorIndex
'    On Error GoTo errhndlr2
'    Exit Function
'
'errhndlr:
'MsgBox "Error before execution!"
'Exit Function
'
'errhndlr2:
'MsgBox "Error after execution!"
'
'End Function

Sub UndoCopyPasteVal(ByVal sh As Object, ByVal Target As Range)

Dim UndoList As String
Dim Awkspc, Bwkspc, Cwkspc, Dwkspc As Range

Set Awkspc = ThisWorkbook.Sheets("Day").Range("ATeamWorkspace")
Set Bwkspc = ThisWorkbook.Sheets("Day").Range("BTeamWorkspace")
Set Cwkspc = ThisWorkbook.Sheets("Night").Range("CTeamWorkspace")
Set Dwkspc = ThisWorkbook.Sheets("Night").Range("DTeamWorkspace")

Select Case sh.Name

Case "Day":

If (Not Intersect(Target, Awkspc) Is Nothing) Then
    GoTo hooked

ElseIf (Not Intersect(Target, Bwkspc) Is Nothing) Then
    GoTo hooked
    
End If
    
Case "Night":

If (Not Intersect(Target, Cwkspc) Is Nothing) Then
    GoTo hooked

ElseIf (Not Intersect(Target, Dwkspc) Is Nothing) Then
    GoTo hooked
    
End If

End Select

GoTo LetsContinue

hooked:
                'MsgBox ("A Team Workspace now access undo list.")
                Application.EnableEvents = False
                On Error Resume Next
                '~~> Get the undo List to capture the last action performed by user
                UndoList = Application.CommandBars("Standard").Controls("&Undo").List(1)
                
                '~~> Check if the last action was not a paste nor an autofill
                If Left(UndoList, 5) <> "Paste" And UndoList <> "Auto Fill" Then GoTo LetsContinue
                
                '~~> Undo the paste that the user did but we are not clearing the
                '~~> clipboard so the copied data is still in memory
                Application.Undo
                
                If UndoList = "Auto Fill" Then Selection.Copy
                
                '~~> Do a pastespecial to preserve formats
                
                '~~> Handle text data copied from a website
                Target.Select
                ActiveSheet.PasteSpecial Format:="Text", Link:=False, _
                DisplayAsIcon:=False
 
                Target.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, _
                SkipBlanks:=False, Transpose:=False
                On Error GoTo 0
 
                '~~> Retain selection of the pasted data
                Union(Target, Selection).Select
                    
LetsContinue:
                'Application.EnableEvents = True
                On Error GoTo 0

End Sub


