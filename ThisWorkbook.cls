VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)

If (ActiveWorkbook.MultiUserEditing = True) And (EICC1 = 1 Or EICC2 = 1) Then
    MsgBox ("EICC Violated BUT Saving is ENABLED for Planning (SHARED File)purposes. Please remove violation ASAP. Please REMEMBER to UNSHARE!")
    Cancel = False
    'SaveAsUI = False

ElseIf (PlanningMode = True) And (EICC1 = 1 Or EICC2 = 1) Then
    MsgBox ("File saved. Planning mode (EICC Violation ignored).")
    Cancel = False
    'SaveAsUI = False

ElseIf (EICC1 = 1 Or EICC2 = 1) Then
    MsgBox ("EICC Violated! Saving is disabled unless violation is removed.")
    Cancel = True
    'SaveAsUI = False
        
Else:
    Cancel = False
    'SaveAsUI = True
    
End If
    
End Sub


Private Sub Workbook_SheetChange(ByVal sh As Object, ByVal Target As Range)

If ActiveWorkbook.MultiUserEditing = True Then
    MsgBox ("Workbook is shared")
    Exit Sub
End If

    Application.ScreenUpdating = False 'disable screen updating to speed up process
    Application.Calculation = xlManual 'disable auto-calc to speed up process
    Application.EnableEvents = False 'IMPORTANT!!! To prevent looping due to automated excel editing
    Application.CutCopyMode = False
    
    If Worksheets("Day").Range("A1").Value = "P" Then
        Application.DisplayStatusBar = True
        Application.StatusBar = "Planning Mode ACTIVE"
        PlanningMode = True
        Worksheets("Day").Range("D12").Value = "Planning Mode"
    End If

'MsgBox sh.Name
    If (Not Intersect(Target, Range("A1")) Is Nothing) And (sh.Name = "Day") Then
        If Worksheets("Day").Range("A1").Value = "P" Then
                Application.DisplayStatusBar = True
                Application.StatusBar = "Planning Mode ACTIVE"
                PlanningMode = True
                Worksheets("Day").Range("D12").Value = "Planning Mode"

        ElseIf Worksheets("Day").Range("A1").Value <> "P" Then
                Application.StatusBar = False
                PlanningMode = False
                Worksheets("Day").Range("D12").Value = ""
        End If
        
    Else: If sh.Name = "Changelog" Then Exit Sub
        
    End If
    
    Call UndoCopyPasteVal(sh, Target)
    
    Select Case PlanningMode
    Case False
        
        'Call UnprotectAll
    
        Call FillColoredBox(sh, Target)
            
        Application.Calculate 'Remember to re-enable Auto calc for find_violate to work correctly!!!
        Call FindViolate(sh)
 
        'Call ProtectAll
    
    Case True
        
        'Call UnprotectAll
    
        Call FillColoredBox(sh, Target)
        
        Application.Calculate 'Remember to re-enable Auto calc for find_violate to work correctly!!!
        Call FindViolate(sh)
        
        'Call ProtectAll
    
    End Select

Application.ScreenUpdating = True 'enable screen updating to return to normal
'Application.CutCopyMode = True
Application.Calculation = xlAutomatic 'reenable auto-calc updating
Application.Calculate
Application.EnableEvents = True  'IMPORTANT!!! To prevent looping due to automated excel editing

End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
reset_commandbars
End Sub

'Private Sub Workbook_Open()
'modify_commandbar
'End Sub

'Private Sub Workbook_WindowActivate(ByVal Wn As Window)
'modify_commandbar
'End Sub

Private Sub Workbook_Deactivate()
reset_commandbars
End Sub

Sub speedupSub()

Application.EnableEvents = False
Application.Calculation = xlCalculationManual
Application.ScreenUpdating = False

End Sub

Sub unspeedSub()

On Error GoTo 0
Application.EnableEvents = True
Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True

End Sub
