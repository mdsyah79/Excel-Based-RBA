Attribute VB_Name = "Module4"
Sub Label55_Click()
    NewMthUserform.Show
End Sub
Sub speedupSub()

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False

End Sub


Sub unspeedSub()

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.Calculate
Application.EnableEvents = True

End Sub

Sub clearWorkspace()

Dim selectedWorkSpace As Range

Dim setTeamCol, bteamcol, ateamcol, cteamcol, dteamcol, whitecol, nofillcol, phcol  As Long
    
Dim confirmdel As VbMsgBoxResult

   'bteamcol = 35
    ateamcol = 36
    bteamcol = 20
    cteamcol = 43
    dteamcol = 33
    whitecol = 2
    nofillcol = -4142
    phcol = 19

speedupSub

        confirmdel = MsgBox("OT Scheduling in progress. Do you really want to clear ALL OT DATA?", vbYesNo, "Confirm Data Delete.")
        If confirmdel = vbCancel Then Exit Sub
        
        Call unHideFrontBackMonth
        
        On Error Resume Next
        Set selectedWorkSpace = Union(Worksheets("Day").Range("BTeamWorkspace"), Worksheets("Day").Range("ATeamWorkspace"))
            
            With selectedWorkSpace
                .ClearContents
                .ClearComments
            End With
            
            With selectedWorkSpace.Interior
             .Pattern = xlNone
             .TintAndShade = 0
             .PatternTintAndShade = 0
            End With
                
        Set selectedWorkSpace = Union(Worksheets("Night").Range("CTeamWorkspace"), Worksheets("Night").Range("DTeamWorkspace"))
            
            With selectedWorkSpace
                .ClearContents
                .ClearComments
            End With
        
            With selectedWorkSpace.Interior
             .Pattern = xlNone
             .TintAndShade = 0
             .PatternTintAndShade = 0
            End With
            
        Set selectedWorkSpace = Worksheets("Day").Range("ATeamWorkspace")
            setTeamCol = ateamcol
        Application.DisplayStatusBar = True
        Application.StatusBar = "Filling I in A Shift Cells..."
        Call fillcolorbox2(selectedWorkSpace, setTeamCol)

        
        Set selectedWorkSpace = Worksheets("Day").Range("BTeamWorkspace")
            setTeamCol = bteamcol
        Application.DisplayStatusBar = True
        Application.StatusBar = "Filling I in B Shift Cells..."
        Call fillcolorbox2(selectedWorkSpace, setTeamCol)
        
        Application.Worksheets("Night").Activate
        
        Set selectedWorkSpace = Worksheets("Night").Range("CTeamWorkspace")
            setTeamCol = cteamcol
        Application.DisplayStatusBar = True
        Application.StatusBar = "Filling I in C Shift Cells..."
        Call fillcolorbox2(selectedWorkSpace, setTeamCol)
        
        Set selectedWorkSpace = Worksheets("Night").Range("DTeamWorkspace")
            setTeamCol = dteamcol
        Application.DisplayStatusBar = True
        Application.StatusBar = "Filling I in D Shift Cells..."
        Call fillcolorbox2(selectedWorkSpace, setTeamCol)
        
        
        Application.DisplayStatusBar = False
        Application.Worksheets("Day").Activate
        Call HideFrontBackMonth

unspeedSub

End Sub

