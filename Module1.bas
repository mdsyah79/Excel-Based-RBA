Attribute VB_Name = "Module1"
Option Explicit

Sub FillColoredBox(sheetchg As Object, shiftwkspace As Range)

    Dim selectedWorkSpace As Range

    Dim setTeamCol, bteamcol, ateamcol, cteamcol, dteamcol, whitecol, nofillcol, phcol  As Long
    
    Dim confirmdel As Integer
    
   'bteamcol = 35
    ateamcol = 36
    bteamcol = 20
    cteamcol = 43
    dteamcol = 33
    whitecol = 2
    nofillcol = -4142
    phcol = 19
    
    On Error GoTo errhandlr
        
    Select Case sheetchg.Name
    
    Case "Day":
    
    If Not Intersect(shiftwkspace, Range("ATeamWorkspace")) Is Nothing Then
        'MsgBox "A Team Sheet Changed"
        Set selectedWorkSpace = Worksheets("Day").Range("ATeamWorkspace")
        setTeamCol = ateamcol
        'If shiftwkspace.Count = 1 Then
        '    If shiftwkspace.Value = "" And shiftwkspace.DisplayFormat.Interior.ColorIndex = setTeamCol Then
        '        shiftwkspace.Value = "W"
        '    End If
        'Else: Call fillcolorbox2(selectedWorkSpace, setTeamCol)
        'End If
        Call fillcolorbox2(selectedWorkSpace, setTeamCol)
        
            
                        
    ElseIf Not Intersect(shiftwkspace, Range("BTeamWorkspace")) Is Nothing Then
        'MsgBox "B Team Sheet Changed"
        Set selectedWorkSpace = Worksheets("Day").Range("BTeamWorkspace")
        setTeamCol = bteamcol
        'If shiftwkspace.Count = 1 Then
        '    If shiftwkspace.Value = "" And shiftwkspace.DisplayFormat.Interior.ColorIndex = setTeamCol Then
        '        shiftwkspace.Value = "W"
        '    End If
        'Else: Call fillcolorbox2(selectedWorkSpace, setTeamCol)
        'End If
        Call fillcolorbox2(selectedWorkSpace, setTeamCol)

                
'    ElseIf Not Intersect(shiftwkspace, Range("MonthYearSpace")) Is Nothing Then
'        confirmdel = MsgBox("OT Scheduling in progress. Do you really want to clear ALL OT DATA?", vbYesNo, "Confirm Data Delete.")
'        If confirmdel = vbNo Then Exit Sub
'
'        Call unHideFrontBackMonth
'
'        On Error Resume Next
'        Set selectedWorkSpace = Union(Worksheets("Day").Range("BTeamWorkspace"), Worksheets("Day").Range("ATeamWorkspace"))
'
'            With selectedWorkSpace
'                .ClearContents
'                .ClearComments
'            End With
'
'            With selectedWorkSpace.Interior
'             .Pattern = xlNone
'             .TintAndShade = 0
'             .PatternTintAndShade = 0
'            End With
'
'        Set selectedWorkSpace = Union(Worksheets("Night").Range("CTeamWorkspace"), Worksheets("Night").Range("DTeamWorkspace"))
'
'            With selectedWorkSpace
'                .ClearContents
'                .ClearComments
'            End With
'
'            With selectedWorkSpace.Interior
'             .Pattern = xlNone
'             .TintAndShade = 0
'             .PatternTintAndShade = 0
'            End With
'
'        Set selectedWorkSpace = Worksheets("Day").Range("ATeamWorkspace")
'            setTeamCol = ateamcol
'        Application.DisplayStatusBar = True
'        Application.StatusBar = "Filling I in A Shift Cells..."
'        Call fillcolorbox2(selectedWorkSpace, setTeamCol)
'
'
'        Set selectedWorkSpace = Worksheets("Day").Range("BTeamWorkspace")
'            setTeamCol = bteamcol
'        Application.DisplayStatusBar = True
'        Application.StatusBar = "Filling I in B Shift Cells..."
'        Call fillcolorbox2(selectedWorkSpace, setTeamCol)
'
'        Application.Worksheets("Night").Activate
'
'        Set selectedWorkSpace = Worksheets("Night").Range("CTeamWorkspace")
'            setTeamCol = cteamcol
'        Application.DisplayStatusBar = True
'        Application.StatusBar = "Filling I in C Shift Cells..."
'        Call fillcolorbox2(selectedWorkSpace, setTeamCol)
'
'        Set selectedWorkSpace = Worksheets("Night").Range("DTeamWorkspace")
'            setTeamCol = dteamcol
'        Application.DisplayStatusBar = True
'        Application.StatusBar = "Filling I in D Shift Cells..."
'        Call fillcolorbox2(selectedWorkSpace, setTeamCol)
'        Application.DisplayStatusBar = False
'
'        Application.Worksheets("Day").Activate
'        Call HideFrontBackMonth
        
    Else: Exit Sub 'MsgBox "No known area in Day Sheet."
        End If
    
    Case "Night":
        If Not Intersect(shiftwkspace, Worksheets("Night").Range("CTeamWorkspace")) Is Nothing Then
            'MsgBox "C Team Sheet Changed"
            Set selectedWorkSpace = Worksheets("Night").Range("CTeamWorkspace")
            setTeamCol = cteamcol
        'If shiftwkspace.Count = 1 Then
        '    If shiftwkspace.Value = "" And shiftwkspace.DisplayFormat.Interior.ColorIndex = setTeamCol Then
        '        shiftwkspace.Value = "W"
        '    End If
        'Else: Call fillcolorbox2(selectedWorkSpace, setTeamCol)
        '    End If
        Call fillcolorbox2(selectedWorkSpace, setTeamCol)
        
            
        
        ElseIf Not Intersect(shiftwkspace, Worksheets("Night").Range("DTeamWorkspace")) Is Nothing Then
            'MsgBox "D Team Sheet Changed"
            Set selectedWorkSpace = Worksheets("Night").Range("DTeamWorkspace")
            setTeamCol = dteamcol
        'If shiftwkspace.Count = 1 Then
        '    If shiftwkspace.Value = "" And shiftwkspace.DisplayFormat.Interior.ColorIndex = setTeamCol Then
        '        shiftwkspace.Value = "W"
        '    End If
        'Else: Call fillcolorbox2(selectedWorkSpace, setTeamCol)
            'End If
        Call fillcolorbox2(selectedWorkSpace, setTeamCol)
        
        
        Else: Exit Sub 'MsgBox "No known area in Night Sheet."
    
        End If
    
    Case Else:
        Exit Sub 'MsgBox "Not within any known range"
    
    End Select
    
Range(shiftwkspace.Address).Select
Unload UserForm1 ' close userform progress bar
Exit Sub

errhandlr:
If Err.Number = 1004 Then
    Sheets(sheetchg.Name).Activate
    Resume Next
Else
    Application.EnableEvents = True
    confirmdel = MsgBox _
    (("The OT Schedule Planning Excel Macro encountered an error based on your last actions. Please email to syah.muslim@seagate.com or Whatsapp 96755114 for help."), _
    vbApplicationModal, "OT EICC Planning Error Handler")
    Exit Sub
End If

Application.DisplayStatusBar = False
End Sub

Sub fillcolorbox2(selectedWorkSpace1 As Range, ByVal cellcolorindex As Long)

Dim rCell As Range
Dim progress_bar_increment_width As Double
Dim counti, progress_bar_count As Long

counti = 0
    
selectedWorkSpace1.Select
With Selection
    .Font.Size = 12
End With
'selectedWorkSpace1.SpecialCells(xlCellTypeBlanks).Select 'select only empty cells to speed up processing time

'##### Progress Bar Module ###
progress_bar_count = Selection.Cells.Count 'Count how many cells selected to feed values to progress bar
progress_bar_increment_width = set_progress_bar_increment(progress_bar_count) 'call function to set progres bar width per increment
Call show_progressbar
Application.ScreenUpdating = True
Application.Wait (Now + TimeValue("00:00:01"))
Application.ScreenUpdating = False
'Application.ScreenUpdating = False
'Progress Bar module
  
        For Each rCell In Selection
        counti = counti + 1
        Call update_progress_bar(progress_bar_increment_width, counti)   'Call to update progress bar
                
            If (rCell.DisplayFormat.Interior.ColorIndex = cellcolorindex) And (rCell.Value = Empty) Then    '--> to use this must remove the "then" <--- Or rCell.DisplayFormat.Interior.ColorIndex = phcol Then

                'If (rCell.Value = "") Then

                    rCell.Font.ColorIndex = cellcolorindex
                    rCell.Value = "W"
                                        

                'End If

            ElseIf (rCell.Value = Empty) Then
                rCell.ClearContents
    
            End If

        Next rCell
        
Unload UserForm1 ' close userform progress bar
Application.ScreenUpdating = True
Application.Wait (Now + TimeValue("00:00:01"))
Application.ScreenUpdating = False
'Application.ScreenUpdating = False

End Sub
'##### Progress Bar Component ###
Sub show_progressbar()

On Error Resume Next
UserForm1.Show
UserForm1.Bar.Width = 0
Application.ScreenUpdating = True
Application.ScreenUpdating = False
Return 'IMPT!!! Please retain this to ensure the sub returns back to main sub.

End Sub
'##### Progress Bar Component ###
Function set_progress_bar_increment(ByVal cell_count As Long) As Double

set_progress_bar_increment = 204 / cell_count

End Function
'##### Progress Bar Component ###
Sub update_progress_bar(increment_width As Double, ByVal counti As Long)

Dim output_percentage As Single
Dim new_width, control_width As Integer

control_width = 204

new_width = increment_width * counti
output_percentage = 100 / 204 * new_width
output_percentage = Int(output_percentage)
UserForm1.Bar.Width = new_width
UserForm1.Bar.Caption = output_percentage & "%"

DoEvents

End Sub
'##### Find Violation in EICC Check Area ###
Sub FindViolate(sheetchg As Object)

Dim rFound As Range

Application.DisplayStatusBar = True
Application.StatusBar = "Auditing Violation..."

'Select Case sheetchg.Name

'Case "Day":
'On Error Resume Next
    With Worksheets("Day").Range("DayTeamEICCspace")

        Set rFound = .Find("VIOLATE", LookIn:=xlValues)
    
        If Not rFound Is Nothing Then
            If PlanningMode = False Then
                MsgBox ("EICC Violated Day Shift! File Saving is disabled.")
                EICC1 = 1
            Else:
                MsgBox ("EICC Violated BUT SAVING is ENABLED for Planning.")
                EICC1 = 1
            End If
        Else: EICC1 = 0
        End If

    End With

'Case "Night":

'########## BLOCK OF CODE FOR JANUARY OT FILE####################################
    With Worksheets("Night").Range("NightTeamEICCspace")

        Set rFound = .Find("VIOLATE", LookIn:=xlValues)

        If Not rFound Is Nothing Then
            MsgBox ("EICC Violated Night Shift! File Saving is disabled.")
            EICC2 = 1

        Else: EICC2 = 0
        End If

    End With
'################################################################################
On Error GoTo 0
'End Select

Application.DisplayStatusBar = False
End Sub
' ##### Hide or show Columns WW0, WW00 and WW7 ###
Sub HideFrontBackMonth()

Dim sheet As Worksheet
Dim day, i, findresult As Range
Dim frontdate, useddate, idate, currentdate, backdate As Date
Dim y, m, d1, d2, currentyear, j As Integer
Dim currentmonth As String

currentmonth = Worksheets("Day").Range("$B$13")
currentyear = Worksheets("Day").Range("$C$13")

useddate = DateValue(currentyear & "-" & currentmonth & "-" & "1")

'MsgBox useddate

y = Year(useddate)
m = Month(useddate)
d1 = 23
frontdate = DateSerial(y, m - 1, 23)
backdate = DateSerial(y, m + 1, 0)

'MsgBox frontdate
'MsgBox backdate

'################### Hide Front Script#######################
        With Worksheets(1).Range("$E$16:$K$16")
            Set findresult = .Find(Format(frontdate, "dd"), LookIn:=xlValues)
            If findresult Is Nothing Then 'IF FIND UNABLE TO FIND DATE
                'MsgBox findresult.Value
                Worksheets(1).Range("$E$16:$K$16").EntireColumn.Hidden = True
                Worksheets(2).Range("$E$16:$K$16").EntireColumn.Hidden = True
                Worksheets(1).Range("$BI$16").EntireColumn.Hidden = True
                Worksheets(2).Range("$BI$16").EntireColumn.Hidden = True
            Else: GoTo ENDHIDEFRONT
            End If
        End With
        
        With Worksheets(1).Range("$L$16:$R$16")
            Set findresult = .Find(Format(frontdate, "dd"), LookIn:=xlValues)
            If findresult Is Nothing Then
                'MsgBox findresult.Value
                Worksheets(1).Range("$L$16:$R$16").EntireColumn.Hidden = True
                Worksheets(2).Range("$L$16:$R$16").EntireColumn.Hidden = True
                Worksheets(1).Range("$BJ$16").EntireColumn.Hidden = True
                Worksheets(2).Range("$BJ$16").EntireColumn.Hidden = True
            End If
        End With
ENDHIDEFRONT:
        
'################## Hide Back Script ###########################
        With Worksheets(1).Range("$BB$16:$BH$16")
            Set findresult = .Find(Format(backdate, "dd"), LookIn:=xlValues)
            If findresult Is Nothing Then
                'MsgBox findresult.Value
                Worksheets(1).Range("$BB$16:$BH$16").EntireColumn.Hidden = True
                Worksheets(2).Range("$BB$16:$BH$16").EntireColumn.Hidden = True
                Worksheets(1).Range("$BP$16").EntireColumn.Hidden = True
                Worksheets(2).Range("$BP$16").EntireColumn.Hidden = True
            End If
        End With


'        For Each i In Worksheets(1).Range("$e$16:$k$16")
'            idate = i.Value
'            MsgBox idate
            'd2 = Format(idate, "d")
            'MsgBox d2
            'MsgBox day(frontdate)
            
'            If idate <> frontdate Then
'                Worksheets(1).Range("$e$16:$k$16").EntireColumn.Hidden = True
'            End If
'        Next i

End Sub

Sub unHideFrontBackMonth()

Dim sht As Worksheet
Dim i, j As Range
'Dim k As Integer

'For Each sht In Worksheets
'MsgBox sht.Name
'    If sht.Index = 3 Then Exit Sub
Set sht = Worksheets("Day")
    Set j = sht.Range("$E:$BP")
'    j.Activate
'    On Error Resume Next
    j.EntireColumn.Hidden = False
    
Set sht = Worksheets("Night")
    Set j = sht.Range("$E:$BP")
'    On Error Resume Next
    j.EntireColumn.Hidden = False
    
    'If j.Columns.Hidden = True Then
    '    j.EntireColumn.Hidden = False
    'End If
        'For Each i In j
        '    If i.Columns.Hidden = True Then
        '        i.EntireColumn.Hidden = False
        '    End If
        'Next i
'Next sht
    
End Sub

Sub unHideTopBottomNamesRow()

Dim sht As Worksheet
Dim i, j As Range
Dim k As VbMsgBoxResult

k = MsgBox("Unhide All sheets? Select (No) for current sheet only?", vbYesNo)

If k = vbNo Then
Set sht = ActiveSheet
    Set j = sht.Range("17:90")
    j.EntireRow.Hidden = False
    
Else
Set sht = Worksheets("Day")
    Set j = sht.Range("17:90")
    j.EntireRow.Hidden = False

Set sht = Worksheets("Night")
    Set j = sht.Range("17:90")
    j.EntireRow.Hidden = False
End If
    
    
End Sub

Sub UnprotectAll()
    Dim sh As Worksheet
    Dim yourPassword As String
    yourPassword = "otadmin"

If Not ActiveWorkbook.MultiUserEditing Then ' disable sheet protect if sharing mode enabled
    For Each sh In ActiveWorkbook.Worksheets
        If Sheets(sh.Name).ProtectContents = True Then
            sh.Unprotect Password:=yourPassword
        End If
    Next sh
End If

End Sub

Sub ProtectAll()
Dim sh As Worksheet
Dim yourPassword As String
    yourPassword = "otadmin"

If Not ActiveWorkbook.MultiUserEditing Then ' disable sheet protect if sharing mode enabled
'    If PlanningMode = False Then
        For Each sh In ActiveWorkbook.Worksheets
            sh.Protect Password:=yourPassword
        Next sh
'    End If
End If

End Sub

