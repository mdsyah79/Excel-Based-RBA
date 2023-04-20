Attribute VB_Name = "Module3"
Option Explicit
Sub modify_commandbar()
    Dim oPasteBtns As office.CommandBarControls
    Dim oPasteBtn As office.CommandBarButton
    Dim nPasteBtn, nInfoBtn As office.CommandBarButton

reset_commandbars 'initialis all commandbars

'   Set oPasteBtns = Excel.Application.CommandBars.FindControls(ID:=755)
'    For Each oPasteBtn In oPasteBtns
'        MsgBox oPasteBtn.Caption
'        oPasteBtn.Delete
'    Next
        Set oPasteBtn = Excel.Application.CommandBars("cell").FindControl(ID:=22)
'        Set nPasteBtn = Excel.Application.CommandBars("cell").Controls.Add(Type:=msoControlButton, ID:=3, before:=3)
'            With nPasteBtn
'                .FaceId = 59 'oPasteBtn.ID = old paste button id, 59 = smiley
'                .Caption = oPasteBtn.Caption
'                .TooltipText = "Paste values only"
'                .Style = oPasteBtn.Style
'                .BeginGroup = oPasteBtn.BeginGroup
'                .Tag = oPasteBtn.Tag
'                .OnAction = "PasteValues"
'            End With
            
        Set nInfoBtn = Excel.Application.CommandBars("cell").Controls.Add(Type:=msoControlButton, ID:=4, before:=4)
            With nInfoBtn
                .FaceId = 926 '926 = question mark
                .Caption = "where are my buttons?"
                '.TooltipText = oPasteBtn.TooltipText
                .Style = oPasteBtn.Style
                '.BeginGroup = oPasteBtn.BeginGroup
                '.Tag = oPasteBtn.Tag
                .OnAction = "InfoBtnContent"
            End With
            
On Error Resume Next
        With Application.CommandBars
            .FindControl(ID:=21).Delete ' make "Paste" button invisible
            .FindControl(ID:=22).Visible = False ' make "Paste" button invisible
            .FindControl(ID:=21437).Delete ' make "Paste" button invisible
            .FindControl(ID:=3624).Delete ' make "Paste" button invisible
            .FindControl(ID:=292).Delete ' make "Paste" button invisible
            '.Controls.Item(1).OnAction = "CutWarning" ' Changes right click cell paste to use Pastevalues macro
        End With
        
        With Application
            .OnKey "^c", "CutCopyPasteDisabled"
            .OnKey "^v", "CutCopyPasteDisabled"
            .OnKey "^x", "CutCopyPasteDisabled"
            .OnKey "+{DEL}", "CutCopyPasteDisabled"
            .OnKey "^{INSERT}", "CutCopyPasteDisabled"
        End With
On Error GoTo 0
        

End Sub

Sub reset_commandbars()
Dim x As office.CommandBar
Dim y As office.CommandBars

Set y = Excel.Application.CommandBars
On Error Resume Next
For Each x In y
    x.Reset
    'MsgBox x.Name
Next x


        With Application
            .OnKey "^c"
            .OnKey "^v"
            .OnKey "^x"
            .OnKey "+{DEL}"
            .OnKey "^{INSERT}"
        End With
On Error GoTo 0
'Application.CommandBars("Cell").Reset
'Application.CommandBars("Row").Reset
'Application.CommandBars("Column").Reset
'Application.CommandBars("Worksheet Menu Bar").Reset
'Application.CommandBars("Standard").Reset

End Sub

Sub test()
'code to list out all control bar caption n id for ident n list in Immediate Window

Dim i, c As Integer
Dim message As Object
Dim x As office.CommandBarControl
Dim y As office.CommandBarControl

'c = Application.CommandBars("cell").Controls.Count
Set message = CreateObject("WScript.Shell")

'Set x = Application.CommandBars("cell").FindControl(id:=22).
'For i = 1 To c
'    message.Popup x.Caption & " " & x.ID, 1, "Quick Message"
''    Debug.Print x.Caption & " " & x.ID & " " & x.Index
'Next i

For i = 1 To c
'MsgBox Application.CommandBars(Application.CommandBars("Cell").Controls(i).Tag)
message.Popup Application.CommandBars("Cell").Controls(i).Caption & " " & Application.CommandBars("cell").Controls(i).ID, 1, "Quick Message"
'MsgBox Application.CommandBars("Cell").Controls(i).Caption & " " & Application.CommandBars("cell").Controls(i).Index
Debug.Print Application.CommandBars("Cell").Controls(i).Caption & " " & Application.CommandBars("cell").Controls(i).ID & " " & Application.CommandBars("cell").Controls(i).Index
'    If Application.CommandBars("Cell").Controls(i).BeginGroup = True Then
'        MsgBox Application.CommandBars("Cell").Controls(i).Caption
'    End If
Next

End Sub

Sub PasteValues()
'
' PasteValues Macro
'
' Keyboard Shortcut: Ctrl+v
'
Application.Calculation = xlCalculationManual
Application.CommandBars.ExecuteMso ("PasteValues")

'On Error Resume Next
'Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
        
'On Error GoTo 0
        
'   MsgBox "paste"

End Sub

Sub InfoBtnContent()
Dim x As VbMsgBoxResult
Dim secretpassword, enteredpassword As String
secretpassword = "otadmin"

x = MsgBox("To protect the contents of the file to prevent errors, certain buttons have been disabled." & vbCr & vbCr & _
"To reEnable please press Retry.", vbRetryCancel)
If x = vbRetry Then
    enteredpassword = InputBox("Enter password to reset commandbar.", "Password Entry.")
    If enteredpassword = secretpassword Then
        reset_commandbars
    Else
        MsgBox ("Wrong password...")
    End If
ElseIf x = vbCancel Then
    Exit Sub
End If

End Sub

Sub CutCopyPasteDisabled()

MsgBox "Shortcut Disabled."

End Sub



'        Private Sub ReplacePasteButtons()
'    On Error GoTo Err_Hnd
'    Dim oPasteBtns As Office.CommandBarControls
'    Dim oPasteBtn As Office.CommandBarButton
'    Dim oNewBtn As Office.CommandBarButton

'    Const lIDPaste_c As Long = 22

'    RestorePasteButtons

'    Set oPasteBtns = Excel.Application.CommandBars.FindControls(ID:=lIDPaste_c)

'    For Each oPasteBtn In oPasteBtns
'        Set oNewBtn = oPasteBtn.Parent.Controls.Add(msoControlButton, Before:=oPasteBtn.Index, Temporary:=True)
'        oNewBtn.FaceId = lIDPaste_c
'        oNewBtn.Caption = oPasteBtn.Caption
'        oNewBtn.TooltipText = oPasteBtn.TooltipText
'        oNewBtn.Style = oPasteBtn.Style
'        oNewBtn.BeginGroup = oPasteBtn.BeginGroup
'        oNewBtn.Tag = m_sTag_c
'        oNewBtn.OnAction = m_sPasteProcedure_c
'        oPasteBtn.Visible = False
'    Next
'    Exit Sub
        

    'Application.CommandBars("Row").Reset
'        With Application.CommandBars("Row")
'            .Controls("Paste").OnAction = "PasteValues" ' Changes right click row paste to use Pastevalues macro
'            .Controls("Paste &Special...").Delete ' removes paste secial option
''        End With

    'Application.CommandBars("Column").Reset
'        With Application.CommandBars("Column")
'            .Controls("Paste").OnAction = "PasteValues" ' Changes right click column paste to use Pastevalues macro
'            .Controls("Paste &Special...").Delete ' removes paste secial option
'        End With

    'Application.CommandBars("Worksheet Menu Bar").Reset
'    With Application.CommandBars("Worksheet Menu Bar").Controls
'        .Item(2).Controls(6).OnAction = "PasteValues"  ' changes edit/paste to use Pastevalues macro
'        .Item(2).Controls(7).Enabled = False ' disable paste special
'    End With

    'Application.CommandBars("Standard").Reset
'    With Application.CommandBars("Standard").Controls.Item(12)
'        .Enabled = False
'    End With
