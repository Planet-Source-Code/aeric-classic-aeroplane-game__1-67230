Attribute VB_Name = "modAI"
Option Explicit

'AI Logic
'1 After NextTurn, Set thinktime
'2 Delay a while then Click Button
'2 Blink Plane while delaying
'3 Select AI
'4 Fly or Move
'5 Next Turn

'For AI Player
'If return > 0 then go else nextturn
Public Function AI(AINum As Integer) As Integer
On Error GoTo AI_Err
'Need Initialize?
'Go = 0
Select Case AINum
    Case 0 '(Stupid Random AI) Tested: No
        AI = AI0 'Level or Type 0
    Case 1 '(Smallest Number AI)
        AI = AI1
    Case 2 '(Smarter AI) Tested: No
        AI = AI2
    Case 3 '(Cruel AI) Tested: No
        AI = AI3
    Case 4 '(Unpatient AI) Tested: No
        AI = AI4
    Case 5 '(Most clever AI) Tested: No
        AI = AI5
    Case Else
        AI = AI1 'Default AI
End Select
'Debug.Print "AI (" & Turn & ") Case " & AINum & " Go = " & Go
Exit Function
AI_Err:
    'AI 5 got many bugs
    MsgBox Error & Chr(13) & "AI (" & Player(Turn) & ") Case " & AINum, vbExclamation, "AI"
End Function

Private Function AI0() As Integer
    Dim a As Integer
    'AI that randomly choose plane (Random)
    'If DieNumber = 6 then it will choose any plane
    'If not 6 then it will randomly choose available plane
        Randomize
        a = (10 * Rnd() Mod 4) + 1
        If frmGame.IsBlinkingPlane(a) = True Then
            AI0 = a
        Else 'if random plane is False
            AI0 = AI1 'choose smallest number
        End If
End Function

Private Function AI1() As Integer
    'Choose smallest PlaneNum
    If frmGame.IsBlinkingPlane(1) = True Then
        AI1 = 1
    ElseIf frmGame.IsBlinkingPlane(2) = True Then
        AI1 = 2
    ElseIf frmGame.IsBlinkingPlane(3) = True Then
        AI1 = 3
    ElseIf frmGame.IsBlinkingPlane(4) = True Then
        'Need to check blncolour(blue,4)? Yes
        AI1 = 4
    Else
        'NextTurn
        AI1 = 0
    End If
End Function

Private Function AI2() As Integer
    'AI that will depart plane left in Base(smallest PlaneNum)
    'This means it must depart at least one plane
    'If all planes are departed then it will choose behind plane
    'This is opposite of Case 4 (If can win also doesn't wan)
    If Steps = 6 Then
        If frmGame.IsInsideBase(1) = True Then
            AI2 = 1
        ElseIf frmGame.IsInsideBase(2) = True Then
            AI2 = 2
        ElseIf frmGame.IsInsideBase(3) = True Then
            AI2 = 3
        ElseIf frmGame.IsInsideBase(4) = True Then
            AI2 = 4
        Else 'All out already
            If frmGame.IsGoal(1) = False Then
                AI2 = 1
            ElseIf frmGame.IsGoal(2) = False Then
                AI2 = 2
            ElseIf frmGame.IsGoal(3) = False Then
                AI2 = 3
            ElseIf frmGame.IsGoal(4) = False Then
                AI2 = 4
            Else
                AI2 = 0 'frmGame.Smallest
            End If
        End If
    Else
        AI2 = frmGame.Smallest
    End If
'    If AI2 = 0 Then AI2 = AI1
'        Debug.Print Go
End Function

Private Function AI3() As Integer
    'AI that like to kick player
    'It will check if there is any plane in front to kick
    'Check DieNumber
    'But if no plane to kick, it follows AI Case 0 or Case 1
    'If Green(1) - Blue(i) < 6 And Green(1) - Blue(i) > 0 Then
    'But if jump then cannot kick
    'Also need to check if plane is enabled
    
    AI3 = frmGame.CheckAnyKick
'    If AI3 = 0 Then AI3 = AI1
    'if no plane to kick, it will try to jump
    If AI3 = 0 Then AI3 = frmGame.CheckAnyJump
    If AI3 = 0 Then AI3 = AI1
End Function

Private Function AI4() As Integer
    'AI that like to win early
    'Always choose in front position plane
    'Select Case Turn 'Turn
'        If frmGame.Largest() > 0 Then 'at least got a plane is enabled
'            AI4 = frmGame.Largest
'        Else
'            Exit Function 'No plane is enabled
'        End If
    'End Select
    
    AI4 = frmGame.Largest
    If AI4 = 0 Then
        If frmGame.IsBlinkingPlane(1) = True Then
            AI4 = 1
        ElseIf frmGame.IsBlinkingPlane(2) = True Then
            AI4 = 2
        ElseIf frmGame.IsBlinkingPlane(3) = True Then
            AI4 = 3
        ElseIf frmGame.IsBlinkingPlane(4) = True Then
            AI4 = 4
        Else
            AI4 = 0
        End If
    End If
'    If AI4 = 0 Then AI4 = AI1
    
End Function

Private Function AI5() As Integer
    'AI first check any plane left in base (similar to Case 2)
    'then it will see whether shortcut is possible
    'or it will see if jump is possible
    'But this AI does not like kick player
    If Steps = 6 Then
        If frmGame.IsInsideBase(1) = True Then
            AI5 = 1
        ElseIf frmGame.IsInsideBase(2) = True Then
            AI5 = 2
        ElseIf frmGame.IsInsideBase(3) = True Then
            AI5 = 3
        ElseIf frmGame.IsInsideBase(4) = True Then
            AI5 = 4
        Else 'All out already
            AI5 = frmGame.CheckAnyShortcut
            If AI5 = 0 Then AI5 = frmGame.CheckAnyJump
            If AI5 = 0 Then AI5 = AI1
        End If
    Else
        AI5 = frmGame.CheckAnyShortcut
        If AI5 = 0 Then AI5 = frmGame.CheckAnyJump
        If AI5 = 0 Then AI5 = AI1
    End If
'    If AI5 = 0 Then AI5 = AI0
End Function
