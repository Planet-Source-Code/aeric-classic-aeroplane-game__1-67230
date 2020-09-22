Attribute VB_Name = "modFunction"
Option Explicit
'General Function

Public Function GenerateAI() As Integer
    Randomize
    GenerateAI = Rnd() * 20 Mod 6
    
    'Debugging statement
    'GenerateAI = 0
End Function

Public Sub RandomTurn()
    'Set Turn as random
    Dim t As Integer
    
    Randomize
    t = (Rnd() * 400) Mod 4
    
    If t = 0 Then
        Turn = BLUE
    ElseIf t = 1 Then
        Turn = GREEN
    ElseIf t = 2 Then
        Turn = RED
    Else
        Turn = YELLOW
    End If
End Sub

Public Sub AppendText(FileName As String, Optional TextValue As String)
    Open App.Path & "\" & FileName & ".txt" For Append As #1
    If TextValue = "" Then
        Write #1, "[Invalid Value]"
    Else
        Write #1, TextValue
    End If
    Close #1
End Sub

Public Sub ReadText(ByVal FileName As String, ByVal LineNo As Integer, ByRef sOutput As String)
    On Error GoTo newfile
    Dim i As Integer
    Open App.Path & "\" & FileName & ".txt" For Input As #2
        If LineNo < 0 Then
            Do Until EOF(2) = True
                Input #2, sOutput
            Loop
        ElseIf LineNo > 0 Then
            For i = 0 To LineNo
                If Not EOF(2) Then
                    Input #2, sOutput
                Else
                    Exit For
                End If
            Next
        Else
            Input #2, sOutput
        End If
    Close
    Exit Sub
newfile:
    AppendText "Error", Now & " " & " Error: " & Error & " ReadText(" & FileName & ") Line Number: " & LineNo
End Sub

Public Function FileExists(strPath As String) As Boolean
Dim lngRetVal As Long

On Error Resume Next
lngRetVal = Len(Dir$(strPath))
If err Or lngRetVal = 0 Then
    FileExists = False
Else
    FileExists = True
End If
End Function
