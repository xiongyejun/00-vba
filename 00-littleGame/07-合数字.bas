Attribute VB_Name = "模块1"
Option Explicit
'暂停线索
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public colorArr(1 To 25, 1 To 2) As Long
Public rngGame As Range
Public selectRng As Range
Public ifStart As Boolean
Public myMax As Long
Const iSleep As Long = 5

Public rngMax As Range
Public rngScore As Range

Sub main()

    ifStart = True
    
    Set rngScore = Range("F1")
    rngScore.value = ""
    
    setColor colorArr
    Set rngGame = Range("game")
    Range("B2:H8").Clear
    With rngGame
        .value = "=INT(RAND()*3) + 1"
        .value = .value
    End With
    myMax = Application.WorksheetFunction.Max(rngGame)
    Set rngMax = Range("D1")
    rngMax.value = myMax
    formatRngGame
    
    Set selectRng = Nothing
    
    With rngGame
        .Font.Size = 30
        .Font.Bold = True
        .ColumnWidth = 8
        .RowHeight = 50
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous '边框
    End With

    Application.EnableEvents = True

End Sub

Sub endGame()
    
    ifStart = False
    Erase colorArr
    Set rngGame = Nothing
    Set selectRng = Nothing
    Set rngMax = Nothing
    Set rngScore = Nothing
    Application.EnableEvents = True
End Sub

Function setColor(Arr() As Long)
    Dim i As Long
    
    For i = 1 To 25
        Arr(i, 1) = Cells(i + 1, 1).Interior.ColorIndex
        Arr(i, 2) = Cells(i + 1, 1).Font.ColorIndex
    Next i
End Function

Function formatRngGame()
    Dim rng As Range
     
    For Each rng In rngGame
        With rng
            .Interior.ColorIndex = colorArr(.value, 1)
            .Font.ColorIndex = colorArr(.value, 2)
        End With
    Next rng
    
    Set rng = Nothing
End Function

Function selectSameRng(rng As Range, offsetRow As Long, offsetCol As Long)
    Dim tempRng As Range
    
    Set tempRng = rng.Offset(offsetRow, offsetCol)
    
    If Not Application.Intersect(tempRng, selectRng) Is Nothing Then Exit Function
    
    If tempRng.value = 0 Or tempRng.value <> rng.value Then
        Exit Function
    Else
        Set selectRng = Union(selectRng, tempRng)
        
        selectSameRng tempRng, 0, -1
        selectSameRng tempRng, 0, 1
        selectSameRng tempRng, 1, 0
        selectSameRng tempRng, -1, 0
    End If

    Set tempRng = Nothing
End Function

Function moveRng()
    Dim iRow As Long
    Dim iCol As Long
    Dim i As Long, j As Long, k As Long
    Dim Arr(1 To 10, 1 To 5) As Long
    Dim iMax As Long
    
    iMax = myMax - 2
    If iMax > 6 Then iMax = 6
    
    Randomize
    For i = 1 To 5
        For j = 1 To 5
            Arr(i, j) = Int(Rnd * iMax + 1)
        Next j
    Next i

    For iCol = 3 To 7
        For i = 10 To 6 Step -1
            Arr(i, iCol - 2) = Cells(i - 3, iCol).value
        Next i
        
        
        k = 10
        For iRow = 7 To 3 Step -1
            
            Do Until Arr(k, iCol - 2) > 0
                k = k - 1
            Loop
            
            Cells(iRow, iCol).value = Arr(k, iCol - 2)
            k = k - 1
        Next iRow
    Next iCol

    formatRngGame
End Function

Function gameOver() As Boolean
    Dim rng As Range
    
    For Each rng In rngGame
        Select Case rng.value
            Case rng.Offset(0, -1).value
                gameOver = False
                Exit Function
            Case rng.Offset(0, 1).value
                gameOver = False
                Exit Function
            Case rng.Offset(1, 0).value
                gameOver = False
                Exit Function
            Case rng.Offset(-1, 0).value
                gameOver = False
                Exit Function
        End Select
    Next rng
    
    gameOver = True
    Set rng = Nothing
End Function

Sub PaiMing()
    Dim Arr, i As Integer, Temp As Integer
    
    Arr = Range("J2:K11").value
    
    For i = 1 To 10
        If Arr(i, 1) < rngMax.value Then Exit For
    Next i
    i = i + 1               '和行号对应
    
    If i < 12 Then
    
        Range("J" & i).value = rngMax.value
        Range("K" & i).value = rngScore.value
    
        Range("J2:K11").Font.ColorIndex = 0
        Cells(i, "J").Resize(1, 2).Font.ColorIndex = 3
        
        Do Until i = 11
            i = i + 1
            Range("J" & i).value = Arr(i - 1, 1)
            Range("K" & i).value = Arr(i - 1, 2)
            
        Loop
    
        Application.DisplayAlerts = False
        ThisWorkbook.Save
        Application.DisplayAlerts = True
    End If
  
    Erase Arr
    
End Sub

