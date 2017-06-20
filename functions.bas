Option Explicit

Function firstIndex(arr, item) As Integer
    Dim i As Integer
    firstIndex = -1
    For i = LBound(arr) To UBound(arr)
        If arr(i) = item Then
            firstIndex = i
            Exit For
        End If
    Next i
End Function

Function lastIndex(arr, item) As Integer
    Dim i As Integer: i = -1
    lastIndex = -1
    For i = UBound(arr) To LBound(arr) Step -1
        If arr(i) = item Then
            lastIndex = i
            Exit For
        End If
    Next i
End Function
Function removeUselessPath(arr)
    
    removeUselessPath = arr
    Dim item
    Dim newArr()
    Dim i, j As Integer
    Dim iStart, iEnd As Integer
    For i = LBound(arr) To UBound(arr)
        iStart = firstIndex(arr, arr(i))
        iEnd = lastIndex(arr, arr(i))
        If iStart <> iEnd Then
            Set removeUselessPath = Nothing
            Dim iNew As Integer: iNew = 0
            ReDim newArr(UBound(arr) - (iEnd - iStart))
            For j = LBound(arr) To UBound(arr)
                If j <= iStart Or j > iEnd Then
                    'Debug.Print "Arry J:", j, arr(j)
                    newArr(iNew) = arr(j)
                    iNew = iNew + 1
                End If
            Next j
            removeUselessPath = newArr
            Exit For
        End If
    Next i
    
    If (Not newArr) <> -1 Then
        For i = LBound(newArr) To UBound(newArr)
            If firstIndex(newArr, newArr(i)) <> lastIndex(newArr, newArr(i)) Then
                removeUselessPath = removeUselessPath(newArr)
                Exit For
            End If
        Next i
    End If
    
End Function


Function generateCells()
    Application.ScreenUpdating = False
    Dim ws As Worksheet: Set ws = Sheet1
    Dim rStart As Integer: rStart = Range("cellStart").Row
    Dim cStart As Integer: cStart = Range("cellStart").Column
    ws.Cells.Clear
    ws.Cells.Interior.Color = vbWhite
    ws.Columns.ColumnWidth = cellWidth
    ws.Rows.RowHeight = cellWidth * ratio
    
    With ws.Range(Range("cellStart"), Range("cellStart").Offset(mazeSize - 1, mazeSize - 1))
        .Borders.LineStyle = 9
        .Interior.Color = unvisited
    End With
    Range("cellStart").Borders(xlEdgeTop).LineStyle = xlNone
    Range("cellStart").Offset(mazeSize - 1, mazeSize - 1).Borders(xlEdgeBottom).LineStyle = xlNone
    Set ws = Nothing
    Application.ScreenUpdating = True
End Function


Function hasUnvisitedCells() As Boolean
    hasUnvisitedCells = False
    Dim ws As Worksheet: Set ws = Sheet1
    Dim rStart As Integer: rStart = Range("cellStart").Row
    Dim cStart As Integer: cStart = Range("cellStart").Column
    Dim rng As Range
    For Each rng In Range(ws.Cells(rStart, cStart), ws.Cells(rStart + mazeSize - 1, cStart + mazeSize - 1))
        If rng.Interior.Color = unvisited Then
            hasUnvisitedCells = True
            Exit For
        End If
    Next
    Set ws = Nothing
End Function

Function getNeighbours(rng As Range)
    Dim neighbour
    Dim neighbours(1 To 4)
    Dim unvisitedNeighbours()
    
    'Neighbour Top
    If rng.Row <> 1 Then
        If rng.Offset(-1, 0).Interior.Color = unvisited Then
            Set neighbours(1) = rng.Offset(-1, 0)
        Else
            neighbours(1) = False
        End If
    End If
    
    'Neighbour Bottom
    If rng.Offset(1, 0).Interior.Color = unvisited Then
        Set neighbours(2) = rng.Offset(1, 0)
    Else
        neighbours(2) = False
    End If
    
    'Neighbour Left
    If rng.Column <> 1 Then
        If rng.Offset(0, -1).Interior.Color = unvisited Then
            Set neighbours(3) = rng.Offset(0, -1)
        Else
            neighbours(3) = False
        End If
    End If
    'Neighbour Right
    If rng.Offset(0, 1).Interior.Color = unvisited Then
        Set neighbours(4) = rng.Offset(0, 1)
    Else
        neighbours(4) = False
    End If
    
    Dim arrSize As Integer: arrSize = 1
    For Each neighbour In neighbours
        If TypeName(neighbour) = "Range" Then
            ReDim Preserve unvisitedNeighbours(1 To arrSize)
            Set unvisitedNeighbours(arrSize) = neighbour
            arrSize = arrSize + 1
        End If
    Next
    'If (Not unvisitedNeighbours) = -1 Then
   
    getNeighbours = unvisitedNeighbours
End Function


Function randomNeighbour(neighbours)
    Randomize
    Dim randomIndex As Integer: randomIndex = Int(UBound(neighbours) * rnd) + 1
    Set randomNeighbour = neighbours(randomIndex)
End Function

Function removeWall(currentCell, neighbour)
    If currentCell.Row = neighbour.Row Then
        If currentCell.Column > neighbour.Column Then   'Neighbour on left
            currentCell.Borders(xlEdgeLeft).LineStyle = xlNone
        Else    'Neighbour on right
            currentCell.Borders(xlEdgeRight).LineStyle = xlNone
        End If
    Else
        If currentCell.Row > neighbour.Row Then 'Neighbour on top
            currentCell.Borders(xlEdgeTop).LineStyle = xlNone
        Else    'Neighbour in bottom
            currentCell.Borders(xlEdgeBottom).LineStyle = xlNone
        End If
    End If

End Function
