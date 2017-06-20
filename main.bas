Option Explicit
Public Const mazeSize = 20
Public Const cellWidth = 4
Public Const ratio = 5.6
Public Const visited = vbGreen
Public Const unvisited = vbYellow
Public Const mazePath = vbCyan
Public Const actualPath = vbYellow
Public Const refreshFrame = 0.001
Sub run()
    mazeGenerator
    pathFinder
End Sub

Sub mazeGenerator()
    Dim arrCells
    arrCells = generateCells
    Dim currentCell As Range: Set currentCell = Range("cellStart")
    currentCell.Select
    currentCell.Interior.Color = visited
    
    Dim neighbour, neighbours
    
    Dim route
    Set route = New Collection
    Dim dTime As Double
    
    Do While hasUnvisitedCells
            'Debug.Print hasUnvisitedCells
loopStart:
    dTime = Timer
    While Timer - dTime < refreshFrame
        DoEvents
    Wend
            On Error GoTo afei:
            neighbours = getNeighbours(currentCell)
            
            If UBound(neighbours) > 0 Then
                'Debug.Print "Has neighbours"
                'Step 1: Choose randomly one of the unvisited neighbours
                Set neighbour = randomNeighbour(neighbours)
                
                'Step 2: Push the current cell to the stack
                route.Add currentCell
                
                'Step 3: Remove the wall between the current cell and the chosen cell
                Call removeWall(currentCell, neighbour)
                
                'Step 4: Make the chosen cell the current cell and mark it as visited
                Set currentCell = neighbour
                currentCell.Select
                currentCell.Interior.Color = visited
            Else
afei:
                'MsgBox "You find the end of the maze"
                
                'Step 1: Pop a cell from the stack
                'Step 2: Make it the current cell
                
                If route.Count > 0 Then
                    Set currentCell = route.item(route.Count)
                    currentCell.Select
                    route.Remove (route.Count)
                Else
                    Set route = Nothing
                    Exit Sub
               End If
               Resume loopStart
            End If
            
    Loop
    
    Range("A1").Select
    Set route = Nothing
End Sub

'Maze Path finder                                                
Sub pathFinder()
    Dim ws As Worksheet: Set ws = Sheet1
    Dim cellPath As New Collection
    
    Dim currentCell As Range: Set currentCell = Range("cellStart")
    cellPath.Add currentCell
    currentCell.Interior.Color = vbRed
    currentCell.Value = "S"
    currentCell.Select
    Dim endCell As Range: Set endCell = currentCell.Offset(mazeSize - 1, mazeSize - 1)
    
    
    While currentCell.Address <> endCell.Address
        Set currentCell = cellMove(currentCell, cellPath)
        cellPath.Add currentCell
        currentCell.Select
        currentCell.Interior.Color = mazePath
    Wend
    endCell.Interior.Color = vbRed
    
    Call mapPath(cellPath)
End Sub

Function cellMove(c As Range, cellPath)
    Dim nextCell As Range
    Dim dir As String: dir = c.Text
    If True Then
        If dir = "S" Then
            If c.Borders(xlEdgeLeft).LineStyle = xlNone Then     'Turn right
                    Set nextCell = c.Offset(0, -1)
                    nextCell.Value = "W"
            Else
                If c.Borders(xlEdgeBottom).LineStyle = xlNone Then 'Go forward
                    Set nextCell = c.Offset(1, 0)
                    nextCell.Value = "S"
                ElseIf c.Borders(xlEdgeRight).LineStyle = xlNone Then  'Turn left
                    Set nextCell = c.Offset(0, 1)
                    nextCell.Value = "E"
                Else    'Go back
                    Set nextCell = c.Offset(-1, 0)
                    nextCell.Value = "N"
                End If
            End If
        ElseIf dir = "W" Then
            If c.Borders(xlEdgeTop).LineStyle = xlNone Then     'Turn right
                    Set nextCell = c.Offset(-1, 0)
                    nextCell.Value = "N"
            
            Else
                If c.Borders(xlEdgeLeft).LineStyle = xlNone Then 'Go forward
                    Set nextCell = c.Offset(0, -1)
                    nextCell.Value = "W"
                ElseIf c.Borders(xlEdgeBottom).LineStyle = xlNone Then  'Turn left
                    Set nextCell = c.Offset(1, 0)
                    nextCell.Value = "S"
                Else    'Go back
                    Set nextCell = c.Offset(0, 1)
                    nextCell.Value = "E"
                End If
            End If
        
        ElseIf dir = "N" Then
            If c.Borders(xlEdgeRight).LineStyle = xlNone Then     'Turn right
                    Set nextCell = c.Offset(0, 1)
                    nextCell.Value = "E"
            Else
                If c.Borders(xlEdgeTop).LineStyle = xlNone Then 'Go forward
                Set nextCell = c.Offset(-1, 0)
                nextCell.Value = "N"
                ElseIf c.Borders(xlEdgeLeft).LineStyle = xlNone Then  'Turn left
                    Set nextCell = c.Offset(0, -1)
                    nextCell.Value = "W"
                Else    'Go back
                    Set nextCell = c.Offset(1, 0)
                    nextCell.Value = "S"
                End If
            End If
        Else
            If c.Borders(xlEdgeBottom).LineStyle = xlNone Then     'Turn right
                    Set nextCell = c.Offset(1, 0)
                    nextCell.Value = "S"
            
            Else
                If c.Borders(xlEdgeRight).LineStyle = xlNone Then 'Go forward
                    Set nextCell = c.Offset(0, 1)
                    nextCell.Value = "E"
                ElseIf c.Borders(xlEdgeTop).LineStyle = xlNone Then  'Turn left
                    Set nextCell = c.Offset(-1, 0)
                    nextCell.Value = "N"
                Else    'Go back
                    Set nextCell = c.Offset(0, -1)
                    nextCell.Value = "W"
                End If
            End If
        End If
        
    End If
    Set cellMove = nextCell
    
End Function

Function hasVisitedNeighbour(c As Range) As Boolean
    hasVisitedNeighbour = False
    If c.Offset(0, 1).Interior.Color = visited Or _
    c.Offset(0, -1).Interior.Color = visited Or _
    c.Offset(-1, 0).Interior.Color = visited Or _
    c.Offset(1, 0).Interior.Color = visited Then
        hasVisitedNeighbour = True
    End If
End Function

Function mapPath(cellPath)
    Sheet1.Cells.ClearContents
    Dim cellArray()
    Dim path
    ReDim cellArray(cellPath.Count - 1)
    Dim index As Integer: index = 0
    For Each path In cellPath
        cellArray(index) = path.Address
        index = index + 1
    Next
    Dim newPath
    newPath = removeUselessPath(cellArray)
    
    Dim i As Integer
        For i = LBound(newPath) To UBound(newPath)
            If i = LBound(newPath) Then
                Range(newPath(i)).Interior.Color = vbRed
                Range(newPath(i)).Value = "Start"
            ElseIf i = UBound(newPath) Then
                Range(newPath(i)).Interior.Color = vbRed
                Range(newPath(i)).Value = "End"
            Else
                Range(newPath(i)).Interior.Color = actualPath
                If Range(newPath(i)).Row = Range(newPath(i - 1)).Row Then
                    If Range(newPath(i)).Column > Range(newPath(i - 1)).Column Then
                        'Arrow right
                        Range(newPath(i)).Value = "E"
                    Else
                        'Arrow left
                        Range(newPath(i)).Value = "W"
                    End If
                Else
                    If Range(newPath(i)).Row > Range(newPath(i - 1)).Row Then
                        'Arrow Down
                        Range(newPath(i)).Value = "S"
                    Else
                        'Arrow Up
                        Range(newPath(i)).Value = "N"
                    End If
                End If
            End If
            
        Next i
End Function
                                                    
                                                    
