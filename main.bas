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
