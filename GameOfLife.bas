Attribute VB_Name = "GameOfLife"
' Set colorIndex for Dead and Alive cells
'
Const DEAD = xlNone ' empty
Const ALIVE = 1          ' black

' run cleanGrid before changing these to delete the grid border
' grid starts at (1,1)
'
Const rowNr As Integer = 50
Const colNr As Integer = 50

Dim RoundNr As Integer, AliveCells As Integer
Dim cell As Range

' Sub GameOfLife()
'
' Plays one match of Conway's Game Of Life
'
' How to play: fill some cells in the grid, then start the sub.
' Game ends when there is a stalemate or no alive cells.
'
'
Sub GameOfLife()
    Dim area As Range, isUpdated As Boolean, chk As Boolean
    
    Set area = Cells(1, 1).Resize(rowNr, colNr)
    RoundNr = 0
    chk = generateStart(area)
    
    If AliveCells = 0 Then
        MsgBox "Fill some cells with color to start playing!", _
                                                vbInformation, "No cells :("
        Exit Sub
    End If
   
    Do
       Application.Wait Now + #12:00:01 AM#  ' 1s artificial delay
       RoundNr = RoundNr + 1
       isUpdated = False
       
        For Each cell In area
            n = CheckNeighbors(cell)
            If n = 3 Then
                If cell.Interior.ColorIndex = DEAD Then
                    cell.ID = "ALIVE"
                    AliveCells = AliveCells + 1
                    isUpdated = True
                End If
            ElseIf n <> 2 Then
                If cell.Interior.ColorIndex = ALIVE Then
                    cell.ID = "DEAD"
                    AliveCells = AliveCells - 1
                    isUpdated = True
                End If
            End If
        Next
        a = UpdateBoard(area)
        Range("A1").Comment.Text "Round " & CStr(RoundNr)
        
        DoEvents
    Loop While AliveCells > 0 And isUpdated
    
    If AliveCells = 0 Then
        MsgBox "This Game Of Life lasted " & CStr(RoundNr + 1) & " rounds.", _
                                                                vbExclamation, "Game Over!"
    Else
        MsgBox "Stalemate at round " & CStr(RoundNr) & ".", _
                                            vbExclamation, "Stalemate!"
   End If
   
End Sub

' Private Function CheckNeighbors(Range) as Integer
'
' counts the alive neighbors of the received cell
' @ returns the number of neighbors
'
' O(M*N), can be further optimized
'
Private Function CheckNeighbors(c As Range) As Integer
    Dim count As Integer
    count = 0
    For i = -1 To 1
    If c.Row + i < 1 Or c.Row + i > rowNr Then GoTo NextRow
        For j = -1 To 1
            If c.Column + j < 1 Or c.Column + j > colNr Then GoTo NextCol
            If i = 0 And j = 0 Then GoTo NextCol ' neighbors only!
            If c.Offset(i, j).Interior.ColorIndex = ALIVE Then
                count = count + 1
            End If
NextCol:
        Next
NextRow:
    Next
    CheckNeighbors = count
End Function

' Private Function UpdateBoard(Range) As Boolean
'
' graphically updates the cells on the board
'
' O(M*N), can be further optimized
'
Private Function UpdateBoard(a As Range) As Boolean
    For Each cell In a
        If cell.ID = "ALIVE" Then
            cell.Interior.ColorIndex = ALIVE
            cell.ID = "BLANK"
        ElseIf cell.ID = "DEAD" Then
            cell.Interior.ColorIndex = DEAD
            cell.ID = "BLANK"
        End If
    Next
    ' Debug.Print "Round " & RoundNr & " complete"
    'Application.Calculate
    UpdateBoard = True
End Function

' Private Function generateStart(Range) As Boolean
'
' Initialises the game area and counts the filled cells
' none = dead, filled = alive
' they are converted to the set colors for the grid
'
Private Function generateStart(a As Range) As Boolean
    ' set the grid layout
    a.RowHeight = 17
    a.ColumnWidth = 2
    a.BorderAround xlDashDot, xlMedium
    Range("A1").AddComment "Start"

    For Each cell In a
        If cell.Interior.ColorIndex = xlNone Then
            cell.Interior.ColorIndex = DEAD
        Else
            cell.Interior.ColorIndex = ALIVE
            AliveCells = AliveCells + 1
        End If
        cell.ID = "BLANK"
    Next
    generateStart = True
End Function

'Sub cleanGrid()
'
' resets the game grid
'
Sub cleanGrid()
    Range("A1").Comment.Delete
    
    For Each cell In Cells(1, 1).Resize(rowNr, colNr)
        cell.Interior.ColorIndex = xlNone
        cell.Borders.LineStyle = xlNone
        cell.ID = "BLANK"
    Next
End Sub
 
