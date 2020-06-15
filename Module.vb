Module Module1

    'grid array is created for 9x9 grid.
    Dim grid(8, 8) As Cell

    'Structure used to hold both the actual value in a particular cell as well as the potential candidates.
    Structure Cell
        Dim value As Integer
        Dim candidates As List(Of Integer)
    End Structure

    'Subroutine used to drawgrid each time in a consistent format.
    Sub drawgrid()
        For y As Integer = 0 To 8
            'Goes through column ensuring spaces between subgrids
            If y = 3 Or y = 6 Then
                Console.WriteLine("_ _ _ _ _ _ _ _ _ _ _ _ _ ") 'Added to help displayed text look more like a Sudoku grid
                Console.WriteLine()
            End If
            'Goes through row ensuring spaces between subgrids
            For x As Integer = 0 To 8
                If x = 3 Or x = 6 Then
                    Console.Write(" | ") 'Added to help displayed text look more like a Sudoku grid
                End If

                'Cells displayed one by one
                Console.Write(" " & grid(x, y).value)
            Next
            Console.WriteLine()
        Next
    End Sub

    'Subroutine that clears all cell values to original blank grid.
    Sub cleargrid()
        'Loops through all values within array going down columns first
        For y As Integer = 0 To 8
            For x As Integer = 0 To 8
                grid(x, y).value = 0
            Next
        Next
    End Sub

    Sub Main()
        'Holds answers to questions asked on screen.
        Dim menuchoice As Integer
        Dim solvechoice As Integer
        Dim gchoice As Integer
        'Holds the desired number for a specific cell selected.
        Dim number As Integer
        'Represents coordinate of cell chosen to modify.
        Dim xchange As Integer
        Dim ychange As Integer
        'Returns boolean value indicating whether puzzle solvable or not.
        Dim possible As Boolean = True
        'Incremeneted for the number of times the solver reaches the last cell (anything over 2 is disregarded).
        Dim possibilities As Integer = 0
        'Used to hold the point at which the first empty cell is located.
        Dim startpoint As Integer = 0
        'Indicates whether the solver backtracked and didn't find another possibility.
        Dim rebacktrack As Boolean = False
        'Holds the x and y values of the cells starting with the lowest candidates.
        Dim OrderedCandidatesX As New List(Of Integer)
        Dim OrderedCandidatesY As New List(Of Integer)
        'Counter used during Heuristic Backtrack to go through the grid in order of the list.
        Dim counter As Integer = 0
        'Creates a stopwatch used to time solvers.
        Dim stopwatch As New Stopwatch

        'Default values assigned to grid and then it is displayed.
        cleargrid()
        drawgrid()

        'Loop used to repeatedly go through the menu options.
        Do
            Console.WriteLine("Enter a value manually [1] Generate Sudoku [2] Solve current puzzle [3] Clear [4] Exit [5]")
            'Exception handling used to prevent invalid inputs.
            Try
                menuchoice = Console.ReadLine
            Catch ex As Exception
                Console.WriteLine()
                Console.ForegroundColor = ConsoleColor.Red
                Console.WriteLine("Error: Invalid input")
                Console.ForegroundColor = ConsoleColor.White
            End Try

            If menuchoice = 1 Then
                Do
                    Do
                        Console.WriteLine("Type the column number (x)")
                        Try
                            xchange = Console.ReadLine
                        Catch ex As Exception
                            Console.ForegroundColor = ConsoleColor.Red
                            Console.WriteLine("Error: Value not valid")
                            Console.ForegroundColor = ConsoleColor.White
                        End Try
                        'Loops until value is within bound of grid.
                    Loop Until xchange <= 9 And xchange >= 1
                    Do
                        Console.WriteLine("Type the column number (y)")
                        Try
                            ychange = Console.ReadLine
                        Catch ex As Exception
                            Console.ForegroundColor = ConsoleColor.Red
                            Console.WriteLine("Error: Value not valid")
                            Console.ForegroundColor = ConsoleColor.White
                        End Try
                        'Loops until value is within bound of grid.
                    Loop Until ychange <= 9 And ychange >= 1


                    Do
                        Try
                            Console.WriteLine("Type a number (1-9) for this cell")
                            number = Console.ReadLine
                        Catch ex As Exception
                            Console.ForegroundColor = ConsoleColor.Red
                            Console.WriteLine("Error: Value not valid")
                            Console.ForegroundColor = ConsoleColor.White
                        End Try
                        'Loops until value is within bound.
                    Loop Until number <= 9 And number >= 0
                    'Checks whether inputted value would cause collisions provided it isn't 0. 
                    'If 0 is entered it essentially clears the cell.
                    If IsValid(xchange - 1, ychange - 1, number) = False And number <> 0 Then
                        Console.ForegroundColor = ConsoleColor.Red
                        Console.WriteLine("Error: Input would cause a collision")
                        Console.ForegroundColor = ConsoleColor.White
                        Console.WriteLine()
                    End If
                Loop Until IsValid(xchange - 1, ychange - 1, number) = True Or number = 0
                'Adds value entered into cell. Subtract 1 seen as the grid is drawn from 0 to 8.
                grid(xchange - 1, ychange - 1).value = number

                drawgrid()
            End If

            'Generation Menu 
            If menuchoice = 2 Then
                Do
                    Try
                        Console.WriteLine("Choose a puzzle difficulty:")
                        Console.WriteLine("Easy [1] Medium [2] Hard [3]")
                        gchoice = Console.ReadLine
                    Catch ex As Exception
                        Console.ForegroundColor = ConsoleColor.Red
                        Console.WriteLine("Error: Value not valid")
                        Console.ForegroundColor = ConsoleColor.White
                    End Try
                Loop Until gchoice = 1 Or gchoice = 2 Or gchoice = 3
                stopwatch.Restart()
                Generation(gchoice)
                stopwatch.Stop()
                Console.WriteLine()
                Console.WriteLine("It took :" & stopwatch.ElapsedMilliseconds & " milliseconds")
                Console.WriteLine(stopwatch.ElapsedTicks & " timer ticks")
                Console.WriteLine()
                drawgrid()
            End If









            'Solving Menu.
            If menuchoice = 3 Then
                Do
                    Console.WriteLine()
                    Console.WriteLine("Choose a method for solving:")
                    Console.WriteLine("Backtrack Greedy Heuristic [1] Rule Based Logic [2]")

                    Try
                        solvechoice = Console.ReadLine
                    Catch ex As Exception
                        Console.WriteLine()
                        Console.ForegroundColor = ConsoleColor.Red
                        Console.WriteLine("Error: Invalid input")
                        Console.ForegroundColor = ConsoleColor.White
                        Console.WriteLine()
                    End Try
                Loop Until solvechoice = 1 Or solvechoice = 2

                If solvechoice = 1 Then
                    stopwatch.Restart()
                    'Before function can run the candidates available for each cell need to be calculated and stored.
                    CellCandidates()
                    'Once candidates calculated, they need to be ordered from least to most in order for the function to go through them.
                    OrderCandidates(OrderedCandidatesX, OrderedCandidatesY)
                    BacktrackHeuristic(OrderedCandidatesX, OrderedCandidatesY, possible, possibilities, counter, startpoint, rebacktrack)
                    stopwatch.Stop()
                    Console.WriteLine()
                    Console.WriteLine("It took :" & stopwatch.ElapsedMilliseconds & " milliseconds")
                    Console.WriteLine(stopwatch.ElapsedTicks & " timer ticks")
                    Console.WriteLine()

                    If possible = True Then
                        Console.WriteLine("Solution found!")
                        'If the program was required to rebacktrack it must have meant no second solution was found.
                        If rebacktrack = True Then
                            Console.WriteLine("The puzzle has a unique solution :)")
                            BacktrackHeuristic(OrderedCandidatesX, OrderedCandidatesY, possible, possibilities, counter, startpoint, rebacktrack)
                        Else
                            Console.WriteLine("The puzzle has more than one solution :/")
                        End If
                        Console.WriteLine(possibilities)
                        Console.WriteLine()
                        drawgrid()
                    Else
                        Console.WriteLine("No solutions found")
                    End If

                    'Variables set to default in case algorithm reused for another task.
                    possible = True
                    possibilities = 0
                    counter = 0
                    startpoint = 0
                    rebacktrack = False
                End If
                If solvechoice = 2 Then
                    stopwatch.Restart()
                    CellCandidates()
                    RuleBased()
                    Console.ForegroundColor = ConsoleColor.Cyan
                    Console.WriteLine("Logic Applied")
                    Console.ForegroundColor = ConsoleColor.White
                    drawgrid()
                    CellCandidates()
                    OrderCandidates(OrderedCandidatesX, OrderedCandidatesY)
                    BacktrackHeuristic(OrderedCandidatesX, OrderedCandidatesY, possible, possibilities, counter, startpoint, rebacktrack)
                    stopwatch.Stop()
                    Console.WriteLine()
                    Console.WriteLine("It took :" & stopwatch.ElapsedMilliseconds & " milliseconds")
                    Console.WriteLine(stopwatch.ElapsedTicks & " timer ticks")
                    Console.WriteLine()

                    If possible = True Then
                        Console.WriteLine("Solution found!")
                        'If the program was required to rebacktrack it must have meant no second solution was found.
                        If rebacktrack = True Then
                            Console.WriteLine("The puzzle has a unique solution :)")
                            BacktrackHeuristic(OrderedCandidatesX, OrderedCandidatesY, possible, possibilities, counter, startpoint, rebacktrack)
                        Else
                            Console.WriteLine("The puzzle has more than one solution :/")
                        End If
                        Console.WriteLine(possibilities)
                        Console.WriteLine()
                        drawgrid()
                    Else
                        Console.WriteLine("No solutions found")
                    End If

                    'Variables set to default in case algorithm reused for another task.
                    possible = True
                    possibilities = 0
                    counter = 0
                    startpoint = 0
                    rebacktrack = False
                End If
            End If

            If menuchoice = 4 Then
                cleargrid()
                Console.ForegroundColor = ConsoleColor.Green
                Console.WriteLine()
                Console.WriteLine("Grid has been cleared!")
                Console.ForegroundColor = ConsoleColor.White
                drawgrid()
            End If

        Loop Until menuchoice = 5
        End
    End Sub

    'Checks for collisions.
    Function IsValid(ByVal X1, ByVal Y1, ByVal i) As Boolean

        Dim nocollisions As Boolean = True
        Dim xpointer As Integer
        Dim ypointer As Integer

        'Checks row and column for collisions
        For a = 0 To 8
            If grid(X1, a).value = i Or grid(a, Y1).value = i Then
                nocollisions = False
            End If
        Next

        'Checks sub-grid for collisions
        If X1 < 3 Then 'Checks to see which box the coordinate lies in based on its x and y values
            xpointer = 0 'Pointer is adjusted in accordance to where on the grid the box starts
        End If
        If X1 > 2 And X1 < 6 Then
            xpointer = 3
        End If
        If X1 > 5 And X1 < 9 Then
            xpointer = 6
        End If

        If Y1 < 3 Then
            ypointer = 0
        End If
        If Y1 > 2 And Y1 < 6 Then
            ypointer = 3
        End If
        If Y1 > 5 And Y1 < 9 Then
            ypointer = 6
        End If

        For y As Integer = ypointer To (ypointer + 2)
            For x As Integer = xpointer To (xpointer + 2)
                If grid(x, y).value = i Then
                    nocollisions = False
                End If
            Next
        Next

        Return nocollisions 'Whether or not the value for the particular cell is valid is returned
    End Function

    'Populates the list of candidates for each cell.
    Sub CellCandidates()
        For x As Integer = 0 To 8
            For y As Integer = 0 To 8
                'A new list is created for each cell.
                grid(x, y).candidates = New List(Of Integer)
                'Cell is only checked for candidates if it is empty.
                If grid(x, y).value = 0 Then
                    For i = 1 To 9
                        If IsValid(x, y, i) = True Then
                            'possible values are checked for cell is a valid input is found it is stored as a possible candidate.
                            grid(x, y).candidates.Add(i)
                        End If
                    Next
                End If
            Next
        Next
    End Sub

    'Orders cells based on candidates.
    Sub OrderCandidates(ByRef OrderedCandidatesX As List(Of Integer), ByRef OrderedCandidatesY As List(Of Integer))
        OrderedCandidatesX.Clear() 'List is cleared to prevent any old data from previous runs causing issues.
        OrderedCandidatesY.Clear()

        For NumberOfCandidates = 0 To 9
            For x = 0 To 8
                For y = 0 To 8
                    'All cell coordinates added to list in order of least number of candidates to most.
                    If grid(x, y).candidates.Count = NumberOfCandidates Then
                        OrderedCandidatesX.Add(x)
                        OrderedCandidatesY.Add(y)
                    End If
                Next
            Next
        Next
    End Sub

    '///Implementation of Heuristic Backtrack.
    Function BacktrackHeuristic(ByVal x As List(Of Integer), ByVal y As List(Of Integer), ByRef possible As Boolean, ByRef possibilities As Integer, ByRef counter As Integer, ByRef startpoint As Integer, ByRef rebacktrack As Boolean) As Boolean

        If grid(x(counter), y(counter)).candidates.Count = 0 Then 'If cell is occupied, it is not modified and the function for the next cell starts
            counter = counter + 1 'Counter increment allows next coordinate in ordered list to be checked.
            Console.WriteLine()
            drawgrid()
            Console.WriteLine()
            'During first run, the startpoint is incremented so that all occupied cells at the beginning are taken into account when setting the startpoint.
            If possibilities = 0 Then
                startpoint = startpoint + 1
            End If




            If counter = 81 Then
                possibilities = possibilities + 1
                Return True 'If backtracking function reaches grid(9,8) solution must have been found.
            End If

            Return BacktrackHeuristic(x, y, possible, possibilities, counter, startpoint, rebacktrack)

        Else
            For i = 0 To grid(x(counter), y(counter)).candidates.Count - 1

                If IsValid(x(counter), y(counter), grid(x(counter), y(counter)).candidates(i)) = True Then 'If value doesn't cause collisions it is tested otherwise it isn't
                    grid(x(counter), y(counter)).value = grid(x(counter), y(counter)).candidates(i) 'value assigned to cell
                    counter = counter + 1 'Next coordinate.

                    If counter = 81 Then 'If backtracking function reaches grid(9,8) solution must have been found.
                        possibilities = possibilities + 1
                        Return True
                    End If

                    'Recursively goes through grid, if value is returned as false then next number is tried.
                    BacktrackHeuristic(x, y, possible, possibilities, counter, startpoint, rebacktrack)

                    If possibilities = 2 Then Return True
                    'If all values have been tried and is still returned as false it retraces its steps
                    counter = counter - 1 'Previous cell in ordered list.
                End If
            Next
            grid(x(counter), y(counter)).value = 0

            'If backtrack back at beginning and a possibility has already been found then there must not be a second solution therefore it must revert back to the first solution.
            If counter - 1 = startpoint And possibilities = 1 Then
                rebacktrack = True
            ElseIf counter - 1 = startpoint And possibilities = 0 Then
                'If backtrack gets back to beginning and still no possibilities there isn't a solution.
                possible = False
            End If
            Return False
        End If
    End Function

    '///Implementation of Rule Based Logic.
    Sub RuleBased()
        Dim repeat As Integer = 0 'Counter stores how many times a value is found in a certain region.
        Dim positionx As Integer 'Holds the coordinates of the place where a value was found.
        Dim positiony As Integer
        Dim xpointer As Integer
        Dim ypointer As Integer


        'Naked Single
        For x = 0 To 8
            For y = 0 To 8
                'Loops through entire grid and checks for any cells with a single candidate which are assigned with the single possibility.
                If grid(x, y).candidates.Count = 1 Then
                    grid(x, y).value = grid(x, y).candidates(0)
                End If
            Next
        Next


        'Hidden Single [Rows]
        For i = 1 To 9
            For y = 0 To 8
                For x = 0 To 8
                    'Loops through rows checking whether numbers are present in possible candidates.
                    If grid(x, y).candidates.Contains(i) Then
                        repeat = repeat + 1
                        positionx = x 'Sets coordiante of value at which valid possibility.
                    End If
                Next
                If repeat = 1 Then 'If there is only one cell within the row that can hold a particular value then it is assigned.
                    grid(positionx, y).value = i
                End If
                'After every check variables need to be cleared to stop them interferring with next loop
                repeat = 0
                positionx = 0
            Next
        Next


        'Hidden Single [Columns]
        For i = 1 To 9 'Works exactly the same as the Hidden Row
            For x = 0 To 8
                For y = 0 To 8
                    If grid(x, y).candidates.Contains(i) Then
                        repeat = repeat + 1
                        positiony = y
                    End If
                Next
                If repeat = 1 Then
                    grid(x, positiony).value = i
                End If
                repeat = 0
                positiony = 0
            Next
        Next


        'Hidden Single [Box]
        For i = 1 To 9
            For box = 0 To 8 'Goes through different subgrids
                For x = 0 To 2
                    For y = 0 To 2
                        xpointer = x + (3 * box)
                        If xpointer > 8 And xpointer < 18 Then 'Ensures loop doesn't exceed array size and correctly goes through boxes
                            xpointer = xpointer - 9
                            ypointer = ypointer + 3
                        End If
                        If xpointer > 17 And xpointer < 28 Then
                            xpointer = xpointer - 18
                            ypointer = xpointer + 6
                        End If
                        If grid(xpointer, xpointer).candidates.Contains(i) Then
                            repeat = repeat + 1
                            'Works exactly the same as Hidden Row/Column except in this case both the x and y positions need to be recorded.
                            positionx = xpointer
                            positiony = xpointer
                        End If
                    Next
                Next
                If repeat = 1 Then
                    'If there is only one cell within a sub grid that can hold a particular value then it is assigned.
                    grid(positionx, positiony).value = i
                End If
                repeat = 0
                positionx = 0
                positiony = 0
            Next
        Next

    End Sub

    '///Generation Algorithm.
    Sub Generation(ByVal gchoice As Integer)
        Randomize()
        CellCandidates()
        Dim OrderedCandidatesX As New List(Of Integer) 'Temporary variables for the Heuristic Backtrack
        Dim OrderedCandidatesY As New List(Of Integer)
        Dim counter As Integer = 0
        Dim possible As Boolean = True
        Dim possibilities As Integer = 0
        Dim rebacktrack As Boolean = False
        Dim startpoint As Integer = 0

        Dim assigned As Boolean = False 'Boolean indicates whether a value has been assigned succesfully.
        Dim x As Integer
        Dim y As Integer
        Dim number As Integer 'Holds randomly generated value for cells.
        Dim blankspaces As Integer = 0 'Holds a value for the number of spaces to be cleared during the generation.

        Dim tempgrid(8, 8) As Integer 'New grid made to hold the original grid whilst it is being modified.


        '20 random cells are populated with random numbers that satisfy the constraint. I chose 20 because it is a large enough number to produce a random puzzle and not too large that it takes too long to populate.
        For i = 1 To 20
            Do
                x = Int(Rnd() * 8) 'Random cell to use generated.
                y = Int(Rnd() * 8)

                number = Int(Rnd() * 8) 'Random number to input into cell generated.
                If grid(x, y).value = 0 Then 'Checks whether input value is valid, if it is then it is assigned. 
                    If IsValid(x, y, number) = True Then
                        grid(x, y).value = number
                        assigned = True
                    End If
                End If
            Loop Until assigned = True 'If it wasn't a valid value then it tries another value in another cell.
            assigned = False
        Next

        'Grid is solved to generate a random solved grid
        CellCandidates()
        OrderCandidates(OrderedCandidatesX, OrderedCandidatesY)
        BacktrackHeuristic(OrderedCandidatesX, OrderedCandidatesY, possible, possibilities, counter, startpoint, rebacktrack)

        For x = 0 To 8 'Tempgrid is populated so that it is identical to the randomly solved grid.
            For y = 0 To 8
                tempgrid(x, y) = grid(x, y).value
            Next
        Next

        If gchoice = 1 Then 'Depending on which difficulty is selected, different numbers of spaces are cleared leabing different number of givens.
            blankspaces = 30
        ElseIf gchoice = 2 Then
            blankspaces = 40
        Else
            blankspaces = 50
        End If

        For i = 1 To blankspaces 'Clears randoms cells and checks for a unique solution.
            Do
                Do
                    x = Int(Rnd() * 8)
                    y = Int(Rnd() * 8)

                Loop Until tempgrid(x, y) <> 0 'If the cell has already been cleared it chooses another random cell until an occupied one is chosen.

                grid(x, y).value = 0 'Cell is cleared.

                'Variables set to default settings before backtrack occurs.
                possible = True
                possibilities = 0
                counter = 0
                startpoint = 0
                rebacktrack = False
                'Cell candidates repopulated after every removal and reordered to keep up to date.
                CellCandidates()
                OrderCandidates(OrderedCandidatesX, OrderedCandidatesY)
                'Heuristic backtrack is used to solve the puzzle.
                BacktrackHeuristic(OrderedCandidatesX, OrderedCandidatesY, possible, possibilities, counter, startpoint, rebacktrack)
                If i = 1 Then rebacktrack = True 'Issue with algorithm caused first cleared cell to return with no solution. First cleared cell would in fact always have a solution.
                'The grid array is overwritten with the tempgrid which held the original fully solved puzzle. This avoids any changes in the grid due to the backtrack algorithm.
                For x2 = 0 To 8
                    For y2 = 0 To 8
                        grid(x2, y2).value = tempgrid(x2, y2)
                    Next
                Next
            Loop Until rebacktrack = True 'If solution isn't unique it tries again.
            grid(x, y).value = 0 'If only one solution then the grid cell is actually cleared.
            tempgrid(x, y) = 0
        Next
    End Sub

End Module
























