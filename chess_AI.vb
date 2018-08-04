Public Class Form1

    ' Extended Project Chess v1.9.
    ' Description: This program plays chess against the player, using minimax with alpha-beta pruning to determine the best move for the program to make.
    ' It can look ahead a ply of 4 in no more than a few seconds. Castling and en passant are not implemented. The player can undo moves.
    ' Finished: 16 Februrary 2013.

    ' Pieces: [1, Pawn] [2, Rook] [3, Knight] [4, Bishop] [5, Queen] [6, King] [+, White] [-, Black] [0, Empty] [99, Invalid]
    ' Board: Index of bottom left corner = 21. Index of bottom right corner = 28. Index of top left corner = 91. Index of bottom left corner = 98.

    ' Structures/classes.
    Public Class ChessMove
        Public startOfMove As Integer
        Public endOfMove As Integer
    End Class
    Public Class ChessBoard
        Public indexes(119) As Integer
        Public materialValue As Integer
        Public whiteIndexes As New List(Of Integer)
        Public blackIndexes As New List(Of Integer)
    End Class

    ' Formatting:
    Dim squareSize As Integer = 64
    Dim squareColor1 As Color = Color.Peru
    Dim squareColor2 As Color = Color.NavajoWhite
    Dim squareColor3 As Color = Color.Green
    Dim squareColor4 As Color = Color.Gold

    ' Arrays and lists:
    Dim visualBoard(7, 7) As Windows.Forms.PictureBox
    Dim gameBoard As New ChessBoard

    ' Controls:
    Dim lblInfo As New Label

    ' Other variables:
    Dim ply As Integer = 4
    Dim playersTurnNow As Boolean = True
    Dim recordOfGameBoards As New List(Of ChessBoard)
    Dim recordOfMoves As New List(Of ChessMove)
    ' The highest values that a board can have, both with a king and without a king.
    Dim highestLegalValue, infinity As Integer
    Dim gameState As String = "In progress"
    Dim pieceValue As New Dictionary(Of Integer, Integer)

    ' Variables that are declared here so that they can be used in multiple subroutines:
    Dim gameBoardCopy As New ChessBoard
    Dim moveNumber As Integer = 1
    Dim pieceIsSelectedByPlayer As Boolean = False
    Dim selectedIndexByPlayer As Integer = 0

    ' Functions involving the form and controls.
    Private Sub Form1_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

        ' Determines what to do upon loading the form.

        formatForm()
        formatOtherStuff()
        createBoard()
        resetGame()

    End Sub
    Private Sub Form1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown

        ' Handles all keys pressed.

        Select Case e.KeyCode
            Case Keys.Escape
                Me.Close()
        End Select

    End Sub
    Private Sub btnUndoMove_Click(sender As System.Object, e As System.EventArgs) Handles btnUndoMove.Click

        undoGameMove()
        Me.ActiveControl = Nothing

    End Sub
    Private Sub btnResetGame_Click(sender As System.Object, e As System.EventArgs) Handles btnResetGame.Click

        resetGame()
        Me.ActiveControl = Nothing

    End Sub
    Private Sub formatForm()

        ' Format the form (sizes, text, buttons, etc).

        ' Size the form (16 and 38 account for the extra size of the borders of the form).
        Me.Size = New System.Drawing.Size(8 * squareSize + 16, 8 * squareSize + 38 + 40)
        ' Add the title to the form.
        Me.Text = "Chess v1.8"
        ' Centre the form.
        Me.Left = (Screen.PrimaryScreen.WorkingArea.Width - Me.Width) / 2
        Me.Top = (Screen.PrimaryScreen.WorkingArea.Height - Me.Height) / 2
        ' Make the border look nicer.
        Me.FormBorderStyle = Windows.Forms.FormBorderStyle.FixedSingle
        ' Don't give the user the option to mess up the sizes by maximizing the form.
        Me.MaximizeBox = False
        ' Format lblInfo.
        lblInfo.Size = New Size(300, 14)
        lblInfo.Location = New System.Drawing.Point(10, 8 * squareSize + 13)
        lblInfo.Text = "Your move."
        lblInfo.BringToFront()
        Me.Controls.Add(lblInfo)
        ' Format buttons.
        btnResetGame.Location = New System.Drawing.Point(Me.Width - 90, 8 * squareSize + 8)
        btnUndoMove.Location = New System.Drawing.Point(Me.Width - 170, 8 * squareSize + 8)

    End Sub
    Private Sub formatOtherStuff()

        ' Format and set up everything about from the form.

        Randomize()
        ' Set up the pieceValue dictionary.
        pieceValue(0) = 0
        pieceValue(1) = 100
        pieceValue(2) = 500
        pieceValue(3) = 300
        pieceValue(4) = 325
        pieceValue(5) = 975
        pieceValue(6) = 100000
        For i = 1 To 6
            pieceValue(-i) = -pieceValue(i)
        Next i
        highestLegalValue = 11025
        infinity = 400000

    End Sub

    ' Functions dealing with the flow of the game.
    Private Sub createBoard()

        ' Create the board, both visually and as storage.

        ' Create the board visually.
        For i = 0 To 7
            For j = 0 To 7
                ' Create and format the new square.
                Dim square As New Windows.Forms.PictureBox
                square.Size = New System.Drawing.Point(squareSize, squareSize)
                square.Location = New System.Drawing.Point(i * squareSize, j * squareSize)
                square.BorderStyle = BorderStyle.FixedSingle
                square.SizeMode = PictureBoxSizeMode.CenterImage

                ' Attach the corresponding board index as a tag to the square.
                ' The 7 is there because the squares are placed in reverse order to how they're stored in the board array. In the
                ' board array, 0 is at the bottom left, as they are placed, 0 is at the top left.
                square.Tag = translateTo1D(i, j)

                ' Add the event handler for clicking.
                AddHandler square.MouseDown, AddressOf squareClicked

                ' Color the square (adds the two digits of the index and colors depending on whether it is odd or even, this saves code).
                If (i + j) Mod 2 = 1 Then
                    square.BackColor = squareColor1
                Else
                    square.BackColor = squareColor2
                End If

                ' Add square to visualBoard array.
                visualBoard(i, j) = square

            Next j
        Next i

        ' Display the board.
        For Each square In visualBoard
            Me.Controls.Add(square)
        Next square

        ' Create the board for storage (the 1 dimensional array). 0 for an empty square, 99 for an invalid square.
        For i = 0 To 119
            If (i <= 20) Or (i >= 99) Or (i Mod 10 = 0) Or (i Mod 10 = 9) Then
                gameBoard.indexes(i) = 99
            Else
                gameBoard.indexes(i) = 0
            End If
        Next i

    End Sub
    Private Sub resetGame()

        ' Sets the game to the initial position.

        setDefaultBoard()
        recordOfGameBoards.Clear()
        recordOfGameBoards.Add(copyOfChessBoard(gameBoard))
        recordOfMoves.Clear()
        moveNumber = 1
        playersTurnNow = True
        pieceIsSelectedByPlayer = False
        gameState = "In progress"
        updateGame()

    End Sub
    Private Sub changeTurn()

        ' Changes the current turn.
        playersTurnNow = Not playersTurnNow
        pieceIsSelectedByPlayer = False
        selectedIndexByPlayer = 0
        If hasNoMoves(gameBoard, playersTurnNow) Then
            If isInCheck(gameBoard, playersTurnNow) Then
                If playersTurnNow Then
                    gameState = "Black win"
                Else
                    gameState = "White win"
                End If
            Else
                gameState = "Stalemate"
            End If
        Else
            If Not playersTurnNow Then
                updateGame()
                calculateComputerMove()
            End If
        End If

    End Sub
    Private Sub updateGame()

        ' Update the visualBoard, update lblInfo and deal with the game ending.

        Application.DoEvents()

        Dim pieceImage As Image
        Dim currentPieceNumber As Integer
        Dim moveList As List(Of ChessMove) = generateMoves(gameBoard, selectedIndexByPlayer, True, True)

        ' Set the images on the squares.
        For i = 0 To 7
            For j = 0 To 7
                currentPieceNumber = gameBoard.indexes(translateTo1D(i, j))
                Select Case currentPieceNumber
                    ' Select the correct image for the piece number.
                    Case 1 : pieceImage = My.Resources.imgWhitePawn
                    Case 2 : pieceImage = My.Resources.imgWhiteRook
                    Case 3 : pieceImage = My.Resources.imgWhiteKnight
                    Case 4 : pieceImage = My.Resources.imgWhiteBishop
                    Case 5 : pieceImage = My.Resources.imgWhiteQueen
                    Case 6 : pieceImage = My.Resources.imgWhiteKing
                    Case -1 : pieceImage = My.Resources.imgBlackPawn
                    Case -2 : pieceImage = My.Resources.imgBlackRook
                    Case -3 : pieceImage = My.Resources.imgBlackKnight
                    Case -4 : pieceImage = My.Resources.imgBlackBishop
                    Case -5 : pieceImage = My.Resources.imgBlackQueen
                    Case -6 : pieceImage = My.Resources.imgBlackKing
                    Case Else : pieceImage = Nothing
                End Select
                ' Physically change the image.
                visualBoard(i, j).Image = pieceImage
                ' Set the color of the square, assuming it won't be a possible move or a previous move.
                If (i + j) Mod 2 = 1 Then
                    visualBoard(i, j).BackColor = squareColor1
                Else
                    visualBoard(i, j).BackColor = squareColor2
                End If
                ' Set the color of the square if it is part of the computer's last move.
                If moveNumber > 2 Then
                    If playersTurnNow Then
                        If (recordOfMoves.Last.startOfMove = translateTo1D(i, j) Or recordOfMoves.Last.endOfMove = translateTo1D(i, j)) Then
                            visualBoard(i, j).BackColor = squareColor4
                        End If
                    End If
                End If
                ' Set the color of the square if it is a move to be displayed as possible to the player.
                If pieceIsSelectedByPlayer Then
                    For Each possibleMove In moveList
                        If possibleMove.endOfMove = translateTo1D(i, j) Then
                            visualBoard(i, j).BackColor = squareColor3
                            Exit For
                        End If
                    Next possibleMove
                End If
            Next j
        Next i

        ' Update lblInfo.
        If playersTurnNow Then
            If isInCheck(gameBoard, True) Then
                lblInfo.Text = "Your turn. You are in check."
            Else
                lblInfo.Text = "Your turn."
            End If
        Else
            If isInCheck(gameBoard, False) Then
                lblInfo.Text = "AI's turn. AI is in check."
            Else
                lblInfo.Text = "AI's turn."
            End If
        End If

        ' Check if the game is over.
        If gameState <> "In progress" Then
            If gameState = "White win" Then
                lblInfo.Text = "Checkmate. You win!"
            ElseIf gameState = "Black win" Then
                lblInfo.Text = "Checkmate. AI wins."
            ElseIf gameState = "Stalemate" Then
                lblInfo.Text = "Stalemate."
            End If
            Select Case MsgBox(lblInfo.Text & " Would you like to play again?", MsgBoxStyle.YesNo, "Game Over")
                Case MsgBoxResult.Yes
                    resetGame()
                Case MsgBoxResult.No
                    Me.Close()
            End Select
        End If

        Application.DoEvents()

    End Sub
    Private Sub squareClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs)

        ' Determines what is to be done when a board square has been clicked.

        If playersTurnNow Then
            Dim clickedIndexByPlayer As Integer = sender.Tag
            ' Determine if a piece is already selected by the player.
            If pieceIsSelectedByPlayer Then
                ' Determine if the clicked piece is the previously selected one.
                If clickedIndexByPlayer = selectedIndexByPlayer Then
                    ' Deselect the piece.
                    pieceIsSelectedByPlayer = False
                Else
                    ' Determine if the clicked square is an available move.
                    Dim availableMove As Boolean = False
                    For Each possibleMove In generateMoves(gameBoard, selectedIndexByPlayer, True, True)
                        If possibleMove.endOfMove = clickedIndexByPlayer Then
                            ' Make the move.
                            availableMove = True
                            Dim finalMove As New ChessMove
                            finalMove.startOfMove = selectedIndexByPlayer
                            finalMove.endOfMove = clickedIndexByPlayer
                            makeFinalMove(finalMove)
                            Exit For
                        End If
                    Next possibleMove
                    ' If the clicked square was not an available move, select the clicked piece if possible.
                    If Not availableMove Then
                        If gameBoard.indexes(clickedIndexByPlayer) > 0 Then
                            pieceIsSelectedByPlayer = True
                            selectedIndexByPlayer = clickedIndexByPlayer
                        End If
                    End If
                End If
            Else
                ' Select the clicked piece if possible.
                If gameBoard.indexes(clickedIndexByPlayer) > 0 Then
                    pieceIsSelectedByPlayer = True
                    selectedIndexByPlayer = clickedIndexByPlayer
                End If
            End If
        End If

        updateGame()

    End Sub

    ' Functions dealing with the board and moves.
    Private Sub setDefaultBoard()

        ' Set up the default chess board.

        ' Remove all pieces and indexes.
        removeAllPieces(gameBoard)
        gameBoard.whiteIndexes.Clear()
        gameBoard.blackIndexes.Clear()
        ' Add white pawns.
        For i = 31 To 38
            addPiece(gameBoard, 1, i)
            gameBoard.whiteIndexes.Add(i)
        Next i
        ' Add the other white pieces.
        addPiece(gameBoard, 2, 21)
        addPiece(gameBoard, 2, 28)
        addPiece(gameBoard, 3, 22)
        addPiece(gameBoard, 3, 27)
        addPiece(gameBoard, 4, 23)
        addPiece(gameBoard, 4, 26)
        addPiece(gameBoard, 5, 24)
        addPiece(gameBoard, 6, 25)
        For i = 21 To 28
            gameBoard.whiteIndexes.Add(i)
        Next i
        ' Add black pawns.
        For i = 81 To 88
            addPiece(gameBoard, -1, i)
            gameBoard.blackIndexes.Add(i)
        Next i
        ' Add the other black pieces.
        addPiece(gameBoard, -2, 91)
        addPiece(gameBoard, -2, 98)
        addPiece(gameBoard, -3, 92)
        addPiece(gameBoard, -3, 97)
        addPiece(gameBoard, -4, 93)
        addPiece(gameBoard, -4, 96)
        addPiece(gameBoard, -5, 94)
        addPiece(gameBoard, -6, 95)
        For i = 91 To 98
            gameBoard.blackIndexes.Add(i)
        Next i
        ' Set the material value.
        gameBoard.materialValue = 0

    End Sub
    Private Sub addPiece(ByRef board As ChessBoard, pieceNumber As Integer, index As Integer)

        ' Inserts the given piece at the given index on the given board.

        board.indexes(index) = pieceNumber

    End Sub
    Private Sub removePiece(ByRef board As ChessBoard, index As Integer)

        ' Replaces the piece at the given index with an empty space on the given board.

        board.indexes(index) = 0

    End Sub
    Private Sub removeAllPieces(ByRef board As ChessBoard)

        ' Remove all pieces and indexes from the given board.
        For x = 0 To 7
            For y = 0 To 7
                removePiece(board, translateTo1D(x, y))
            Next y
        Next x
        board.whiteIndexes.Clear()
        board.blackIndexes.Clear()
        board.materialValue = 0

    End Sub
    Private Sub movePiece(ByRef board As ChessBoard, move As ChessMove)

        ' Moves the given piece on the given board.
        ' This updates the material count and the list of piece indexes for each side, and checks for pawn promotion.

        Dim startOfMove As Integer = move.startOfMove
        Dim endOfMove As Integer = move.endOfMove
        Dim startPiece As Integer = board.indexes(startOfMove)
        Dim endPiece As Integer = board.indexes(endOfMove)

        ' Update the index lists.
        If startPiece > 0 Then
            ' White/human move.
            If endPiece <> 0 Then
                board.materialValue -= pieceValue(endPiece)
                ' Remove the black index if that piece is taken by this move.
                board.blackIndexes.Remove(endOfMove)
            End If
            ' Change the white index to the new index the piece has been moved to.
            board.whiteIndexes(board.whiteIndexes.FindIndex(Function(index) index.Equals(startOfMove))) = endOfMove
        ElseIf startPiece < 0 Then
            ' Black/computer move.
            If endPiece <> 0 Then
                board.materialValue -= pieceValue(endPiece)
                ' Remove the white index if that piece is taken by this move.
                board.whiteIndexes.Remove(endOfMove)
            End If
            ' Change the black index to the new index the piece has been moved to.
            board.blackIndexes(board.blackIndexes.FindIndex(Function(index) index.Equals(startOfMove))) = endOfMove
        End If

        ' Make the move.
        addPiece(board, startPiece, endOfMove)
        removePiece(board, startOfMove)

        ' Check for pawn promotion.
        If startPiece = 1 Then
            If endOfMove >= 91 Then
                addPiece(board, 5, endOfMove)
                board.whiteIndexes.Add(endOfMove)
                board.materialValue += 875
                Exit Sub
            End If
        ElseIf startPiece = -1 Then
            If endOfMove <= 28 Then
                addPiece(board, -5, endOfMove)
                board.blackIndexes.Add(endOfMove)
                board.materialValue -= 875
                Exit Sub
            End If
        End If

    End Sub
    Private Sub undoMove(ByRef board As ChessBoard, move As ChessMove, endPiece As Integer, pawnPromotion As Boolean)

        ' Undoes a chess move. The piece that may have been taken must be known.

        Dim startPiece As Integer = board.indexes(move.endOfMove)
        If pawnPromotion Then
            startPiece = startPiece / Math.Abs(startPiece)
        End If
        Dim startIndex As Integer = move.startOfMove
        Dim endIndex As Integer = move.endOfMove

        ' The moved piece goes back to the square it moved from.
        addPiece(board, startPiece, startIndex)
        ' The moved piece's index changes.
        If startPiece > 0 Then
            board.whiteIndexes(board.whiteIndexes.FindIndex(Function(index) index.Equals(endIndex))) = startIndex
        ElseIf startPiece < 0 Then
            board.blackIndexes(board.blackIndexes.FindIndex(Function(index) index.Equals(endIndex))) = startIndex
        End If

        ' The captured piece goes back to where it was taken from.
        addPiece(board, endPiece, endIndex)
        ' The captured piece's index is added.
        If endPiece > 0 Then
            board.whiteIndexes.Add(move.endOfMove)
        ElseIf endPiece < 0 Then
            board.blackIndexes.Add(move.endOfMove)
        End If

        board.materialValue += pieceValue(endPiece)

    End Sub
    Private Sub placeNewPiece(ByRef board As ChessBoard, pieceNumber As Integer, index As Integer)

        ' Place a new piece at an index, changing the whiteIndex or blackIndex list and material value as appropriate.

        addPiece(board, pieceNumber, index)
        If pieceNumber > 0 Then
            board.whiteIndexes.Add(index)
        Else
            board.blackIndexes.Add(index)
        End If
        board.materialValue += pieceValue(pieceNumber)

    End Sub
    Private Function moveIsPawnPromotion(board As ChessBoard, move As ChessMove)

        ' Determines whether the given mode is a pawn promotion.

        Dim start As Integer = move.startOfMove
        If start >= 81 Then
            If board.indexes(move.startOfMove) = 1 Then
                Return True
            End If
        ElseIf start <= 38 Then
            If board.indexes(move.startOfMove) = -1 Then
                Return True
            End If
        End If

        Return False

    End Function
    Private Sub makeFinalMove(move As ChessMove)

        ' Make the given move on the game board.

        movePiece(gameBoard, move)
        recordOfGameBoards.Add(copyOfChessBoard(gameBoard))
        recordOfMoves.Add(move)
        moveNumber += 1
        changeTurn()

    End Sub
    Private Sub undoGameMove()

        ' Takes the game back a move, until it's the player's move again.
        If playersTurnNow And moveNumber <> 1 Then
            recordOfGameBoards.RemoveRange(recordOfGameBoards.Count - 2, 2)
            gameBoard = copyOfChessBoard(recordOfGameBoards.Last)
            recordOfMoves.RemoveRange(recordOfMoves.Count - 2, 2)
            moveNumber -= 2
        ElseIf (Not playersTurnNow) Then
            recordOfGameBoards.RemoveRange(recordOfGameBoards.Count - 1, 1)
            gameBoard = copyOfChessBoard(recordOfGameBoards.Last)
            recordOfMoves.RemoveRange(recordOfMoves.Count - 1, 1)
            moveNumber -= 1
        End If

        pieceIsSelectedByPlayer = False
        playersTurnNow = True
        updateGame()

    End Sub
    Private Function generateMoves(board As ChessBoard, index As Integer, humanPlayer As Boolean, legalOnly As Boolean)

        ' Generate moves for a given piece on a given board. Returns a list of moves.

        Dim endOfMoveList As New List(Of Integer)
        Dim piece As Integer = board.indexes(index)
        Dim temp As Integer

        Select Case piece
            Case 1
                ' White pawn
                ' Forward one if no piece is there.
                If board.indexes(index + 10) = 0 Then
                    endOfMoveList.Add(index + 10)
                    ' Forward two if the pawn is yet to move.
                    If index >= 31 And index <= 38 Then
                        If board.indexes(index + 20) = 0 Then
                            endOfMoveList.Add(index + 20)
                        End If
                    End If
                End If
                ' Diagonally forward one if any opposite piece is there.
                If (board.indexes(index + 9)) * (piece) < 0 Then
                    endOfMoveList.Add(index + 9)
                End If
                If (board.indexes(index + 11)) * (piece) < 0 Then
                    endOfMoveList.Add(index + 11)
                End If
            Case -1
                ' Black pawn
                ' Backward one if no piece is there.
                If board.indexes(index - 10) = 0 Then
                    endOfMoveList.Add(index - 10)
                    ' Forward two if the pawn is yet to move.
                    If index >= 81 And index <= 88 Then
                        If board.indexes(index - 20) = 0 Then
                            endOfMoveList.Add(index - 20)
                        End If
                    End If
                End If
                ' Diagonally backward one if any opposite piece is there.
                temp = board.indexes(index - 9)
                If (temp * (piece) < 0) And (Not temp = 99) Then
                    endOfMoveList.Add(index - 9)
                End If
                temp = board.indexes(index - 11)
                If (temp * (piece) < 0) And (Not temp = 99) Then
                    endOfMoveList.Add(index - 11)
                End If
            Case 3, -3
                ' Knight
                temp = board.indexes(index - 21)
                If Not (temp = 99) Then
                    If temp * piece <= 0 Then endOfMoveList.Add(index - 21)
                End If
                temp = board.indexes(index - 19)
                If Not (temp = 99) Then
                    If temp * piece <= 0 Then endOfMoveList.Add(index - 19)
                End If
                temp = board.indexes(index - 12)
                If Not (temp = 99) Then
                    If temp * piece <= 0 Then endOfMoveList.Add(index - 12)
                End If
                temp = board.indexes(index - 8)
                If Not (temp = 99) Then
                    If temp * piece <= 0 Then endOfMoveList.Add(index - 8)
                End If
                temp = board.indexes(index + 8)
                If Not (temp = 99) Then
                    If temp * piece <= 0 Then endOfMoveList.Add(index + 8)
                End If
                temp = board.indexes(index + 12)
                If Not (temp = 99) Then
                    If temp * piece <= 0 Then endOfMoveList.Add(index + 12)
                End If
                temp = board.indexes(index + 19)
                If Not (temp = 99) Then
                    If temp * piece <= 0 Then endOfMoveList.Add(index + 19)
                End If
                temp = board.indexes(index + 21)
                If Not (temp = 99) Then
                    If temp * piece <= 0 Then endOfMoveList.Add(index + 21)
                End If
            Case 6, -6
                ' King
                temp = board.indexes(index - 11)
                If Not (temp = 99) Then
                    If temp * piece <= 0 Then endOfMoveList.Add(index - 11)
                End If
                temp = board.indexes(index - 10)
                If Not (temp = 99) Then
                    If temp * piece <= 0 Then endOfMoveList.Add(index - 10)
                End If
                temp = board.indexes(index - 9)
                If Not (temp = 99) Then
                    If temp * piece <= 0 Then endOfMoveList.Add(index - 9)
                End If
                temp = board.indexes(index - 1)
                If Not (temp = 99) Then
                    If temp * piece <= 0 Then endOfMoveList.Add(index - 1)
                End If
                temp = board.indexes(index + 1)
                If Not (temp = 99) Then
                    If temp * piece <= 0 Then endOfMoveList.Add(index + 1)
                End If
                temp = board.indexes(index + 9)
                If Not (temp = 99) Then
                    If temp * piece <= 0 Then endOfMoveList.Add(index + 9)
                End If
                temp = board.indexes(index + 10)
                If Not (temp = 99) Then
                    If temp * piece <= 0 Then endOfMoveList.Add(index + 10)
                End If
                temp = board.indexes(index + 11)
                If Not (temp = 99) Then
                    If temp * piece <= 0 Then endOfMoveList.Add(index + 11)
                End If
            Case 2, -2
                'Rook
                ' Check to the left.
                For i = 1 To 7
                    temp = board.indexes(index - i)
                    If (temp = 99) Or (temp * piece > 0) Then
                        Exit For
                    ElseIf (temp * piece < 0) Then
                        endOfMoveList.Add(index - i)
                        Exit For
                    Else
                        endOfMoveList.Add(index - i)
                    End If
                Next i
                ' Check to the right.
                For i = 1 To 7
                    temp = board.indexes(index + i)
                    If (temp = 99) Or (temp * piece > 0) Then
                        Exit For
                    ElseIf (temp * piece < 0) Then
                        endOfMoveList.Add(index + i)
                        Exit For
                    Else
                        endOfMoveList.Add(index + i)
                    End If
                Next i
                ' Check down.
                For i = 1 To 7
                    temp = board.indexes(index - (10 * i))
                    If (temp = 99) Or (temp * piece > 0) Then
                        Exit For
                    ElseIf (temp * piece < 0) Then
                        endOfMoveList.Add(index - (10 * i))
                        Exit For
                    Else
                        endOfMoveList.Add(index - (10 * i))
                    End If
                Next i
                ' Check up.
                For i = 1 To 7
                    temp = board.indexes(index + (10 * i))
                    If (temp = 99) Or (temp * piece > 0) Then
                        Exit For
                    ElseIf (temp * piece < 0) Then
                        endOfMoveList.Add(index + (10 * i))
                        Exit For
                    Else
                        endOfMoveList.Add(index + (10 * i))
                    End If
                Next i
            Case 4, -4
                ' Bishop.
                ' Check down/left.
                For i = 1 To 7
                    temp = board.indexes(index - (11 * i))
                    If (temp = 99) Or (temp * piece > 0) Then
                        Exit For
                    ElseIf (temp * piece < 0) Then
                        endOfMoveList.Add(index - (11 * i))
                        Exit For
                    Else
                        endOfMoveList.Add(index - (11 * i))
                    End If
                Next i
                ' Check down/right
                For i = 1 To 7
                    temp = board.indexes(index - (9 * i))
                    If (temp = 99) Or (temp * piece > 0) Then
                        Exit For
                    ElseIf (temp * piece < 0) Then
                        endOfMoveList.Add(index - (9 * i))
                        Exit For
                    Else
                        endOfMoveList.Add(index - (9 * i))
                    End If
                Next i
                ' Check up/left.
                For i = 1 To 7
                    temp = board.indexes(index + (9 * i))
                    If (temp = 99) Or (temp * piece > 0) Then
                        Exit For
                    ElseIf (temp * piece < 0) Then
                        endOfMoveList.Add(index + (9 * i))
                        Exit For
                    Else
                        endOfMoveList.Add(index + (9 * i))
                    End If
                Next i
                ' Check up/right.
                For i = 1 To 7
                    temp = board.indexes(index + (11 * i))
                    If (temp = 99) Or (temp * piece > 0) Then
                        Exit For
                    ElseIf (temp * piece < 0) Then
                        endOfMoveList.Add(index + (11 * i))
                        Exit For
                    Else
                        endOfMoveList.Add(index + (11 * i))
                    End If
                Next i
            Case 5, -5
                ' Queen (the code for the rook and bishop is simply copied here).

                ' Check to the left.
                For i = 1 To 7
                    temp = board.indexes(index - i)
                    If (temp = 99) Or (temp * piece > 0) Then
                        Exit For
                    ElseIf (temp * piece < 0) Then
                        endOfMoveList.Add(index - i)
                        Exit For
                    Else
                        endOfMoveList.Add(index - i)
                    End If
                Next i
                ' Check to the right.
                For i = 1 To 7
                    temp = board.indexes(index + i)
                    If (temp = 99) Or (temp * piece > 0) Then
                        Exit For
                    ElseIf (temp * piece < 0) Then
                        endOfMoveList.Add(index + i)
                        Exit For
                    Else
                        endOfMoveList.Add(index + i)
                    End If
                Next i
                ' Check down.
                For i = 1 To 7
                    temp = board.indexes(index - (10 * i))
                    If (temp = 99) Or (temp * piece > 0) Then
                        Exit For
                    ElseIf (temp * piece < 0) Then
                        endOfMoveList.Add(index - (10 * i))
                        Exit For
                    Else
                        endOfMoveList.Add(index - (10 * i))
                    End If
                Next i
                ' Check up.
                For i = 1 To 7
                    temp = board.indexes(index + (10 * i))
                    If (temp = 99) Or (temp * piece > 0) Then
                        Exit For
                    ElseIf (temp * piece < 0) Then
                        endOfMoveList.Add(index + (10 * i))
                        Exit For
                    Else
                        endOfMoveList.Add(index + (10 * i))
                    End If
                Next i
                ' Check down/left.
                For i = 1 To 7
                    temp = board.indexes(index - (11 * i))
                    If (temp = 99) Or (temp * piece > 0) Then
                        Exit For
                    ElseIf (temp * piece < 0) Then
                        endOfMoveList.Add(index - (11 * i))
                        Exit For
                    Else
                        endOfMoveList.Add(index - (11 * i))
                    End If
                Next i
                ' Check down/right
                For i = 1 To 7
                    temp = board.indexes(index - (9 * i))
                    If (temp = 99) Or (temp * piece > 0) Then
                        Exit For
                    ElseIf (temp * piece < 0) Then
                        endOfMoveList.Add(index - (9 * i))
                        Exit For
                    Else
                        endOfMoveList.Add(index - (9 * i))
                    End If
                Next i
                ' Check up/left.
                For i = 1 To 7
                    temp = board.indexes(index + (9 * i))
                    If (temp = 99) Or (temp * piece > 0) Then
                        Exit For
                    ElseIf (temp * piece < 0) Then
                        endOfMoveList.Add(index + (9 * i))
                        Exit For
                    Else
                        endOfMoveList.Add(index + (9 * i))
                    End If
                Next i
                ' Check up/right.
                For i = 1 To 7
                    temp = board.indexes(index + (11 * i))
                    If (temp = 99) Or (temp * piece > 0) Then
                        Exit For
                    ElseIf (temp * piece < 0) Then
                        endOfMoveList.Add(index + (11 * i))
                        Exit For
                    Else
                        endOfMoveList.Add(index + (11 * i))
                    End If
                Next i
        End Select

        ' Convert to endOfMove's in to full moves to return.
        Dim moveList As New List(Of ChessMove)
        For Each endOfMove In endOfMoveList
            Dim move As New ChessMove
            move.startOfMove = index
            move.endOfMove = endOfMove
            moveList.Add(move)
        Next endOfMove

        ' Remove any moves that will cause check if required.
        If legalOnly Then
            moveList.RemoveAll(Function(x) willCauseCheck(board, x, humanPlayer) = True)
        End If

        Return moveList

    End Function
    Private Function generateAllMoves(board As ChessBoard, humanPlayer As Boolean, legalOnly As Boolean)

        ' Returns a list of all possible legal moves for a certain player.
        Dim moveList As New List(Of ChessMove)

        If humanPlayer Then
            For Each index In board.whiteIndexes
                For Each possibleMove In generateMoves(board, index, True, legalOnly)
                    moveList.Add(possibleMove)
                Next possibleMove
            Next index
        ElseIf Not humanPlayer Then
            For Each index In board.blackIndexes
                For Each possibleMove In generateMoves(board, index, False, legalOnly)
                    moveList.Add(possibleMove)
                Next possibleMove
            Next index
        End If

        Return moveList

    End Function
    Private Function willCauseCheck(board As ChessBoard, move As ChessMove, humanPlayer As Boolean)

        ' Determines whether a given move will cause the given player to become in check.

        Dim copyBoard As ChessBoard = copyOfChessBoard(board)
        movePiece(copyBoard, move)
        If isInCheck(copyBoard, humanPlayer) Then
            Return True
        Else
            Return False
        End If

    End Function
    Private Function isInCheck(board As ChessBoard, humanPlayer As Boolean)

        ' Checks if a given player is in check on a given board. This is done by tracing outward from the king.

        Dim temp, sign, kingIndex As Integer

        ' Set the appropriate sign depending on the player.
        If humanPlayer Then
            sign = 1
        Else
            sign = -1
        End If
        ' Find the appropriates player's king on the board.
        If humanPlayer Then
            For Each index In board.whiteIndexes
                If board.indexes(index) = 6 Then
                    kingIndex = index
                    Exit For
                End If
            Next
        ElseIf Not humanPlayer Then
            For Each index In board.blackIndexes
                If board.indexes(index) = -6 Then
                    kingIndex = index
                    Exit For
                End If
            Next
        End If

        ' Check horizontally and vertically for rooks and queens.
        ' Check down.
        For i = 1 To 7
            temp = board.indexes(kingIndex - (10 * i))
            If (temp = 99) Or (temp * sign > 0) Then
                Exit For
            ElseIf (temp * sign < 0) Then
                If (temp) * (-sign) = 2 Or (temp) * (-sign) = 5 Then
                    Return True
                Else
                    Exit For
                End If
            End If
        Next i
        ' Check up.
        For i = 1 To 7
            temp = board.indexes(kingIndex + (10 * i))
            If (temp = 99) Or (temp * sign > 0) Then
                Exit For
            ElseIf (temp * sign < 0) Then
                If (temp) * (-sign) = 2 Or (temp) * (-sign) = 5 Then
                    Return True
                Else
                    Exit For
                End If
            End If
        Next i
        ' Check left.
        For i = 1 To 7
            temp = board.indexes(kingIndex - (i))
            If (temp = 99) Or (temp * sign > 0) Then
                Exit For
            ElseIf (temp * sign < 0) Then
                If (temp) * (-sign) = 2 Or (temp) * (-sign) = 5 Then
                    Return True
                Else
                    Exit For
                End If
            End If
        Next i
        ' Check right.
        For i = 1 To 7
            temp = board.indexes(kingIndex + (i))
            If (temp = 99) Or (temp * sign > 0) Then
                Exit For
            ElseIf (temp * sign < 0) Then
                If (temp) * (-sign) = 2 Or (temp) * (-sign) = 5 Then
                    Return True
                Else
                    Exit For
                End If
            End If
        Next i

        ' Check diagonally for bishops and queens.
        ' Check down/left.
        For i = 1 To 7
            temp = board.indexes(kingIndex - (11 * i))
            If (temp = 99) Or (temp * sign > 0) Then
                Exit For
            ElseIf (temp * sign < 0) Then
                If (temp) * (-sign) = 4 Or (temp) * (-sign) = 5 Then
                    Return True
                Else
                    Exit For
                End If
            End If
        Next i
        ' Check down/right
        For i = 1 To 7
            temp = board.indexes(kingIndex - (9 * i))
            If (temp = 99) Or (temp * sign > 0) Then
                Exit For
            ElseIf (temp * sign < 0) Then
                If (temp) * (-sign) = 4 Or (temp) * (-sign) = 5 Then
                    Return True
                Else
                    Exit For
                End If
            End If
        Next i
        ' Check up/left.
        For i = 1 To 7
            temp = board.indexes(kingIndex + (9 * i))
            If (temp = 99) Or (temp * sign > 0) Then
                Exit For
            ElseIf (temp * sign < 0) Then
                If (temp) * (-sign) = 4 Or (temp) * (-sign) = 5 Then
                    Return True
                Else
                    Exit For
                End If
            End If
        Next i
        ' Check up/right.
        For i = 1 To 7
            temp = board.indexes(kingIndex + (11 * i))
            If (temp = 99) Or (temp * sign > 0) Then
                Exit For
            ElseIf (temp * sign < 0) Then
                If (temp) * (-sign) = 4 Or (temp) * (-sign) = 5 Then
                    Return True
                Else
                    Exit For
                End If
            End If
        Next i

        ' Check immediately diagonally for pawns.
        temp = board.indexes(kingIndex + (sign * 9))
        If (temp) * (-sign) = 1 Then
            Return True
        End If
        temp = board.indexes(kingIndex + (sign * 11))
        If (temp) * (-sign) = 1 Then
            Return True
        End If

        ' Check adjacently for a king.
        temp = board.indexes(kingIndex - 11)
        If (temp) * (-sign) = 6 Then
            Return True
        End If
        temp = board.indexes(kingIndex - 10)
        If (temp) * (-sign) = 6 Then
            Return True
        End If
        temp = board.indexes(kingIndex - 9)
        If (temp) * (-sign) = 6 Then
            Return True
        End If
        temp = board.indexes(kingIndex - 1)
        If (temp) * (-sign) = 6 Then
            Return True
        End If
        temp = board.indexes(kingIndex + 1)
        If (temp) * (-sign) = 6 Then
            Return True
        End If
        temp = board.indexes(kingIndex + 9)
        If (temp) * (-sign) = 6 Then
            Return True
        End If
        temp = board.indexes(kingIndex + 10)
        If (temp) * (-sign) = 6 Then
            Return True
        End If
        temp = board.indexes(kingIndex + 11)
        If (temp) * (-sign) = 6 Then
            Return True
        End If

        ' Check knightly for knights.
        temp = board.indexes(kingIndex - 21)
        If (temp) * (-sign) = 3 Then
            Return True
        End If
        temp = board.indexes(kingIndex - 19)
        If (temp) * (-sign) = 3 Then
            Return True
        End If
        temp = board.indexes(kingIndex - 12)
        If (temp) * (-sign) = 3 Then
            Return True
        End If
        temp = board.indexes(kingIndex - 8)
        If (temp) * (-sign) = 3 Then
            Return True
        End If
        temp = board.indexes(kingIndex + 8)
        If (temp) * (-sign) = 3 Then
            Return True
        End If
        temp = board.indexes(kingIndex + 12)
        If (temp) * (-sign) = 3 Then
            Return True
        End If
        temp = board.indexes(kingIndex + 19)
        If (temp) * (-sign) = 3 Then
            Return True
        End If
        temp = board.indexes(kingIndex + 21)
        If (temp) * (-sign) = 3 Then
            Return True
        End If

        ' Return false if a check hasn't already been recognised.
        Return False

    End Function
    Private Function hasNoMoves(board As ChessBoard, humanPlayer As Boolean)

        ' Determines whether a given player has no moves available on a given board.

        Dim moveList As List(Of ChessMove) = generateAllMoves(board, humanPlayer, True)
        If moveList.Count = 0 Then
            Return True
        Else
            Return False
        End If

    End Function

    ' Functions dealing with calculating the computer's move.
    Private Sub calculateComputerMove()

        ' Calculates the move the computer will make.

        gameBoardCopy = copyOfChessBoard(gameBoard)
        Dim bestMoves As New List(Of ChessMove)
        Dim checkmateMoves As New List(Of ChessMove)
        Dim bestValue As Integer = infinity

        ' Go through every possible move to work out which one yields the minimum (best for the computer) value.
        For Each possibleMove In generateAllMoves(gameBoardCopy, False, True)

            ' Test whether the move is a pawn promotion, so that it can be undone if it is.
            Dim pawnPromotion As Boolean = False
            If moveIsPawnPromotion(gameBoardCopy, possibleMove) Then pawnPromotion = True
            Dim tempEndPiece As Integer = gameBoardCopy.indexes(possibleMove.endOfMove)
            movePiece(gameBoardCopy, possibleMove)

            ' Do the alpha-beta stuff.
            Dim value As Integer = alphaBeta(1, -infinity, bestValue + 1, True)
            If value < bestValue Then
                bestMoves.Clear()
                bestMoves.Add(possibleMove)
                bestValue = value
            ElseIf value = bestValue Then
                bestMoves.Add(possibleMove)
            End If
            If value < -highestLegalValue Then checkmateMoves.Add(possibleMove)

            ' Undo the move.
            undoMove(gameBoardCopy, possibleMove, tempEndPiece, pawnPromotion)

            Application.DoEvents()
            ' Stop calculating if the move has been undone, and it is now the player's move.
            If playersTurnNow Then Exit For

        Next

        ' Make the move.
        If Not playersTurnNow Then
            If checkmateMoves.Count > 0 Then
                makeFinalMove(bestFinalMove(checkmateMoves))
            Else
                makeFinalMove(bestFinalMove(bestMoves))
            End If
        End If

    End Sub
    Private Function bestFinalMove(moves As List(Of ChessMove))

        ' Determines the best move out of a list of moves by the criteria that a capture is better than just moving forward which is better than not moving forward. This
        ' will give the AI a bit more direction at the start of the game.

        Dim bestMoves As New List(Of ChessMove)

        ' Check if any moves lead to pawn promotion.
        For Each m In moves
            If moveIsPawnPromotion(gameBoard, m) Then
                bestMoves.Add(m)
            End If
        Next

        If bestMoves.Count <> 0 Then
            Return bestMoves(Math.Round(Rnd() * (bestMoves.Count - 1)))
        End If

        ' Check if any moves lead to capture.
        For Each m In moves
            If gameBoard.indexes(m.endOfMove) <> 0 Then
                bestMoves.Add(m)
            End If
        Next m

        If bestMoves.Count <> 0 Then
            Return bestMoves(Math.Round(Rnd() * (bestMoves.Count - 1)))
        End If

        ' Check if any moves go forward.
        For Each m In moves
            If Math.Floor(m.endOfMove / 10) < Math.Floor(m.startOfMove / 10) Then
                bestMoves.Add(m)
            End If
        Next

        If bestMoves.Count <> 0 Then
            Return bestMoves(Math.Round(Rnd() * (bestMoves.Count - 1)))
        Else
            Return moves(Math.Round(Rnd() * (moves.Count - 1)))
        End If

    End Function
    Private Function alphaBeta(depth As Integer, alpha As Integer, beta As Integer, playersTurn As Boolean)

        ' Recursively executes a minimax evaluation of the game tree with alpha-beta pruning.

        Dim value As Integer

        ' If the ply has been reached, return the evaluation value.
        If (depth = ply) Then Return evaluateBoard(gameBoardCopy)

        ' Work out the possible moves.
        Dim possibleMoves As New List(Of ChessMove)
        If depth <= 1 Then
            possibleMoves = generateAllMoves(gameBoardCopy, playersTurn, True)
            ' If stalemate or checkmate.
            If possibleMoves.Count = 0 Then
                If playersTurn Then
                    If isInCheck(gameBoardCopy, True) Then : Return -infinity
                    Else : Return 0
                    End If
                Else
                    If isInCheck(gameBoardCopy, False) Then : Return infinity
                    Else : Return 0
                    End If
                End If
            End If
        Else
            possibleMoves = generateAllMoves(gameBoardCopy, playersTurn, False)
        End If

        ' Alpha-beta evaluate stuff.
        If playersTurn Then
            For Each possibleMove In possibleMoves
                ' Make move on board.
                Dim pawnPromotion As Boolean = False
                If moveIsPawnPromotion(gameBoardCopy, possibleMove) Then pawnPromotion = True
                Dim tempEndPiece As Integer = gameBoardCopy.indexes(possibleMove.endOfMove)
                movePiece(gameBoardCopy, possibleMove)
                ' Evaluate deeper.
                value = alphaBeta(depth + 1, alpha, beta, Not playersTurn)
                If value > alpha Then alpha = value ' We have found a better best move.
                If alpha > beta Then
                    undoMove(gameBoardCopy, possibleMove, tempEndPiece, pawnPromotion)
                    Return alpha ' Cut off.
                End If
                ' Undo move on board.
                undoMove(gameBoardCopy, possibleMove, tempEndPiece, pawnPromotion)
            Next possibleMove
            Return alpha ' This is our best move.
        Else
            For Each possibleMove In possibleMoves
                ' Make move on board.
                Dim pawnPromotion As Boolean = False
                If moveIsPawnPromotion(gameBoardCopy, possibleMove) Then pawnPromotion = True
                Dim tempEndPiece As Integer = gameBoardCopy.indexes(possibleMove.endOfMove)
                movePiece(gameBoardCopy, possibleMove)
                ' Evaluate deeper.
                value = alphaBeta(depth + 1, alpha, beta, Not playersTurn)
                If value < beta Then beta = value ' Opponent has found a better worse move.
                If alpha > beta Then
                    undoMove(gameBoardCopy, possibleMove, tempEndPiece, pawnPromotion)
                    Return beta ' Cut off.
                End If
                ' Undo move on board.
                undoMove(gameBoardCopy, possibleMove, tempEndPiece, pawnPromotion)
            Next possibleMove
            Return beta ' This is the opponent's best move.
        End If

    End Function
    Private Function evaluateBoard(board As ChessBoard)

        ' Performs an evaluation of the current board, a negative score being good for black, and a positive score being good for white.
        ' At present, this function only counts the material of either side to determine a value.

        Return board.materialValue

    End Function

    ' Utilities.
    Private Function copyOfChessBoard(board As ChessBoard)

        ' Return a copy of the given board. 

        Dim copyBoard As New ChessBoard

        For i = 0 To 20 : copyBoard.indexes(i) = 99 : Next
        For i = 21 To 98 : copyBoard.indexes(i) = board.indexes(i) : Next
        For i = 99 To 119 : copyBoard.indexes(i) = 99 : Next

        copyBoard.whiteIndexes.AddRange(board.whiteIndexes)
        copyBoard.blackIndexes.AddRange(board.blackIndexes)
        copyBoard.materialValue = board.materialValue

        Return copyBoard

    End Function
    Private Function translateTo1D(x As Integer, y As Integer)

        ' Translate the 2D coordinates to the 1D coordinate.

        Return x - 10 * y + 91

    End Function

    ' Functions that are useful for testing/debugging.
    Private Function boardToText(board As ChessBoard)

        ' Translates a ChessBoard in to a string.

        Dim msg As String = ""
        For y = 0 To 7
            For x = 0 To 7
                Dim currentPiece As Integer = board.indexes(translateTo1D(x, y))
                If currentPiece >= 0 Then
                    msg += "   " & currentPiece & "   "
                Else
                    msg += "  " & currentPiece & "  "
                End If
            Next
            msg += vbNewLine
        Next
        msg.Remove(msg.Count - 1)

        Return msg

    End Function
    Private Sub setCustomBoard()

        ' Set up a custom board for testing purposes.

        removeAllPieces(gameBoard)

        ' Whites:
        placeNewPiece(gameBoard, 6, 31)
        placeNewPiece(gameBoard, 1, 54)
        ' Blacks:
        placeNewPiece(gameBoard, -6, 98)
        placeNewPiece(gameBoard, -1, 34)
        placeNewPiece(gameBoard, -2, 65)


    End Sub

End Class