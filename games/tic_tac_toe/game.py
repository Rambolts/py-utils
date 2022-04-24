from enum import Enum

class Cell(Enum):
    EMPTY = ' ',
    X     = 'X',
    O     = 'O'

class Game(object):
    def __init__(self, x_player=None, o_player=None):
        self.board = [
            [Cell.EMPTY, Cell.EMPTY, Cell.EMPTY],
            [Cell.EMPTY, Cell.EMPTY, Cell.EMPTY],
            [Cell.EMPTY, Cell.EMPTY, Cell.EMPTY],
        ]
        self.is_x_turn = True
        self.x_player = x_player
        self.o_player = o_player

    def _check_winner(self):
        # check rows
        for row in self.board:
            if len(set(row)) == 1 and row[0] != Cell.EMPTY:
                return row[0]

        # check columns
        for column in [*zip(*self.board)]:
            if len(set(column)) == 1 and column[0] != Cell.EMPTY:
                return column[0]

        # check diagonals
        size = len(self.board)
        major_diagonal = set()
        minor_diagonal = set()
        for i in range(size):
            major_diagonal.add(self.board[i][i])
            minor_diagonal.add(self.board[i][size-i-1])

        if len(major_diagonal) == 1 and self.board[0][0] != Cell.EMPTY:
            return self.board[0][0]

        if len(minor_diagonal) == 1 and self.board[0][size-1] != Cell.EMPTY:
            return self.board[0][size-1]

    def _check_draw(self) -> bool:
        for row in self.board:
            for cell in row:
                if cell == Cell.EMPTY:
                    return False
        return True

    def make_turn(self, turn: int, piece: Cell):
        size = len(self.board)
        i = turn // size
        j = turn % size
        self.board[i][j] = piece
        self.is_x_turn = not self.is_x_turn

    def print_board(self):
        pass

    def is_game_over(self):
        winner = self._check_winner()
        if winner is not None:
            return winner
        return self._check_draw()

    def print_winner(self, winner):
        print(winner)
        pass

    def play(self):
        self.print_board()
        while not (winner := self.is_game_over()):
            if self.is_x_turn:
                turn = self.x_player.get_turn(self.board)
                piece = Cell.X
            else:
                turn = self.o_player.get_turn(self.board)
                piece = Cell.O
            self.make_turn(turn, piece)
            self.print_board()
        self.print_winner(winner)