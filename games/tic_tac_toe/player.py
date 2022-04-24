from abc import ABC, abstractmethod
from tic_tac_toe.game import Cell
import string
import random

class Player(ABC):
    def __init__(self, name=None, frontend=None):
        self.name = name
        self.frontend = frontend

    @abstractmethod
    def get_turn(self, board) -> int:  # Python 3.5+
        pass


class RandomPlayer(Player):
    def __init__(self):
        random_name = "".join([random.choice(string.ascii_letters)
                               for _ in range(8)])
        super().__init__(name=random_name)

    def get_turn(self, board):
        available_cells = []
        for i, row in enumerate(board):
            for j, column in enumerate(row):
                if board[i][j] == Cell.EMPTY:
                    cell_index = i * len(board) + j
                    available_cells.append(cell_index)
        return random.choice(available_cells)