from tic_tac_toe.game import Game
from tic_tac_toe.player import RandomPlayer

def main():
    x_player = RandomPlayer()
    o_player = RandomPlayer()
    game = Game(x_player=x_player, o_player=o_player)
    game.play()

if __name__ == "__main__":
    main()