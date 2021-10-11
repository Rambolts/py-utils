import random

options = {'r':'rock', 'p':'paper', 's':'scisor'}

def compare(hand, robot_hand):
    print(f'{options[hand]} VS {options[robot_hand]}')
    if(hand == robot_hand):
        print('Draw')
    elif (hand=='r' and robot_hand=='s') or (hand=='s' and robot_hand=='p') or (hand=='p' and robot_hand=='r'):
        print('U Win!')
    else:
        print('U Loose!')

def main():
    print('SELECT YOUR HAND:\n\'r\', \'p\' or \'s\'')
    hand = str.lower(input())
    
    while hand in options.keys():
        robot_hand = random.choice(list(options.keys()))
        compare(hand, robot_hand)
        hand = str.lower(input())


if __name__ == '__main__':
    #main()
    try:
        main()
    except Exception as e:
        print('error: ', e)