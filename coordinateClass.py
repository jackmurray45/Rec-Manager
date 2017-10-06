from collections import defaultdict
def split_values(coord:str)-> [str, str]:
    letters = ''
    nums = ''
    for i in coord:
        try:
            int(i)
            nums+=i
        except:
            letters+=i
    return [letters, nums]


def alphabet_location(letter:str)-> int:
    alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
    spots = defaultdict(int)
    for i in range(len(alphabet)):
        for j in letter:
            if j.upper() == alphabet[i]:
                spots[j] = i
    return spots





class Coordinate:
    def __init__(self, coord='A1'):
        self.coord = split_values(coord)
        self.alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
        self.letter_locations = alphabet_location(self.coord[0])

    def value(self):
        return self.coord[0]+self.coord[1]

    def move_down(self):
        self.coord[1] = str(int(self.coord[1])+1)
        return self.coord[0]+self.coord[1]

    def move_up(self):
        if self.coord[1] != '1':
            self.coord[1] = str(int(self.coord[1])-1)
        return self.coord[0]+self.coord[1]

    def move_right(self):
        old = self.coord[0]
        self.coord[0] = self.alphabet[self.letter_locations[self.coord[0]]+1]
        self.letter_locations[self.coord[0]] = self.letter_locations[old]+1
        del self.letter_locations[old]
        return self.coord[0]+self.coord[1]

    def move_left(self):
        if self.coord[0] != 'A':
            old = self.coord[0]
            self.coord[0] = self.alphabet[self.letter_locations[self.coord[0]]-1]
            self.letter_locations[self.coord[0]] = self.letter_locations[old]-1
            del self.letter_locations[old]
        return self.coord[0]+self.coord[1]

    def set_far_left(self):
        self.coord[0]='A'

    def print(self):
        print(self.coord[0]+self.coord[1])
        
            
        


    
        
        
