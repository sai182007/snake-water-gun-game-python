import random
import win32com.client as wincom
speaker=wincom.Dispatch("SAPI.SpVoice")
comp=0
you=0
ties=0


def win(comp,you):
    if comp>you:
        print(f"======================(you LOSE!, with {you} rounds and {ties} ties.)======================")
        speaker.speak(f"you LOSE!, with {you} rounds and {ties} ties.")
    elif comp<you:
        print(f"======================(you WIN!, with {you} rounds {ties} ties.)======================")
        speaker.speak(f"you WIN!, with {you} rounds {ties} ties.")
    else:
        print(f"======================(its a DRAW!, with {you} rounds {ties} ties.)======================")
        speaker.speak(f"its a DRAW!, with {you} rounds {ties} ties.")
    

def play_game():
    for i in range(1,11):
        choices=["snake","water","gun"]
        comp_game_choice=random.choice(choices)
        print(f"==============Round {i}==============")
        entered_game_choice=input("Enter your choice between ('snake', 'water' or 'gun') :").lower()
        if comp_game_choice==entered_game_choice:
            print("its a tie!!!")
            speaker.speak("its a tie!!!")
            global ties
            ties=ties+1
        elif (entered_game_choice=="snake" and comp_game_choice=="water")or\
            (entered_game_choice=="water" and comp_game_choice=="gun")or\
            (entered_game_choice=="gun" and comp_game_choice=="snake"):
            print("you won this round!")
            speaker.speak("you won this round!")
            global you
            you=you+1
        else:
            print("you lose this round!")
            speaker.speak("you lose this round!")
            global comp
            comp=comp+1
        

while True:
    print("---------------------------------------------------------------------")
    print("do you wanna play the game?(enter your choice here as 'yes' or 'no'): ")
    speaker.speak("do you wanna play the game?(enter your choice here as 'yes' or 'no'")
    print("---------------------------------------------------------------------")
    entered_choice=input("ENTER HERE:")
    if entered_choice=="yes":
        play_game()
        win(comp,you)
    elif entered_choice=="no":
        print("okie, you are out!")
        speaker.speak("okie, you are out!")
        break
    else:
        speaker.speak("wrong choices are being entered.")
        raise ValueError("wrong choices are being entered.")