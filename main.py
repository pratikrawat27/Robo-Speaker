import win32com.client as w

if __name__ == '__main__':
    print("Welcome to Robo Speaker")
    while True:
        c = input("Write something to pronounce or to speak : ")
        if c == "END":
            speaker.Speak("Chalo BYE.")
            print("Program Ended.")
            break
        speaker = w.Dispatch("SAPI.SpVoice")
        speaker.Speak(c)