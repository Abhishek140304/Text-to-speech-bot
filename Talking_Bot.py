import win32com.client
import sys

def main():
    try:
        speaker = win32com.client.Dispatch("SAPI.SpVoice")
    except Exception as e:
        print("Failed initialising the SAPI Engine: ")
        print(f"details: {e}")
        sys.exit(1)

    speaker.speak("This is the basic speech bot, press q to exit")
    while True:
        text = input()
        if text.lower() == 'q':
            break
        
        if text.strip():
            speaker.Speak(text)

    speaker.speak("Exiting the program")
    

if __name__=="__main__":
    main()