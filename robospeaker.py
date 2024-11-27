import win32com.client as wincl

speaker_number = 1
spk = wincl.Dispatch("SAPI.SpVoice")
vcs = spk.GetVoices()
SVSFlag = 11
print(vcs.Item (speaker_number) .GetAttribute ("Name")) # speaker name
spk.Voice
spk.SetVoice(vcs.Item(speaker_number)) # set voice (see Windows Text-to-Speech settings)

print("Hi, I am your Robot!")
while True:
    try:
        user_input = input("Enter anything you want me to speak (type 'exit' to quit): ")
        if user_input.lower() == 'exit':
            break
        spk.Speak(user_input)
    except Exception as e:
        print(f"Error: {e}")

