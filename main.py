import docx
import pyttsx3

engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[0].id)

def speak(audio):
    engine.say(audio)
    engine.runAndWait()

def ReadingTextDocuments():
    fileName = input("Enter the file name with the extension: ")
    doc = docx.Document(fileName)

    completedText = []

    for paragraph in doc.paragraphs:
        completedText.append(paragraph.text)
    print('\n' .join(completedText))
    return '\n' .join(completedText)
    
speak(ReadingTextDocuments())
