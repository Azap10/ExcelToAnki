# Flashcard Factory

A small Windows desktop app that helps create Anki Chinese<->English flashcards using PDF OCR and hotkey workflows

## Setup

```powershell
python -m pip install -r requirements.txt
python app.py
```

## Using the app

1. The app contains 3 tabs for adding flashcards and a settings tab.
2. The first tab is for selecting words. You can upload a PDF and highlight the area of the PDF you'd like to select from. If the PDF does not have plaintext the Flashcard Factory will use OCR to recognize the word. Alternatively, you can simply type in the Chinese word/phrase to add it to the list.
3. The second tab is for flashcard setup. Your word list will be processed with the CC-CEDICT Chinese/English dictionary. You can create the desired definition using hotkeys.
4. The third tab is for adding to your Anki collection. Select a target deck, preview the cards you have made, and add.