# Excel to Anki

A small Windows desktop app that filters vocabulary cards from an Excel workbook and enters them into Anki's **Add** window.

## Setup

```powershell
python -m pip install -r requirements.txt
python app.py
```

The workbook must include these columns:

- `Chinese`
- `Pinyin`
- `English`
- `Lesson`
- `Character`

## Using the app

1. Open Anki's **Add** window and arrange its fields.
2. Launch `app.py`, choose the workbook, select the lesson and card type, and preview the matching rows.
3. Check the four screen coordinates. The defaults preserve the positions used by the original script.
4. Click **Start sending cards**. You have three seconds to focus Anki. Do not move the Anki window while the app is running.
5. Press **Stop** to stop after the current card.

The app uses screen automation, so coordinate positions are specific to your display layout. Test first with a workbook that contains a single card.
