# Translation Review

A PyQt5-based GUI application for evaluating and scoring Arabic translations of English text. The application reads Excel files with alternating content/score rows and allows reviewers to score multiple translation models.

## Features

- **Excel Integration**: Load and save scores directly to Excel files
- **Multi-Model Evaluation**: Score translations from 4 different models:
  - Tencent LM
  - Tencent RAG
  - Gemma V2
  - Gemma V3
- **Scoring System**: Rate each translation on a scale of 0-10
- **JSON Export**: Automatically saves evaluation data to JSON files for further analysis
- **RTL Support**: Proper right-to-left text display for Arabic translations
- **Easy Navigation**: Move between rows with Previous/Next buttons

## Installation

### Prerequisites

using the requirements:

```bash
python pip install -r requirements.txt
```

## Excel File Format

The application expects Excel files with a specific structure:

- **Row 1**: Headers (column titles)
- **Row 2**: Empty
- **Row 3**: First content row (English text + 4 Arabic translations)
- **Row 4**: Empty (reserved for scores)
- **Row 5**: Second content row
- **Row 6**: Empty (for scores)
- And so on...

### Column Structure

| Column | Content |
|--------|---------|
| A | English text |
| B | Translation 1 (Tencent LM) |
| C | Translation 2 (Tencent RAG) |
| D | Translation 3 (Gemma V2) |
| E | Translation 4 (Gemma V3) |

## Usage

### Running the Application

```bash
python evaluation_gui_qt.py
```

Or make it executable:

```bash
chmod +x evaluation_gui_qt.py
./evaluation_gui_qt.py
```

### Workflow

1. **Load Excel File**: Click "📁 Load Excel..." to open your Excel file
2. **Select Sheet**: Choose the worksheet from the dropdown if multiple sheets exist
3. **Navigate Rows**: Use the row selector or Previous/Next buttons to move between content rows
4. **Review Translations**: Read the English text and compare the 4 Arabic translations
5. **Score**: Select a score (0-10) for each translation using the radio buttons
6. **Save**: Click "💾 Save Scores" to write scores back to Excel and update JSON files

### Keyboard Shortcuts

- Navigate using the row spinner or Previous/Next buttons
- Scores persist in the Excel file and can be edited later

## Output Files

The application generates three JSON files for analysis:

### 1. `evaluated_rows.json`
Tracks which rows have been evaluated per sheet:
```json
{
  "Sheet1": [0, 1, 2, 5]
}
```

### 2. `model_scores.json`
Contains aggregate statistics for each model:
```json
{
  "Tencent LM": {
    "total_score": 85,
    "count": 10,
    "average": 8.5,
    "comments_count": 0
  }
}
```

### 3. `detailed_comments.json`
Stores individual evaluation records:
```json
[
  {
    "sheet": "Sheet1",
    "row_index": 0,
    "model": "Tencent LM",
    "score": 9,
    "comment": "",
    "english_text": "Hello world",
    "translation": "مرحبا بالعالم"
  }
]
```

## Interface Overview

- **Top Bar**: File loading, sheet selection, row navigation, and save controls
- **English Text Section**: Displays the source English text
- **Translation Sections**: Four expandable sections showing each model's Arabic translation with scoring options
- **Score Radio Buttons**: 0-10 rating scale for each translation
- **Info Bar**: Shows current sheet, row number, and Excel row index

## Features in Detail

### Automatic Score Persistence
- Scores are saved both to Excel (in the row below the content) and to JSON files
- Previously scored rows display their existing scores when loaded
- Re-scoring a row updates both Excel and JSON data

### Multi-Sheet Support
- Handle Excel files with multiple worksheets
- Switch between sheets without reloading the file
- Independent scoring for each sheet