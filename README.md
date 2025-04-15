# 🔢 KG Career Correction System (2021)

A Python-based GUI application to semi-automatically analyze and score career planning submissions using natural language similarity, NG-word detection, and duplication analysis.

---

## 🌍 Overview

This tool was developed to support instructors in reviewing large-scale student career reports efficiently. It parses CSV submissions, identifies copy-paste behavior, counts flagged words (NG words), and generates result sheets for teachers with customizable thresholds.

Originally designed for use at Kwansei Gakuin University, this application supports Japanese input analysis using Janome and Levenshtein similarity.

---

## 🌐 Main Features

- ✏️ GUI interface for file input and NG-word list control
- ⚡ Real-time Levenshtein similarity comparison among student submissions
- 🔢 Threshold adjustment for similarity detection
- 🌐 Morphological analysis with Janome (replacing MeCab)
- 🔍 NG-word detection in answers
- 🎓 Automatic teacher assignment based on student ID prefix
- 📎 Export results to `.csv` with detailed flags

---

## 📄 Tech Stack

- Language: Python 3.x
- GUI: Tkinter
- NLP: Janome, pykakasi
- Data: pandas, csv
- Similarity: python-Levenshtein
- Platform: Windows desktop (supports Excel CP932 encoding)

---

## 🚀 How to Run

1. Place your `report-answer-*.csv` and `teacher_list.xlsx` in the same directory.
2. Run:

```bash
python main.py
```

or just double-click:
```bash
run_program.bat
```

3. From the GUI:
   - Select the `.csv` to process
   - Add NG words
   - Adjust similarity threshold
   - Start the checking process

4. Output: `result.csv` with detailed annotations for teachers

---

## 🌐 File Outputs

- `converted-report-*.csv`: Normalized input
- `result.csv`: Final output for grading

Each row includes:
- Teacher name
- Student ID
- Keywords (1-5)
- Copied-from info (if similar answer found)
- NG word count and content
- Retake flag (same student found more than once)

---

## ⚖️ Detection Logic

- **Similarity Check**: Uses Levenshtein distance, normalized to a 0.0~1.0 range
- **Copy Grouping**: Assigns shared "Pattern" ID for highly similar submissions
- **NG Word Check**: Flags if any of the registered terms appear in keyword answers
- **Retake Check**: Flags the same student ID appearing twice

---

## 📁 Project Structure

```
kg-career-correction-system/
├── main.py                # Main GUI & logic controller
├── downloader.py          # (Optional) Data loader
├── run_program.bat        # Batch launcher for Windows
├── teacher_list.xlsx      # Excel mapping for teacher assignments
├── data/                  # Temporary CSV storage & results
└── venv/                  # Virtual environment (not committed to Git)
```

> ✅ You can launch the app via `python main.py` or by double-clicking `run_program.bat`

---

## 📄 License

MIT License

---

Created with care by [@kamekingdom](https://github.com/kamekingdom)
