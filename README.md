# 🎓 Automatic Certificate Generation using MATLAB

This project automates the generation of participation certificates using MATLAB by reading participant details (like Name and Topic) from an Excel file and placing them on a certificate template image.

---

## 📌 Project Features

- ✅ Reads participant data from an Excel file (`.xlsx`)
- ✅ Automatically inserts names and topics into a certificate template
- ✅ Saves each certificate as a `.png` file with a unique filename
- ✅ Simple and customizable using MATLAB's `insertText()` function

---

## 🛠️ Requirements

- MATLAB (R2015b or later recommended)
- Image Processing Toolbox
- Certificate Template (PNG/JPG)
- Excel file (`Registration.xlsx`) with the following format:


### Example Excel Format:

| S.No | Name     | Topic     |
|------|----------|-----------|
| 1    | Alice    | MATLAB    |
| 2    | Bob      | Python    |
| 3    | Charlie  | Simulink  |

---

Visual Layout of the Certificate:
  ------------------------------------------------
|                                              |
|      Certificate of Participation            |
|                                              |
|                Bob                          ← (800, 520)
|                                              |
|           Participated in Python            ← (720, 820)
|                                              |
------------------------------------------------

## SAMPLE OUTPUT
--------------------------------------------
|                                          |
|       Certificate of Participation       |
|                                          |
|                Bob                      |
|                                          |
|          Participated in Python         |
|                                          |
--------------------------------------------

💡 Customization
Change text position by editing the position array in the script.
Modify font size, color, or text alignment via insertText() options.
Save as .pdf instead of .png using exportgraphics() or print().

📜 License
This project is open-source and free to use for academic and non-commercial purposes.

