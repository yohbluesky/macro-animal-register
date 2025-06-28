# 🐾 Animal Registration Macro (Excel VBA)

This project is a simple automation created in Excel using VBA macros.

The goal is to simulate a basic **animal registration system** where data entered in the "BASE" worksheet is automatically copied and organized into the "BANCO DE DADOS" worksheet using relative references and macro recording.

## 📌 Features

- Uses recorded macro logic with manual adjustments
- Moves data from cells `B2` to `B5` (on sheet "BASE")
- Pastes the data in the next available rows in "BANCO DE DADOS"
- Clears the input fields after saving the data
- Demonstrates usage of:
  - `Range`, `Select`, `Offset`
  - `Sheets` navigation
  - Basic data organization logic

## 📄 VBA Code Example

```vba
Sub Cadastro2()
    Range("B2").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("B25").Select
    Sheets("BANCO DE DADOS").Select
    Range("E1").Select
    Selection.End(xlDown).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlUp).Select
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Paste

    Sheets("BASE").Select
    Range("B3").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("BANCO DE DADOS").Select
    Range("B1048576").Select
    Selection.End(xlUp).Select
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Paste

    Sheets("BASE").Select
    Range("B4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("BANCO DE DADOS").Select
    Range("C1048576").Select
    Selection.End(xlUp).Select
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Paste

    Sheets("BASE").Select
    Range("B5").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("BANCO DE DADOS").Select
    Range("D1048576").Select
    Selection.End(xlUp).Select
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Paste

    Sheets("BASE").Select
    Range("B2:B5").Select
    Selection.ClearContents
    Range("D5").Select
End Sub

