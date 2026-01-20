# Duplicate-detection-using-Macro

Process to Identify Possible Duplicate Cases

1. Concatenate fields in the Parameter column of the Master Dump according to the specified logic.
After concatenation, replace formulas with values using Paste Special.

2. Sort the Master Dump based on the Parameter column.

3. In the Finding column, apply the formula designed to flag potential duplicates by marking them with “a”.

4. Sort the Master Dump by the Finding column to bring all entries marked “a” to the top.
Copy the corresponding concatenated values (from the Parameter column) for all rows marked “a” (e.g., rows 2–10) into a new sheet for VLOOKUP processing.

5. In the Master Dump, next to the Parameter column, apply VLOOKUP to the copied values and then Paste Special → Values.

6. Sort the Master Dump using two columns:


VLOOKUP result
Actual Invoice Number

7. Delete all rows with #N/A in the VLOOKUP column.
The remaining rows represent the possible duplicate cases for the current parameter.

8. Copy all identified duplicate rows from the VLOOKUP column down to the last matching row into the Possible Duplicate Output file.

9. Repeat Steps 1–8 for all six parameters, continuously appending each parameter’s duplicate results into the same output file.



Post‑Processing of the Combined Output

10. After completing all six parameter checks, perform the steps below to insert a separator row between each duplicate pair/case:

Next to the Parameter column, enter a formula based on the parameter logic (e.g., 1=2) and convert formulas to values.
Apply a filter and select only “False”.
Select the entire filtered rows, right‑click, and choose Insert Row to create separation between duplicate cases for easier analysis.



1) Excel VBA Macro — “Possible Duplicate Detection (End‑to‑End)”

How to use

1. Save your workbook as .xlsm.
2. Press ALT+F11 → Insert → Module → paste the macro below.
3. Adjust the configuration block at the top (sheet names, header names, parameter column names if yours differ).
4. Close the editor → Developer → Macros → run FindPossibleDuplicates.

![Image](https://github.com/user-attachments/assets/15ae41b0-c6ac-4b36-b37c-de83b2c4be56)
![Image](https://github.com/user-attachments/assets/31fd5993-3b5e-4ece-a799-1b4380761e9e)
![Image](https://github.com/user-attachments/assets/f30f7f5b-e10e-42c1-a53d-a45a23ed9456)
![Image](https://github.com/user-attachments/assets/d5270c10-2a13-4af2-b2e1-be6f526e4ac7)
![Image](https://github.com/user-attachments/assets/399355bb-9160-4013-8d89-f2f6c2a52846)
![Image](https://github.com/user-attachments/assets/5a4d78d5-572a-4687-ba32-89fd4fe96007)

#MACRO FOR 5 CHAR
Option Explicit

'==========================
' Helper: Strip Accents
'==========================
Private Function StripAccents(ByVal s As String) As String
    Dim i As Long
    Dim fromChars As Variant, toChars As Variant

    fromChars = Array( _
        "à", "â", "ä", "á", "ã", "æ", "ç", "é", "è", "ê", "ë", "î", "ï", "í", "ô", "ö", "ó", "õ", "ù", "û", "ü", "ú", "ÿ", "ñ", _
        "À", "Â", "Ä", "Á", "Ã", "Æ", "Ç", "É", "È", "Ê", "Ë", "Î", "Ï", "Í", "Ô", "Ö", "Ó", "Õ", "Ù", "Û", "Ü", "Ú", "Ÿ", "Ñ" _
    )
    toChars = Array( _
        "a", "a", "a", "a", "a", "ae", "c", "e", "e", "e", "e", "i", "i", "i", "o", "o", "o", "o", "u", "u", "u", "u", "y", "n", _
        "A", "A", "A", "A", "A", "AE", "C", "E", "E", "E", "E", "I", "I", "I", "O", "O", "O", "O", "U", "U", "U", "U", "Y", "N" _
    )

    For i = LBound(fromChars) To UBound(fromChars)
        s = Replace(s, fromChars(i), toChars(i))
    Next i
    StripAccents = s
End Function

'==========================
' Helper: Remove spaces & special characters
' Keeps only A–Z, a–z, 0–9
'==========================
Private Function KeepAlnumOnly(ByVal s As String) As String
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    With re
        .Global = True
        .IgnoreCase = True
        .Pattern = "[^A-Za-z0-9]"   ' remove anything not A–Z, a–z, 0–9
    End With
    KeepAlnumOnly = re.Replace(s, "")
End Function

'==========================
' MAIN: Process Column H -> Column M
' Order: Read H -> Strip Accents -> Remove spaces/specials -> UPPER -> Remove "THE" at start -> LEFT 5 -> Write M
'==========================
Sub Process_H_To_M_InRequestedOrder()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim srcVal As Variant
    Dim cleaned As String

    Set ws = ActiveSheet  ' or ThisWorkbook.Worksheets("Sheet1")

    ' Determine last used row in Column H
    lastRow = ws.Cells(ws.Rows.Count, "H").End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "No data found in Column H.", vbInformation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    For r = 2 To lastRow   ' start at row 2 (skip header)
        srcVal = ws.Cells(r, "H").Value2

        If Len(Trim$(CStr(srcVal))) > 0 Then
            cleaned = CStr(srcVal)

            ' 1) Remove accents
            cleaned = StripAccents(cleaned)

            ' 2) Remove spaces & special characters (keep only A–Z, a–z, 0–9)
            cleaned = KeepAlnumOnly(cleaned)

            ' 3) Convert to UPPERCASE
            cleaned = UCase$(cleaned)

            ' 4) Remove "THE" if at the start (case-insensitive)
            '    Since we already uppercased, a simple LEFT check is enough
            If Left$(cleaned, 3) = "THE" Then
                cleaned = Mid$(cleaned, 4)
            End If

            ' 5) Keep only the first 5 characters
            cleaned = Left$(cleaned, 5)

            ' 6) Output to Column M
            ws.Cells(r, "M").Value = cleaned
        Else
            ws.Cells(r, "M").ClearContents
        End If
    Next r

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    MsgBox "Done: H ? M with steps in requested order.", vbInformation
End Sub
<img width="272" height="2442" alt="image" src="https://github.com/user-attachments/assets/15f369aa-fe0c-4f38-a8db-6abd7103bdf3" />



![Image](https://github.com/user-attachments/assets/5f009ab6-7c94-456c-8513-50d79b4ad480)
![Image](https://github.com/user-attachments/assets/1136a848-35c6-4c99-9c5c-f1c7c13c2c9b)


Output of duplicate 
![Image](https://github.com/user-attachments/assets/d9385c76-46ef-4d62-8fee-34a7cd92899a)
