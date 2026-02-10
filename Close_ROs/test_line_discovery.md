# Test Cases for DiscoverLineLetters Function

## Overview
The `DiscoverLineLetters()` function dynamically discovers which line letters (A, B, C, etc.) are present on the RO Detail screen by reading the LC column.

## Test Scenarios

### Test 1: Consecutive Lines (A, B, C)
**Setup:** Screen has line letters A, B, C in rows 7, 8, 9
**Expected Result:** Function returns Array("A", "B", "C")
**Status:** Should work with existing code

### Test 2: Non-Consecutive Lines (A, C) - Missing B
**Setup:** Screen has line letters A, C in rows 7, 8 (B is skipped)
**Expected Result:** Function returns Array("A", "C")
**Status:** This is the primary bug fix scenario
**Verification:** 
- Function reads row 7, finds "A", adds to array
- Function reads row 8, finds "C", adds to array
- Function reads row 9, finds empty or non-letter, increments consecutiveEmptyCount
- Function reads row 10, finds empty or non-letter, consecutiveEmptyCount = 2, exits loop
- Returns Array("A", "C")

### Test 3: Single Line (A only)
**Setup:** Screen has only line letter A in row 7
**Expected Result:** Function returns Array("A")
**Status:** Should work with consecutiveEmptyCount logic

### Test 4: No Lines Found
**Setup:** Screen has no line letters in expected rows
**Expected Result:** Function returns default Array("A", "B", "C") with warning logged
**Status:** Fallback behavior to maintain backward compatibility

### Test 5: Multiple Lines (A through F)
**Setup:** Screen has line letters A, B, C, D, E, F
**Expected Result:** Function returns Array("A", "B", "C", "D", "E", "F")
**Status:** Should work up to maxLinesToCheck limit

### Test 6: Non-Consecutive with Gap in Middle (A, B, D, E)
**Setup:** Screen has line letters A, B, D, E (C is missing)
**Expected Result:** Function returns Array("A", "B", "D", "E")
**Status:** Should work - consecutive empty counter only stops after 2 consecutive empties

## Implementation Notes

### Screen Coordinate Assumptions
- **Row 7** is the first data row (after LC header on row 6)
- **Column 1** contains the line letter
- These coordinates may need adjustment based on actual screen layout

### Robustness Features
1. **Consecutive Empty Detection**: Stops reading after 2 consecutive non-letter rows
2. **Error Handling**: Uses On Error Resume Next for ReadScreen calls
3. **Validation**: Only accepts A-Z characters
4. **Fallback**: Returns default array if no letters found
5. **Logging**: Logs discovered letters for debugging

## Usage in Scripts

### Close_ROs_Pt1.vbs
```vbscript
' Before (hardcoded):
commands = Array("A", "B", "C")

' After (dynamic discovery):
commands = DiscoverLineLetters()
```

### Close_ROs_Pt2.vbs
```vbscript
' Before (hardcoded):
AddStory bzhao, "B"
AddStory bzhao, "C"

' After (dynamic discovery):
Dim lineLetters, i
lineLetters = DiscoverLineLetters()
For i = 0 To UBound(lineLetters)
    If UCase(lineLetters(i)) <> "A" Then
        AddStory bzhao, lineLetters(i)
    End If
Next
```

## Future Enhancements

### If Screen Coordinates Are Incorrect
If the hardcoded row 7/column 1 coordinates don't match the actual screen layout:

1. **Option 1**: Make coordinates configurable (add constants at top of file)
2. **Option 2**: Search for "LC" header first, then read below it
3. **Option 3**: Use the alternative approach mentioned in the issue: try each line letter sequentially and check for "Line X does not exist" error

### Alternative Discovery Method (Not Implemented Yet)
```vbscript
Function DiscoverLineLettersAlternative()
    ' Try each letter A-J and check for error messages
    Dim testLetters, i, lineLetter, screenContent
    testLetters = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J")
    Dim validLetters()
    Dim validCount: validCount = 0
    
    For i = 0 To UBound(testLetters)
        lineLetter = testLetters(i)
        ' Try to navigate to line (implementation would vary by screen)
        ' Check if error message appears
        ' If no error, add to validLetters array
    Next
    
    DiscoverLineLettersAlternative = validLetters
End Function
```
