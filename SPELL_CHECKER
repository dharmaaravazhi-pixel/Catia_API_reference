Public Type SpellIssue
    originalWord    As String
    suggestion      As String
    contextText     As String
    sheetIndex      As Long
    viewIndex       As Long
    textIndex       As Long
    tableIndex      As Long
    rowIndex        As Long
    colIndex        As Long
    isTable         As Boolean
End Type

Public issues() As SpellIssue
Public issueCount As Long
Public uniqueWords() As String
Public uniqueCount As Long

Public CATIA As Object
Public objWord As Object
Public oTempDoc As Object
Public dictCache As Object

Public DICT_FOLDER As String
Public DICT_FILE As String

Sub CATMain()

    DICT_FOLDER = Environ("USERPROFILE") & "\CATIA_SpellChecker\"
    DICT_FILE = DICT_FOLDER & "CustomDictionary.txt"

    Err.Clear
    
    On Error Resume Next
    
    Set CATIA = GetObject(, "CATIA.Application")
    
    If Err.Number <> 0 Or CATIA Is Nothing Then
        MsgBox "Could not connect to CATIA. Please ensure CATIA is running.", vbCritical, "Connection Error"
        Exit Sub
    End If
    
    Dim activeDoc As Object
    Set activeDoc = CATIA.ActiveDocument
    
    If Err.Number <> 0 Or activeDoc Is Nothing Then
        MsgBox "Please open a 2D Drawing before running the Spell Checker.", vbCritical, "No Document Found"
        Exit Sub
    End If
    
    On Error GoTo 0

    If TypeName(activeDoc) <> "DrawingDocument" Then
        MsgBox "This macro only works on 2D Drawing (.CATDrawing) files.", vbExclamation, "Invalid Document"
        Exit Sub
    End If

    On Error Resume Next
    Set objWord = CreateObject("Word.Application")
    
    If Err.Number <> 0 Or objWord Is Nothing Then
        MsgBox "Failed to connect to Microsoft Word.", vbCritical, "Dependency Error"
        Exit Sub
    End If
    
    objWord.Visible = False
    
    Set oTempDoc = objWord.Documents.Add
    
    Set dictCache = CreateObject("Scripting.Dictionary")
    dictCache.CompareMode = 1
    
    ReDim issues(100)
    ReDim uniqueWords(100)
    
    On Error GoTo ErrorHandler

    Dim customDict() As String
    customDict = LoadCustomDictionary()

    issueCount = 0

    Dim objSheets As Object, objSheet As Object
    Dim objViews As Object, objView As Object
    Dim objTexts As Object, objText As Object
    Dim objTables As Object, objTable As Object

    Dim s As Long, i As Long, j As Long
    Dim tb As Long, r As Long, c As Long
    Dim currentString As String

    Set objSheets = activeDoc.sheets

    For s = 1 To objSheets.count
        Set objSheet = objSheets.Item(s)
        Set objViews = objSheet.views

        For i = 1 To objViews.count
            Set objView = objViews.Item(i)

            Set objTexts = objView.texts
            For j = 1 To objTexts.count
                Set objText = objTexts.Item(j)
                currentString = objText.Text
                If Len(Trim(currentString)) > 0 Then
                    CollectIssues currentString, objWord, customDict, s, i, j, 0, 0, 0, False, issues, issueCount
                End If
            Next

            Set objTables = objView.tables
            If Not objTables Is Nothing Then
                For tb = 1 To objTables.count
                    Set objTable = objTables.Item(tb)
                    For r = 1 To objTable.NumberOfRows
                        For c = 1 To objTable.NumberOfColumns
                            currentString = objTable.GetCellString(r, c)
                            If Len(Trim(currentString)) > 0 Then
                                CollectIssues currentString, objWord, customDict, s, i, 0, tb, r, c, True, issues, issueCount
                            End If
                        Next
                    Next
                Next
            End If
        Next
    Next

    If issueCount = 0 Then
        MsgBox "No spelling errors found!", vbInformation, "Spell Check Complete"
        GoTo SafeExit
    End If

    uniqueCount = 0
    Dim k As Long, d As Long, isDup As Boolean

    For k = 0 To issueCount - 1
        isDup = False
        For d = 0 To uniqueCount - 1
            If LCase(uniqueWords(d)) = LCase(issues(k).originalWord) Then
                isDup = True
                Exit For
            End If
        Next d
        If Not isDup Then

            If uniqueCount > UBound(uniqueWords) Then
                ReDim Preserve uniqueWords(UBound(uniqueWords) + 100)
            End If
            uniqueWords(uniqueCount) = issues(k).originalWord
            uniqueCount = uniqueCount + 1
        End If
    Next k

    Dim summaryMsg As String
    summaryMsg = "Found " & uniqueCount & " unique misspelled word(s)" & vbCrLf & _
                 "(" & issueCount & " total occurrence(s) across drawing)" & vbCrLf & String(40, "-") & vbCrLf

    For k = 0 To uniqueCount - 1
        Dim dispSuggestion As String
        dispSuggestion = ""
        Dim m As Long
        For m = 0 To issueCount - 1
            If LCase(issues(m).originalWord) = LCase(uniqueWords(k)) Then
                dispSuggestion = issues(m).suggestion
                Exit For
            End If
        Next m

        summaryMsg = summaryMsg & (k + 1) & ".  " & uniqueWords(k)
        If dispSuggestion <> "" Then
            summaryMsg = summaryMsg & "   ->   " & dispSuggestion
        Else
            summaryMsg = summaryMsg & "   ->   (no suggestion)"
        End If
        summaryMsg = summaryMsg & vbCrLf
    Next k

    summaryMsg = summaryMsg & String(40, "-") & vbCrLf & vbCrLf & _
                 "Choose an action:" & vbCrLf & vbCrLf & _
                 "[Yes]     = Accept All suggestions at once" & vbCrLf & _
                 "[No]      = Review One by One" & vbCrLf & _
                 "[Cancel]  = Do nothing, exit"

    Dim userChoice As Integer
    userChoice = MsgBox(summaryMsg, vbYesNoCancel + vbQuestion, "CATIA Spell Checker")

    If userChoice = vbCancel Then
        MsgBox "Spell check cancelled. No changes made.", vbExclamation, "Cancelled"
        GoTo SafeExit
    End If

    Dim correctionsMade As Long
    correctionsMade = 0
    Dim finalWord As String, userInput As String, addChoice As Integer

    For k = 0 To uniqueCount - 1
        Dim currentOriginal As String
        Dim currentSuggestion As String

        currentOriginal = uniqueWords(k)
        currentSuggestion = ""
        
        For m = 0 To issueCount - 1
            If LCase(issues(m).originalWord) = LCase(currentOriginal) Then
                currentSuggestion = issues(m).suggestion
                Exit For
            End If
        Next m

        finalWord = currentOriginal

        If userChoice = vbYes Then
            If currentSuggestion <> "" And currentSuggestion <> "(check manually)" Then
                finalWord = currentSuggestion
            Else
                GoTo NextUnique
            End If

        ElseIf userChoice = vbNo Then
            Dim oneByOneMsg As String
            oneByOneMsg = "Word " & (k + 1) & " of " & uniqueCount & vbCrLf & String(30, "-") & vbCrLf & "Misspelled:  " & currentOriginal & vbCrLf
            If currentSuggestion <> "" And currentSuggestion <> "(check manually)" Then
                oneByOneMsg = oneByOneMsg & "Suggested:   " & currentSuggestion & vbCrLf
            Else
                oneByOneMsg = oneByOneMsg & "No suggestion — please type correction." & vbCrLf
            End If
            oneByOneMsg = oneByOneMsg & vbCrLf & "Type correction or leave as-is to skip:" & vbCrLf & "(Press Cancel to stop reviewing)"

            If currentSuggestion <> "" And currentSuggestion <> "(check manually)" Then
                userInput = InputBox(oneByOneMsg, "Spell Checker - One by One", currentSuggestion)
            Else
                userInput = InputBox(oneByOneMsg, "Spell Checker - One by One", currentOriginal)
            End If

            If StrPtr(userInput) = 0 Then
                MsgBox "Review stopped. " & correctionsMade & " correction(s) applied so far.", vbExclamation, "Stopped"
                GoTo SafeExit
            End If

            If userInput = "" Or userInput = currentOriginal Then
                addChoice = MsgBox("Add """ & currentOriginal & """ to custom dictionary?" & vbCrLf & vbCrLf & "It will be ignored in all future spell check runs.", vbYesNo + vbQuestion, "Add to Dictionary?")
                If addChoice = vbYes Then
                    AddToCustomDictionary currentOriginal
                    MsgBox """" & currentOriginal & """ added to dictionary.", vbInformation, "Dictionary Updated"
                End If
                GoTo NextUnique
            End If
            finalWord = userInput
        End If

        If finalWord = currentOriginal Then GoTo NextUnique

        Dim n As Long
        For n = 0 To issueCount - 1
            If LCase(issues(n).originalWord) = LCase(currentOriginal) Then
                ApplyCorrection activeDoc, issues(n), finalWord
                correctionsMade = correctionsMade + 1
            End If
        Next n

NextUnique:
    Next k

    MsgBox "Spell check complete!" & vbCrLf & correctionsMade & " occurrence(s) corrected.", vbInformation, "Done"

SafeExit:
    On Error Resume Next
    If Not oTempDoc Is Nothing Then
        oTempDoc.Close False
        Set oTempDoc = Nothing
    End If
    If Not objWord Is Nothing Then
        objWord.Quit
        Set objWord = Nothing
    End If
    Set dictCache = Nothing
    CATIA.RefreshDisplay = True
    Exit Sub

ErrorHandler:
    MsgBox "An unexpected error occurred during execution: " & vbCrLf & Err.Description, vbCritical, "Macro Error"
    Resume SafeExit
End Sub

Sub CollectIssues(ByVal inputStr As String, ByVal objWord As Object, ByRef customDict() As String, _
                  ByVal sIdx As Long, ByVal vIdx As Long, ByVal tIdx As Long, ByVal tbIdx As Long, _
                  ByVal rIdx As Long, ByVal cIdx As Long, ByVal isTable As Boolean, _
                  ByRef issues() As SpellIssue, ByRef issueCount As Long)

    Dim regEx As Object, matches As Object, match As Object
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Global = True
    regEx.IgnoreCase = True
    regEx.Pattern = "\b[A-Za-z][A-Za-z\-']*\b"

    Set matches = regEx.Execute(inputStr)

    Dim singleWord As String
    Dim wordToCheck As String
    Dim suggestion As String

    For Each match In matches
        singleWord = match.Value
        wordToCheck = LCase(singleWord)

        If IsInCustomDictionary(singleWord, customDict) Then GoTo NextWord

        If IsEngineeringWord(singleWord) Then GoTo NextWord
        
        If dictCache.Exists(wordToCheck) Then
            If dictCache(wordToCheck) = "OK" Then
                GoTo NextWord
            Else
                suggestion = dictCache(wordToCheck)
                GoTo LogIssue
            End If
        End If

        If objWord.CheckSpelling(wordToCheck) Then
            dictCache.Add wordToCheck, "OK"
            GoTo NextWord
        Else
            suggestion = GetSuggestion(wordToCheck)
            
            If suggestion = "" Or InStr(1, suggestion, Left(wordToCheck, 4), vbTextCompare) = 0 Then
                If Right(wordToCheck, 1) = "s" And Len(wordToCheck) > 3 Then
                    Dim singularWord As String
                    singularWord = Left(wordToCheck, Len(wordToCheck) - 1)
                    Dim singularSuggestion As String
                    singularSuggestion = GetSuggestion(singularWord)
                    If singularSuggestion <> "" Then suggestion = singularSuggestion & "s"
                End If
            End If

            If suggestion <> "" Then suggestion = UCase(suggestion)
            
            dictCache.Add wordToCheck, IIf(suggestion <> "", suggestion, "(check manually)")
        End If

LogIssue:
        If issueCount > UBound(issues) Then
            ReDim Preserve issues(UBound(issues) + 100)
        End If

        issues(issueCount).originalWord = singleWord
        issues(issueCount).suggestion = suggestion
        issues(issueCount).contextText = inputStr
        issues(issueCount).sheetIndex = sIdx
        issues(issueCount).viewIndex = vIdx
        issues(issueCount).textIndex = tIdx
        issues(issueCount).tableIndex = tbIdx
        issues(issueCount).rowIndex = rIdx
        issues(issueCount).colIndex = cIdx
        issues(issueCount).isTable = isTable
        issueCount = issueCount + 1

NextWord:
    Next match
End Sub


Function LoadCustomDictionary() As String()
    Dim words() As String
    Dim count As Long
    count = 0

    ReDim words(2000)

    If Dir(DICT_FILE) = "" Then
        LoadCustomDictionary = words
        Exit Function
    End If

    Dim fileNum As Integer
    fileNum = FreeFile

    On Error Resume Next
    Open DICT_FILE For Input As #fileNum

    Dim lineText As String

    Do While Not EOF(fileNum)
        Line Input #fileNum, lineText
        lineText = Trim(lineText)
        If Len(lineText) > 0 And count <= 1999 Then
            words(count) = UCase(lineText)
            count = count + 1
        End If
    Loop

    Close #fileNum

    If count = 0 Then
        ReDim words(0)
    Else
        ReDim Preserve words(count - 1)
    End If

    LoadCustomDictionary = words
End Function

Function IsInCustomDictionary(ByVal word As String, ByRef dict() As String) As Boolean
    Dim i As Long
    On Error Resume Next
    For i = 0 To UBound(dict)
        If UCase(Trim(word)) = dict(i) Then
            IsInCustomDictionary = True
            Exit Function
        End If
    Next i
    IsInCustomDictionary = False
End Function

Sub AddToCustomDictionary(ByVal word As String)
    On Error Resume Next
    If Dir(DICT_FOLDER, vbDirectory) = "" Then
        MkDir DICT_FOLDER
    End If
    Dim fileNum As Integer
    fileNum = FreeFile
    Open DICT_FILE For Append As #fileNum
    Print #fileNum, UCase(Trim(word))
    Close #fileNum
End Sub

Function IsEngineeringWord(ByVal word As String) As Boolean
    Dim w As String
    w = UCase(Trim(word))
    
    Dim ch As String
    Dim hasLetter As Boolean
    Dim hasDigit As Boolean
    Dim ci As Integer
    hasLetter = False
    hasDigit = False

    For ci = 1 To Len(w)
        ch = Mid(w, ci, 1)
        If ch >= "A" And ch <= "Z" Then hasLetter = True
        If ch >= "0" And ch <= "9" Then hasDigit = True
    Next ci

    If hasLetter And hasDigit Then
        IsEngineeringWord = True
        Exit Function
    End If

    IsEngineeringWord = False
End Function

Sub ApplyCorrection(ByVal activeDoc As Object, ByRef issue As SpellIssue, ByVal newWord As String)
    On Error Resume Next

    Dim objSheet, objView, objText, objTable
    Set objSheet = activeDoc.sheets.Item(issue.sheetIndex)
    Set objView = objSheet.views.Item(issue.viewIndex)

    Dim oldText As String
    Dim newText As String

    If issue.isTable Then
        Set objTable = objView.tables.Item(issue.tableIndex)
        oldText = objTable.GetCellString(issue.rowIndex, issue.colIndex)
        newText = Join(Split(oldText, issue.originalWord, -1, vbTextCompare), newWord)
        objTable.SetCellString issue.rowIndex, issue.colIndex, newText
    Else
        Set objText = objView.texts.Item(issue.textIndex)
        oldText = objText.Text
        
        Const catUnderline = 2
        
        Dim lines() As String
        lines = Split(oldText, vbLf)
        
        Dim uLines() As Long
        ReDim uLines(LBound(lines) To UBound(lines))
        
        Dim currentPos As Long
        currentPos = 1
        Dim i As Long, j As Long
        Dim firstNonSpace As Long
        Dim checkPos As Long
        
        For i = LBound(lines) To UBound(lines)
            If Len(lines(i)) > 0 Then

                firstNonSpace = 0
                For j = 1 To Len(lines(i))
                    If Mid(lines(i), j, 1) <> " " And Mid(lines(i), j, 1) <> vbCr Then
                        firstNonSpace = j - 1
                        Exit For
                    End If
                Next j
                
                checkPos = currentPos + firstNonSpace
                uLines(i) = objText.GetParameterOnSubString(catUnderline, checkPos, 1)
                
                currentPos = currentPos + Len(lines(i)) + 1
            Else
                uLines(i) = 0
                currentPos = currentPos + 1
            End If
        Next i
        
        newText = Join(Split(oldText, issue.originalWord, -1, vbTextCompare), newWord)
        objText.Text = newText
        
        objText.SetParameterOnSubString catUnderline, 1, Len(newText), 0
        
        Dim newLines() As String
        newLines = Split(newText, vbLf)
        currentPos = 1
        
        For i = LBound(newLines) To UBound(newLines)
            If Len(newLines(i)) > 0 Then

                If i <= UBound(uLines) Then
                    If uLines(i) = 1 Then
                        objText.SetParameterOnSubString catUnderline, currentPos, Len(newLines(i)), 1
                    End If
                End If
                currentPos = currentPos + Len(newLines(i)) + 1
            Else
                currentPos = currentPos + 1
            End If
        Next i
    End If
End Sub

Function GetSuggestion(ByVal word As String) As String
    Dim oRange As Object
    Dim oSuggestions As Object
    Dim result As String
    result = ""
    
    On Error Resume Next
    
    Set oRange = oTempDoc.Range
    oRange.Text = word

    If oTempDoc.SpellingErrors.count > 0 Then
        Set oSuggestions = oTempDoc.SpellingErrors(1).GetSpellingSuggestions
        If oSuggestions.count > 0 Then
            result = oSuggestions(1).Name
        End If
    End If

    Set oRange = Nothing
    Set oSuggestions = Nothing
    GetSuggestion = result
End Function
