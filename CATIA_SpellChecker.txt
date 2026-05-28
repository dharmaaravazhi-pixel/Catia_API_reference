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
    isDimension     As Boolean
    dimPart         As Integer
    locationString  As String
End Type

Public issues() As SpellIssue
Public issueCount As Long
Public uniqueWords() As String
Public uniqueCount As Long

Public CATIA As Object

' Set to True for development (requires references to MS Word, Scripting Runtime, and VBScript Regular Expressions 5.5)
' Set to False for distribution
#Const EARLY_BINDING = False

#If EARLY_BINDING Then
    Public objWord As word.Application
    Public oTempDoc As word.Document
    Public dictCache As Scripting.Dictionary
    Public globalRegExExt As RegExp
    Public globalRegExRep As RegExp
#Else
    Public objWord As Object
    Public oTempDoc As Object
    Public dictCache As Object
    Public globalRegExExt As Object
    Public globalRegExRep As Object
#End If

Public DICT_FOLDER As String
Public DICT_FILE As String
Public AUTOCORRECT_FILE As String

Sub CATMain()

    DICT_FOLDER = Environ("USERPROFILE") & "\CATIA_SpellChecker\"
    DICT_FILE = DICT_FOLDER & "CustomDictionary.txt"
    AUTOCORRECT_FILE = DICT_FOLDER & "AutoCorrect.txt"

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
#If EARLY_BINDING Then
    Set objWord = New word.Application
    Set dictCache = New Scripting.Dictionary
    Set globalRegExExt = New RegExp
    Set globalRegExRep = New RegExp
#Else
    Set objWord = CreateObject("Word.Application")
    Set dictCache = CreateObject("Scripting.Dictionary")
    Set globalRegExExt = CreateObject("VBScript.RegExp")
    Set globalRegExRep = CreateObject("VBScript.RegExp")
#End If
    
    If Err.Number <> 0 Or objWord Is Nothing Then
        MsgBox "Failed to connect to Microsoft Word.", vbCritical, "Dependency Error"
        Exit Sub
    End If
    
    objWord.Visible = False
    
    Set oTempDoc = objWord.Documents.Add
    
    dictCache.CompareMode = 1
    
    globalRegExExt.Global = True
    globalRegExExt.IgnoreCase = True
    globalRegExExt.Pattern = "\b[A-Za-zÀ-ÿØ][A-Za-zÀ-ÿØ0-9\-\']*\b"

    globalRegExRep.Global = False
    globalRegExRep.IgnoreCase = True
    
#If EARLY_BINDING Then
    Dim autoCorrectDict As Scripting.Dictionary
    Set autoCorrectDict = New Scripting.Dictionary
#Else
    Dim autoCorrectDict As Object
    Set autoCorrectDict = CreateObject("Scripting.Dictionary")
#End If
    autoCorrectDict.CompareMode = 1 ' Case-insensitive compare
    LoadAutoCorrectDictionary AUTOCORRECT_FILE, autoCorrectDict
    
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
    Dim tb As Long, R As Long, C As Long
    Dim currentString As String

    Dim sel As Object
    Set sel = activeDoc.Selection
    
    CATIA.Interactive = False
    CATIA.RefreshDisplay = False
    
    Dim visProp As Object
    Dim showState As Long ' Changed from Integer to Long to prevent Type Mismatch in GetShow
    Dim startCount As Long

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
                    Dim locStr As String
                    locStr = objSheet.Name & ", " & objView.Name & ", " & objText.Name
                    
                    startCount = issueCount
                    CollectIssues currentString, objWord, customDict, autoCorrectDict, s, i, j, 0, 0, 0, False, False, 0, locStr, issues, issueCount
                    
                    If issueCount > startCount Then
                        sel.Clear
                        sel.Add objText
                        Set visProp = sel.VisProperties
                        visProp.GetShow showState
                        If showState = 1 Then
                            issueCount = startCount
                        End If
                        Set visProp = Nothing
                    End If
                End If
                Set objText = Nothing
            Next

            Set objTables = objView.tables
            If Not objTables Is Nothing Then
                For tb = 1 To objTables.count
                    Set objTable = objTables.Item(tb)
                    
                    startCount = issueCount
                    
                    For R = 1 To objTable.NumberOfRows
                        For C = 1 To objTable.NumberOfColumns
                            currentString = objTable.GetCellString(R, C)
                            If Len(Trim(currentString)) > 0 Then
                                locStr = objSheet.Name & ", " & objView.Name & ", " & objTable.Name & " R" & CStr(R) & "C" & CStr(C)
                                CollectIssues currentString, objWord, customDict, autoCorrectDict, s, i, 0, tb, R, C, True, False, 0, locStr, issues, issueCount
                            End If
                        Next
                    Next
                    
                    If issueCount > startCount Then
                        sel.Clear
                        sel.Add objTable
                        Set visProp = sel.VisProperties
                        visProp.GetShow showState
                        If showState = 1 Then
                            issueCount = startCount
                        End If
                        Set visProp = Nothing
                    End If
                    
                    Set objTable = Nothing
                Next
            End If
            
            Dim objDims As Object
            Dim objDim As Object
            Dim dimValue As Object
            Dim strBefore As String, strAfter As String, strUpper As String, strLower As String
            Dim strPrefix As String, strSuffix As String
            Dim dimStr As String

            On Error Resume Next
            Set objDims = objView.Dimensions
            If Not objDims Is Nothing Then
                For j = 1 To objDims.count
                    Set objDim = objDims.Item(j)
                    
                    startCount = issueCount
                    
                    Set dimValue = objDim.GetValue
                    Err.Clear
                    
                    Dim fakeType As Integer
                    fakeType = dimValue.FakeDimType
                    If fakeType = 2 Then
                        dimStr = dimValue.GetFakeDimValue()
                        If Len(Trim(dimStr)) > 0 And Err.Number = 0 Then
                            locStr = objSheet.Name & ", " & objView.Name & ", " & objDim.Name & " (Fake Text)"
                            CollectIssues dimStr, objWord, customDict, autoCorrectDict, s, i, j, 0, 0, 0, False, True, 5, locStr, issues, issueCount
                        End If
                    End If
                    Err.Clear
                    
                    strBefore = "": strAfter = "": strUpper = "": strLower = ""
                    dimValue.GetBaultText 1, strBefore, strAfter, strUpper, strLower
                    If Err.Number = 0 Then
                        If Len(Trim(strBefore)) > 0 Then
                            locStr = objSheet.Name & ", " & objView.Name & ", " & objDim.Name & " (Before)"
                            CollectIssues strBefore, objWord, customDict, autoCorrectDict, s, i, j, 0, 0, 0, False, True, 1, locStr, issues, issueCount
                        End If
                        If Len(Trim(strAfter)) > 0 Then
                            locStr = objSheet.Name & ", " & objView.Name & ", " & objDim.Name & " (After)"
                            CollectIssues strAfter, objWord, customDict, autoCorrectDict, s, i, j, 0, 0, 0, False, True, 2, locStr, issues, issueCount
                        End If
                        If Len(Trim(strUpper)) > 0 Then
                            locStr = objSheet.Name & ", " & objView.Name & ", " & objDim.Name & " (Upper)"
                            CollectIssues strUpper, objWord, customDict, autoCorrectDict, s, i, j, 0, 0, 0, False, True, 3, locStr, issues, issueCount
                        End If
                        If Len(Trim(strLower)) > 0 Then
                            locStr = objSheet.Name & ", " & objView.Name & ", " & objDim.Name & " (Lower)"
                            CollectIssues strLower, objWord, customDict, autoCorrectDict, s, i, j, 0, 0, 0, False, True, 4, locStr, issues, issueCount
                        End If
                    End If
                    Err.Clear
                    
                    strPrefix = "": strSuffix = ""
                    dimValue.GetPSText 1, strPrefix, strSuffix
                    If Err.Number = 0 Then
                        If Len(Trim(strPrefix)) > 0 Then
                            locStr = objSheet.Name & ", " & objView.Name & ", " & objDim.Name & " (Prefix)"
                            CollectIssues strPrefix, objWord, customDict, autoCorrectDict, s, i, j, 0, 0, 0, False, True, 6, locStr, issues, issueCount
                        End If
                        If Len(Trim(strSuffix)) > 0 Then
                            locStr = objSheet.Name & ", " & objView.Name & ", " & objDim.Name & " (Suffix)"
                            CollectIssues strSuffix, objWord, customDict, autoCorrectDict, s, i, j, 0, 0, 0, False, True, 7, locStr, issues, issueCount
                        End If
                    End If
                    Err.Clear
                    
                    If issueCount > startCount Then
                        sel.Clear
                        sel.Add objDim
                        Set visProp = sel.VisProperties
                        visProp.GetShow showState
                        If showState = 1 Then
                            issueCount = startCount
                        End If
                        Set visProp = Nothing
                    End If
                    
                    Set dimValue = Nothing
                    Set objDim = Nothing
                Next j
            End If
            Set objDims = Nothing
            On Error GoTo ErrorHandler
            Set objTexts = Nothing
            Set objTables = Nothing
            Set objView = Nothing
        Next
        Set objViews = Nothing
        Set objSheet = Nothing
    Next
    
    sel.Clear
    CATIA.Interactive = True
    CATIA.RefreshDisplay = True

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

    ' Only show first 20 in summary to avoid MsgBox overflow
    Dim limit As Long
    limit = uniqueCount
    If limit > 20 Then limit = 20
    
    For k = 0 To limit - 1
        Dim dispSuggestion As String
        Dim locs As String
        dispSuggestion = ""
        locs = ""
        Dim m As Long
        For m = 0 To issueCount - 1
            If LCase(issues(m).originalWord) = LCase(uniqueWords(k)) Then
                If dispSuggestion = "" Then dispSuggestion = issues(m).suggestion
                If locs = "" Then
                    locs = issues(m).locationString
                Else
                    locs = locs & " | " & issues(m).locationString
                End If
            End If
        Next m
        
        If Len(locs) > 80 Then
            locs = Left(locs, 75) & "..."
        End If

        summaryMsg = summaryMsg & (k + 1) & ".  " & uniqueWords(k)
        If dispSuggestion <> "" Then
            summaryMsg = summaryMsg & "   ->   " & dispSuggestion
        Else
            summaryMsg = summaryMsg & "   ->   (no suggestion)"
        End If
        summaryMsg = summaryMsg & "  [" & locs & "]" & vbCrLf
    Next k
    
    If uniqueCount > 20 Then
        summaryMsg = summaryMsg & "... and " & (uniqueCount - 20) & " more." & vbCrLf
    End If

    summaryMsg = summaryMsg & String(40, "-") & vbCrLf & vbCrLf & _
                 "Choose an action:" & vbCrLf & vbCrLf & _
                 "[Yes]     = Accept All suggestions automatically" & vbCrLf & _
                 "[No]      = Review One by One (Show exact locations)" & vbCrLf & _
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
    
    If userChoice = vbYes Then
        ' Apply all unique words
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
            If currentSuggestion <> "" And currentSuggestion <> "(check manually)" Then
                Dim n As Long
                For n = 0 To issueCount - 1
                    If LCase(issues(n).originalWord) = LCase(currentOriginal) Then
                        ApplyCorrection activeDoc, issues(n), currentSuggestion
                        correctionsMade = correctionsMade + 1
                    End If
                Next n
            End If
        Next k
    ElseIf userChoice = vbNo Then
        ' Review individually occurrence by occurrence
        For k = 0 To issueCount - 1
            Dim oneByOneMsg As String
            oneByOneMsg = "Occurrence " & (k + 1) & " of " & issueCount & vbCrLf & String(40, "-") & vbCrLf & _
                          "Word:        " & issues(k).originalWord & vbCrLf
            If issues(k).suggestion <> "" And issues(k).suggestion <> "(check manually)" Then
                oneByOneMsg = oneByOneMsg & "Suggested:   " & issues(k).suggestion & vbCrLf
            Else
                oneByOneMsg = oneByOneMsg & "Suggested:   (no suggestion)" & vbCrLf
            End If
            
            oneByOneMsg = oneByOneMsg & "Location:    " & issues(k).locationString & vbCrLf & vbCrLf & _
                                        "Type correction or leave as-is to skip:" & vbCrLf & _
                                        "(Press Cancel to stop reviewing)"
            
            If issues(k).suggestion <> "" And issues(k).suggestion <> "(check manually)" Then
                userInput = InputBox(oneByOneMsg, "Review One by One", issues(k).suggestion)
            Else
                userInput = InputBox(oneByOneMsg, "Review One by One", issues(k).originalWord)
            End If

            If StrPtr(userInput) = 0 Then
                MsgBox "Review stopped. " & correctionsMade & " correction(s) applied so far.", vbExclamation, "Stopped"
                GoTo SafeExit
            End If

            If userInput = "" Or userInput = issues(k).originalWord Then
                addChoice = MsgBox("Add """ & issues(k).originalWord & """ to custom dictionary?" & vbCrLf & vbCrLf & "It will be ignored in all future spell check runs.", vbYesNo + vbQuestion, "Add to Dictionary?")
                If addChoice = vbYes Then
                    AddToCustomDictionary issues(k).originalWord
                    MsgBox """" & issues(k).originalWord & """ added to dictionary.", vbInformation, "Dictionary Updated"
                End If
            Else
                ApplyCorrection activeDoc, issues(k), userInput
                correctionsMade = correctionsMade + 1
            End If
        Next k
    End If

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
    If Not CATIA Is Nothing Then
        CATIA.Interactive = True
        CATIA.RefreshDisplay = True
    End If
    Exit Sub

ErrorHandler:
    MsgBox "An unexpected error occurred during execution: " & vbCrLf & Err.Description, vbCritical, "Macro Error"
    Resume SafeExit
End Sub

Sub CollectIssues(ByVal inputStr As String, ByVal objWord As Object, ByRef customDict() As String, _
                  ByVal autoCorrectDict As Object, _
                  ByVal sIdx As Long, ByVal vIdx As Long, ByVal tIdx As Long, ByVal tbIdx As Long, _
                  ByVal rIdx As Long, ByVal cIdx As Long, ByVal isTable As Boolean, ByVal isDim As Boolean, ByVal dimPrt As Integer, ByVal locStr As String, _
                  ByRef issues() As SpellIssue, ByRef issueCount As Long)

#If EARLY_BINDING Then
    Dim matches As MatchCollection, match As match
#Else
    Dim matches As Object, match As Object
#End If

    Set matches = globalRegExExt.Execute(inputStr)

    Dim singleWord As String
    Dim wordToCheck As String
    Dim suggestion As String
    Dim logThisIssue As Boolean

    For Each match In matches
        singleWord = match.Value
        wordToCheck = LCase(singleWord)
        logThisIssue = False
        suggestion = ""

        If Not IsInCustomDictionary(singleWord, customDict) And Not IsEngineeringWord(singleWord) Then
            If autoCorrectDict.Exists(wordToCheck) Then
                suggestion = autoCorrectDict(wordToCheck)
                logThisIssue = True
            ElseIf dictCache.Exists(wordToCheck) Then
                If dictCache(wordToCheck) <> "OK" Then
                    suggestion = dictCache(wordToCheck)
                    logThisIssue = True
                End If
            Else
                If objWord.CheckSpelling(wordToCheck) Then
                    dictCache.Add wordToCheck, "OK"
                Else
                    suggestion = GetSuggestion(wordToCheck)
                    
                    If suggestion = "" Or InStr(1, suggestion, Left(wordToCheck, 4), vbTextCompare) = 0 Then
                        Dim singularWord As String, singularSuggestion As String
                        If Right(wordToCheck, 3) = "ies" And Len(wordToCheck) > 4 Then
                            singularWord = Left(wordToCheck, Len(wordToCheck) - 3) & "y"
                            singularSuggestion = GetSuggestion(singularWord)
                            If singularSuggestion <> "" Then
                                If Right(singularSuggestion, 1) = "y" Then
                                    suggestion = Left(singularSuggestion, Len(singularSuggestion) - 1) & "ies"
                                Else
                                    suggestion = singularSuggestion & "s"
                                End If
                            End If
                        ElseIf Right(wordToCheck, 2) = "es" And Len(wordToCheck) > 4 Then
                            singularWord = Left(wordToCheck, Len(wordToCheck) - 2)
                            singularSuggestion = GetSuggestion(singularWord)
                            If singularSuggestion <> "" Then suggestion = singularSuggestion & "es"
                        ElseIf Right(wordToCheck, 1) = "s" And Len(wordToCheck) > 3 Then
                            singularWord = Left(wordToCheck, Len(wordToCheck) - 1)
                            singularSuggestion = GetSuggestion(singularWord)
                            If singularSuggestion <> "" Then suggestion = singularSuggestion & "s"
                        End If
                    End If
        
                    If suggestion <> "" Then suggestion = UCase(suggestion)
                    dictCache.Add wordToCheck, IIf(suggestion <> "", suggestion, "(check manually)")
                    logThisIssue = True
                End If
            End If
        End If

        If logThisIssue Then
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
            issues(issueCount).isDimension = isDim
            issues(issueCount).dimPart = dimPrt
            issues(issueCount).locationString = locStr
            issueCount = issueCount + 1
        End If
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

Sub LoadAutoCorrectDictionary(ByVal filePath As String, ByRef dict As Object)
    On Error Resume Next
    If Dir(filePath) = "" Then
        If Dir(DICT_FOLDER, vbDirectory) = "" Then MkDir DICT_FOLDER
        Dim tempNum As Integer
        tempNum = FreeFile
        Open filePath For Output As #tempNum
        Print #tempNum, "prt=PART"
        Print #tempNum, "assy=ASSEMBLY"
        Print #tempNum, "dwg=DRAWING"
        Print #tempNum, "matl=MATERIAL"
        Close #tempNum
    End If

    Dim fileNum As Integer
    fileNum = FreeFile
    Open filePath For Input As #fileNum
    Dim lineText As String, parts() As String
    Do While Not EOF(fileNum)
        Line Input #fileNum, lineText
        lineText = Trim(lineText)
        If Len(lineText) > 0 And InStr(lineText, "=") > 0 Then
            parts = Split(lineText, "=", 2)
            dict(Trim(parts(0))) = Trim(parts(1))
        End If
    Loop
    Close #fileNum
End Sub

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
    
    ' Check for single letter hyphenation (e.g., A-A)
    If Len(w) = 3 And Mid(w, 2, 1) = "-" Then
        If Left(w, 1) >= "A" And Left(w, 1) <= "Z" And Right(w, 1) >= "A" And Right(w, 1) <= "Z" Then
            IsEngineeringWord = True
            Exit Function
        End If
    End If
    
    Dim ch As String
    Dim hasLetter As Boolean
    Dim hasDigit As Boolean
    Dim hasSpecialEngChar As Boolean
    Dim ci As Integer
    hasLetter = False
    hasDigit = False
    hasSpecialEngChar = False

    For ci = 1 To Len(w)
        ch = Mid(w, ci, 1)
        If ch >= "A" And ch <= "Z" Then hasLetter = True
        If ch >= "0" And ch <= "9" Then hasDigit = True
        If ch = "Ø" Or ch = "°" Then hasSpecialEngChar = True
    Next ci

    If (hasLetter And hasDigit) Or hasSpecialEngChar Then
        IsEngineeringWord = True
        Exit Function
    End If

    IsEngineeringWord = False
End Function

Function EscapeRegex(ByVal str As String) As String
    Dim res As String
    res = str
    res = Replace(res, "\", "\\")
    res = Replace(res, "^", "\^")
    res = Replace(res, "$", "\$")
    res = Replace(res, "*", "\*")
    res = Replace(res, "+", "\+")
    res = Replace(res, "?", "\?")
    res = Replace(res, ".", "\.")
    res = Replace(res, "(", "\(")
    res = Replace(res, ")", "\)")
    res = Replace(res, "|", "\|")
    res = Replace(res, "{", "\{")
    res = Replace(res, "}", "\}")
    res = Replace(res, "[", "\[")
    res = Replace(res, "]", "\]")
    EscapeRegex = res
End Function

Sub ApplyCorrection(ByVal activeDoc As Object, ByRef issue As SpellIssue, ByVal newWord As String)
    On Error Resume Next

    Dim objSheet, objView, objText, objTable, objDim, dimValue
    Set objSheet = activeDoc.sheets.Item(issue.sheetIndex)
    Set objView = objSheet.views.Item(issue.viewIndex)

    Dim oldText As String
    Dim newText As String
    globalRegExRep.Pattern = "\b" & EscapeRegex(issue.originalWord) & "\b"

    If issue.isDimension Then
        Set objDim = objView.Dimensions.Item(issue.textIndex)
        Set dimValue = objDim.GetValue
        If issue.dimPart = 5 Then
            oldText = dimValue.GetFakeDimValue()
            newText = globalRegExRep.Replace(oldText, newWord)
            dimValue.SetFakeDimValue newText
        ElseIf issue.dimPart >= 1 And issue.dimPart <= 4 Then
            Dim sBef As String, sAft As String, sUp As String, sLow As String
            dimValue.GetBaultText 1, sBef, sAft, sUp, sLow
            If issue.dimPart = 1 Then
                newText = globalRegExRep.Replace(sBef, newWord)
                sBef = newText
            ElseIf issue.dimPart = 2 Then
                newText = globalRegExRep.Replace(sAft, newWord)
                sAft = newText
            ElseIf issue.dimPart = 3 Then
                newText = globalRegExRep.Replace(sUp, newWord)
                sUp = newText
            ElseIf issue.dimPart = 4 Then
                newText = globalRegExRep.Replace(sLow, newWord)
                sLow = newText
            End If
            dimValue.SetBaultText 1, sBef, sAft, sUp, sLow
        ElseIf issue.dimPart = 6 Or issue.dimPart = 7 Then
            Dim sPre As String, sSuf As String
            dimValue.GetPSText 1, sPre, sSuf
            If issue.dimPart = 6 Then
                newText = globalRegExRep.Replace(sPre, newWord)
                sPre = newText
            ElseIf issue.dimPart = 7 Then
                newText = globalRegExRep.Replace(sSuf, newWord)
                sSuf = newText
            End If
            dimValue.SetPSText 1, sPre, sSuf
        End If
        Set dimValue = Nothing
        Set objDim = Nothing
    ElseIf issue.isTable Then
        Set objTable = objView.tables.Item(issue.tableIndex)
        oldText = objTable.GetCellString(issue.rowIndex, issue.colIndex)
        newText = globalRegExRep.Replace(oldText, newWord)
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
        
        newText = globalRegExRep.Replace(oldText, newWord)
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
    
    Set objText = Nothing
    Set objTable = Nothing
    Set objView = Nothing
    Set objSheet = Nothing
End Sub

Function GetSuggestion(ByVal word As String) As String
    Dim oRange As Object
    Dim oSuggestions As Object
    Dim Result As String
    Result = ""
    
    On Error Resume Next
    
    Set oRange = oTempDoc.Range
    oRange.Text = word

    If oTempDoc.SpellingErrors.count > 0 Then
        Set oSuggestions = oTempDoc.SpellingErrors(1).GetSpellingSuggestions
        If oSuggestions.count > 0 Then
            Result = oSuggestions(1).Name
        End If
    End If

    Set oRange = Nothing
    Set oSuggestions = Nothing
    GetSuggestion = Result
End Function
