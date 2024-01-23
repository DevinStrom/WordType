Dim originalText As String
Dim cc As ContentControl
Dim startTime As Double

Sub InitializeTypingExercise()
    ' Clear all text on the page
    ActiveDocument.Range.Delete

    ' Set the original paragraph
    originalText = GetRandomParagraph(ThisDocument.Path & "\Paragraphs.txt")

    ' Insert the original paragraph with size 18 font
    Selection.Font.Size = 18
    Selection.TypeText originalText

    ' Insert a line break
    Selection.TypeParagraph

    ' Insert a Content Control for the user to type in
    Set cc = ActiveDocument.ContentControls.Add(wdContentControlRichText, Selection.Range)
    ' Lock the Content Control for editing
    cc.LockContentControl = True
    ' set font size for content
    cc.Range.Font.Size = 18

    ' Set the cursor inside the Content Control and collapse it
    cc.Range.Select
    cc.Range.Collapse Direction:=wdCollapseEnd
    
    ' Set up the event handler for SelectionChange
    ThisDocument.Application.OnTime Now, "StartMonitoring"
    
    ' Record the start time
    startTime = Timer
End Sub

Function GetRandomParagraph(externalDocPath As String) As String
    Dim fileNumber As Integer
    fileNumber = FreeFile

    ' Open the file for reading
    Open externalDocPath For Input As fileNumber

    ' Initialize a Collection to store lines
    Dim lines As New Collection
    Dim tempStr As String

    ' Read all lines into the Collection
    Do Until EOF(fileNumber)
        Line Input #fileNumber, tempStr
        If Trim(tempStr) <> "" Then
            lines.Add tempStr
        End If
    Loop

    ' Close the file
    Close fileNumber

    ' Print information for debugging
    Debug.Print "Number of paragraphs: " & lines.Count

    ' Select a random line
    Dim randomIndex As Integer
    Randomize
    If lines.Count > 0 Then
        randomIndex = Int(lines.Count * Rnd) + 1
        Debug.Print "Selected index: " & randomIndex
        GetRandomParagraph = lines(randomIndex)
    Else
        GetRandomParagraph = ""
    End If
End Function

Sub StartMonitoring()
    ' Set up the event handler for SelectionChange
    ThisDocument.Application.OnTime Now + 0.3 / 86400, "CheckUserInput"
End Sub


Sub CheckUserInput()
    ' Check the user's input against the original text
    Dim userInput As String
    userInput = cc.Range.Text

    ' Set font size and clear previous highlighting
    cc.Range.Font.Size = 18
    cc.Range.Font.Color = RGB(0, 0, 0) ' Black

    ' Iterate through the characters and highlight differences
    Dim i As Long
    For i = 1 To Len(userInput)
        If Mid(userInput, i, 1) <> Mid(originalText, i, 1) Then
            ' Highlight the non-matching character in red
            cc.Range.Characters(i).Font.Color = RGB(255, 0, 0) ' Red
        End If
    Next i

    ' Check if the user has typed the entire paragraph
    If userInput = originalText Then
        ' Calculate the elapsed time
        Dim elapsedTime As Double
        elapsedTime = Timer - startTime

        ' Calculate words per minute (WPM)
        Dim wordsTyped As Integer
        wordsTyped = UBound(Split(originalText, " ")) + 1 ' Counting words based on spaces
        Dim wpm As Double
        wpm = wordsTyped / (elapsedTime / 60) ' Divide by time taken in minutes

        ' Display the results in the document
        cc.Range.Text = "Time taken: " & Format(elapsedTime, "0.0") & " seconds" & vbCrLf & _
                       "Words per Minute: " & Format(wpm, "0.0") & vbCrLf & vbCrLf & _
                       "originalText:" & vbCrLf & originalText & vbCrLf & vbCrLf & _
                       "userInput:" & vbCrLf & userInput

        
        ' Pause for 3 seconds before starting a new typing exercise
        Dim endTime As Double
        endTime = Timer + 3 ' Set the end time for the loop
        Do While Timer < endTime
            DoEvents ' Allow other processes to execute
        Loop
        

        ' Reset the timer and start a new typing exercise
        startTime = Timer
        cc.Range.Text = "" ' Clear the user's input
        ThisDocument.Application.OnTime Now + 0.3 / 86400, "StartMonitoring"
        RemoveAllContentControls
        InitializeTypingExercise
        
    Else
        ' Set up the event handler for SelectionChange
        ThisDocument.Application.OnTime Now + 0.3 / 86400, "CheckUserInput"
    End If
End Sub

Sub RemoveAllContentControls()
    Dim cc As ContentControl
    For Each cc In ActiveDocument.ContentControls
        If cc.LockContentControl Then
            ' If locked, unlock it first
            cc.LockContentControl = False
        End If
        cc.Delete
    Next cc
End Sub

Sub DisableProofing()
    ActiveDocument.SpellingChecked = False
    ActiveDocument.GrammarChecked = False
    ActiveDocument.ShowSpellingErrors = False
    ActiveDocument.ShowGrammaticalErrors = False
    Options.AutoFormatAsYouTypeReplaceHyperlinks = False
    Options.AutoFormatAsYouTypeApplyBorders = False
    Options.AutoFormatAsYouTypeApplyBulletedLists = False
    Options.AutoFormatAsYouTypeApplyNumberedLists = False
    Options.AutoFormatAsYouTypeApplyTables = False
End Sub

Sub EnableProofing()
    ActiveDocument.SpellingChecked = True
    ActiveDocument.GrammarChecked = True
    ActiveDocument.ShowSpellingErrors = True
    ActiveDocument.ShowGrammaticalErrors = True
    Options.AutoFormatAsYouTypeReplaceHyperlinks = True
    Options.AutoFormatAsYouTypeApplyBorders = True
    Options.AutoFormatAsYouTypeApplyBulletedLists = True
    Options.AutoFormatAsYouTypeApplyNumberedLists = True
    Options.AutoFormatAsYouTypeApplyTables = True
End Sub
