Sub ExtractRequirementsWithTags()
    ' Layout in the doc:   <requirement sentence>. [R] #tag
    ' Pulls each [R] into Excel with:
    '   Col A: nearest heading above, INCLUDING its number (e.g. "1.1.2.3 Sample text")
    '   Col B: the requirement = the sentence text BEFORE [R]
    '   Col C: the #tag(s) that follow [R]   (e.g. "#sampletag")
    On Error GoTo ErrorHandler

    Const SearchTextR As String = "[R]"

    Dim srcDoc As Document
    Dim xlApp As Object, xlBook As Object, xlSheet As Object
    Dim xlRow As Long
    Dim aRng As Range, headRng As Range, reqRng As Range, tagRng As Range
    Dim headNum As String, headingText As String
    Dim reqText As String, afterText As String, tagText As String
    Dim re As Object, matches As Object, m As Object

    Set srcDoc = ActiveDocument

    ' Regex to grab tags like #sampletag, #SPM, #TechLead
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.Pattern = "#\w+"

    ' --- Excel setup ---
    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = xlApp.Workbooks.Add
    Set xlSheet = xlBook.Worksheets(1)
    xlApp.Visible = True

    xlSheet.Cells(1, 1).Value = "Heading"
    xlSheet.Cells(1, 2).Value = "Requirement"
    xlSheet.Cells(1, 3).Value = "Tag(s)"
    xlRow = 2

    ' --- Find every [R] in the document ---
    Set aRng = srcDoc.Range
    With aRng.Find
        .ClearFormatting
        .Text = SearchTextR
        .Forward = True
        .Wrap = wdFindStop
        Do While .Execute
            ' --- Requirement = sentence BEFORE [R] (paragraph start up to [R]) ---
            Set reqRng = aRng.Duplicate
            reqRng.Start = aRng.Paragraphs(1).Range.Start
            reqRng.End = aRng.Start
            reqText = Trim(reqRng.Text)

            ' --- Tag(s) = text AFTER [R], to end of its paragraph ---
            Set tagRng = aRng.Duplicate
            tagRng.Start = aRng.End
            tagRng.End = aRng.Paragraphs(1).Range.End
            afterText = tagRng.Text
            tagText = ""
            If re.Test(afterText) Then
                Set matches = re.Execute(afterText)
                For Each m In matches
                    If tagText = "" Then
                        tagText = m.Value
                    Else
                        tagText = tagText & ", " & m.Value
                    End If
                Next m
            End If

            ' --- Heading above, with its multilevel number ---
            Set headRng = aRng.GoTo(What:=wdGoToHeading, Which:=wdGoToPrevious)
            If Not headRng Is Nothing Then
                headRng.End = headRng.Paragraphs(1).Range.End - 1
                headNum = headRng.ListFormat.ListString
                headingText = Trim(headNum & " " & Trim(headRng.Text))
            Else
                headingText = "No Heading"
            End If

            ' --- Write the row ---
            xlSheet.Cells(xlRow, 1).Value = headingText
            xlSheet.Cells(xlRow, 2).Value = reqText
            xlSheet.Cells(xlRow, 3).Value = tagText
            xlRow = xlRow + 1

            aRng.Collapse Direction:=wdCollapseEnd
        Loop
    End With

    xlSheet.Columns("A:C").AutoFit
    MsgBox "Done. " & (xlRow - 2) & " requirement(s) found.", vbInformation
    Exit Sub

ErrorHandler:
    MsgBox "Error: " & Err.Description, vbExclamation
    If Not xlApp Is Nothing Then xlApp.Visible = True
End Sub
