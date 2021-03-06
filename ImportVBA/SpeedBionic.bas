Attribute VB_Name = "SpeedBionic"
'Follow me on IG @GPFSye

Sub SpeedReaderSelection()

'Allows you to highlight specific parts of the document and run the Macro on that
'Use this for performance, faster than converting the entire document

Application.ScreenUpdating = False

Dim c As Long
Dim objRegex As RegExp
Dim matches1, matches2, matches3 As MatchCollection
Dim fnd As Match
Dim Coll As New Collection

    Set objRegex = New RegExp
    Set myRange = Selection.Range
    Set myRangeSel = myRange.Duplicate
    
    ' Great excuse to play with regex
    With objRegex
        .Pattern = "\b[A-Za-z]{1}(?=[A-Za-z]\b)|\b[A-Za-z](?=[A-Za-z]{0,2}?\b)"
        .Global = True
        .IgnoreCase = True
        Set matches1 = .Execute(myRange)
    End With
    
    For Each fnd In matches1
        On Error Resume Next
            Coll.Add fnd, fnd
        On Error GoTo 0
    Next fnd
        
    With objRegex
        .Pattern = "\b(?:[A-Za-z]{4}(?=[A-Za-z]{2}[A-Za-z]?\b)|[A-Za-z]{5}(?=[A-Za-z]{2}[A-Za-z]?\b)|[A-Za-z]{7}(?=[A-Za-z]{3}[A-Za-z]?\b)|[A-Za-z]{5}(?=[A-Za-z]{3}[A-Za-z]?\b)|[A-Za-z]{7}(?=[A-Za-z]{4}[A-Za-z]?\b)|[A-Za-z]{8}(?=[A-Za-z]{5}[A-Za-z]?\b)|[A-Za-z]{9}(?=[A-Za-z]{6}[A-Za-z]?\b)|[A-Za-z]{9}(?=[A-Za-z]{8}[A-Za-z]?\b)|[A-Za-z]{10}(?=[A-Za-z]{9}[A-Za-z]?\b))"
        .Global = True
        .IgnoreCase = True
        Set matches2 = .Execute(myRange)
    End With
    
    For Each fnd In matches2
        On Error Resume Next
            Coll.Add fnd, fnd
        On Error GoTo 0
    Next fnd
    
    With objRegex
        .Pattern = "\b(?:[A-Za-z]{3}(?=[A-Za-z]{2}[A-Za-z]?\b)|[A-Za-z]{2}(?=[A-Za-z]{1}[A-Za-z][A-Za-z]?\b))"
        .Global = True
        .IgnoreCase = True
        Set matches3 = .Execute(myRange)
    End With
    
    For Each fnd In matches3
        On Error Resume Next
            Coll.Add fnd, fnd
        On Error GoTo 0
    Next fnd
 
    With myRange.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Replacement.Font.Bold = True
        .Format = True
        .Forward = True
        .Wrap = wdFindStop
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .MatchPrefix = True
        For Each fnd In Coll
            On Error Resume Next
                Debug.Print fnd & " " & c & "/" & Coll.Count
                c = c + 1
                .Text = fnd
                .Replacement.Text = "^&"
                myRange.Find.Execute Replace:=wdReplaceAll
            On Error GoTo 0
        Next fnd
    End With

myRange.ParagraphFormat.LineSpacing = LinesToPoints(2)

Application.ScreenUpdating = True

End Sub

'Follow me on IG @GPFSye

Sub SpeedReaderSelectionRoundDown()

'Allows you to highlight specific parts of the document and run the Macro on that
'Use this for performance, faster than converting the entire document

Application.ScreenUpdating = False

Dim c As Long
Dim objRegex As RegExp
Dim matches As MatchCollection
Dim fnd As Match
Dim Coll As New Collection

    Set objRegex = New RegExp
    Set myRange = Selection.Range
    Set myRangeSel = myRange.Duplicate
    
    ' Great excuse to play with regex
    With objRegex
        .Pattern = "\b(?:[A-Za-z]{1}(?=[A-Za-z]{0}[A-Za-z]?[A-Za-z]\b)|[A-Za-z]{2}(?=[A-Za-z]{1}[A-Za-z]?[A-Za-z]\b)|[A-Za-z]{3}(?=[A-Za-z]{2}[A-Za-z]?[A-Za-z]\b)|[A-Za-z]{4}(?=[A-Za-z]{3}[A-Za-z]?[A-Za-z]\b)|[A-Za-z]{5}(?=[A-Za-z]{4}[A-Za-z]?[A-Za-z]\b)|[A-Za-z]{6}(?=[A-Za-z]{5}[A-Za-z]?[A-Za-z]\b)|[A-Za-z]{7}(?=[A-Za-z]{6}[A-Za-z]?[A-Za-z]\b)|[A-Za-z]{8}(?=[A-Za-z]{7}[A-Za-z]?[A-Za-z]\b)|[A-Za-z]{9}(?=[A-Za-z]{8}[A-Za-z]?[A-Za-z]\b)|[A-Za-z]{10}(?=[A-Za-z]{9}[A-Za-z]?[A-Za-z]\b))"
        .Global = True
        .IgnoreCase = True
        Set matches = .Execute(myRange)
    End With
    
    For Each fnd In matches
        On Error Resume Next
            Coll.Add fnd, fnd
        On Error GoTo 0
    Next fnd

    With myRange.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Replacement.Font.Bold = True
        .Format = True
        .Forward = True
        .Wrap = wdFindStop
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .MatchPrefix = True
        For Each fnd In Coll
            On Error Resume Next
                Debug.Print fnd & " " & c & "/" & Coll.Count
                c = c + 1
                .Text = fnd
                .Replacement.Text = "^&"
                myRange.Find.Execute Replace:=wdReplaceAll
            On Error GoTo 0
        Next fnd
    End With

myRange.ParagraphFormat.LineSpacing = LinesToPoints(2)

Application.ScreenUpdating = True

End Sub

Option Explicit

Private Sub SpeedReaderFullDoc()

'Follow me on IG @GPFSye

' Ctrl + G opens the console log so you can track progress
' On larger documents, this will freeze, but don't worry, the code is solid, VBA is just slow, it'll take awhile
' but it'll get there lol

Application.ScreenUpdating = False

Dim i, c As Long
Dim objRegex As RegExp
Dim matches1, matches2, matches3 As MatchCollection
Dim fnd As Match
Dim Coll As New Collection

    Set objRegex = New RegExp
    Set myRange = ActiveDocument.Content
    
    ' Great excuse to play with regex
    With objRegex
        .Pattern = "\b[A-Za-z]{1}(?=[A-Za-z]\b)|\b[A-Za-z](?=[A-Za-z]{0,2}?\b)"
        .Global = True
        .IgnoreCase = True
        Set matches1 = .Execute(myRange)
    End With
    
    For Each fnd In matches1
        On Error Resume Next
            Coll.Add fnd, fnd
        On Error GoTo 0
    Next fnd
        
    With objRegex
        .Pattern = "\b(?:[A-Za-z]{4}(?=[A-Za-z]{2}[A-Za-z]?\b)|[A-Za-z]{5}(?=[A-Za-z]{2}[A-Za-z]?\b)|[A-Za-z]{7}(?=[A-Za-z]{3}[A-Za-z]?\b)|[A-Za-z]{5}(?=[A-Za-z]{3}[A-Za-z]?\b)|[A-Za-z]{7}(?=[A-Za-z]{4}[A-Za-z]?\b)|[A-Za-z]{8}(?=[A-Za-z]{5}[A-Za-z]?\b)|[A-Za-z]{9}(?=[A-Za-z]{6}[A-Za-z]?\b)|[A-Za-z]{9}(?=[A-Za-z]{8}[A-Za-z]?\b)|[A-Za-z]{10}(?=[A-Za-z]{9}[A-Za-z]?\b))"
        .Global = True
        .IgnoreCase = True
        Set matches2 = .Execute(myRange)
    End With
    
    For Each fnd In matches2
        On Error Resume Next
            Coll.Add fnd, fnd
        On Error GoTo 0
    Next fnd
    
    With objRegex
        .Pattern = "\b(?:[A-Za-z]{3}(?=[A-Za-z]{2}[A-Za-z]?\b)|[A-Za-z]{2}(?=[A-Za-z]{1}[A-Za-z][A-Za-z]?\b))"
        .Global = True
        .IgnoreCase = True
        Set matches3 = .Execute(myRange)
    End With
    
    For Each fnd In matches3
        On Error Resume Next
            Coll.Add fnd, fnd
        On Error GoTo 0
    Next fnd
 
    With ActiveDocument.Range.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Replacement.Font.Bold = True
        .Format = True
        .Forward = True
        .Wrap = wdFindStop
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .MatchPrefix = True
        For Each fnd In Coll
            On Error Resume Next
            Debug.Print fnd & " " & c & "/" & Coll.Count
            c = c + 1
            .Text = fnd
            .Replacement.Text = "^&"
            .Execute Replace:=wdReplaceAll
            On Error GoTo 0
        Next fnd
    End With

myRange.ParagraphFormat.LineSpacing = LinesToPoints(2)

Application.ScreenUpdating = True

End Sub


Private Sub SpeedReaderFullDocRoundDown()

'Follow me on IG @GPFSye

' Ctrl + G opens the console log so you can track progress
' On larger documents, this will freeze, but don't worry, the code is solid, VBA is just slow, it'll take awhile
' but it'll get there lol

Application.ScreenUpdating = False

Dim i, c As Long
Dim objRegex As RegExp
Dim matches As MatchCollection
Dim fnd As Match
Dim Coll As New Collection

    Set objRegex = New RegExp
    Set myRange = ActiveDocument.Content
    
    ' Great excuse to play with regex
    With objRegex
        .Pattern = "\b(?:[A-Za-z]{1}(?=[A-Za-z]{0}[A-Za-z]?[A-Za-z]\b)|[A-Za-z]{2}(?=[A-Za-z]{1}[A-Za-z]?[A-Za-z]\b)|[A-Za-z]{3}(?=[A-Za-z]{2}[A-Za-z]?[A-Za-z]\b)|[A-Za-z]{4}(?=[A-Za-z]{3}[A-Za-z]?[A-Za-z]\b)|[A-Za-z]{5}(?=[A-Za-z]{4}[A-Za-z]?[A-Za-z]\b)|[A-Za-z]{6}(?=[A-Za-z]{5}[A-Za-z]?[A-Za-z]\b)|[A-Za-z]{7}(?=[A-Za-z]{6}[A-Za-z]?[A-Za-z]\b)|[A-Za-z]{8}(?=[A-Za-z]{7}[A-Za-z]?[A-Za-z]\b)|[A-Za-z]{9}(?=[A-Za-z]{8}[A-Za-z]?[A-Za-z]\b)|[A-Za-z]{10}(?=[A-Za-z]{9}[A-Za-z]?[A-Za-z]\b))"
        .Global = True
        .IgnoreCase = True
        Set matches = .Execute(myRange)
    End With
    
    For Each fnd In matches
        On Error Resume Next
            Coll.Add fnd, fnd
        On Error GoTo 0
    Next fnd
 
    With ActiveDocument.Range.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Replacement.Font.Bold = True
        .Format = True
        .Forward = True
        .Wrap = wdFindStop
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .MatchPrefix = True
        For Each fnd In Coll
            On Error Resume Next
            Debug.Print fnd & " " & c & "/" & Coll.Count
            c = c + 1
            .Text = fnd
            .Replacement.Text = "^&"
            .Execute Replace:=wdReplaceAll
            On Error GoTo 0
        Next fnd
    End With

myRange.ParagraphFormat.LineSpacing = LinesToPoints(2)

Application.ScreenUpdating = True

End Sub

