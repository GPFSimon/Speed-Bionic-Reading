VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Loading 
   Caption         =   "Loading..."
   ClientHeight    =   2220
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   4580
   OleObjectBlob   =   "Loading.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Loading"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
#If VBA7 And Win64 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As LongPtr)
#Else
    Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
#End If


Private Sub Frame1_Click()

End Sub

Private Sub Label3_Click()

End Sub

Private Sub Label2_Click()

End Sub

Private Sub UserForm_Click()

End Sub

Option Explicit

Private Sub UserForm_Activate()

Application.ScreenUpdating = False

Dim i, c As Long
Dim objRegex As RegExp
Dim matches, matches1, matches2, matches3 As MatchCollection
Dim fnd As Match
Dim Coll As New Collection

Me.Repaint

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
    
    i = 33
    Label1.Width = i
    Frame1.Repaint
    
    For Each fnd In matches3
        On Error Resume Next
            Coll.Add fnd, fnd
        On Error GoTo 0
    Next fnd
    
    i = 66
    Label1.Width = i
    Frame1.Repaint
 
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
i = 200
Label1.Width = i
Frame1.Repaint

Application.ScreenUpdating = True
Sleep 1000
Unload Me


End Sub
