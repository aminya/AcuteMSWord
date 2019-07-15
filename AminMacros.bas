Attribute VB_Name = "AminMacros"

Sub MathType2Word()
Attribute MathType2Word.VB_ProcData.VB_Invoke_Func = "Normal.AminMacros.MathType2Word"
' By Amin Yahyaabadi
' MathType2Word Macro: to convert Mathtype Equations to Microsoft Word Equations
'
'
    Application.Run MacroName:="MathTypeCommands.UILib.MTCommand_TeXToggle"
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "(\\\[)(*)(\\\])"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchKashida = False
        .MatchDiacritics = False
        .MatchAlefHamza = False
        .MatchControl = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute
    Selection.Cut
    Selection.OMaths.Add Range:=Selection.Range
    Selection.paste
    Selection.OMaths.BuildUp
    
    Selection.Find.ClearFormatting
    With Selection.Find
        .Text = "$*$"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchKashida = False
        .MatchDiacritics = False
        .MatchAlefHamza = False
        .MatchControl = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Cut
    Selection.OMaths.Add Range:=Selection.Range
    Selection.paste
    Selection.OMaths.BuildUp
    
End Sub


Sub codePaste()
Attribute codePaste.VB_ProcData.VB_Invoke_Func = "Normal.AminMacros.codePaste"
' By Amin Yahyaabadi
' codePaste Macro: to paste a code snippet into MS word and make its background gray
'
'
  Dim oRng As Word.Range
  Set oRng = Selection.Range
  oRng.paste
  oRng.Select
  
    With Selection.ParagraphFormat
        With .Shading
            .Texture = wdTexture10Percent
            .ForegroundPatternColor = wdColorAutomatic
            .BackgroundPatternColor = wdColorWhite
        End With
        .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
        .Borders(wdBorderRight).LineStyle = wdLineStyleNone
        .Borders(wdBorderTop).LineStyle = wdLineStyleNone
        .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
        .Borders(wdBorderHorizontal).LineStyle = wdLineStyleNone
        With .Borders
            .DistanceFromTop = 1
            .DistanceFromLeft = 4
            .DistanceFromBottom = 1
            .DistanceFromRight = 4
            .Shadow = False
        End With
    End With
    With Options
        .DefaultBorderLineStyle = wdLineStyleSingle
        .DefaultBorderLineWidth = wdLineWidth050pt
        .DefaultBorderColor = wdColorAutomatic
    End With
End Sub



Sub pasteSelected()
'
'A basic Word macro coded by Greg Maxey
  Dim oRng As Word.Range
  Set oRng = Selection.Range
  oRng.paste
  oRng.Select
End Sub