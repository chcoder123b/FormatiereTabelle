VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub FormatiereTabelle()
    Dim rng As Range
    Dim tbl As Table
    Dim row As row
    Dim cell As cell
    Dim i As Integer
    Dim rows As Long
    Dim cols As Long
    Dim maxCols As Long
    
    ' Set the selected range
    Set rng = Selection.Range

    ' Check if there is a selection
    If rng.Text = "" Then
        MsgBox "Bitte einen Textbereich markieren", vbExclamation
        Exit Sub
    End If

    ' Calculate number of rows based on the number of paragraphs
    rows = rng.Paragraphs.Count
    
    ' Calculate the maximum number of columns based on the paragraphs
    For i = 1 To rows
        cols = UBound(Split(rng.Paragraphs(i).Range.Text, vbTab)) + 1
        If cols > maxCols Then
            maxCols = cols
        End If
    Next i
    
    ' Convert the selected text to a table
    Set tbl = rng.ConvertToTable(Separator:=wdSeparateByTabs, NumRows:=rows, NumColumns:=maxCols)

    ' Format the entire table
    With tbl
        ' Set font and size
        .Range.Font.Name = "Arial"
        .Range.Font.Size = 10

        ' Apply shading to every second row
        For i = 1 To .rows.Count Step 2
            .rows(i).Shading.BackgroundPatternColor = RGB(221, 221, 221)
        Next i

        ' Format the first row
        With .rows(1)
            .Range.Font.Bold = True
            .Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
        End With

        ' Align the first cell to the left
        .cell(1, 1).Range.ParagraphFormat.Alignment = wdAlignParagraphLeft

        ' Set spacing and alignment in each cell
        For Each cell In .Range.Cells
            With cell
                .VerticalAlignment = wdCellAlignVerticalCenter
                With .Range.ParagraphFormat
                    .SpaceBefore = 6
                    .SpaceAfter = 6
                    .LeftIndent = InchesToPoints(0.1)
                    .RightIndent = InchesToPoints(0.1)
                End With
            End With
        Next cell

        ' Set borders for the entire table
        With .Borders
            .InsideLineStyle = wdLineStyleSingle
            .OutsideLineStyle = wdLineStyleSingle
        End With
    End With
End Sub
