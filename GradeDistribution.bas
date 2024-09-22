Attribute VB_Name = "GradeDistribution"
Option Explicit
Private CUSTOMCOLOR As Long
Private Sub InitCustomColor()
    CUSTOMCOLOR = RGB(255, 0, 0)
End Sub

Sub GenerateGradeDistributionLinear()
    Call InitCustomColor
    Dim maxPoints As Double
    maxPoints = GetMaxPoints()

    Dim MyTable As table
    Set MyTable = CreateTable()

    Dim gradeDistribution() As Variant
    ReDim gradeDistribution(1 To 6, 1 To 5)

    ' Calculate linear grading scale values
    Dim i As Integer
    For i = 1 To 6
        gradeDistribution(i, 1) = Int((maxPoints - (i - 1) * maxPoints / 6) / 5 * 10) * 5 / 10
        gradeDistribution(i, 2) = Int((maxPoints - (i) * maxPoints / 6) / 5 * 10) * 5 / 10 + 0.5
        If i = 6 Then gradeDistribution(i, 2) = 0
        gradeDistribution(i, 3) = gradeDistribution(i, 1) - gradeDistribution(i, 2)
        gradeDistribution(i, 4) = gradeDistribution(i, 1) / maxPoints
        gradeDistribution(i, 5) = gradeDistribution(i, 2) / maxPoints
    Next i

    PopulateTable MyTable, gradeDistribution, 6
End Sub
Sub GenerateGradeDistributionFiftyPercent()
    Call InitCustomColor
    Dim maxPoints As Double
    maxPoints = GetMaxPoints()

    Dim MyTable As table
    Set MyTable = CreateTable()

    Dim gradeDistribution() As Variant
    ReDim gradeDistribution(1 To 6, 1 To 5)

    Dim mid As Double
    mid = maxPoints / 2
    Dim upperRange As Double
    Dim downerRanger As Double
    upperRange = maxPoints - mid - 1.5
    downerRanger = mid - 1

    ' Calculate fifty percent grading scale values
    Dim i As Integer
    i = 0
    Do
        If upperRange > 0 Then
            gradeDistribution(i + 1, 3) = gradeDistribution(i + 1, 3) + 0.5
            upperRange = upperRange - 0.5
            i = (i + 1) Mod 4
        Else
            Exit Do
        End If
    Loop
    gradeDistribution(1, 1) = maxPoints
    gradeDistribution(1, 2) = gradeDistribution(1, 1) - gradeDistribution(1, 3)
    gradeDistribution(5, 3) = Int(downerRanger + 0.5) / 2
    gradeDistribution(6, 3) = Int(downerRanger) / 2

    For i = 1 To 6
        If i > 1 Then
            gradeDistribution(i, 1) = gradeDistribution(i - 1, 2) - 0.5
            gradeDistribution(i, 2) = gradeDistribution(i, 1) - gradeDistribution(i, 3)
        End If
        gradeDistribution(i, 4) = gradeDistribution(i, 1) / maxPoints
        gradeDistribution(i, 5) = gradeDistribution(i, 2) / maxPoints
    Next i

    PopulateTable MyTable, gradeDistribution, 7
End Sub
Function GetMaxPoints() As Double
    Dim inputString As String
    Do
        inputString = InputBox("Gib die Gesamtpunktzahl ein:", "Gesamtpunktzahl eingeben")
        If IsNumeric(inputString) Then
            GetMaxPoints = CDbl(inputString)
            Exit Function
        Else
            MsgBox "Ungültige Eingabe. Bitte eine Zahl eingeben."
        End If
    Loop
End Function
Function CreateTable() As table
    Dim MyTable As table
    Set MyTable = ActiveDocument.Tables.Add(Range:=Selection.Range, NumRows:=7, NumColumns:=7)
    With MyTable
        .Range.Font.COLOR = CUSTOMCOLOR
        .Rows(1).Borders(wdBorderBottom).LineStyle = wdLineStyleDouble
        .Rows(1).Borders(wdBorderBottom).COLOR = CUSTOMCOLOR
        .Rows(1).Range.Font.Bold = True
        .Columns(1).Borders(wdBorderRight).LineStyle = wdLineStyleDouble
        .Columns(1).Borders(wdBorderRight).COLOR = CUSTOMCOLOR
        .Spacing = 0
        .Select
        Selection.ParagraphFormat.SpaceAfter = 0
        With Selection.ParagraphFormat
            .LineSpacingRule = wdLineSpaceMultiple
            .LineSpacing = 12
        End With
        .Columns(2).Select
        Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
        .Columns(3).Select
        Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
        Dim i As Integer
        For i = 4 To .Columns.Count
            .Columns(i).Select
            Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        Next i
        .Columns(5).Width = CentimetersToPoints(0.5)
        ' Header row
        .Rows(1).Cells(2).Range.Text = "Von"
        .Rows(1).Cells(3).Range.Text = "Bis"
        .Rows(1).Cells(4).Range.Text = "Schritt"
        .Rows(1).Cells(6).Range.Text = "% Von"
        .Rows(1).Cells(7).Range.Text = "% Bis"
    End With
    Set CreateTable = MyTable
End Function
Sub PopulateTable(MyTable As table, gradeDistribution() As Variant, fiftyPercentColumn As Integer)
    Dim i As Integer
    For i = 1 To 6
        MyTable.Rows(i + 1).Cells(1).Range.Text = Format(i, "0")
        MyTable.Cell(i + 1, 1).Range.Font.Bold = True
        MyTable.Rows(i + 1).Cells(2).Range.Text = Format(gradeDistribution(i, 1), "0.0")
        MyTable.Rows(i + 1).Cells(3).Range.Text = Format(gradeDistribution(i, 2), "0.0")
        MyTable.Rows(i + 1).Cells(4).Range.Text = Format(gradeDistribution(i, 3), "0.0P.")
        MyTable.Rows(i + 1).Cells(6).Range.Text = FormatPercent(gradeDistribution(i, 4), 0)
        MyTable.Rows(i + 1).Cells(7).Range.Text = FormatPercent(gradeDistribution(i, 5), 0)
    Next i
    MyTable.Cell(5, fiftyPercentColumn).Range.Font.Bold = True ' Set the 50% cell to bold
    MyTable.Columns.AutoFit
    MyTable.Columns(5).Width = CentimetersToPoints(0.5)
End Sub
