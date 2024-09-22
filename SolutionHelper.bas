Attribute VB_Name = "SolutionHelper"
Option Explicit
Private Const SOL = "LOESUNG"
Private Const SPEC = "ANGABE"
Private CUSTOMCOLOR As Long
Private Sub InitCustomColor()
    CUSTOMCOLOR = RGB(255, 0, 0)
End Sub
Sub ExportAll()
' Exports student and sample solution version of the document as PDF.
' Characters for the sample solution are created in RED RGB(255,0,0).
    Call ExportLK
    Call ExportSuS
End Sub
Sub ExportLK()
    Call Exporter(True)
End Sub
Sub ExportSuS()
    Call Exporter(False)
End Sub
Sub ChangeCustomToWhite()
    Call Changer(False)
End Sub
Sub ChangeWhiteToCustom()
    Call Changer(True)
End Sub
Sub Exporter(teacher As Boolean)
    Call InitCustomColor
' teacher = True: ExportLK
' teacher = False: ExportSuS
    Dim var As String
    If teacher Then
        var = SOL
    Else
        var = SPEC
    End If
    ' Show Save As dialog and select location + filename
    Dim SaveAsDlg As FileDialog
    Set SaveAsDlg = Application.FileDialog(msoFileDialogSaveAs)
    With SaveAsDlg
        .InitialView = msoFileDialogViewList
        .InitialFileName = ActiveDocument.Path & Application.PathSeparator & Split(ActiveDocument.Name, ".")(0) & "_" & var & ".pdf"
        .FilterIndex = 7
        .Title = "Speichern unter ... (Exportdatei f r " & var & ")"

    End With
    ' Should be saved?
    If SaveAsDlg.Show <> 0 Then
        Application.StatusBar = "PDF-Export f r " & var & " l uft ..."
        DoEvents
            Changer (teacher) ' Hide sample solutions
            SaveAsDlg.Execute ' Save PDF
            ChangeWhiteToCustom ' Show sample solutions
        Application.StatusBar = ""
        MsgBox "PDF-Export (" & var & ") erfolgreich abgeschlossen."
    Else
        MsgBox "PDF-Export (" & var & ") abgebrochen."
    End If
End Sub
Sub Changer(colorBool As Boolean)
    Call InitCustomColor
    ' colorBool = True:  WHITE -> customColor
    ' colorBool = False: customColor -> WHITE
    
    DoEvents
    ActiveDocument.Bookmarks.Add _
    Name:="temp", Range:=Selection.Range ' Save current cursor position
    ' Replace body text
    ActiveDocument.Select ' Select whole document
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    If colorBool Then
        Selection.Find.Font.COLOR = RGB(255, 255, 255)
        Selection.Find.Replacement.Font.COLOR = CUSTOMCOLOR
    Else
        Selection.Find.Font.COLOR = CUSTOMCOLOR
        Selection.Find.Replacement.Font.COLOR = RGB(255, 255, 255)
    End If
    Selection.Find.Format = True
    Selection.Find.Execute Replace:=wdReplaceAll
    ' Replace textboxes
    Dim longCount As Long
    For longCount = 1 To ActiveDocument.Shapes.Count
        ActiveDocument.Shapes.Range(longCount).Select
        Selection.Find.Format = True
        Selection.Find.Execute Replace:=wdReplaceAll
    Next
    ' Replace shapes
    Dim shape As shape
    For Each shape In ActiveDocument.Shapes
        ' fill
        If shape.Fill.ForeColor.RGB = CUSTOMCOLOR Then
            If colorBool Then
                shape.Fill.Transparency = 0
            Else
                shape.Fill.Transparency = 1
            End If
        End If
        
        ' lines
        If shape.line.ForeColor.RGB = CUSTOMCOLOR Then
            If colorBool Then
                shape.line.Transparency = 0
            Else
                shape.line.Transparency = 1
            End If
        End If
    Next shape
    
    ' Replace table lines
    Dim aTable As table
    Dim Row As Integer
    Dim column As Integer
    Dim newColor, oldColor As Long
    If colorBool Then
        newColor = CUSTOMCOLOR
        oldColor = RGB(255, 255, 255)
    Else
        newColor = RGB(255, 255, 255)
        oldColor = CUSTOMCOLOR
    End If
    
    For Each aTable In ActiveDocument.Tables
        For Row = 1 To aTable.Rows.Count
            For column = 1 To aTable.Columns.Count
                On Error Resume Next
                With aTable.Cell(Row, column)
                    If .Borders(wdBorderTop).Visible And .Borders(wdBorderTop).COLOR = oldColor Then
                        .Borders(wdBorderTop).COLOR = newColor
                    End If
                    If .Borders(wdBorderBottom).Visible And .Borders(wdBorderBottom).COLOR = oldColor Then
                        .Borders(wdBorderBottom).COLOR = newColor
                    End If
                    If .Borders(wdBorderLeft).Visible And .Borders(wdBorderLeft).COLOR = oldColor Then
                        .Borders(wdBorderLeft).COLOR = newColor
                    End If
                    If .Borders(wdBorderRight).Visible And .Borders(wdBorderRight).COLOR = oldColor Then
                        .Borders(wdBorderRight).COLOR = newColor
                    End If
                End With
            Next column
        Next Row
    Next
        
    ' Replace header and footer
    ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
    Selection.Find.Format = True
    Selection.Find.Execute Replace:=wdReplaceAll
    ' Restore old cursor position
    ActiveDocument.Bookmarks("temp").Select
    ActiveDocument.Bookmarks("temp").delete
End Sub
