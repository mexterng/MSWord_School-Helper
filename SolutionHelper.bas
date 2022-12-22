Attribute VB_Name = "SolutionHelper"
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
Sub ChangeRedToWhite()
    Call Changer(False)
End Sub
Sub ChangeWhiteToRed()
    Call Changer(True)
End Sub
Sub Exporter(teacher As Boolean)
' teacher = True: ExportLK
' teacher = False: ExportSuS
    Dim var As String
    If teacher Then
        var = "LK"
    Else
        var = "SuS"
    End If
    ' Show Save As dialog and select location + filename
    Dim SaveAsDlg As FileDialog
    Set SaveAsDlg = Application.FileDialog(msoFileDialogSaveAs)
    With SaveAsDlg
        .InitialView = msoFileDialogViewList
        .InitialFileName = ActiveDocument.Path & Application.PathSeparator & Split(ActiveDocument.Name, ".")(0) & "_" & var & ".pdf"
        .FilterIndex = 7
        .Title = "Speichern unter ... (Exportdatei für " & var & ")"

    End With
  ' Should be saved?
    If SaveAsDlg.Show <> 0 Then
        Application.StatusBar = "PDF-Export für " & var & " läuft ..."
        DoEvents
            ChangeRedToWhite ' Hide sample solutions
            SaveAsDlg.Execute ' Save PDF
            ChangeWhiteToRed ' Show sample solutions
        Application.StatusBar = ""
        MsgBox "PDF-Export (" & var & ") erfolgreich abgeschlossen."
    Else
        MsgBox "PDF-Export (" & var & ") abgebrochen."
    End If
End Sub
Sub Changer(color As Boolean)
' color = True:  WHITE -> RED
' color = False: RED -> WHITE
    DoEvents
    ActiveDocument.Bookmarks.Add _
    Name:="temp", Range:=Selection.Range ' Save current cursor position
  ' Replace body text
    ActiveDocument.Select ' Select whole document
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    If color Then
        Selection.Find.Font.color = RGB(255, 255, 255)
        Selection.Find.Replacement.Font.color = RGB(255, 0, 0)
    Else
        Selection.Find.Font.color = RGB(255, 0, 0)
        Selection.Find.Replacement.Font.color = RGB(255, 255, 255)
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
  ' Replace header and footer
    ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
    Selection.Find.Format = True
    Selection.Find.Execute Replace:=wdReplaceAll
  ' Restore old cursor position
    ActiveDocument.Bookmarks("temp").Select
    ActiveDocument.Bookmarks("temp").Delete
End Sub
