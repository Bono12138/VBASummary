Sub CopyTables()
    Dim Source As Document
    Dim Target As Document
    Dim tbl As Table
    Dim tr As Range

    Set Source = ActiveDocument
    Set Target = Documents.Add

    For Each tbl In Source.Tables
        tr.Collapse wdCollapseEnd
        tr.FormattedText = tbl.Range.FormattedText
        tr.Collapse wdCollapseEnd
        tr.Text = vbCrLf
    Next
End Sub