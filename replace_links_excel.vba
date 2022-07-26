Sub test()
    Dim hLink As Hyperlink
    Dim wSheet As Worksheet

    For Each wSheet In Worksheets
       For Each hLink In wSheet.Hyperlinks
            hLink.Address = Replace(hLink.Address, "\\mysrv001\", "\\mysrv002\")
        Next hLink
    Next
End Sub
