Sub Run_Tracker():

    Dim sht As Worksheet

    For Each sht In ThisWorkbook.Worksheets

        Call stock_tracker(sht)

    Next sht

End Sub
