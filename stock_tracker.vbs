Sub stock_tracker(WSheet As Worksheet):
    'This sub routine is used to tally in each sheet based on Ticker (A*) and Volume (G*)
    Dim ws As Worksheet
    Set ws = WSheet

    'Initialize first Variables (Header in row 1)
    Dim curr As String
    Dim curr_total As Double
    Dim curr_open As Double
    Dim curr_closed As Double
    Dim curr_percent As Double 'To be used later

    'Initializing Max and Min Global Values
    Dim max_inc As Double
    Dim max_dec As Double
    Dim max_vol As Double

    curr = ""
    curr_total = 0
    curr_open = 0
    curr_closed = 0

    max_inc = 0
    max_dec = 0
    max_vol = 0

    'Initialize Destination (I*, J*, K*, L*)
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"

    'Formating Destination 1. USE COLORINDEX
    Dim red As Integer
    Dim green As Integer

    red = 3
    green = 4

    'Initialize Destination (N2:N4,O2:O4)
    ws.Range("N3").Value = "Greatest % Increase"
    ws.Range("N4").Value = "Greatest % Decrease"
    ws.Range("N5").Value = "Greatest Total Volume"
    ws.Range("O2").Value = "Ticker"
    ws.Range("P2").Value = "Value"

    ws.Range("P3:P4").NumberFormat = "0.00%"

    'Formating Destination 2
    'Range("N2:P2", "N2:N5").Borders (xlDiagonalUp)
    'Range("N2:P2", "N2:N5").Borders (xlInsideHorizontal)

    'Looping Logic
    Dim i As Long 'Counter
    Dim row_max As Long

    Dim j As Long 'Row counter for results
    j = 2

    row_max = ws.Rows.Count 'Max Row Variable

    For i = 1 To (row_max - 1)
        'Initalizing curr
        If curr = "" Then
            curr = ws.Cells(i + 1, 1).Value 'Ticker
            curr_total = ws.Cells(i + 1, 7) 'Volume
            curr_open = ws.Cells(i + 1, 3) 'Open Price

        'Updating Stock Volume Total
        ElseIf (curr = ws.Cells(i + 1, 1).Value) Then
            curr_total = curr_total + ws.Cells(i + 1, 7).Value
            curr_closed = ws.Cells(i + 1, 6)

        'Different Stock. Store data and Start new
        ElseIf Not (curr = ws.Cells(i + 1, 1).Value) Then
            'Save the Data
            ws.Cells(j, 9).Value = curr 'Ticker
            ws.Cells(j, 12).Value = curr_total 'Volume total

            If (curr_open = 0) Then 'Divide by 0 Error
                curr_percent = 0
            Else
                curr_percent = (curr_closed - curr_open) / curr_open
            End If

            ws.Cells(j, 10).Value = curr_closed - curr_open 'Yearly Change
            ws.Cells(j, 11).Value = curr_percent 'Percent Change

            'Second conditional for formatting
            If (ws.Cells(j, 10).Value > 0) Then
                ws.Cells(j, 10).Interior.ColorIndex = green
            Else
                ws.Cells(j, 10).Interior.ColorIndex = red
            End If

            ws.Cells(j, 11).NumberFormat = "0.00%" 'Format to Percent
            j = j + 1 'increment j

            'Set MAX or MIN
            If (curr_percent > max_inc) Then
                max_inc = curr_percent
                ws.Range("O3").Value = curr
                ws.Range("P3").Value = max_inc
            End If

            If (curr_percent < max_dec) Then
                max_dec = curr_percent
                ws.Range("O4").Value = curr
                ws.Range("P4").Value = max_dec
            End If

            If (curr_total > max_tot) Then
                max_tot = curr_total
                ws.Range("O5").Value = curr
                ws.Range("P5").Value = max_tot
            End If

            'Start new Ticker and Total
            curr = ws.Cells(i + 1, 1).Value
            curr_total = ws.Cells(i + 1, 7).Value
            curr_open = ws.Cells(i + 1, 3).Value

        End If

    Next i

End Sub
