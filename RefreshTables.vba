Private Sub Refresh_Click()

    For Each objConnection In ThisWorkbook.Connections
        'Get current background-refresh value
        bBackground = objConnection.OLEDBConnection.BackgroundQuery

        'Temporarily disable background-refresh
        objConnection.OLEDBConnection.BackgroundQuery = False

        'Refresh this connection
        objConnection.Refresh

        'Set background-refresh value back to original value
        objConnection.OLEDBConnection.BackgroundQuery = bBackground
    Next

    MsgBox "Finished refreshing all data connections"
dup
End Sub
