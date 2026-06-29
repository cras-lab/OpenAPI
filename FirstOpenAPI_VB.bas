Sub Weather_To_Cell()

    Dim http As Object
    Dim url As String

    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")

    url = "https://api.open-meteo.com/v1/forecast" & _
          "?latitude=37.5665" & _
          "&longitude=126.9780" & _
          "&current=temperature_2m,relative_humidity_2m,weather_code"

    http.Open "GET", url, False
    http.send

    Range("A1").Value = "Status"
    Range("B1").Value = http.Status

    Range("A2").Value = "URL"
    Range("B2").Value = url

    Range("A3").Value = "Response"
    Range("B3").Value = http.responseText

    Columns("B").ColumnWidth = 80
    Range("B3").WrapText = True

End Sub

Sub Weather_Params()

    Dim http As Object
    Dim params As Object
    Dim url As String

    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    Set params = CreateObject("Scripting.Dictionary")

    params.Add "latitude", "37.5665"
    params.Add "longitude", "126.9780"
    params.Add "current", "temperature_2m%2Crelative_humidity_2m%2Cweather_code"

    url = "https://api.open-meteo.com/v1/forecast?" & BuildQuery(params)

    Debug.Print url

    http.Open "GET", url, False
    http.send

    Range("A10").Value = "Status"
    Range("B10").Value = http.Status

    Range("A12").Value = "URL"
    Range("B12").Value = url

    Range("A13").Value = "Response"
    Range("B13").Value = http.responseText

    Columns("B").ColumnWidth = 80
    Range("B13").WrapText = True

End Sub

Function BuildQuery(params As Object) As String

    Dim key As Variant
    Dim result As String

    For Each key In params.Keys
        result = result & key & "=" & params(key) & "&"
    Next key

    If Len(result) > 0 Then
        result = Left(result, Len(result) - 1)
    End If

    BuildQuery = result

End Function
