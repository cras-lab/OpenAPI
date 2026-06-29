{
  "nbformat": 4,
  "nbformat_minor": 0,
  "metadata": {
    "colab": {
      "provenance": [],
      "authorship_tag": "ABX9TyOTUnsy5J6miYEbXv7yCJgt",
      "include_colab_link": true
    },
    "kernelspec": {
      "name": "python3",
      "display_name": "Python 3"
    },
    "language_info": {
      "name": "python"
    }
  },
  "cells": [
    {
      "cell_type": "markdown",
      "metadata": {
        "id": "view-in-github",
        "colab_type": "text"
      },
      "source": [
        "<a href=\"https://colab.research.google.com/github/cras-lab/OpenAPI/blob/main/FirstOpenAPI_VB.bas\" target=\"_parent\"><img src=\"https://colab.research.google.com/assets/colab-badge.svg\" alt=\"Open In Colab\"/></a>"
      ]
    },
    {
      "cell_type": "code",
      "execution_count": null,
      "metadata": {
        "id": "9U3KMAE1yGK1"
      },
      "outputs": [],
      "source": [
        "Sub Weather_To_Cell()\n",
        "\n",
        "    Dim http As Object\n",
        "    Dim url As String\n",
        "\n",
        "    Set http = CreateObject(\"WinHttp.WinHttpRequest.5.1\")\n",
        "\n",
        "    url = \"https://api.open-meteo.com/v1/forecast\" & _\n",
        "          \"?latitude=37.5665\" & _\n",
        "          \"&longitude=126.9780\" & _\n",
        "          \"&current=temperature_2m,relative_humidity_2m,weather_code\"\n",
        "\n",
        "    http.Open \"GET\", url, False\n",
        "    http.send\n",
        "\n",
        "    Range(\"A1\").Value = \"Status\"\n",
        "    Range(\"B1\").Value = http.Status\n",
        "\n",
        "    Range(\"A2\").Value = \"URL\"\n",
        "    Range(\"B2\").Value = url\n",
        "\n",
        "    Range(\"A3\").Value = \"Response\"\n",
        "    Range(\"B3\").Value = http.responseText\n",
        "\n",
        "    Columns(\"B\").ColumnWidth = 80\n",
        "    Range(\"B3\").WrapText = True\n",
        "\n",
        "End Sub\n",
        "\n",
        "Sub Weather_Params()\n",
        "\n",
        "    Dim http As Object\n",
        "    Dim params As Object\n",
        "    Dim url As String\n",
        "\n",
        "    Set http = CreateObject(\"MSXML2.ServerXMLHTTP.6.0\")\n",
        "    Set params = CreateObject(\"Scripting.Dictionary\")\n",
        "\n",
        "    params.Add \"latitude\", \"37.5665\"\n",
        "    params.Add \"longitude\", \"126.9780\"\n",
        "    params.Add \"current\", \"temperature_2m%2Crelative_humidity_2m%2Cweather_code\"\n",
        "\n",
        "    url = \"https://api.open-meteo.com/v1/forecast?\" & BuildQuery(params)\n",
        "\n",
        "    Debug.Print url\n",
        "\n",
        "    http.Open \"GET\", url, False\n",
        "    http.send\n",
        "\n",
        "    Range(\"A10\").Value = \"Status\"\n",
        "    Range(\"B10\").Value = http.Status\n",
        "\n",
        "    Range(\"A12\").Value = \"URL\"\n",
        "    Range(\"B12\").Value = url\n",
        "\n",
        "    Range(\"A13\").Value = \"Response\"\n",
        "    Range(\"B13\").Value = http.responseText\n",
        "\n",
        "    Columns(\"B\").ColumnWidth = 80\n",
        "    Range(\"B13\").WrapText = True\n",
        "\n",
        "End Sub\n",
        "\n",
        "Function BuildQuery(params As Object) As String\n",
        "\n",
        "    Dim key As Variant\n",
        "    Dim result As String\n",
        "\n",
        "    For Each key In params.Keys\n",
        "        result = result & key & \"=\" & params(key) & \"&\"\n",
        "    Next key\n",
        "\n",
        "    If Len(result) > 0 Then\n",
        "        result = Left(result, Len(result) - 1)\n",
        "    End If\n",
        "\n",
        "    BuildQuery = result\n",
        "\n",
        "End Function\n"
      ]
    }
  ]
}