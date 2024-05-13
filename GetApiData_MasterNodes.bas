
Sub GetApiData_MasterNodes()
    Dim httpRequest As Object
    Set httpRequest = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    ' Définissez l'URL de l'API pour le portefeuille sur Meria
    Dim apiUrl As String
    apiUrl = "https://api.meria.com/v1/masternodes"
    
    With httpRequest
        .Open "GET", apiUrl, False
        ' Définir l'en-tête API-KEY avec votre clé API pour l'authentification
        .SetRequestHeader "API-KEY", "Your API KEY"
        .Send
    End With
    
    ' Vérifier la réponse
    If httpRequest.Status = 200 Then
        Dim response As String
        response = httpRequest.responseText
        'MsgBox response

        ' Parser la réponse JSON
        Dim json As Object
        Set json = JsonConverter.ParseJson(response)
       
        
       ' Vérifier si la propriété "success" est vraie
        If json("success") = True Then
            Dim datas() As Variant
            ReDim datas(json("data").Count, 7)
            Dim data As Variant
            
            Dim i As Long
            i = 0
                    For Each data In json("data")
                      datas(i, 0) = data("id")
                      datas(i, 1) = data("status")
                      datas(i, 2) = data("collateral")
                      datas(i, 3) = data("currencyCode")
                      datas(i, 4) = data("reward")
                      datas(i, 5) = ConvertUnixToDate(data("startDate"))
                            ' Correction de la vérification de la valeur Null
                            ' Utilisez IsNull pour vérifier correctement si data("endDate") est Null
                            If IsNull(data("endDate")) Then
                                datas(i, 6) = "N/A"  ' Mettre une valeur par défaut ou laisser vide
                            Else
                                datas(i, 6) = ConvertUnixToDate(data("endDate"))
                            End If
                      i = i + 1
            Next data
          'change "API M MasterNodes" par le nom de votre feuille          
         Sheets("API M MasterNodes").Range(Cells(3, 1), Cells(json("data").Count + 1, 7)) = datas
            
    End If
   End If
End Sub
Function ConvertUnixToDate(ByVal unixTimestamp As Long) As Date
    ' Conversion du timestamp Unix en date Excel
    ConvertUnixToDate = DateAdd("s", unixTimestamp, "1970-01-01")
End Function


