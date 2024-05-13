
Sub GetApiData_Operations()
    Dim httpRequest As Object
    Set httpRequest = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    ' Définissez l'URL de l'API pour le portefeuille sur Meria
    Dim apiUrl As String
    apiUrl = "https://api.meria.com/v1/operations"
    
    With httpRequest
        .Open "GET", apiUrl, False
        ' Définir l'en-tête API-KEY avec votre clé API pour l'authentification
        .SetRequestHeader "API-KEY", "Your API Key"
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
            ReDim datas(json("data").Count, 10)
            Dim data As Variant
            
            Dim i As Long
            i = 0
            For Each data In json("data")
                     ' Vérifier si l'objet data est Null
                     If Not IsNull(data) Then
                      datas(i, 0) = data("id")
                      datas(i, 1) = data("type")
                      datas(i, 2) = data("status")
                      datas(i, 3) = ConvertUnixToDate(data("Date"))
                            ' Correction de la vérification de la valeur Null
                            ' Utilisez IsNull pour vérifier correctement si data("endDate") est Null
                            If IsNull(data("endDate")) Then
                                datas(i, 6) = "N/A"  ' Mettre une valeur par défaut ou laisser vide
                            Else
                                datas(i, 6) = ConvertUnixToDate(data("endDate"))
                            End If
                      datas(i, 4) = data("sourceCurrencyCode")
                      datas(i, 5) = data("sourceAmount")
                      datas(i, 6) = data("destinationCurrencyCode")
                      datas(i, 7) = data("destinationAmount")
                      datas(i, 8) = data("memo")
                      
                      i = i + 1
                    Else
                    ' Optionnel: sortir de la boucle si un élément indispensable est manquant ou juste continuer
                    ' Exit For  ' Décommentez pour quitter la boucle en cas de data Null
                    MsgBox "Un élément de la liste est vide"
                    End If
            Next data
            
        If i > 0 Then
            Sheets("API M OP").Range(Cells(3, 1), Cells(json("data").Count + 1, 9)) = datas
            MsgBox "Succès"
        Else
           MsgBox "Aucune donnée pour l'affichage"
        End If
        
         
         
    End If
   End If
End Sub
Function ConvertUnixToDate(ByVal unixTimestamp As Long) As Date
    ' Conversion du timestamp Unix en date Excel
    ConvertUnixToDate = DateAdd("s", unixTimestamp, "1970-01-01")
End Function




