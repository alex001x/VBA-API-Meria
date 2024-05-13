
Sub GetApiData_Adresses()
    Dim httpRequest As Object
    Set httpRequest = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    ' Définissez l'URL de l'API pour le portefeuille sur Meria
    Dim apiUrl As String
    apiUrl = "https://api.meria.com/v1/walletAddresses"
    
    With httpRequest
        .Open "GET", apiUrl, False
        ' Définir l'en-tête API-KEY avec votre clé API pour l'authentification
        .SetRequestHeader "API-KEY", "Your Key API"
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
            ReDim datas(json("data").Count, 6)
            Dim data As Variant
            
            Dim i As Long
            i = 0
            For Each data In json("data")
                     ' Vérifier si l'objet data est Null
                     If Not IsNull(data) Then
                      datas(i, 0) = data("id")
                      datas(i, 1) = data("currencyCode")
                      datas(i, 2) = data("name")
                      datas(i, 3) = data("adress")
                      datas(i, 4) = data("isBEP2")
                      datas(i, 5) = data("memo")
                      
                      i = i + 1
                    Else
                    ' Optionnel: sortir de la boucle si un élément indispensable est manquant ou juste continuer
                    ' Exit For  ' Décommentez pour quitter la boucle en cas de data Null
                    MsgBox "Un élément de la liste est vide"
                    End If
            Next data
            
        If i > 0 Then
            Sheets("API M Adresse").Range(Cells(3, 1), Cells(json("data").Count + 1, 6)) = datas
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
