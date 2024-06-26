Sub GetApiData_Wallet()
    Dim httpRequest As Object
    Set httpRequest = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    ' Définissez l'URL de l'API pour le portefeuille sur Meria
    Dim apiUrl As String
    apiUrl = "https://api.meria.com/v1/wallets"
    
    With httpRequest
        .Open "GET", apiUrl, False
        ' Définir l'en-tête API-KEY avec votre clé API pour l'authentification
        .SetRequestHeader "API-KEY", "your key api "
        .Send
    End With
    
    ' Vérifier la réponse
    If httpRequest.Status = 200 Then
        Dim response As String
        response = httpRequest.responseText
        'MsgBox response

        ' Parser la réponse JSON , necessite d'ajouter le module JsonConverteur dans vos modules, disponible dans ce projet VBA-API-Meria/JsonConverteur.bas
        Dim json As Object
        Set json = JsonConverter.ParseJson(response)   
        
       ' Vérifier si la propriété "success" est vraie
        If json("success") = True Then
            ' Déclaration variable
            Dim datas As Variant
            ReDim datas(json("data").Count, 2)
            Dim data As Dictionary
       
            Dim i As Long
            i = 0
                    For Each data In json("data")
                      datas(i, 0) = data("currencyCode")
                      datas(i, 1) = data("balance")
                      
                      i = i + 1
            Next data
                    
            ' Remplacer "mysheet" par le nom de votre feuille 
            ' range(cells(ligne, colonne) emplacement 1 celule pour insérer les données
            Sheets("mysheet").Range(Cells(3, 1), Cells(json("data").Count + 1, 2)) = datas

            else 
            msgbox "error ! quelque chose n'a pas fonctionné"
    End If
   End If
End Sub
