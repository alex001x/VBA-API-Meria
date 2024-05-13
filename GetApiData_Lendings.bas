
Sub GetApiData_Lendings()
    Dim httpRequest As Object
    Set httpRequest = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    ' Définissez l'URL de l'API pour le staking sur Meria
    Dim apiUrl As String
    apiUrl = "https://api.meria.com/v1/lendings"
    
    With httpRequest
        .Open "GET", apiUrl, False
        ' Définir l'en-tête API-KEY avec votre clé API pour l'authentification
        .SetRequestHeader "API-KEY", "your key api"
        .Send
    End With

' Message dans la barre de statut
    Application.StatusBar = "En cours de récupération des données..."

    ' Vérifier la réponse
    If httpRequest.Status = 200 Then
        Dim response As String
        response = httpRequest.responseText
        'MsgBox response
       
        ' Parser la réponse JSON
        Dim json As Object
        Set json = JsonConverter.ParseJson(response)
       
        'MsgBox json("success")
        'MsgBox json("")
        'MsgBox json("amount")
        'MsgBox json("data")
        
        
       ' Vérifier si la propriété "success" est vraie
        If json("success") = True Then
            Dim datas() As Variant
            Dim varDetails() As Variant
            Dim creditDetails() As Variant
            Dim dataCount As Long
            dataCount = json("data").Count
            'Redim le tableau pour 5 colonnes (currencyCode, amount, reward, lockedReward, startDate)+ résum variation et credit
            ReDim datas(1 To dataCount, 1 To 8)
            
            ' Initialisation pour le comptage des détails
            Dim varCount As Long
            varCount = 0
            Dim creditCount As Long
            creditCount = 0
        
            Dim i As Long, j As Long, k As Long
            i = 1
            For Each data In json("data")
                datas(i, 1) = data("currencyCode")
                datas(i, 2) = data("amount")
                datas(i, 3) = data("reward")
                datas(i, 4) = data("lockedReward")
                datas(i, 5) = data("startDate")
                datas(i, 6) = ConvertUnixToDate(data("startDate")) ' Convertir le timestamp Unix en date Excel
                
                    ' Agregation des données pour variations
                    datas(i, 7) = data("variations").Count ' Compter le nombre de variations
                    
                    ' Calcul de la somme des montants dans credits
                    Dim sumCredits As Double
                    sumCredits = 0
                    For Each credit In data("credits")
                        sumCredits = sumCredits + credit("amount")
                    Next credit
                    datas(i, 8) = sumCredits ' Somme des montants des crédits
 
                ' Préparation pour variations et crédits
                varCount = varCount + data("variations").Count
                creditCount = creditCount + data("credits").Count
        
                ' Message dans la barre de statut
                Application.StatusBar = i & " Données en cours de récupération..."
                
                i = i + 1
            Next data
            
            ' Redimensionnement des tableaux pour variations et crédits
            ReDim varDetails(1 To varCount, 1 To 5)
            ReDim creditDetails(1 To creditCount, 1 To 5)
            
            ' Reset des indices pour remplissage détaillé
            i = 1
            j = 1
            k = 1
            ' Remplissage des tableaux détaillés
        For Each data In json("data")
            For Each variation In data("variations")
                varDetails(j, 1) = data("currencyCode")
                varDetails(j, 2) = variation("amount")
                varDetails(j, 3) = ConvertUnixToDate(variation("date"))
                varDetails(j, 4) = ConvertUnixToDate(variation("effectiveDate"))
                varDetails(j, 5) = variation("applied")
                ' Message dans la barre de statut
                Application.StatusBar = j & " Données en cours de récupération..."
                j = j + 1
            Next variation
            
            For Each credit In data("credits")
                creditDetails(k, 1) = data("currencyCode")
                creditDetails(k, 2) = credit("amount")
                creditDetails(k, 3) = ConvertUnixToDate(credit("date"))
                creditDetails(k, 4) = credit("released")
                ' Message dans la barre de statut
                Application.StatusBar = k & " Données en cours de récupération..."
                k = k + 1
            Next credit
        Next data
        
        
        ' Ecrire les donées dans Excel , remplacer "API Meria Lendings" par le nom de votre feuille
            With ThisWorkbook.Sheets("API Meria Lendings")
                .Range(Cells(3, 1), .Cells(dataCount + 1, 8)).Value = datas
            
            ' Écriture des détails de variations
            .Range(.Cells(3, 11), .Cells(varCount + 1, 15)).Value = varDetails
            
            ' Écriture des détails de crédits
            .Range(.Cells(3, 18), .Cells(creditCount + 1, 21)).Value = creditDetails
            
            End With
        
        ' Réinitialiser la barre de statut
        Application.StatusBar = False
        
        ' Afficher une boîte de message de succès
        MsgBox "Données reçu avec succès", vbInformation, "Succès"
    Else
        Application.StatusBar = "Échec de la récupération des données."
        MsgBox "Erreur lors de la récupération des données. Vérifiez votre connexion et vos paramètres API.", vbCritical, "Erreur"
        
        End If
    End If
End Sub
Function ConvertUnixToDate(ByVal unixTimestamp As Long) As Date
    ' Conversion du timestamp Unix en date Excel
    ConvertUnixToDate = DateAdd("s", unixTimestamp, "1970-01-01")
End Function
