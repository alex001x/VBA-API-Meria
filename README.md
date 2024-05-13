-- Français--

Suite à de nombreuses demandes, je partage un code VBA pour connecter l'API Meria.com à votre feuille Excel. 
Manquant de temps, je vais essayer d'ajouter progressivement toutes les API. 
Vous pourrez ensuite modifier le code selon vos besoins.
D'autre API seront traités tel que l'accès à un référenciel pour obtenir la valeur en temps réel des actifs soit depuis coinMarketCap ou autre...
Laisser moi vos commentaires.

Dans votre fichier Excel, 
ajoutez un bouton depuis le menu développeur : Dev → Insérer → Bouton. 
Dessinez la zone du bouton sur votre feuille. Une fenêtre intitulée "Affecter une macro" apparaîtra, 
cliquez sur "Nouveau". Un nouveau module sera proposé et les lignes Sub seront générées automatiquement. 
Copiez et collez le code  "GetAPI de votre choix" dans votre fichier module, en remplaçant les lignes Sub générées. 
Ajouter Votre clé API (Depuis votre compte Meria, vous pouvez générer une clé API) Attention ne partager pas votre clé !
Notez que pour chaque fichier *.bas, un nouveau module est nécessaire.
N oubliez pas de créer un nouveau module et de copier coller JsonConverteur, Indispensable pour parser vos données Json
Revenez ensuite sur votre feuille Excel et cliquez sur le bouton pour exécuter.

A partir des informations reçu vous pouvez concevoir un tableau de bord personnaliser, appliquer des graphiques, indiquer vos points de réfénrences ou d'actions, analyser votre moyenne d'achat par CurrencyCode, etc...

Tous les codes API peuvent être écrits dans le même fichier module, avec des noms de Sub différents. Sélectionnez les Sub correspondants pour vos boutons.

--English--

Following numerous requests, I am sharing a VBA code to connect the Meria.com API to your Excel sheet. Lacking time, I will try to gradually add all the APIs. You can then modify the code according to your needs. Other APIs will be addressed such as accessing a repository to obtain the real-time value of assets either from CoinMarketCap or another source... Please leave me your comments.

In your Excel file, add a button from the developer menu: Dev → Insert → Button. Draw the button area on your sheet. A window titled "Assign Macro" will appear, click "New". A new module will be proposed and the Sub lines will be generated automatically. Copy and paste the "GetAPI of your choice" code into your module file, replacing the generated Sub lines. Add your API key (From your Meria account, you can generate an API key) but be careful not to share your key! Note that for each *.bas file, a new module is necessary. Do not forget to create a new module and copy and paste JsonConverter, essential for parsing your Json data. Then return to your Excel sheet and click the button to execute.

From the information received, you can design a customized dashboard, apply graphs, indicate your reference or action points, analyze your average purchase by CurrencyCode, etc...

All API codes can be written in the same module file, with different Sub names. Select the corresponding Subs for your buttons.

Alexandre OLIVEIRA
;)
