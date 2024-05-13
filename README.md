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
