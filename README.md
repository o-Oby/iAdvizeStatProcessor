# iAdvizeStatProcessor

**iAdvizeStatProcessor** est un outil puissant et convivial conçu pour extraire et convertir les données statistiques des conversations sur iAdvize en un fichier Excel structuré. Grâce à cet outil, vous pouvez facilement analyser les performances de vos interactions clients avec des indicateurs clés tels que :

- Le nombre total de conversations cumulées chaque mois
- La durée moyenne de traitement des conversations
- Nombre d'heures totales : La formule utilisée par iAdvizeStatProcessor convertit d'abord la durée moyenne de traitement en minutes décimales. Ensuite, elle multiplie cette valeur par le nombre total de conversations pour obtenir le temps total en minutes, qui est ensuite converti en heures.

## Fonctionnalités

- Extraction et conversion des données des conversations en fichier Excel.
- Calcul automatique des heures totales de traitement par mois et par an.
- Interface utilisateur simple : copier-coller les données et générer le rapport d'un simple clic.
- Génération d'un tableau au format Excel propre et structuré, avec des formules calculées automatiquement.

## Utilisation

### Prérequis

- Node.js
- npm (Node Package Manager)
- PM2 (pour la gestion du processus en production)

### Installation

1. Clonez le dépôt :

   ```bash
   git clone https://github.com/o-Oby/iAdvizeStatProcessor.git
   cd iAdvizeStatProcessor
   ```

2. Installez les dépendances :

   ```bash
   npm install
   ```

3. Installez PM2 globalement si ce n'est pas déjà fait :

   ```bash
   npm install -g pm2
   ```

### Lancement de l'application

1. Démarrez l'application avec PM2 :

   ```bash
   pm2 start index.js --name iadvize-stat-organizer
   ```

2. Accédez à l'application dans votre navigateur à l'adresse :

   ```
   http://localhost:3001
   ```

### Comment utiliser iAdvizeStatProcessor

1. **Copier et coller** : Copiez les statistiques des conversations (Nombre de conversations, Mois, Durée moyenne de traitement) depuis iAdvize et collez-les dans la zone de texte de l'application.
2. **Répétition de l'opération** : Répétez l'opération pour toutes les pages de données afin d'avoir toutes les données dans la zone de texte.
3. **Générer le fichier Excel** : Cliquez sur "Convertir en Excel" pour générer un fichier Excel structuré et formaté avec toutes les données. Le fichier indiquera le mois et l'année de la période concernée, le nombre de conversations effectuées dans le mois, la durée moyenne de traitement par conversation et calculera ensuite le nombre d'heures totales effectives par mois puis par année.

## Dépendances

Ce projet utilise les modules suivants :

- [express](https://www.npmjs.com/package/express)
- [body-parser](https://www.npmjs.com/package/body-parser)
- [cors](https://www.npmjs.com/package/cors)
- [fs](https://nodejs.org/api/fs.html)
- [path](https://nodejs.org/api/path.html)
- [url](https://nodejs.org/api/url.html)
- [write-excel-file](https://www.npmjs.com/package/write-excel-file)

## Support

Pour toute question ou assistance, n'hésitez pas à me contacter.

## Licence

Ce projet est sous licence Apache. Voir le fichier [LICENSE](LICENSE) pour plus de détails.
