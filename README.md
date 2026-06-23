# gg_sheet_calendar
Google Sheet Script that automatically creates events in your google agenda :)



# 📅 Google Sheets → Google Calendar Sync

Automatiser la création, modification et suppression d'événements Google Calendar 
directement depuis Google Sheets avec cases à cocher !

## ✨ Fonctionnalités

- ✅ **Créer** des événements en cochant une case
- ✏️ **Modifier** des événements existants
- 🗑️ **Supprimer** des événements
- 📅 Support des événements **journée entière**
- 👥 Ajout automatique d'**invités**
- 🎨 Coloration automatique des lignes (succès/erreur/attente)
- 🔄 Traitement en lot avec `runAllPending()`
- 🌍 Headers multilingues (FR/EN)

## 🚀 Installation rapide

### 1. Copier le template

👉 **[Cliquez ici pour copier le template Google Sheets](https://docs.google.com/spreadsheets/d/1dK_MdRAN0NXJ0je8OXFyUReraSN4Fst0ogPEkDMGVHk/copy)**



### 2. Ajouter le script
/// NORMALEMENT VOUS POUVEZ COPIER L'APP SCRIPT MAIS SI CE N'EST PAS LE CAS SUIVEZ LES ETAPES SUIVANTES : ///
1. Dans votre copie, allez dans **Extensions → Apps Script**
2. Supprimez le code existant
3. Copiez-collez le contenu de [`src/Code.gs`](src/Code.gs)
4. Sauvegardez (Ctrl+S)

### 3. Configurer le trigger

1. Dans l'éditeur Apps Script, exécutez la fonction `setup()`
2. Autorisez les permissions demandées (Google Calendar + Sheets)

## 📊 Structure du template

| Colonne | Description | Obligatoire |
|---------|-------------|:-----------:|
| **Title** | Titre de l'événement | ✅ |
| **Start Date** | Date de début (YYYY-MM-DD ou format local) | ✅ |
| **Start Time** | Heure de début (HH:MM) | ❌ |
| **End Date** | Date de fin | ❌ |
| **End Time** | Heure de fin | ❌ |
| **All Day** | Événement journée entière (TRUE/FALSE) | ❌ |
| **Description** | Description de l'événement | ❌ |
| **Location** | Lieu | ❌ |
| **Guests** | Emails des invités (séparés par , ou ;) | ❌ |
| **Calendar ID** | ID du calendrier cible (vide = calendrier par défaut de l'endroit ou se trouve le ggsheet - votre drive gmail ) | ❌ |
| **Status** | Statut automatique (Created/Modified/Deleted/Error) | Auto |
| **Event ID** | ID de l'événement créé | Auto |
| **Processing Log** | Journal de traitement | Auto |
| **Last Updated** | Horodatage ou message d'erreur | Auto |
| **Create** | ☑️ Cocher pour créer | Action |
| **Modify** | ☑️ Cocher pour modifier | Action |
| **Delete** | ☑️ Cocher pour supprimer | Action |

## 📖 Utilisation

### Créer un événement
1. Remplissez une ligne avec au minimum **Title** et **Start Date**
2. Cochez la case **Create** ✅
3. L'événement est créé automatiquement, la case se décoche

### Modifier un événement
1. Modifiez les données de la ligne (l'**Event ID** doit exister)
2. Cochez la case **Modify** ✅

### Supprimer un événement
1. Cochez la case **Delete** ✅ sur la ligne (l'**Event ID** doit exister)

### Traitement en lot
Exécutez `runAllPending()` depuis Apps Script pour traiter toutes les lignes cochées.

## 🔧 Configuration avancée

### Utiliser un calendrier spécifique

1. Ouvrez Google Calendar
2. Paramètres du calendrier → Intégrer le calendrier
3. Copiez l'**ID de l'agenda**
4. Collez-le dans la colonne **Calendar ID**

### Headers alternatifs supportés

Le script reconnaît automatiquement plusieurs variantes :

| Canonical | Alternatives acceptées |
|-----------|------------------------|
| title | titre, sujet, summary |
| start date | date debut, date début, date |
| start time | heure debut, heure début, time |
| create | ready, à créer, go, créer |
| ... | (voir code source pour la liste complète) |

## ⚠️ Limitations

- Maximum ~50 événements par exécution (quotas Google)
- Les invités reçoivent une notification par défaut
- Le script nécessite les autorisations Calendar et Sheets

## 🐛 Dépannage

| Problème | Solution |
|----------|----------|
| "Event not found" | L'Event ID est invalide ou l'événement a été supprimé |
| "Calendrier introuvable" | Vérifiez le Calendar ID |
| Pas de réaction au clic | Exécutez `setup()` pour créer le trigger |
| Erreur d'autorisation | Ré-exécutez `setup()` et acceptez les permissions |

## 🤝 Contribution

Les contributions sont les bienvenues ! N'hésitez pas à ouvrir une issue ou une PR.
