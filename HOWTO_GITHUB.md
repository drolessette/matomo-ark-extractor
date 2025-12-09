# 📖 Comment obtenir l'exécutable Windows

Ce guide explique comment créer votre dépôt GitHub et obtenir l'exécutable `.exe` compilé automatiquement.

## Étape 1 : Créer un compte GitHub (si nécessaire)

1. Allez sur https://github.com
2. Cliquez "Sign up" et suivez les instructions

## Étape 2 : Créer un nouveau dépôt

1. Connectez-vous à GitHub
2. Cliquez le bouton vert **"New"** ou allez sur https://github.com/new
3. Remplissez :
   - **Repository name** : `matomo-ark-extractor`
   - **Description** : `Extraction des statistiques ARK depuis Matomo`
   - Cochez **Public**
   - ⚠️ Ne cochez PAS "Add a README file"
4. Cliquez **"Create repository"**

## Étape 3 : Uploader les fichiers

### Option A : Via l'interface web (le plus simple)

1. Sur la page de votre nouveau dépôt vide, cliquez **"uploading an existing file"**
2. Glissez-déposez TOUS les fichiers du dossier `matomo-ark-extractor` :
   - `app.py`
   - `requirements.txt`
   - `README.md`
   - `LICENSE`
   - `.gitignore`
   - `CHANGELOG.md`
   - Le dossier `.github` (avec son contenu)
3. En bas, tapez un message : `Initial commit`
4. Cliquez **"Commit changes"**

### Option B : Avec GitHub Desktop

1. Téléchargez GitHub Desktop : https://desktop.github.com
2. Connectez-vous avec votre compte
3. Clone votre dépôt vide
4. Copiez les fichiers dans le dossier cloné
5. Commit et Push

## Étape 4 : Vérifier la compilation

1. Allez dans l'onglet **"Actions"** de votre dépôt
2. Vous devriez voir un workflow en cours d'exécution (rond jaune)
3. Attendez qu'il devienne vert ✅ (environ 5-10 minutes)
4. Cliquez dessus pour voir les détails

## Étape 5 : Télécharger l'exécutable

### Méthode 1 : Depuis les Artifacts (sans release)

1. Dans **Actions**, cliquez sur le dernier workflow réussi
2. En bas de la page, section **"Artifacts"**
3. Cliquez sur **"MatomoARKExtractor-Windows"** pour télécharger
4. Dézippez et lancez `MatomoARKExtractor.exe`

### Méthode 2 : Créer une Release (recommandé)

Pour avoir un lien permanent et facile à partager :

1. Allez dans l'onglet **"Releases"** (colonne de droite)
2. Cliquez **"Create a new release"**
3. Cliquez **"Choose a tag"** et tapez `v1.0.0`
4. Cliquez **"Create new tag: v1.0.0 on publish"**
5. Titre : `Version 1.0.0`
6. Cliquez **"Publish release"**
7. Attendez que le workflow se termine
8. Rafraîchissez la page : l'exe apparaît dans les Assets !

## 🎉 C'est terminé !

Vous pouvez maintenant :
- Télécharger `MatomoARKExtractor.exe`
- Le copier sur n'importe quel PC Windows
- L'exécuter directement (aucune installation requise)

## 🔄 Mises à jour

Pour mettre à jour l'application :
1. Modifiez les fichiers sur GitHub
2. Créez un nouveau tag (ex: `v1.1.0`)
3. Une nouvelle release sera créée automatiquement

---

💡 **Astuce** : Partagez simplement le lien de votre page Releases avec vos collègues !
