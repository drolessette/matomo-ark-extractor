# 📚 Matomo ARK Extractor

[![Build Windows EXE](https://github.com/YOUR_USERNAME/matomo-ark-extractor/actions/workflows/build.yml/badge.svg)](https://github.com/YOUR_USERNAME/matomo-ark-extractor/actions/workflows/build.yml)
[![Release](https://img.shields.io/github/v/release/YOUR_USERNAME/matomo-ark-extractor)](https://github.com/YOUR_USERNAME/matomo-ark-extractor/releases)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](https://opensource.org/licenses/MIT)

**Application Windows pour extraire les statistiques de consultation des ressources ARK depuis les exports XML Matomo.**

Conçu pour les Bibliothèques spécialisées de la Ville de Paris.

![Screenshot](screenshot.png)

---

## ✨ Fonctionnalités

- 📂 **Import XML Matomo** : Sélection simple du fichier d'export
- 📊 **Extraction complète** : Visites, pages vues, temps de consultation, taux de rebond...
- 🔍 **Récupération métadonnées** : Titre, auteur, type de document (optionnel)
- 📈 **Export Excel formaté** : Fichier horodaté avec tableaux, filtres, résumé et Top 20
- 🎨 **Interface moderne** : Design sombre élégant avec CustomTkinter
- ⚡ **Portable** : Exécutable Windows autonome, aucune installation requise

---

## 📥 Téléchargement

### 👉 [Télécharger la dernière version (Windows .exe)](../../releases/latest)

1. Téléchargez `MatomoARKExtractor.exe` depuis les [Releases](../../releases)
2. Double-cliquez pour lancer l'application
3. Aucune installation nécessaire !

---

## 🚀 Utilisation

1. **Lancez** `MatomoARKExtractor.exe`
2. **Cliquez** sur "Parcourir" pour sélectionner votre fichier XML Matomo
3. **Cochez/décochez** l'option de récupération des métadonnées
4. **Cliquez** sur "Extraire et générer l'Excel"
5. **Le fichier Excel** est créé dans le même dossier que le XML

### Format du fichier XML

Le fichier doit être un export XML de Matomo contenant des URLs avec des identifiants ARK :
```
https://bibliotheques-specialisees.paris.fr/ark:/73873/pf0000856602
```

### Fichier Excel généré

Le fichier `stats_matomo_ark_YYYYMMDD_HHMMSS.xlsx` contient :

| Feuille | Contenu |
|---------|---------|
| **Statistiques ARK** | Tableau complet avec toutes les métriques |
| **Résumé** | Statistiques globales et par type |
| **Top 20** | Classement des ressources les plus consultées |

---

## 🛠️ Compilation depuis les sources

### Prérequis
- Python 3.10+
- Git

### Instructions

```bash
# Cloner le dépôt
git clone https://github.com/YOUR_USERNAME/matomo-ark-extractor.git
cd matomo-ark-extractor

# Installer les dépendances
pip install -r requirements.txt

# Lancer l'application
python app.py

# Compiler en .exe (optionnel)
pip install pyinstaller
pyinstaller --onefile --windowed --icon=icon.ico --name=MatomoARKExtractor app.py
```

---

## 🔧 GitHub Actions

L'exécutable Windows est compilé automatiquement via GitHub Actions :

- **À chaque push sur `main`** : Build de test
- **À chaque tag `v*`** : Création d'une Release avec l'exe

Pour créer une nouvelle release :
```bash
git tag v1.0.0
git push origin v1.0.0
```

---

## 📁 Structure du projet

```
matomo-ark-extractor/
├── app.py                    # Application principale
├── requirements.txt          # Dépendances Python
├── README.md                 # Documentation
├── LICENSE                   # Licence MIT
├── icon.ico                  # Icône de l'application
├── .gitignore               # Fichiers ignorés
└── .github/
    └── workflows/
        └── build.yml        # CI/CD GitHub Actions
```

---

## 📋 Données extraites

### Métriques Matomo
- Nombre de visites
- Visiteurs uniques
- Pages vues (hits)
- Temps total passé
- Temps moyen par page
- Taux de rebond
- Taux de sortie

### Métadonnées (si option activée)
- Titre du document
- Auteur
- Type de ressource (Fonds iconographique, Notice bibliographique...)

---

## ⚠️ Notes

- Le scraping des métadonnées dépend de la disponibilité du site
- Certains sites utilisent JavaScript dynamique, les métadonnées peuvent être incomplètes
- Pour des métadonnées complètes, préférez une extraction directe depuis la base Portfolio

---

## 🤝 Contribution

Les contributions sont les bienvenues !

1. Fork le projet
2. Créez une branche (`git checkout -b feature/amelioration`)
3. Committez (`git commit -m 'Ajout fonctionnalité'`)
4. Push (`git push origin feature/amelioration`)
5. Ouvrez une Pull Request

---

## 📄 Licence

MIT License - voir [LICENSE](LICENSE)

---

## 👥 Auteurs

- **CCPID** - Centre de Coordination des Projets en Informatique Documentaire
- **Bibliothèques de la Ville de Paris**

---

<p align="center">
  <i>Fait avec ❤️ pour les bibliothèques parisiennes</i>
</p>
