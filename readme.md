# 🐞 EXCELARBUG

**EXCELARBUG** est un outil Windows simple et rapide permettant d’extraire automatiquement chaque feuille d’un fichier Excel en fichiers Excel séparés, tout en organisant les résultats dans une arborescence claire par utilisateur et par date.

---

## 🎯 Fonctionnalités

* Interface graphique simple (aucune ligne de commande)
* Extraction automatique de **toutes les feuilles Excel**
* Création d’un fichier `.xlsx` par feuille
* Classement automatique par :

  * Utilisateur
  * Date du jour
* Compatible Windows
* Version exécutable `.exe` disponible

---

## 📂 Structure des fichiers générés

Les fichiers extraits sont enregistrés dans :

```
Documents/bugExcel/{user}/{date}/fichier-excel-extraite/
```

### Exemple :

```
Documents/
└── bugExcel/
    └── nari/
        └── 2025-12-15/
            └── fichier-excel-extraite/
                ├── Feuille1.xlsx
                ├── Feuille2.xlsx
                └── Feuille3.xlsx
```

---

## ▶️ Utilisation

1. Lancer **EXCELARBUG.exe**
2. Cliquer sur **« Sélectionner le fichier Excel »**
3. Choisir un fichier `.xlsx`
4. Les feuilles sont automatiquement extraites et enregistrées

---

## 🧰 Technologies utilisées

* Python 3
* pandas
* openpyxl
* Tkinter (interface graphique)
* PyInstaller (génération du `.exe`)

---

## 🛠️ Développement (version Python)

### Prérequis

```bash
pip install pandas openpyxl
```

### Lancer le projet

```bash
python extract_excel.py
```

---

## 📦 Génération du fichier EXE

```bash
pyinstaller --onefile --windowed extract_excel.py
```

Le fichier final se trouve dans :

```
dist/extract_excel.exe
```

---

## ⚠️ Remarques

* Certains antivirus peuvent bloquer l’exécutable (comportement normal avec PyInstaller)
* Ajouter une exception si nécessaire
* Fonctionne uniquement avec des fichiers `.xlsx`

---

## 🚀 Évolutions possibles

* Support de plusieurs fichiers Excel
* Barre de progression
* Fichier log des extractions
* Installateur Windows (.msi / setup.exe)
* Support `.xls`

---

## 👤 Auteur

Projet interne – **NARIHY**
Outil de debug et d’audit Excel


