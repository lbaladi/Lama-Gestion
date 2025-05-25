# Projet VBA/SQL LAMA-GESTION

## Introduction

Dans le cadre de notre stage chez **LAMA-GESTION**, une boutique spécialisée en gestion de fonds, nous avons développé un outil d'automatisation du reporting financier.  
Ce projet s'appuie sur **Excel VBA**, interconnecté avec **Microsoft Access via SQL**, et permet de :

- Centraliser les données des clients et des fonds dans une base Access cohérente et sécurisée ;
- Générer automatiquement des reportings dynamiques (par fonds, client, gérant, boutique, région…).

Ce document détaille les fonctionnalités, l’architecture technique ainsi que les instructions d’utilisation de notre outil.

---

## Prérequis techniques

Pour utiliser l’outil, il faut disposer des éléments suivants :

- Microsoft Excel (avec macros activées – `.xlsm`)
- Microsoft Access
- Références VBA activées :
  - Microsoft ActiveX Data Objects
  - Microsoft Access Object Library
- Microsoft Outlook installé pour l’envoi automatique des reportings

---

## Structure du projet

Voici une brève description de l’architecture des fichiers du projet :

- `data/` : contient les données de base (listings marchés, données géographiques, indices, etc.)
- `data_fonds/` : fichiers Excel liés à chaque fonds géré (Alpha, Omega, etc.)
- `release/` : application principale à ouvrir (`Projet Lama-Gestion.xlsm`)
- `reporting/` : rapports PDF générés par fonds et pour les clients
- `src/` : code source VBA (modules, formulaires utilisateurs)
- `Notice Utilisateur.pdf` : guide d’utilisation de l’application
- `logo.jpg` : logo du projet

---

## Fonctionnalités principales

La feuille macro du projet contient plusieurs boutons, chacun étant lié à une macro spécifique.  
L’ensemble des macros est réparti en 3 modules : `DataBse`, `Manipulation_dataBase` et `Reporting`.

### 1. Création et modification de la base de données Access

- **Bouton** : `Création DATABASE`  
  **Module** : `DataBse`  
  **Macro** : `create_database()`  

  > Crée la base de données Access à partir des fichiers fournis. Étape cruciale à effectuer en premier (sauf "Rapprochement Listing").

- **Bouton** : `Modification DATABASE`  
  **Module** : `DataBse`  
  **Macro** : `modifDB()`  

  > Modifie la structure de la base pour permettre une meilleure manipulation.  
  La base obtenue contient 8 tables :

  - `Parts_actifs` : actifs et leur poids dans chaque fonds.
  - `Pilotage_investisseurs` : infos clients et montants investis.
  - `Pilotage_fonds` : infos sur chaque fonds.
  - `Rendements_actifs1/2/3/4` : rendements quotidiens des titres (2019-2024).
  - `Rdts_actifs` : rendements mensuels dérivés des 4 tables précédentes.

---

### 2. Gestion des clients et investissements

- **Boutons** : `Nouveau client`, `Supprimer client`, `Nouvel investissement`, `Supprimer investissement`  
  **Module** : `Manipulation_dataBase`  
  **Macros** : `NewClient`, `SupClient`, `NewInvest`, `SupInvest`  

  > Les macros affichent un **UserForm** puis appellent des procédures (`AddProc`, `DeleteProc`, `DepotProc`, `RetraitProc`) selon l’opération choisie.  
  Les données sont modifiées dans `pilotage_investisseurs` et `pilotage_fonds`.

---

### 3. Rapprochement des listings Nasdaq

- **Bouton** : `Rapprochement listing`  
  **Module** : `Reporting`  
  **Macro** : `macro_rappro()`  

  > Automatisation du rapprochement entre anciens et nouveaux listings NASDAQ.  
  Résultat affiché dans une feuille Excel récapitulative.

---

### 4. Génération de reportings dynamiques

Une interface **UserForm** permet :

- Génération automatique du reporting
- Exportation PDF
- Envoi via Outlook (`projetLGLM@outlook.fr` par défaut)

> Pour modifier l’adresse de réception, changer la variable `adresse`.

#### - Par boutique

- **Bouton** : `Reporting Boutique`  
  **Module** : `Reporting`  
  **Macro** : `reportingBoutique()`  

  > Génère un reporting global : actifs sous gestion, infos par fonds, récapitulatif client.

#### - Par fonds

- **Bouton** : `Reporting Fonds`  
  **Module** : `Reporting`  
  **Macro** : `reportingFonds()`  

  > L’utilisateur choisit un fonds via un UserForm → génération + envoi du reporting.

#### - Meilleurs clients

- **Bouton** : `Reporting meilleurs clients`  
  **Module** : `Reporting`  
  **Macro** : `reportingTopClient()`  

  > Focus sur les 10 meilleurs clients (montants déposés, infos personnelles, répartitions par fonds).

#### - Par client

- **Bouton** : `Reporting Clients`  
  **Module** : `Reporting`  
  **Macro** : `reportingClient()`  

  > Sélection d’un ou plusieurs clients via une ListBox → génération + envoi des reportings.

---

## Utilisation des UserForms

Tous les **UserForms** fonctionnent sur le même modèle :  
Remplir les champs → cliquer sur **Valider** ou **Annuler** pour quitter.

### 1. `UserFormNewC` – Ajout d’un client

- Infos d’identité à gauche
- Montants et parts allouées à droite (en décimal : `0.5` pour moitié)
- Ne pas remplir une TextBox = pas d’investissement sur ce fonds

### 2. `UserFormSupC` – Suppression d’un client

- Saisie du nom et prénom dans 2 champs

### 3. `UserFormDeposit` / `UserFormWithdraw` – Ajout / retrait de fonds

- Même structure que `UserFormNewC`
- Ajouter : somme déposée + répartition
- Retirer : fonctionnement identique mais soustraction

### 4. `UserFormRpClient` – Reporting client

- Présente une **ListBox** avec tous les clients
- Sélection simple ou multiple → envoi des reportings

### 5. `UserForm1` – Reporting fonds

- Présente une **ComboBox** des fonds
- Saisie directe ou sélection via liste déroulante

---
