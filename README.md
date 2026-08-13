#  Tableau de Bord Excel : Gestion de Stock Dynamique et Analytique

![Excel](https://img.shields.io/badge/Microsoft_Excel-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white)
![Power Pivot](https://img.shields.io/badge/Power_Pivot-Green?style=for-the-badge)
![VBA](https://img.shields.io/badge/VBA-Automation-blue?style=for-the-badge)

##  Aperçu du Projet

Ce projet propose une solution clé en main pour transformer des données brutes de stock (entrées, sorties et inventaire) en un **tableau de bord interactif, visuel et décisionnel**. 

Il résout la problématique de dispersion des informations en centralisant le suivi des flux physiques de marchandises afin de mesurer en temps réel l'état des stocks. L'objectif commercial est d'offrir aux responsables logistiques et opérationnels une **visibilité immédiate sur les indicateurs clés de performance (KPIs)**, facilitant ainsi l'anticipation des approvisionnements et la prévention des ruptures.

---

###  Aperçu de l'interface

![Aperçu du Tableau de Bord](assets/dashboard_preview.png)
*Aperçu du tableau de bord de gestion de stock dynamique.*

---

##  Fonctionnalités Clés & Analyses

*  **Suivi des KPIs fondamentaux :** Affichage dynamique du stock total disponible (quantité), de la valeur financière globale, ainsi que des flux d'entrées (achats) et de sorties (ventes).
*  **Interactivité & Filtrage Croisé :** Segments (*slicers*) dynamiques interconnectés permettant de filtrer instantanément l'ensemble des indicateurs par **catégorie**, **référence de marchandise** ou **mois**.
*  **Visualisations Avancées :**
  * **Graphiques en secteur :** Répartition du stock et des valeurs par catégorie.
  * **Graphiques en barres personnalisés :** Volume mensuel des transactions.
  * **Graphiques combinés (aires dégradées) :** Évolution financière des flux d'entrées/sorties au fil du temps.
  * **Classement Top 3 (Barres horizontales) :** Identification automatique des marchandises les plus vendues.
*  **Alertes de Réapprovisionnement :** Module d'alerte visuelle identifiant immédiatement les articles en stock faible ou en rupture complète.
* **Ergonomie & Navigation :** Boutons de navigation vers les feuilles *Opérations* et *Inventaire*, accompagnés d'un bouton d'actualisation instantanée en un clic.

---

##  Compétences & Techniques Excel Utilisées

### 1. Nettoyage & Modélisation des Données
* **Préparation :** Extraction des valeurs uniques et suppression des doublons sous l'onglet *Données* pour créer des tables de dimensions propres.
* **Power Pivot & Modèle de Données :** Création de relations entres plusieurs tables sources (Achats/Entrées et Ventes/Sorties) via des clés (`Catégorie`, `Référence`, `Mois`) permettant d'appliquer un segment unique sur des TCD indépendants.

### 2. Logique & Formules
* **Croisement de données :** `RECHERCHEV` (`VLOOKUP`) pour l'importation fiable des données.
* **Logique décisionnelle :** Conditions `SI` (`IF`) imbriquées pour déterminer l'état dynamique des articles (*"Non disponible"* ou *"Stock faible"*).
* **Mise en valeur graphique :** Utilisation de la fonction `MAX` pour la gestion dynamique des éléments mis en évidence dans les visuels.
* **Dynamisme visuel :** Liaison directe des cartes KPI (formes géométriques / zones de texte) aux cellules de résultats des TCD.

### 3. Design & Interface Utilisateur (UI)
* **Format "Application" :** Masquage du quadrillage Excel et application d'un arrière-plan uni pour un rendu épuré.
* **Design sur-mesure :** Segments personnalisés en *Flat Design* (bordures et en-têtes masqués) et intégration d'icônes dédiées par KPI.

### 4. Automatisation VBA
* **Navigation :** Macros enregistrées associées à des boutons pour naviguer facilement entre les onglets.
* **Actualisation automatique :** Script VBA sur-mesure permettant de rafraîchir l'ensemble des TCD du classeur en un clic tout en mettant à jour la date de dernière modification (cellule `R1`).

---

##  Structure du Dépôt

```text
├── data/
│   ├── inventaire.xlsx          # Base de données de l'inventaire initial
│   ├── entrees_achats.xlsx      # Historique des flux d'entrées
│   └── sorties_ventes.xlsx      # Historique des flux de sorties
│
├── dashboard/
│   └── Tableau_de_Bord_Gestion_Stock.xlsx # Classeur final (Modèle de données, TCD, VBA, UI)
│
└── assets/
    ├── icons/                   # Pictogrammes des cartes KPIs
    └── dashboard_preview.png    # Capture d'écran du tableau de bord pour ce README
