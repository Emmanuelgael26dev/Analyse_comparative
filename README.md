# Analyse comparative de marché

Comparaison des prix de trois enseignes d'Abidjan (Marché Lagune, Baobab Shop, Comptoir Ivoire) pour aider un client à décider où acheter de l'électronique et de l'électroménager.

> **Projet académique** : DA Python, L2 Génie logiciel, séance 3.
> **Auteur** : Kouassi Emmanuel Gaël.
> Toutes les données sont **fictives** : les enseignes et les prix n'existent pas.

## Objectif

Le programme collecte trois sources de formats différents, les rapproche d'un catalogue commun, compare les prix réellement payés et exploite un historique de plus de 20 000 lignes.

| Enseigne | Type | Fichier | Format |
|---|---|---|---|
| Marché Lagune | site marchand | `marche_lagune.csv` | CSV, séparateur `;`, export « Excel français » |
| Baobab Shop | application mobile | `baobab_shop.json` | JSON renvoyé par l'API |
| Comptoir Ivoire | boutique d'Adjamé | `comptoir_ivoire.xlsx` | classeur Excel |

Le **prix payé** se calcule différemment selon l'enseigne :

- **Marché Lagune** : le prix affiché, livraison incluse.
- **Baobab Shop** : `amount` + `delivery_fee`.
- **Comptoir Ivoire** : prix catalogue diminué de la remise (retrait en boutique).

## Les 8 étapes réalisées

1. **Chargement** : les trois fichiers (CSV, JSON, Excel) sont lus dans une même structure de données (classe `Offre`).
2. **Nettoyage** : les anomalies sont repérées, traitées et listées dans un journal (`sorties/journal_nettoyage.txt`).
3. **Rapprochement** : chaque article est associé à un `produit_id` du catalogue (normalisation des libellés, synonymes, tolérance aux fautes de frappe, rejet des contradictions comme 128 Go / 256 Go).
4. **Comparatif** : pour chaque produit présent dans au moins deux enseignes, prix payé, enseigne la moins chère, écart en FCFA et en %. Export dans `sorties/comparatif_prix.csv`.
5. **Synthèse** : nombre de fois où chaque enseigne est la moins chère, écart moyen par rapport au prix le plus bas, trois produits où il y a le plus à gagner.
6. **Panier de rentrée** : coût dans chaque enseigne, puis en achetant chaque article au prix le plus bas.
7. **Historique** : prix de la semaine du Black Friday (24 au 30 novembre 2025) comparés au prix moyen de l'année.
8. **Complexité** : temps d'exécution mesuré et justification des structures de données.

Un article en rupture de stock ou sans prix exploitable n'est jamais retenu comme « le moins cher ».

## Structure du projet

```
.
├── Analyse_comparative.py       # programme principal
├── requirements.txt             # dépendances
├── README.md
├── donnees/                     # données fournies pour l'exercice
│   ├── catalogue_reference.csv
│   ├── marche_lagune.csv
│   ├── baobab_shop.json
│   ├── comptoir_ivoire.xlsx
│   └── historique_prix.csv
└── sorties/                     # créé automatiquement à l'exécution
    ├── comparatif_prix.csv
    └── journal_nettoyage.txt
```

Les fichiers de données peuvent aussi être placés directement à côté du script, sans dossier `donnees`.

## Prérequis

- Python 3.8 ou supérieur (testé avec Python 3.12 et 3.14).
- La bibliothèque `openpyxl` pour lire le fichier Excel. Le reste du programme utilise uniquement la bibliothèque standard de Python.

## Installation

```bash
# 1. Cloner le dépôt
git clone https://github.com/<Emmanuelgael26dev>/<Analyse_comparative>.git
cd <Analyse_comparative>

# 2. (Recommandé) créer un environnement virtuel
python -m venv venv

# Windows (PowerShell)
.\venv\Scripts\Activate.ps1
# Linux / macOS
source venv/bin/activate

# 3. Installer les dépendances
pip install -r requirements.txt
```

## Exécution

```bash
python Analyse_comparative.py
```

Les résultats des étapes 1 à 8 s'affichent dans le terminal. Le fichier `sorties/comparatif_prix.csv` (séparateur `;`, ouvrable dans Excel) et le journal `sorties/journal_nettoyage.txt` sont créés dans le dossier `sorties/`.

## Résultats obtenus

**Chargement et rapprochement** : 28 articles (Marché Lagune), 26 (Baobab Shop) et 22 (Comptoir Ivoire), tous rapprochés d'un `produit_id`. 23 produits sont disponibles dans au moins deux enseignes.

**Nombre de fois où l'enseigne est la moins chère** (un ex æquo compte pour chaque enseigne) :

| Enseigne | Fois la moins chère | Écart moyen / prix le plus bas |
|---|---|---|
| Marché Lagune | 6 | 7 771 FCFA (3,8 %) |
| Baobab Shop | 9 | 2 702 FCFA (4,7 %) |
| Comptoir Ivoire | 10 | 6 667 FCFA (2,8 %) |

**Trois produits où il y a le plus à gagner** : HP 250 G9 (52 800 FCFA), iPhone 13 128 Go (26 500 FCFA), Climatiseur LG 1 CV (26 000 FCFA).

**Panier de rentrée** :

| Marché Lagune | Baobab Shop | Comptoir Ivoire | Au plus bas par article |
|---|---|---|---|
| 504 300 FCFA | 491 500 FCFA | 484 075 FCFA | 477 575 FCFA |

**Black Friday** (prix de la semaine comparé à la moyenne annuelle) : Marché Lagune −11,3 %, Baobab Shop −11,7 %, Comptoir Ivoire +3,8 %. Chez Comptoir Ivoire, le Black Friday n'est donc pas le bon moment pour acheter.

## Principaux problèmes de données traités

- **Marché Lagune** : BOM UTF-8, deux lignes en doublon exact (`ML-1000`, `ML-1147`), prix écrits en texte (`102 900 FCFA`), « Prix sur demande » et prix vide, articles en rupture.
- **Prix aberrant** : le Tecno Spark 20 est affiché à 779 000 FCFA chez Marché Lagune. Comparé à l'historique récent (environ 78 000 FCFA), un zéro en trop est probable : la valeur retenue est 77 900 FCFA, et la correction est consignée dans le journal.
- **Baobab Shop** : prix `null`, stock à 0, titres en anglais et parfois sans marque (`Electric Kettle 1.7L`).
- **Comptoir Ivoire** : ligne de titre avant les en-têtes, prix enregistré en texte avec une espace insécable, faute de frappe (`Samsun`), noms abrégés (`Frigo`, `Clim`, `DD externe`).

## Choix de structures de données (complexité)

| Besoin | Structure | Intérêt |
|---|---|---|
| Catalogue, offres par (enseigne, produit) | `dict` | accès en O(1) |
| Détection des doublons | `set` / `dict` | test d'appartenance en O(1) |
| Rapprochement des libellés | index inversé (`dict` mot → produits) | on ne compare qu'aux produits qui partagent un mot |
| Historique de 20 876 lignes | `dict` (enseigne, produit) → listes triées par date + `bisect` | recherche d'une période en O(log n) |
| Comptage des victoires | `Counter` | un seul passage |
| Trois plus gros écarts | `heapq.nlargest` | O(n log 3) au lieu d'un tri complet |

Sur 300 recherches de prix dans l'historique, la méthode dictionnaire + `bisect` est environ 700 fois plus rapide qu'un parcours de liste complet (de l'ordre de 0,2 ms contre 150 ms). Les temps varient selon la machine.

## Hypothèses

- Les frais de livraison de Baobab Shop sont comptés par article, donc par unité achetée.
- Chez Comptoir Ivoire, les frais de déplacement vers la boutique ne sont pas comptés.
- Pour l'historique, chaque produit est comparé à sa propre moyenne annuelle, puis la moyenne des écarts est calculée par enseigne : les historiques ne sont pas homogènes (livraison incluse chez Marché Lagune, hors livraison chez Baobab Shop, hors remise chez Comptoir Ivoire).
