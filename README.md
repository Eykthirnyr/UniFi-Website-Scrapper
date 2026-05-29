# Ubiquiti EU Store Scraper

Scraper headless Python du store Ubiquiti Europe ([eu.store.ui.com](https://eu.store.ui.com)).  
Balaye **17 catégories**, gère toutes les variantes produit et exporte un fichier **XLSX professionnel** (2 feuilles) + un CSV de secours.

---

## Fonctionnalités

- **17 catégories** couvertes : Cloud Gateways, Switching, WiFi, Caméras, Door Access, Storage, Câbles, SFP, PoE...
- **Variantes multi-dimensionnelles** : longueurs de câbles, couleurs (swatches), capacités HDD (8 / 16 / 24 TB), formats Indoor/Outdoor… chaque combinaison = une ligne distincte
- **Déduplication stricte** par URL (y compris `?variant=` pour les câbles)
- **Prix EU** interprétés correctement : `2.000,00 €` → `2000.00` (virgule décimale, point milliers, `&nbsp;`)
- **Double layout HTML** géré : cartes produit principales (`sc-13lrnwl-*`) et cartes accessoires (`sc-bw6p3d-*`)
- **JSON-LD** exploité en priorité pour SKU / nom / prix (source la plus fiable)
- Export **XLSX** formaté : en-têtes colorés, lignes alternées, codes couleur disponibilité, hyperliens, filtres automatiques, volets figés, feuille Résumé par catégorie
- Export **CSV** de secours (UTF-8 BOM, compatible Excel)
- Gestion des **PermissionError** Excel : sauvegarde horodatée si le fichier est ouvert
- Mode `--debug` : sauvegarde un snapshot HTML par catégorie pour diagnostiquer les sélecteurs

---

## Prérequis

| Dépendance | Version testée |
|---|---|
| Python | ≥ 3.9 |
| [playwright](https://playwright.dev/python/) | ≥ 1.40 |
| [openpyxl](https://openpyxl.readthedocs.io/) | ≥ 3.1 |

```bash
pip install playwright openpyxl
playwright install chromium
```

---

## Utilisation

```bash
# Lancement standard (génère ubiquiti_products.xlsx dans le même dossier)
python ubiquiti_scraper.py

# Chemin de sortie personnalisé
python ubiquiti_scraper.py --output D:\exports\produits.xlsx

# Sans variantes — plus rapide, une seule ligne par fiche produit
python ubiquiti_scraper.py --no-variants

# Mode debug — sauvegarde un snapshot HTML par catégorie
python ubiquiti_scraper.py --debug
```

---

## Sorties

### `ubiquiti_products.xlsx`

**Feuille 1 — Catalogue**

| Colonne | Description |
|---|---|
| Code Produit | SKU Ubiquiti (ex. `U6-PLUS`, `UACC-HDD-E-16TB`) |
| Nom Commercial | Nom affiché dans le store |
| Variante | Label de la variante cliquée (longueur, couleur, capacité…) |
| Catégorie | Catégorie store d'origine |
| Prix (EUR) | Prix HT en euros (format numérique `#,##0.00 "€"`) |
| Description | Description courte extraite de la carte produit |
| Disponible | `Oui` (vert) / `Non` (rouge) |
| Sold Out | `Oui` (rouge) / `Non` (vert) |
| Lien Produit | Hyperlien cliquable → fiche produit EU store |

**Feuille 2 — Résumé** : tableau récapitulatif par catégorie (total articles, disponibles, sold-out, prix min/max).

### `ubiquiti_products.csv`

Même données, encodage UTF-8 BOM, séparateur virgule. Compatible Excel (double-clic direct).

---

## Architecture technique

```
ubiquiti_scraper.py
│
├── Phase 1 — Pages catégorie
│   ├── Playwright headless Chromium (bloque images/polices/médias)
│   ├── Scroll progressif (lazy-load) + clic "Load More"
│   ├── _JS_CATEGORY_CARDS  : extrait les cartes via DOM (2 layouts)
│   └── Marque les fiches avec variantes pour la Phase 2
│
├── Phase 2 — Fiches produit + variantes
│   ├── _JS_PRODUCT_DATA    : JSON-LD + DOM fallbacks (SKU, prix, description)
│   ├── Clics sur boutons variante (.sc-14eyr3g-0 et .sc-1uupl4m-6 swatches)
│   ├── Détection URL : path changé (HDD) vs ?variant= seulement (câbles)
│   └── BFS + ensembles `queued`/`processed` (anti-boucle)
│
├── write_xlsx()            : export XLSX formaté (openpyxl)
└── write_csv()             : export CSV de secours
```

### Sélecteurs CSS clés (mai 2025)

| Sélecteur | Rôle |
|---|---|
| `a.sc-13lrnwl-14` | Carte produit principale (WiFi, Switch, GW, Storage…) |
| `a.sc-bw6p3d-7` | Carte produit accessoires / câbles |
| `.sc-13lrnwl-4` | Nom commercial (dans la carte) |
| `.sc-13lrnwl-5` | SKU / code produit (dans la carte) |
| `.sc-13lrnwl-23` | Prix HT — WiFi, Switch, GW, HDD |
| `.sc-13lrnwl-19` | Prix HT — câbles et accessoires |
| `.sc-bw6p3d-16` | Prix HT — fallback accessoires Layout 2 |
| `.sc-y2swsu-4` | Prix HT — fiche produit détaillée |
| `.sc-14eyr3g-0 button[title]` | Boutons variante (type, longueur, capacité) |
| `.sc-1uupl4m-6 button[title]` | Swatches couleur (fiche produit) |
| `.sc-1cuaoi0-0` | Badge "Sold Out" |
| `button[label]` | Bouton d'action (`Add to Cart` / `Select` / `Sold Out`) |

> **Note** : les classes `sc-*` sont générées par styled-components. Le préfixe (ex. `sc-13lrnwl`) est stable par composant ; le suffixe numérique peut changer lors de mises à jour du store.

---

## Catégories couvertes

| Catégorie | URL |
|---|---|
| Cloud Gateways | `/category/all-cloud-gateways` |
| Switching | `/category/all-switching` |
| WiFi | `/category/all-wifi` |
| Cameras & NVR | `/category/all-cameras-nvrs` |
| Door Access | `/category/all-door-access` |
| Integrations | `/category/all-integrations` |
| Advanced Hosting | `/category/all-advanced-hosting` |
| Acc. Cables & DACs | `/category/accessories-cables-dacs` |
| Acc. Modules Fiber | `/category/accessories-modules-fiber` |
| Acc. SFP | `/category/accessories-sfp-liberation-day` |
| Acc. Storage | `/category/accessories-storage` |
| Acc. Rack Mount | `/category/accessories-rack-mount` |
| Acc. PoE & Power | `/category/accessories-poe-power` |
| Acc. Access Point | `/category/accessories-access-point` |
| Acc. Camera | `/category/accessories-camera` |
| Acc. Door Access | `/category/accessories-door-access` |
| Acc. Installations | `/category/accessories-installations` |

---

## Dépannage

| Symptôme | Cause probable | Solution |
|---|---|---|
| `0 cartes` sur une catégorie | Sélecteur CSS obsolète (mise à jour store) | Lancer avec `--debug` et inspecter le snapshot HTML |
| Prix `—` dans l'Excel | Nouveau composant price sur le store | Identifier la classe `sc-*` dans le DOM et l'ajouter à `_JS_CATEGORY_CARDS` / `_JS_PRODUCT_DATA` |
| `PermissionError` à la sauvegarde | Fichier XLSX ouvert dans Excel | Fermer Excel ou utiliser le fichier horodaté créé automatiquement |
| Timeout sur une page | Réseau lent ou bot-detection | Augmenter les timeouts dans `page.goto()` ou ajouter un `time.sleep()` |
| Variantes manquantes | Nouveau composant swatch | Vérifier si un sélecteur `.sc-*-6 button[title]` est apparu dans le DOM |

---

## Avertissement légal

Ce scraper est destiné à un usage **interne / professionnel** (veille tarifaire, catalogue interne).  
Consultez les [Conditions d'utilisation de Ubiquiti](https://www.ui.com/legal/termsofservice/) avant tout usage.  
L'accès intensif peut déclencher des protections anti-bot ; utilisez `--no-variants` pour une exécution plus légère.

---

*Scraper développé et testé sur eu.store.ui.com — mai 2025.*
