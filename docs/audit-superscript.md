# Audit detaille — SuperScript/scripts/Superscript.jsx

**Date** : 2026-02-09
**Fichier audite** : `SuperScript/scripts/Superscript.jsx` (4003 lignes)
**Version** : 1.0 beta 10

---

## 1. Structure generale (lignes 1-14)

Le script est encapsule dans une IIFE `(function() { ... })()` avec `"use strict"` a la ligne 15. En ExtendScript (ES3), `"use strict"` est une simple expression string sans aucun effet. Ce n'est pas un bug mais c'est trompeur.

---

## 2. CONFIG (lignes 27-52)

- `APOSTROPHES_TO_REPLACE` (ligne 45) declare 4 variantes : `"'"`, `"&#39;"`, `"&apos;"`, `"&#8217;"` mais seule `"'"` est effectivement traitee dans `replaceApostrophes()` (ligne 1068). Les entites HTML ne seront jamais presentes dans un document InDesign. Cette liste est du **code mort**.

---

## 3. Utilities (lignes 131-466)

### BUG ES3 — `isStyleInList` (ligne 412)

`Array.isArray(styleList)` — methode ES5.1, **non disponible en ES3/ExtendScript**. Levera `TypeError` si `styleList` n'est pas un tableau.

**Correction** : remplacer par `Object.prototype.toString.call(styleList) === '[object Array]'`

### Code duplique — `collectParagraphStyles`

`getParagraphStyles` (ligne 241) et `UIBuilder.createDialog` (ligne 2042) contiennent la meme logique de collecte recursive des styles de paragraphe.

---

## 4. Corrections (lignes 472-1641)

### `moveNotes` (lignes 597-598) — inconsistance

`app.findGrepPreferences = null;` au lieu de `NothingEnum.nothing`. Fonctionne dans la plupart des versions mais inconsistant avec le reste du script.

### `fixValueRanges` Cas 8 (ligne 911) — pattern trop large

`([A-Z])-([A-Z])` matche tout tiret entre deux majuscules, incluant des faux positifs comme `A-B` en mathematiques, codes, acronymes.

### BUG — `fixDashIncises` (ligne 1268) — indexOf avec 3 arguments

```javascript
paraText.indexOf(ENDASH, 0, fermantPosition)
```

`String.prototype.indexOf` n'accepte que 2 arguments. Le troisieme argument est **silencieusement ignore**. L'intention est de limiter la recherche a `[0, fermantPosition]`, mais la recherche se fait sur toute la chaine. Les tirets fermants isoles ne sont **jamais traites**.

### `replaceApostrophes` (lignes 1052-1055) — NothingEnum en majuscules

`NothingEnum.NOTHING` au lieu de `NothingEnum.nothing`. Bug potentiel selon la version d'InDesign.

### `formatNumbers` (lignes 1291-1640) — multiples problemes

**Risque de collision de marqueurs** : `YEAR_MARK_`, `YEAR_ENDASH_`, `###`, `SPECIAL_NUMBER_` sont utilises comme marqueurs temporaires. Si le texte contient deja ces chaines, corruption.

**Boucles infinies potentielles** (lignes 1483, 1489, 1506, 1521) :
```javascript
while (doc.changeGrep().length > 0) {}
```
Si un remplacement produit un resultat qui matche a nouveau, boucle infinie.

**Modification de contenu pendant l'iteration** (ligne 1378) : modifier `largeNumbers[ln].contents` pendant l'iteration peut provoquer des decalages d'index.

---

## 5. UIBuilder (lignes 1713-2358)

### Double enregistrement onClick (lignes 1973-1976 et 1988-1991)

`exposantOrdinalStyleOpt.checkbox.onClick` est enregistre deux fois avec le meme code.

### Compteur de progression hardcode (ligne 2515)

`ProgressBar.update(12, "Termine !")` — utilise la valeur `12` au lieu du compteur `progress`.

---

## 6. SieclesModule (lignes 2604-3877)

### BUG CRITIQUE — `appliquerGrepPartout` utilise `app.activeDocument` (ligne 2987)

```javascript
var doc = app.activeDocument;
```

Au lieu d'utiliser le parametre `doc` passe a `formaterSiecles`/`formaterOrdinaux`/`formaterReferences`. Le mode "tous les documents" ne fonctionne pas pour le SieclesModule.

### BUG — `formaterOrdinaux` et `formaterReferences` utilisent `app.activeDocument`

Lignes 3408, 3414, 3419, 3429, 3527, 3557-3566, 3599, 3612, 3655, 3697-3699, 3819, 3843, 3857, 3871 : utilisent directement `app.activeDocument` au lieu du parametre `doc`.

### Code mort — `initializeUI` (ligne 2777) et `getOptions` (ligne 2890)

Ces methodes ne sont jamais appelees. Vestiges d'une ancienne architecture.

### Doublons dans MOTS_ORDINAUX

`"section"` (3 fois), `"trimestre"` (2 fois), `"classe"` (2 fois), `"phase"` (2 fois), `"division"` (2 fois).

### `savedPreferences` inefficace (lignes 3059-3064)

Sauvegarde des references aux objets singleton de preferences InDesign. Quand ces objets sont reinitialises, la "sauvegarde" pointe vers les memes objets modifies.

---

## 7. Organisation du code

### `applyMasterToLastPageStandalone` defini apres `main()` (ligne 3958 vs 3932)

La fonction est definie apres l'appel de `main()`. Fonctionne grace au hoisting JavaScript mais organisation confuse.

---

## Resume des bugs par gravite

### Critiques

| # | Ligne | Description |
|---|-------|-------------|
| 1 | 2987 | `appliquerGrepPartout` utilise `app.activeDocument` au lieu du `doc` parametre |
| 2 | 3408+ | `formaterOrdinaux` et `formaterReferences` utilisent directement `app.activeDocument` |
| 3 | 1268 | `indexOf` avec 3 arguments — tirets fermants isoles jamais traites |

### Majeurs

| # | Ligne | Description |
|---|-------|-------------|
| 4 | 412 | `Array.isArray` non disponible en ES3/ExtendScript |
| 5 | 1052 | `NothingEnum.NOTHING` (majuscule) au lieu de `NothingEnum.nothing` |
| 6 | 1483+ | Boucles `while (doc.changeGrep().length > 0)` — risque de boucle infinie |
| 7 | 2515 | Compteur de progression hardcode a `12` au lieu de `progress` |
| 8 | 911 | Pattern `([A-Z])-([A-Z])` trop large dans `fixValueRanges` |

### Mineurs

| # | Ligne | Description |
|---|-------|-------------|
| 9 | 1988 | Double enregistrement du onClick |
| 10 | 2042 | Duplication de `collectParagraphStyles` |
| 11 | 2777 | `SieclesModule.initializeUI` — code mort |
| 12 | 2890 | `SieclesModule.getOptions` — code mort |
| 13 | 45 | `APOSTROPHES_TO_REPLACE` — config inutilisee |
| 14 | 3059 | Sauvegarde de preferences inefficace |
| 15 | 597 | `null` au lieu de `NothingEnum.nothing` |
| 16 | 3288 | Conversion en lowercase des chiffres romains |
