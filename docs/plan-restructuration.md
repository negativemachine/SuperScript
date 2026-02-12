# Plan de restructuration complète de SuperScript

## Contexte

SuperScript est actuellement un script monolithique FR-only (~3940 lignes). L'objectif est de le restructurer pour :
1. **Supporter 7 profils linguistiques** (FR-FR, FR-CH, EN-US, EN-UK, DE, ES, IT) via des fichiers JSON externes dans `dictionary/`
2. **Internationaliser l'UI** (FR/EN) via un module I18n
3. **Sauvegarder/charger les préférences** utilisateur via un fichier `superscript-config.json`
4. **Préserver 100% des traitements existants** — aucune régression fonctionnelle

La spec typographique de référence est `docs/typography/guide_micro-typo_multi.md`.

---

## Fichiers modifiés

- **`SuperScript/src/Superscript.jsx`** — script principal (~3940 lignes), toutes les modifications
- **`SuperScript/dictionary/lang-fr-FR.json`** — nouveau, profil linguistique français France
- **`SuperScript/dictionary/lang-fr-CH.json`** — nouveau, profil linguistique français Suisse
- **`SuperScript/dictionary/lang-en-US.json`** — nouveau, profil linguistique anglais US
- **`SuperScript/dictionary/lang-en-UK.json`** — nouveau, profil linguistique anglais UK
- **`SuperScript/dictionary/lang-de.json`** — nouveau, profil linguistique allemand
- **`SuperScript/dictionary/lang-es.json`** — nouveau, profil linguistique espagnol
- **`SuperScript/dictionary/lang-it.json`** — nouveau, profil linguistique italien

## Patterns réutilisés (depuis Markdown-Import)

- **`safeJSON`** (Markdown-Import:340-402) — stringify/parse ES3-compatible
- **`I18n` IIFE** (Markdown-Import:26-334) — pattern translation dict + `__()` + `detectInDesignLanguage()`
- **`saveConfiguration()`** (Markdown-Import:1254-1323) — sérialisation style `.name`, `saveDlg()`
- **`loadConfiguration()`** (Markdown-Import:1330-1415) — `findStyleByName()` reverse mapping
- **`autoLoadConfig()`** (Markdown-Import:1098-1193) — recherche récursive `superscript-config.json`

---

## Phase 1 : safeJSON + I18n (fondations, sans changement de comportement)

### 1a. Ajouter le module `safeJSON`
- Insérer après la fermeture de `CONFIG` (après ligne ~52), avant `ErrorHandler`
- Port exact de Markdown-Import:340-397 (stringify récursif + parse avec validation regex + Function constructor)
- Ajouter le guard `if (typeof JSON === 'undefined') { JSON = safeJSON; }`

### 1b. Ajouter le module `I18n`
- Insérer après `safeJSON`, avant `ErrorHandler`
- IIFE identique au pattern Markdown-Import avec :
  - `detectInDesignLanguage()` via `app.locale`
  - `__()` avec substitution `%s`/`%d`
  - `setLanguage()` / `getLanguage()`
  - Dictionnaires FR/EN couvrant **toutes** les ~60 chaînes UI hardcodées actuelles :
    - Onglets : "Corrections", "Espaces et retours", "Styles", "Formatages", "Mise en page"
    - Checkboxes : toutes les 25+ options dans UIBuilder
    - Panels : "Définition des styles", "Application des styles", etc.
    - Messages : "Corrections appliquées...", "Erreur...", alertes, progress bar texts
    - Boutons : "Appliquer", "Annuler"

### 1c. Migrer UIBuilder vers I18n
- Remplacer tous les strings littéraux FR dans UIBuilder par des appels `I18n.__(key)`
- Remplacer les messages dans `Processor.processDocuments()` (alertes)
- Remplacer les textes de la ProgressBar dans `applyCorrections()`
- Le comportement reste **identique** — la détection auto sélectionne FR si InDesign est en FR

**Commit**: `feat(superscript): add safeJSON and I18n modules, internationalize UI`

---

## Phase 2 : LanguageProfile + fichiers `dictionary/`

### 2a. Créer `SuperScript/dictionary/lang-fr-FR.json`
Extraire toutes les données FR-spécifiques de `SieclesModule.CONFIG` et `Corrections` :
```json
{
  "meta": {
    "id": "fr-FR",
    "label": "Français (France)",
    "labelEN": "French (France)"
  },
  "punctuation": {
    "spaceBeforeSemicolon": "~<",
    "spaceBeforeColon": "~S",
    "spaceBeforeExclamation": "~<",
    "spaceBeforeQuestion": "~<",
    "spaceInsideOpenQuote": "~S",
    "spaceInsideCloseQuote": "~S",
    "spaceBeforePercent": "~S"
  },
  "dashes": {
    "inciseDash": "~_",
    "inciseSpace": "~S",
    "replaceCadratinWithDemiCadratin": true,
    "rangeIntervalDash": "~="
  },
  "quotes": {
    "level1Open": "\u00AB",
    "level1Close": "\u00BB",
    "level2Open": "\u201C",
    "level2Close": "\u201D"
  },
  "centuries": {
    "enabled": true,
    "useRomanNumerals": true,
    "useSuperscript": true,
    "useSmallCaps": true,
    "suffixes": ["e", "er"],
    "wordForCentury": "si\u00E8cle",
    "wordForCenturyPlural": "si\u00E8cles",
    "abbreviation": "s."
  },
  "ordinals": {
    "enabled": true,
    "useSuperscript": true,
    "corrections": {
      "1\u00E8re": "1re",
      "1\u00E8res": "1res",
      "2\u00E8me": "2e",
      "3i\u00E8me": "3e"
    },
    "titleAbbreviations": [
      {"abbr": "Me", "super": "e"},
      {"abbr": "Mgr", "super": "gr"},
      {"abbr": "Dr", "super": "r"},
      {"abbr": "Mme", "super": "me"},
      {"abbr": "Mlle", "super": "lle"}
    ]
  },
  "numbers": {
    "thousandsSeparator": "~<",
    "decimalSeparator": ",",
    "replacePointWithComma": true,
    "addThousandsSpaces": true,
    "excludeYears": true,
    "yearRange": [0, 2050]
  },
  "data": {
    "motsAmbigus": ["vie", "ive", "xie", "ie"],
    "motsOrdinaux": ["...tous les mots actuels de SieclesModule.CONFIG.MOTS_ORDINAUX..."],
    "motsAvantOrdinaux": ["...tous les mots actuels de SieclesModule.CONFIG.MOTS_AVANT_ORDINAUX..."],
    "motsOeuvres": ["...tous les mots actuels de SieclesModule.CONFIG.MOTS_OEUVRES..."],
    "titresPersonnes": ["...tous les mots actuels de SieclesModule.CONFIG.TITRES_PERSONNES..."],
    "nomsPremier": ["...tous les mots actuels de SieclesModule.CONFIG.NOMS_PREMIER..."],
    "abreviationsRefs": ["...tous les mots actuels de SieclesModule.CONFIG.ABREVIATIONS_REFS..."],
    "abreviationsVolumes": ["...tous les mots actuels de SieclesModule.CONFIG.ABREVIATIONS_VOLUMES..."],
    "abreviationsTemporelles": ["...tous les mots actuels de SieclesModule.CONFIG.ABREVIATIONS_TEMPORELLES..."],
    "abreviationsNumeros": ["...tous les mots actuels de SieclesModule.CONFIG.ABREVIATIONS_NUMEROS..."],
    "unitesMesure": ["...tous les mots actuels de SieclesModule.CONFIG.UNITES_MESURE..."],
    "abreviationsDirection": ["...tous les mots actuels de SieclesModule.CONFIG.ABREVIATIONS_DIRECTION..."],
    "titresAppellations": ["...tous les mots actuels de SieclesModule.CONFIG.TITRES_APPELLATIONS..."]
  },
  "italicExpressions": [
    "ad libitum", "ad valorem", "alma mater", "alter ego", "..."
  ],
  "hardHyphens": {
    "apostrophe": "apos~-trophe",
    "aujourd'hui": "aujour~-d'hui",
    "..."
  }
}
```

### 2b. Créer les autres fichiers `dictionary/lang-*.json`
- **lang-fr-CH.json** : copie de fr-FR avec `spaceBeforeColon: "~<"`, `inciseSpace: " "` (espace normale, sécable), `spaceInsideOpenQuote/CloseQuote: "~<"`
- **lang-en-US.json** : pas de spaces avant ponctuation, em dash sans espaces, point décimal, virgule milliers, pas d'exposant siècles, ordinaux 1st/2nd/3rd/4th sans exposant
- **lang-en-UK.json** : comme en-US mais en dash avec espaces pour incises, guillemets simples en niveau 1
- **lang-de.json** : point ordinal `1.`, `„..."` guillemets, point séparateur milliers
- **lang-es.json** : `¿?` `¡!`, `siglo XIX`, `1.º`/`1.ª` ordinaux
- **lang-it.json** : `XIX secolo`, `1º`/`1ª` sans point, `«...»` guillemets

### 2c. Ajouter le module `LanguageProfile` dans Superscript.jsx
- Insérer après I18n, avant ErrorHandler
- Responsabilités :
  - `load(langId)` : lit `dictionary/lang-{langId}.json` via `File`, parse avec `safeJSON`
  - `get(path)` : accès aux données du profil chargé (ex: `LanguageProfile.get("punctuation.spaceBeforeColon")`)
  - `getList(path)` : retourne un tableau (ex: `LanguageProfile.get("data.motsOrdinaux")`)
  - `merge(userOverrides)` : fusionne des overrides utilisateur par-dessus le profil de base
  - `getCurrentId()` : retourne l'id du profil chargé
  - `getAvailableProfiles()` : scanne `dictionary/` pour lister les profils disponibles
- Résolution du chemin `dictionary/` : relatif au fichier script (`$.fileName`)

### 2d. Ajouter un sélecteur de langue à l'UI
- Ajouter un dropdown "Profil linguistique :" en haut du dialogue, avant les onglets
- Peuplé par `LanguageProfile.getAvailableProfiles()`
- Par défaut : `fr-FR` si InDesign en français, `en-US` sinon
- Le changement de profil recharge les données et rafraîchit les valeurs par défaut dans le dialogue

**Commit**: `feat(superscript): add language profiles and dictionary/ folder`

---

## Phase 3 : Adaptation du moteur de correction au profil linguistique

### 3a. Adapter `Corrections.fixTypoSpaces()`
- Au lieu de recevoir un `spaceType` unique, lire du profil :
  - `punctuation.spaceBeforeSemicolon`, `spaceBeforeColon`, `spaceBeforeExclamation`, `spaceBeforeQuestion`
  - FR-FR : `~<` avant `;?!`, `~S` avant `:`
  - FR-CH : `~<` partout
  - EN/DE/ES/IT : skip (pas d'espace avant ponctuation)
- Adapter le REGEX `CHARACTER_BEFORE_DOUBLE_PUNCTUATION` pour exclure `:` si la langue ne met pas d'espace avant `:`
- Adapter les guillemets : `spaceInsideOpenQuote`/`spaceInsideCloseQuote`

### 3b. Adapter `Corrections.replaceDashes()`
- Lire `dashes.replaceCadratinWithDemiCadratin` du profil
- FR-FR : remplacer `—` → `–` (comportement actuel)
- EN-US : pas de remplacement (les em dashes sont corrects)
- EN-UK/DE/ES/IT : remplacer `—` → `–`

### 3c. Adapter `Corrections.fixDashIncises()`
- Lire `dashes.inciseDash` et `dashes.inciseSpace` du profil
- FR-FR : `~S` autour de `–` ou `—`
- FR-CH : espace normale (sécable) autour de `–`
- EN-US : pas d'espace autour de `—`
- EN-UK/DE : espace normale autour de `–`

### 3d. Adapter `Corrections.formatNumbers()`
- Lire `numbers.thousandsSeparator`, `numbers.decimalSeparator`, `numbers.replacePointWithComma`
- FR : `~<` milliers, `,` décimale, remplacement `.` → `,`
- EN : `,` milliers, `.` décimale, pas de remplacement
- DE/ES/IT : `.` milliers, `,` décimale

### 3e. Adapter SieclesModule
- `formaterSiecles()` : si `centuries.enabled === false`, skip
- `formaterOrdinaux()` : lire `data.motsOrdinaux` depuis profil au lieu de `CONFIG.MOTS_ORDINAUX`
- `formaterReferences()` : lire `data.titresPersonnes`, `data.motsOeuvres` depuis profil
- `formaterEspaces()` : lire `data.abreviationsRefs` etc. depuis profil
- Tous les tableaux hardcodés dans `SieclesModule.CONFIG` deviennent des lectures depuis `LanguageProfile`

### 3f. Adapter les options du dialogue
- Si le profil a `centuries.enabled === false` (EN, DE, ES, IT) : désactiver/masquer les options SieclesModule
- Si le profil n'a pas d'espaces avant ponctuation : désactiver l'option fixTypoSpaces
- Adapter les labels des checkboxes en fonction du profil (ex: "Replace em dashes with en dashes" en EN)

**Commit**: `feat(superscript): adapt correction engine to language profiles`

---

## Phase 4 : ConfigManager — sauvegarde/chargement des préférences

### 4a. Ajouter le module `ConfigManager`
- Insérer après LanguageProfile
- `save(options)` : sérialise les options utilisateur en JSON, `saveDlg()`, écriture UTF-8
  - Sauvegarde : profil sélectionné, toutes les cases cochées/décochées, noms des styles sélectionnés, overrides utilisateur
- `load(styles)` : `File.openDialog()`, parse JSON, `findStyleByName()` reverse mapping
- `autoLoad(styles)` : recherche récursive de `superscript-config.json` (max 3 niveaux depuis le dossier du document)
- Fichier : `superscript-config.json`

### 4b. Ajouter les boutons Save/Load au dialogue
- Barre de configuration en haut du dialogue (comme Markdown-Import)
- Boutons "Enregistrer" / "Charger" (ou "Save" / "Load" via I18n)
- Indicateur "Config détectée" si autoLoad trouve un fichier
- Au chargement d'une config : mettre à jour toutes les valeurs du dialogue

### 4c. Structure du fichier `superscript-config.json`
```json
{
  "version": 1,
  "languageProfile": "fr-FR",
  "styles": {
    "noteStyle": "Superscript",
    "italicStyle": "Italic",
    "smallCapsStyle": "Small caps",
    "capitalsStyle": "Large Capitals",
    "superscriptStyle": "Superscript"
  },
  "corrections": {
    "removeSpacesBeforePunctuation": true,
    "fixDoubleSpaces": true,
    "fixTypoSpaces": true,
    "fixDashIncises": false,
    "removeDoubleReturns": true,
    "removeSpacesStartParagraph": true,
    "removeSpacesEndParagraph": true,
    "removeTabs": true,
    "moveNotes": true,
    "applyNoteStyle": true,
    "replaceDashes": true,
    "fixIsolatedHyphens": true,
    "fixValueRanges": true,
    "convertEllipsis": true,
    "replaceApostrophes": true
  },
  "formatting": {
    "formatSiecles": true,
    "formatOrdinaux": true,
    "formatReferences": true,
    "formatEspaces": true,
    "formatNumbers": true,
    "addSpaces": true,
    "useComma": true,
    "excludeYears": true
  },
  "layout": {
    "enableStyleAfter": true,
    "triggerStyles": ["Heading 1", "Heading 2"],
    "targetStyle": "First paragraph",
    "applyMasterToLastPage": false,
    "masterName": null
  }
}
```

**Commit**: `feat(superscript): add ConfigManager with save/load support`

---

## Phase 5 : Nettoyage et finalisation

### 5a. Supprimer les données hardcodées devenues redondantes
- Retirer les tableaux de `SieclesModule.CONFIG` qui sont maintenant lus depuis les JSON
- Garder uniquement les constantes structurelles universelles dans `CONFIG` (REGEX patterns, SCRIPT_TITLE)
- Garder `STYLES_PETITES_CAPITALES`, `STYLES_CAPITALES`, `STYLES_EXPOSANT` pour l'auto-détection (universels)

### 5b. Documentation
- Mettre à jour le commentaire d'en-tête du script avec la nouvelle version
- Documenter les nouveaux modules dans le JSDoc existant

### 5c. Mise à jour de CLAUDE.md
- Ajouter la section SuperScript restructuré avec la nouvelle architecture

**Commit**: `refactor(superscript): remove redundant hardcoded data, update docs`

---

## Ordre d'implémentation recommandé

1. **Phase 1** (safeJSON + I18n) — peut se tester immédiatement : l'UI doit s'afficher identiquement en FR
2. **Phase 2** (LanguageProfile + dictionary/) — les fichiers JSON sont créés, le chargement fonctionne
3. **Phase 3** (adaptations moteur) — les corrections respectent le profil sélectionné
4. **Phase 4** (ConfigManager) — save/load fonctionnel
5. **Phase 5** (nettoyage) — suppression du code mort, docs

Chaque phase produit un commit indépendant et testable.

---

## Vérification

- **Phase 1** : ouvrir le script dans InDesign FR → l'UI doit être identique. Ouvrir en EN → l'UI doit être en anglais.
- **Phase 2** : sélectionner un profil dans le dropdown → les données doivent se charger sans erreur.
- **Phase 3** : exécuter le script avec profil fr-FR → résultat identique à l'actuel. Exécuter avec fr-CH → espaces fines partout (pas de `~S` avant `:`). Exécuter avec en-US → pas d'espaces avant ponctuation, pas de siècles formatés.
- **Phase 4** : sauvegarder une config → fichier JSON lisible. Charger → options restaurées.
- **Phase 5** : aucune régression par rapport à Phase 3.
