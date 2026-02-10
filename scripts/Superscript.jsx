/**
  * SuperScript — Automated multilingual typographic correction for InDesign
  *
  * Corrects typography, spaces, punctuation, dashes, quotes, apostrophes,
  * centuries, ordinals, references, and number formatting according to
  * language-specific rules defined in external JSON profiles (dictionary/).
  *
  * Supports 7 language profiles: FR-FR, FR-CH, EN-US, EN-UK, DE, ES, IT.
  * Bilingual UI (FR/EN) via I18n module, auto-detected from InDesign locale.
  * Save/load user preferences via ConfigManager (superscript-config.json).
  *
  * Architecture:
  *   safeJSON       — ES3-compatible JSON stringify/parse
  *   I18n           — Bilingual UI translations (FR/EN)
  *   LanguageProfile — Loads dictionary/lang-*.json typographic rules
  *   ConfigManager  — Save/load user preferences to JSON
  *   ErrorHandler   — Error handling with context
  *   Utilities      — Document validation, style helpers
  *   Corrections    — 21 typographic correction methods
  *   ProgressBar    — Non-blocking palette progress indicator
  *   UIBuilder      — 5-tab dialog with profile selector and config bar
  *   Processor      — Orchestrates corrections on active document
  *   SieclesModule  — Centuries, ordinals, references, non-breaking spaces
  *
  * @version 2.0
  * @license AGPL
  * @author entremonde / Spectral lab
  * @website https://lab.spectral.art
  */

(function() {
    "use strict";
    
    // Vérifier si l'application est disponible
    if (typeof app === "undefined") {
        alert("Cannot access InDesign application.");
        return;
    }
    
    /**
     * Constantes de configuration
     * @private
     */
    var CONFIG = {
        DEFAULT_STYLES: {
            SUPERSCRIPT: "Superscript",
            ITALIC: "Italic",
            FIRST_PARAGRAPH: "First paragraph"
        },
        REGEX: {
            FOOTNOTE_PATTERN: "[,;.?!~e\u00BB\\s]+~F",
            SPACE_BEFORE_FOOTNOTE: "[ \t\u00A0\u2000-\u200A\u202F\u205F\u3000]+(?=~F)",
            SPACE_BEFORE_POINT: "[ \t\u00A0\u2000-\u200A\u202F\u205F\u3000]+(?=\\.)",
            SPACE_BEFORE_COMMA: "[ \t\u00A0\u2000-\u200A\u202F\u205F\u3000]+(?=,)",
            DOUBLE_SPACES: "[ \t\u00A0\u2000-\u200A\u202F\u205F\u3000]{2,}",
            SPACE_AFTER_OPENING_QUOTE: "(?<=«)[ \t\u00A0\u2000-\u200A\u202F\u205F\u3000]",
            SPACE_BEFORE_CLOSING_QUOTE: "[ \t\u00A0\u2000-\u200A\u202F\u205F\u3000](?=»)",
            SPACE_BEFORE_DOUBLE_PUNCTUATION: "[ \t\u00A0\u2000-\u200A\u202F\u205F\u3000](?=[:;!?])",
            CHARACTER_BEFORE_DOUBLE_PUNCTUATION: "([^\\s])(?=[:;!?])",
            DOUBLE_RETURNS: "\\r[ \t\u00A0\u200B\u202F\u2060]*\\r",
            TRIPLE_DOTS: "\\.{3}",
            APOSTROPHES_TO_REPLACE: ["'", "&#39;", "&apos;", "&#8217;"]
        },
        SPACE_TYPES: [
            { labelKey: "spaceTypeFine", value: "~<" },
            { labelKey: "spaceTypeNonBreaking", value: "~S" },
        ],
        SCRIPT_TITLE: "Superscript"
    };

    // =========================================================================
    // safeJSON — ES3-compatible JSON stringify/parse
    // =========================================================================

    var safeJSON = {
        /**
         * Converts an object to a JSON string
         * @param {Object} obj - The object to stringify
         * @return {String} The JSON string
         */
        stringify: function(obj) {
            var t = typeof obj;
            if (t !== "object" || obj === null) {
                if (t === "string") return '"' + obj.replace(/\\/g, '\\\\').replace(/"/g, '\\"').replace(/\n/g, '\\n').replace(/\r/g, '\\r').replace(/\t/g, '\\t') + '"';
                if (t === "number" || t === "boolean") return String(obj);
                return "null";
            }
            if (obj instanceof Array) {
                var items = [];
                for (var i = 0; i < obj.length; i++) {
                    items.push(safeJSON.stringify(obj[i]));
                }
                return "[" + items.join(",") + "]";
            }
            var pairs = [];
            for (var key in obj) {
                if (obj.hasOwnProperty(key)) {
                    pairs.push('"' + key + '":' + safeJSON.stringify(obj[key]));
                }
            }
            return "{" + pairs.join(",") + "}";
        },

        /**
         * Parses a JSON string into an object
         * @param {String} str - The JSON string to parse
         * @return {Object} The parsed object
         */
        parse: function(str) {
            if (!/^[\],:{}\s]*$/.test(str.replace(/\\["\\\/bfnrtu]/g, '@')
                .replace(/"[^"\\\n\r]*"|true|false|null|-?\d+(?:\.\d*)?(?:[eE][+\-]?\d+)?/g, ']')
                .replace(/(?:^|:|,)(?:\s*\[)+/g, ''))) {
                throw new Error("Invalid JSON");
            }
            try {
                return (new Function('return ' + str))();
            } catch (e) {
                throw new Error("JSON parse error: " + e.message);
            }
        }
    };

    if (typeof JSON === 'undefined') {
        JSON = safeJSON;
    }

    // =========================================================================
    // I18n — Internationalization module
    // =========================================================================

    var I18n = (function() {
        var currentLanguage = 'en';

        var translations = {
            'en': {
                // App-level
                'errorInDesignAccess': 'Cannot access the InDesign application.',
                'errorUnrecoverable': 'An unrecoverable error has occurred.',
                'errorFatal': 'Fatal error: %s',
                'errorScriptHalted': 'Script halted due to a fatal error',
                'errorObjectUndefined': "Object '%s' is undefined or null",
                'errorInContext': 'Error',
                'errorContextIn': ' in ',
                'errorLine': ' (line ',
                'errorInDesignUnavailable': 'The InDesign application is not accessible',
                'errorDocumentsUnavailable': 'The documents collection is not accessible',
                'errorNoDocumentOpen': 'Please open a document before running this script.',
                'errorInvalidDocument': 'Error: Invalid document',
                'errorInvalidMasterName': 'Error: Invalid master name',
                'errorMasterNotFound': "Master '%s' not found in document %s",
                'errorApplyMaster': 'Error applying master page: %s',
                'errorRequiredStyles': 'Required character styles are not defined. Please select valid styles.',

                // Dialog
                'dialogTitle': 'SuperScript',
                'tabCorrections': 'Corrections',
                'tabSpaces': 'Spaces & Returns',
                'tabStyles': 'Styles',
                'tabFormatting': 'Formatting',
                'tabPageLayout': 'Page Layout',

                // Corrections tab
                'moveNotesLabel': 'Move footnote references before punctuation',
                'convertEllipsisLabel': 'Convert ... to ellipsis character (\u2026)',
                'replaceApostrophesLabel': 'Replace straight apostrophes with typographic apostrophes',
                'replaceDashesLabel': 'Replace em dashes with en dashes',
                'fixIsolatedHyphensLabel': 'Convert isolated hyphens to en dashes',
                'fixValueRangesLabel': 'Convert hyphens to en dashes in value ranges',

                // Spaces tab
                'fixTypoSpacesLabel': 'Fix typographic spaces:',
                'fixDashIncisesLabel': 'Fix spaces around \u2013 parenthetical dashes \u2013 :',
                'fixDoubleSpacesLabel': 'Fix multiple spaces',
                'removeDoubleReturnsLabel': 'Remove double returns',
                'removeSpacesBeforePunctuationLabel': 'Remove spaces before periods, commas and footnotes',
                'removeSpacesStartParagraphLabel': 'Remove spaces at beginning of paragraphs',
                'removeSpacesEndParagraphLabel': 'Remove spaces at end of paragraphs',
                'removeTabsLabel': 'Remove tabs',
                'formatEspacesLabel': 'Add non-breaking spaces in page references (p.\u00A054)',

                // Styles definition panel
                'styleDefinitionPanel': 'Style definitions',
                'noteStyleLabel': 'Footnote references:',
                'italicStyleLabel': 'Italic:',
                'smallCapsStyleLabel': 'Small caps:',
                'capitalStyleLabel': 'Capitals:',
                'superscriptStyleLabel': 'Superscripts:',

                // Styles application panel
                'styleApplicationPanel': 'Style application',
                'applyNoteStyleLabel': 'Apply style to footnote references',
                'applyItalicStyleLabel': 'Apply style to italic text',
                'applyExposantStyleLabel': 'Apply style to superscript text',

                // Formatting tab
                'formatSieclesLabel': 'Format centuries (XIVth century)',
                'formatOrdinauxLabel': 'Format ordinal expressions (IInd International)',
                'formatReferencesLabel': 'Format work parts and proper names (Volume III, Louis XIV)',
                'formatNumbersLabel': 'Format numbers',
                'numberSettingsPanel': 'Number formatting options',
                'addSpacesLabel': 'Add spaces between thousands (12345 \u2192 12\u2009345)',
                'excludeYearsLabel': 'Exclude potential years (numbers between 0 and 2050)',
                'useCommaLabel': 'Replace dots with commas (3.14 \u2192 3,14)',

                // Page layout tab
                'enableStyleAfterLabel': 'Conditional style application',
                'triggerStylesPanel': 'Trigger styles (multiple selection)',
                'targetStyleLabel': 'Style to apply to the following paragraph:',
                'applyMasterToLastPageLabel': 'Apply a master to the last facing page',
                'masterLabel': 'Master to apply:',

                // Space types
                'spaceTypeFine': 'Thin non-breaking space (~<)',
                'spaceTypeNonBreaking': 'Non-breaking space (~S)',

                // Default/placeholder
                'defaultStyle': '[Default style]',
                'noStylesAvailable': '[No styles available]',
                'noMastersAvailable': '[No masters available]',

                // Buttons
                'helpTooltip': 'Show SuperScript help',
                'cancelButton': 'Cancel',
                'applyButton': 'Apply',
                'closeButton': 'Close',

                // Help dialog
                'helpDialogTitle': 'SuperScript Help',
                'helpDialogHeader': 'SuperScript - User Guide',
                'helpContent': 'SUPERSCRIPT USER GUIDE\n\n'
                    + 'OVERVIEW\n\n'
                    + 'SuperScript is an InDesign script that automates typographic corrections in layout documents. It quickly fixes spaces, punctuation, quotes, apostrophes, etc.\n\n'
                    + 'TAB "SPACES & RETURNS"\n\n'
                    + '\u2022 Fix typographic spaces: Adds thin non-breaking or non-breaking spaces before or after certain characters according to typographic rules.\n'
                    + '\u2022 Fix multiple spaces: Replaces space sequences with a single space.\n'
                    + '\u2022 Remove double returns: Eliminates empty paragraphs.\n'
                    + '\u2022 Remove spaces before periods, commas and footnotes: Removes unwanted spaces.\n\n'
                    + 'TAB "STYLES"\n\n'
                    + '\u2022 Style definitions: Select styles to use for formatting footnote references, italic text and superscripts.\n'
                    + '\u2022 Style application: Enable options to automatically apply these styles.\n\n'
                    + 'TAB "FORMATTING"\n\n'
                    + '\u2022 Move footnote references: Places footnotes before punctuation.\n'
                    + '\u2022 Replace em dashes: Converts em dashes to en dashes.\n'
                    + '\u2022 Convert ... to ellipsis: Replaces three dots with the typographic ellipsis character.\n'
                    + '\u2022 Replace straight apostrophes: Uses typographic apostrophes.\n'
                    + '\u2022 Format centuries, ordinals and references: Options for formatting Roman numerals.\n\n'
                    + 'TAB "PAGE LAYOUT"\n\n'
                    + '\u2022 Conditional style application: Automatically applies a style to the paragraph following selected styles.\n'
                    + '\u2022 Apply master to last page: Useful for chapter or document endings.\n\n'
                    + 'For more information, visit our website: https://lab.spectral.art',

                // Progress bar
                'progressTitle': 'Applying typographic corrections',
                'progressRemoveSpacesBeforePunctuation': 'Removing spaces before punctuation...',
                'progressFixDoubleSpaces': 'Fixing double spaces...',
                'progressFixTypoSpaces': 'Fixing typographic spaces...',
                'progressFixDashIncises': 'Fixing spaces around parenthetical dashes...',
                'progressRemoveDoubleReturns': 'Removing double returns...',
                'progressRemoveSpacesStart': 'Removing spaces at beginning of paragraphs...',
                'progressRemoveSpacesEnd': 'Removing spaces at end of paragraphs...',
                'progressRemoveTabs': 'Removing tabs...',
                'progressMoveNotes': 'Moving footnote references...',
                'progressApplyNoteStyle': 'Applying style to footnote references...',
                'progressReplaceDashes': 'Replacing em dashes...',
                'progressFixIsolatedHyphens': 'Fixing isolated hyphens...',
                'progressFixValueRanges': 'Fixing value ranges...',
                'progressApplyItalicStyle': 'Applying italic style...',
                'progressApplyExposantStyle': 'Applying superscript style...',
                'progressConvertEllipsis': 'Converting ellipsis...',
                'progressReplaceApostrophes': 'Replacing apostrophes...',
                'progressApplyConditionalStyles': 'Applying conditional styles...',
                'progressApplyMasterToLastPage': 'Applying master to last page...',
                'progressFormatSiecles': 'Formatting centuries and ordinal expressions...',
                'progressFormatNumbers': 'Formatting numbers...',
                'progressComplete': 'Done!',

                // Alerts
                'successCorrectionsApplied': 'Corrections applied to the active document.',

                // SieclesModule errors
                'errorUnknown': 'Unknown error',
                'errorReplaceApostrophes': 'Error in replaceApostrophes: %s',
                'errorFormatSiecles': 'Error formatting centuries: %s',
                'errorFormatOrdinaux': 'Error formatting ordinal expressions: %s',
                'errorFormat1er': "Error formatting '1er' occurrences: %s",
                'errorFormatReferences': 'Error formatting work references and titles: %s',
                'errorFormatEspaces': 'Error formatting non-breaking spaces for references: %s',
                'errorMainFunction': 'main function',

                // Language profile selector
                'languageProfileLabel': 'Language profile:',
                'languageProfileNone': '[No profiles available]',

                // ConfigManager
                'saveConfigButton': 'Save',
                'loadConfigButton': 'Load',
                'configDetected': 'Config detected',
                'configNotDetected': '',
                'saveConfigTitle': 'Save SuperScript Configuration',
                'loadConfigTitle': 'Load SuperScript Configuration',
                'configSaved': 'Configuration saved successfully.',
                'configLoaded': 'Configuration loaded successfully.',
                'errorSaveConfig': 'Error saving configuration: %s',
                'errorLoadConfig': 'Error loading configuration: %s',
                'errorParseConfig': 'Error parsing configuration file: %s',
                'errorOpenConfig': 'Could not open configuration file.',

            },
            'fr': {
                // App-level
                'errorInDesignAccess': 'Impossible d\'acc\u00E9der \u00E0 l\'application InDesign.',
                'errorUnrecoverable': 'Une erreur irr\u00E9cup\u00E9rable s\'est produite.',
                'errorFatal': 'Erreur fatale\u2009: %s',
                'errorScriptHalted': 'Arr\u00EAt du script suite \u00E0 une erreur fatale',
                'errorObjectUndefined': "Objet '%s' est undefined ou null",
                'errorInContext': 'Erreur',
                'errorContextIn': ' dans ',
                'errorLine': ' (ligne ',
                'errorInDesignUnavailable': 'L\'application InDesign n\'est pas accessible',
                'errorDocumentsUnavailable': 'La collection de documents n\'est pas accessible',
                'errorNoDocumentOpen': 'Veuillez ouvrir un document avant d\'ex\u00E9cuter ce script.',
                'errorInvalidDocument': 'Erreur\u2009: Document invalide',
                'errorInvalidMasterName': 'Erreur\u2009: Nom de gabarit invalide',
                'errorMasterNotFound': "Gabarit '%s' introuvable dans le document %s",
                'errorApplyMaster': 'Erreur lors de l\'application du gabarit\u2009: %s',
                'errorRequiredStyles': 'Les styles de caract\u00E8re requis ne sont pas d\u00E9finis. Veuillez s\u00E9lectionner des styles valides.',

                // Dialog
                'dialogTitle': 'SuperScript',
                'tabCorrections': 'Corrections',
                'tabSpaces': 'Espaces et retours',
                'tabStyles': 'Styles',
                'tabFormatting': 'Formatages',
                'tabPageLayout': 'Mise en page',

                // Corrections tab
                'moveNotesLabel': 'D\u00E9placer les appels de notes de bas de page',
                'convertEllipsisLabel': 'Convertir ... en points de suspension (\u2026)',
                'replaceApostrophesLabel': 'Remplacer les apostrophes droites par les apostrophes typographiques',
                'replaceDashesLabel': 'Remplacer les tirets cadratin par des tirets demi-cadratin',
                'fixIsolatedHyphensLabel': 'Transformer les tirets isol\u00E9s en tirets demi-cadratin',
                'fixValueRangesLabel': 'Transformer les tirets en tirets demi-cadratin dans les intervalles de valeurs',

                // Spaces tab
                'fixTypoSpacesLabel': 'Corriger les espaces typographiques\u2009:',
                'fixDashIncisesLabel': 'Corriger les espaces des \u2013 incises \u2013\u2009:',
                'fixDoubleSpacesLabel': 'Corriger les espaces multiples',
                'removeDoubleReturnsLabel': 'Supprimer les doubles retours \u00E0 la ligne',
                'removeSpacesBeforePunctuationLabel': 'Supprimer les espaces avant les points, virgules et notes',
                'removeSpacesStartParagraphLabel': 'Supprimer les espaces en d\u00E9but de paragraphe',
                'removeSpacesEndParagraphLabel': 'Supprimer les espaces en fin de paragraphe',
                'removeTabsLabel': 'Supprimer les tabulations',
                'formatEspacesLabel': 'Ajouter espaces ins\u00E9cables dans les r\u00E9f\u00E9rences de page (p.\u00A054)',

                // Styles definition panel
                'styleDefinitionPanel': 'D\u00E9finition des styles',
                'noteStyleLabel': 'Appels de notes\u2009:',
                'italicStyleLabel': 'Italique\u2009:',
                'smallCapsStyleLabel': 'Petites capitales\u2009:',
                'capitalStyleLabel': 'Capitales\u2009:',
                'superscriptStyleLabel': 'Exposants\u2009:',

                // Styles application panel
                'styleApplicationPanel': 'Application des styles',
                'applyNoteStyleLabel': 'Appliquer un style aux appels de notes',
                'applyItalicStyleLabel': 'Appliquer un style au texte en italique',
                'applyExposantStyleLabel': 'Appliquer un style au texte en exposant',

                // Formatting tab
                'formatSieclesLabel': 'Formater les si\u00E8cles (XIV\u1D49 si\u00E8cle)',
                'formatOrdinauxLabel': 'Formater les expressions ordinales (II\u1D49 Internationale)',
                'formatReferencesLabel': 'Formater parties d\'\u0153uvres et noms propres (Tome III, Louis XIV)',
                'formatNumbersLabel': 'Formater les nombres',
                'numberSettingsPanel': 'Options de formatage des nombres',
                'addSpacesLabel': 'Ajouter des espaces entre les milliers (12345 \u2192 12\u2009345)',
                'excludeYearsLabel': 'Exclure les ann\u00E9es potentielles (nombres entre 0 et 2050)',
                'useCommaLabel': 'Remplacer les points par des virgules (3.14 \u2192 3,14)',

                // Page layout tab
                'enableStyleAfterLabel': 'Application conditionnelle de styles',
                'triggerStylesPanel': 'Styles d\u00E9clencheurs (s\u00E9lection multiple)',
                'targetStyleLabel': 'Style \u00E0 appliquer au paragraphe suivant\u2009:',
                'applyMasterToLastPageLabel': 'Appliquer un gabarit \u00E0 la derni\u00E8re page en vis-\u00E0-vis',
                'masterLabel': 'Gabarit \u00E0 appliquer\u2009:',

                // Space types
                'spaceTypeFine': 'Espace fine ins\u00E9cable (~<)',
                'spaceTypeNonBreaking': 'Espace ins\u00E9cable (~S)',

                // Default/placeholder
                'defaultStyle': '[Style par d\u00E9faut]',
                'noStylesAvailable': '[Aucun style disponible]',
                'noMastersAvailable': '[Aucun gabarit disponible]',

                // Buttons
                'helpTooltip': 'Afficher l\'aide de SuperScript',
                'cancelButton': 'Annuler',
                'applyButton': 'Appliquer',
                'closeButton': 'Fermer',

                // Help dialog
                'helpDialogTitle': 'Aide de SuperScript',
                'helpDialogHeader': 'SuperScript - Guide d\'utilisation',
                'helpContent': 'GUIDE D\'UTILISATION DE SUPERSCRIPT\n\n'
                    + 'PR\u00C9SENTATION\n\n'
                    + 'SuperScript est un script pour InDesign qui automatise les corrections typographiques dans les documents de mise en page. Il permet de corriger rapidement les espaces, les ponctuations, les guillemets, les apostrophes, etc.\n\n'
                    + 'ONGLET "ESPACES ET RETOURS"\n\n'
                    + '\u2022 Corriger les espaces typographiques\u2009: Ajoute des espaces fines ins\u00E9cables ou des espaces ins\u00E9cables avant ou apr\u00E8s certains caract\u00E8res selon les r\u00E8gles typographiques fran\u00E7aises.\n'
                    + '\u2022 Corriger les espaces multiples\u2009: Remplace les s\u00E9quences d\'espaces par une seule espace.\n'
                    + '\u2022 Supprimer les doubles retours\u2009: \u00C9limine les paragraphes vides.\n'
                    + '\u2022 Supprimer les espaces avant les points, virgules et notes\u2009: Retire les espaces ind\u00E9sirables.\n\n'
                    + 'ONGLET "STYLES"\n\n'
                    + '\u2022 D\u00E9finition des styles\u2009: S\u00E9lectionnez les styles \u00E0 utiliser pour mettre en forme les appels de notes, le texte en italique et les exposants.\n'
                    + '\u2022 Application des styles\u2009: Activez les options pour appliquer automatiquement ces styles.\n\n'
                    + 'ONGLET "FORMATAGES"\n\n'
                    + '\u2022 D\u00E9placer les appels de notes\u2009: Place les notes avant la ponctuation.\n'
                    + '\u2022 Remplacer les tirets cadratin\u2009: Convertit les tirets cadratin en tirets demi-cadratin.\n'
                    + '\u2022 Convertir ... en points de suspension\u2009: Remplace trois points par le caract\u00E8re typographique correspondant.\n'
                    + '\u2022 Remplacer les apostrophes droites\u2009: Utilise des apostrophes typographiques.\n'
                    + '\u2022 Formatage des si\u00E8cles, ordinaux et r\u00E9f\u00E9rences\u2009: Options pour la mise en forme des chiffres romains.\n\n'
                    + 'ONGLET "MISE EN PAGE"\n\n'
                    + '\u2022 Application conditionnelle de styles\u2009: Applique automatiquement un style au paragraphe qui suit les styles s\u00E9lectionn\u00E9s.\n'
                    + '\u2022 Appliquer un gabarit \u00E0 la derni\u00E8re page\u2009: Utile pour la fin des chapitres ou des documents.\n\n'
                    + 'Pour plus d\'informations, visitez notre site web\u2009: https://lab.spectral.art',

                // Progress bar
                'progressTitle': 'Application des corrections typographiques',
                'progressRemoveSpacesBeforePunctuation': 'Suppression des espaces avant ponctuation...',
                'progressFixDoubleSpaces': 'Correction des doubles espaces...',
                'progressFixTypoSpaces': 'Correction des espaces typographiques...',
                'progressFixDashIncises': 'Correction des espaces autour des incises...',
                'progressRemoveDoubleReturns': 'Suppression des doubles retours...',
                'progressRemoveSpacesStart': 'Suppression des espaces en d\u00E9but de paragraphe...',
                'progressRemoveSpacesEnd': 'Suppression des espaces en fin de paragraphe...',
                'progressRemoveTabs': 'Suppression des tabulations...',
                'progressMoveNotes': 'D\u00E9placement des notes de bas de page...',
                'progressApplyNoteStyle': 'Application du style aux notes de bas de page...',
                'progressReplaceDashes': 'Remplacement des tirets cadratin...',
                'progressFixIsolatedHyphens': 'Correction des tirets isol\u00E9s...',
                'progressFixValueRanges': 'Correction des intervalles de valeurs...',
                'progressApplyItalicStyle': 'Application du style italique...',
                'progressApplyExposantStyle': 'Application du style aux exposants...',
                'progressConvertEllipsis': 'Conversion des points de suspension...',
                'progressReplaceApostrophes': 'Remplacement des apostrophes...',
                'progressApplyConditionalStyles': 'Application des styles conditionnels...',
                'progressApplyMasterToLastPage': 'Application du gabarit \u00E0 la derni\u00E8re page...',
                'progressFormatSiecles': 'Formatage des si\u00E8cles et expressions ordinales...',
                'progressFormatNumbers': 'Formatage des nombres...',
                'progressComplete': 'Termin\u00E9\u2009!',

                // Alerts
                'successCorrectionsApplied': 'Corrections appliqu\u00E9es au document actif.',

                // SieclesModule errors
                'errorUnknown': 'Erreur inconnue',
                'errorReplaceApostrophes': 'Erreur dans replaceApostrophes\u2009: %s',
                'errorFormatSiecles': 'Erreur lors du formatage des si\u00E8cles\u2009: %s',
                'errorFormatOrdinaux': 'Erreur lors du formatage des expressions ordinales\u2009: %s',
                'errorFormat1er': 'Erreur lors du formatage des occurrences de \'1er\'\u2009: %s',
                'errorFormatReferences': 'Erreur lors du formatage des r\u00E9f\u00E9rences d\'\u0153uvres et titres\u2009: %s',
                'errorFormatEspaces': 'Erreur lors du formatage des espaces ins\u00E9cables pour les r\u00E9f\u00E9rences\u2009: %s',
                'errorMainFunction': 'fonction principale',

                // Language profile selector
                'languageProfileLabel': 'Profil linguistique\u2009:',
                'languageProfileNone': '[Aucun profil disponible]',

                // ConfigManager
                'saveConfigButton': 'Enregistrer',
                'loadConfigButton': 'Charger',
                'configDetected': 'Config d\u00E9tect\u00E9e',
                'configNotDetected': '',
                'saveConfigTitle': 'Enregistrer la configuration SuperScript',
                'loadConfigTitle': 'Charger la configuration SuperScript',
                'configSaved': 'Configuration enregistr\u00E9e avec succ\u00E8s.',
                'configLoaded': 'Configuration charg\u00E9e avec succ\u00E8s.',
                'errorSaveConfig': 'Erreur lors de l\'enregistrement de la configuration\u2009: %s',
                'errorLoadConfig': 'Erreur lors du chargement de la configuration\u2009: %s',
                'errorParseConfig': 'Erreur lors de l\'analyse du fichier de configuration\u2009: %s',
                'errorOpenConfig': 'Impossible d\'ouvrir le fichier de configuration.',

            }
        };

        function __(key) {
            var lang = currentLanguage;
            var langDict = translations[lang] || translations['en'];
            var str = langDict[key] || translations['en'][key] || key;
            if (arguments.length > 1) {
                var args = [];
                for (var i = 1; i < arguments.length; i++) {
                    args.push(arguments[i]);
                }
                str = str.replace(/%[sd]/g, function(match) {
                    if (!args.length) return match;
                    var arg = args.shift();
                    if (match === '%s') return String(arg);
                    if (match === '%d') return parseInt(arg, 10);
                    return match;
                });
            }
            return str;
        }

        function setLanguage(lang) {
            if (translations[lang]) {
                currentLanguage = lang;
            }
        }

        function getLanguage() {
            return currentLanguage;
        }

        function detectInDesignLanguage() {
            try {
                var locale = '';
                if (typeof app !== 'undefined' && app.hasOwnProperty('locale')) {
                    locale = String(app.locale).toLowerCase();
                }
                return (locale.indexOf('fr') !== -1) ? 'fr' : 'en';
            } catch (e) {
                return 'en';
            }
        }

        currentLanguage = detectInDesignLanguage();

        return {
            __: __,
            setLanguage: setLanguage,
            getLanguage: getLanguage,
            detectLanguage: detectInDesignLanguage
        };
    })();

    // =========================================================================
    // LanguageProfile — Language profile loader for typographic rules
    // =========================================================================

    var LanguageProfile = (function() {
        var currentProfile = null;
        var currentId = null;
        var dictionaryPath = null;

        /**
         * Resolves the dictionary/ folder path relative to the script file
         * @return {String} Absolute path to the dictionary folder
         */
        function resolveDictionaryPath() {
            if (dictionaryPath) return dictionaryPath;
            try {
                var scriptFile = new File($.fileName);
                var scriptsFolder = scriptFile.parent;
                var projectFolder = scriptsFolder.parent;
                var dictFolder = new Folder(projectFolder.fsName + "/dictionary");
                if (dictFolder.exists) {
                    dictionaryPath = dictFolder.fsName;
                    return dictionaryPath;
                }
                // Fallback: try sibling of scripts/
                dictFolder = new Folder(scriptsFolder.fsName + "/../dictionary");
                if (dictFolder.exists) {
                    dictionaryPath = dictFolder.fsName;
                    return dictionaryPath;
                }
            } catch (e) {
                // Ignore errors during path resolution
            }
            return null;
        }

        /**
         * Loads a language profile from the dictionary/ folder
         * @param {String} langId - Language ID (e.g., "fr-FR", "en-US")
         * @return {Boolean} True if the profile was loaded successfully
         */
        function load(langId) {
            var dictPath = resolveDictionaryPath();
            if (!dictPath) return false;

            var filePath = dictPath + "/lang-" + langId + ".json";
            var file = new File(filePath);

            if (!file.exists) return false;

            try {
                file.open("r");
                file.encoding = "UTF-8";
                var content = file.read();
                file.close();

                var profile = safeJSON.parse(content);
                if (profile && profile.meta && profile.meta.id) {
                    currentProfile = profile;
                    currentId = profile.meta.id;
                    return true;
                }
            } catch (e) {
                try { file.close(); } catch (ignore) {}
            }
            return false;
        }

        /**
         * Gets a value from the current profile by dot-separated path
         * @param {String} path - Dot-separated path (e.g., "punctuation.spaceBeforeColon")
         * @param {*} defaultValue - Value to return if path not found
         * @return {*} The value at the path, or defaultValue
         */
        function get(path, defaultValue) {
            if (!currentProfile) return (defaultValue !== undefined) ? defaultValue : null;
            var parts = path.split(".");
            var obj = currentProfile;
            for (var i = 0; i < parts.length; i++) {
                if (obj === null || obj === undefined || typeof obj !== "object") {
                    return (defaultValue !== undefined) ? defaultValue : null;
                }
                obj = obj[parts[i]];
            }
            if (obj === undefined) {
                return (defaultValue !== undefined) ? defaultValue : null;
            }
            return obj;
        }

        /**
         * Gets an array value from the current profile
         * @param {String} path - Dot-separated path to the array
         * @return {Array} The array, or empty array if not found
         */
        function getList(path) {
            var val = get(path, []);
            if (Object.prototype.toString.call(val) === "[object Array]") {
                return val;
            }
            return [];
        }

        /**
         * Returns the ID of the currently loaded profile
         * @return {String|null} The profile ID (e.g., "fr-FR"), or null
         */
        function getCurrentId() {
            return currentId;
        }

        /**
         * Returns the full loaded profile object
         * @return {Object|null} The profile object, or null
         */
        function getProfile() {
            return currentProfile;
        }

        /**
         * Scans the dictionary/ folder and returns all available profiles
         * @return {Array} Array of {id, label, labelEN} objects
         */
        function getAvailableProfiles() {
            var profiles = [];
            var dictPath = resolveDictionaryPath();
            if (!dictPath) return profiles;

            var dictFolder = new Folder(dictPath);
            var files = dictFolder.getFiles("lang-*.json");

            for (var i = 0; i < files.length; i++) {
                try {
                    var f = files[i];
                    f.open("r");
                    f.encoding = "UTF-8";
                    var content = f.read();
                    f.close();

                    var profile = safeJSON.parse(content);
                    if (profile && profile.meta) {
                        profiles.push({
                            id: profile.meta.id,
                            label: profile.meta.label,
                            labelEN: profile.meta.labelEN
                        });
                    }
                } catch (e) {
                    try { files[i].close(); } catch (ignore) {}
                }
            }

            // Sort profiles: fr-FR first, then alphabetically by id
            profiles.sort(function(a, b) {
                if (a.id === "fr-FR") return -1;
                if (b.id === "fr-FR") return 1;
                if (a.id < b.id) return -1;
                if (a.id > b.id) return 1;
                return 0;
            });

            return profiles;
        }

        /**
         * Gets the default profile ID based on InDesign's locale
         * @return {String} Default profile ID
         */
        function getDefaultProfileId() {
            var lang = I18n.getLanguage();
            return (lang === 'fr') ? 'fr-FR' : 'en-US';
        }

        return {
            load: load,
            get: get,
            getList: getList,
            getCurrentId: getCurrentId,
            getProfile: getProfile,
            getAvailableProfiles: getAvailableProfiles,
            getDefaultProfileId: getDefaultProfileId
        };
    })();

    // =========================================================================
    // ConfigManager — Save/Load user preferences
    // =========================================================================

    /**
     * Manages saving and loading of user configuration files
     * @private
     */
    var ConfigManager = (function() {
        var CONFIG_FILENAME = "superscript-config.json";
        var CONFIG_VERSION = 1;

        /**
         * Searches recursively for config files
         * @param {Folder} folder - Starting folder
         * @param {number} maxDepth - Maximum recursion depth
         * @param {number} currentDepth - Current recursion depth
         * @return {Array} Array of found File objects
         */
        function findConfigFilesRecursively(folder, maxDepth, currentDepth) {
            if (currentDepth > maxDepth) return [];
            var files = [];
            try {
                var configFiles = folder.getFiles(CONFIG_FILENAME);
                if (configFiles && configFiles.length > 0) {
                    for (var i = 0; i < configFiles.length; i++) {
                        files.push(configFiles[i]);
                    }
                }
                var subfolders = folder.getFiles(function(file) {
                    return file instanceof Folder;
                });
                if (subfolders && subfolders.length > 0) {
                    for (var j = 0; j < subfolders.length; j++) {
                        var subfiles = findConfigFilesRecursively(subfolders[j], maxDepth, currentDepth + 1);
                        for (var k = 0; k < subfiles.length; k++) {
                            files.push(subfiles[k]);
                        }
                    }
                }
            } catch (e) {
                // Continue silently
            }
            return files;
        }

        /**
         * Finds a style index by name in a style name array
         * @param {Array} styleNames - Array of style name strings
         * @param {string} name - Name to find
         * @return {number} Index or -1 if not found
         */
        function findStyleIndexByName(styleNames, name) {
            if (!name) return -1;
            for (var i = 0; i < styleNames.length; i++) {
                if (styleNames[i] === name) return i;
            }
            return -1;
        }

        /**
         * Finds a paragraph style index by name
         * @param {Array} paraStyleNames - Array of paragraph style names
         * @param {string} name - Style name to find
         * @return {number} Index or -1
         */
        function findParaStyleIndexByName(paraStyleNames, name) {
            if (!name) return -1;
            for (var i = 0; i < paraStyleNames.length; i++) {
                if (paraStyleNames[i] === name) return i;
            }
            return -1;
        }

        /**
         * Automatically loads configuration from near the active document
         * @return {Object|null} Parsed config data or null
         */
        function autoLoad() {
            try {
                if (typeof app === 'undefined' || !app.activeDocument || !app.activeDocument.saved) {
                    return null;
                }
                var docPath = app.activeDocument.filePath;
                if (!docPath) return null;

                var folder = new Folder(docPath);
                var files = findConfigFilesRecursively(folder, 3, 0);

                if (!files || files.length === 0) return null;

                var configFile = files[0];
                configFile.encoding = "UTF-8";

                if (configFile.open("r")) {
                    try {
                        var content = configFile.read();
                        configFile.close();
                        var configData = safeJSON.parse(content);
                        return configData;
                    } catch (e) {
                        return null;
                    }
                }
            } catch (e) {
                // Silently fail
            }
            return null;
        }

        /**
         * Saves configuration to a user-selected file
         * @param {Object} configData - Configuration object to save
         * @return {boolean} True if saved successfully
         */
        function save(configData) {
            try {
                var defaultPath = "";
                if (typeof app !== 'undefined' && app.activeDocument && app.activeDocument.saved) {
                    defaultPath = app.activeDocument.filePath + "/";
                }
                var defaultFile = new File(defaultPath + "superscript-config");
                var saveFile = defaultFile.saveDlg(
                    I18n.__("saveConfigTitle"),
                    "JSON files:*.json"
                );
                if (!saveFile) return false;

                // Ensure .json extension
                if (!saveFile.name.match(/\.json$/i)) {
                    saveFile = new File(saveFile.absoluteURI + ".json");
                }

                configData.version = CONFIG_VERSION;

                saveFile.encoding = "UTF-8";
                if (saveFile.open("w")) {
                    try {
                        saveFile.write(safeJSON.stringify(configData));
                        saveFile.close();
                        return true;
                    } catch (e) {
                        alert(I18n.__("errorSaveConfig", e.message));
                        return false;
                    }
                } else {
                    alert(I18n.__("errorOpenConfig"));
                    return false;
                }
            } catch (e) {
                alert(I18n.__("errorSaveConfig", e.message));
                return false;
            }
        }

        /**
         * Loads configuration from a user-selected file
         * @return {Object|null} Parsed config data or null
         */
        function load() {
            try {
                var openFile = File.openDialog(
                    I18n.__("loadConfigTitle"),
                    "JSON files:*.json"
                );
                if (!openFile) return null;

                openFile.encoding = "UTF-8";
                if (openFile.open("r")) {
                    try {
                        var content = openFile.read();
                        openFile.close();
                        var configData = safeJSON.parse(content);
                        return configData;
                    } catch (e) {
                        alert(I18n.__("errorParseConfig", e.message));
                        return null;
                    }
                } else {
                    alert(I18n.__("errorOpenConfig"));
                    return null;
                }
            } catch (e) {
                alert(I18n.__("errorLoadConfig", e.message));
                return null;
            }
        }

        /**
         * Collects current dialog state into a serializable config object
         * @param {Object} controls - Object with references to all dialog controls
         * @return {Object} Config data ready for serialization
         */
        function collectFromDialog(controls) {
            var config = {
                version: CONFIG_VERSION,
                languageProfile: controls.languageProfileId || null,
                styles: {
                    noteStyle: (controls.noteStyleOpt.checkbox.value && controls.noteStyleOpt.dropdown.selection)
                        ? controls.noteStyleOpt.dropdown.selection.text : null,
                    italicStyle: (controls.cbItalicStyle.checkbox.value && controls.cbItalicStyle.dropdown.selection)
                        ? controls.cbItalicStyle.dropdown.selection.text : null,
                    smallCapsStyle: (controls.romainsStyleOpt.checkbox.value && controls.romainsStyleOpt.dropdown.selection)
                        ? controls.romainsStyleOpt.dropdown.selection.text : null,
                    capitalsStyle: (controls.romainsMajStyleOpt.checkbox.value && controls.romainsMajStyleOpt.dropdown.selection)
                        ? controls.romainsMajStyleOpt.dropdown.selection.text : null,
                    superscriptStyle: (controls.exposantOrdinalStyleOpt.checkbox.value && controls.exposantOrdinalStyleOpt.dropdown.selection)
                        ? controls.exposantOrdinalStyleOpt.dropdown.selection.text : null
                },
                corrections: {
                    removeSpacesBeforePunctuation: controls.cbRemoveSpacesBeforePunctuation.value,
                    fixDoubleSpaces: controls.cbFixSpaces.value,
                    fixTypoSpaces: controls.fixTypoSpacesOpt.checkbox.value,
                    fixTypoSpacesType: (controls.fixTypoSpacesOpt.dropdown.selection)
                        ? controls.fixTypoSpacesOpt.dropdown.selection.index : 0,
                    fixDashIncises: controls.fixDashIncisesOpt.checkbox.value,
                    fixDashIncisesType: (controls.fixDashIncisesOpt.dropdown.selection)
                        ? controls.fixDashIncisesOpt.dropdown.selection.index : 0,
                    removeDoubleReturns: controls.cbDoubleReturns.value,
                    removeSpacesStartParagraph: controls.cbRemoveSpacesStartParagraph.value,
                    removeSpacesEndParagraph: controls.cbRemoveSpacesEndParagraph.value,
                    removeTabs: controls.cbRemoveTabs.value,
                    moveNotes: controls.cbMoveNotes.value,
                    applyNoteStyle: controls.applyNoteStyleOpt.value,
                    replaceDashes: controls.cbDashes.value,
                    fixIsolatedHyphens: controls.cbFixIsolatedHyphens.value,
                    fixValueRanges: controls.cbFixValueRanges.value,
                    convertEllipsis: controls.cbEllipsis.value,
                    replaceApostrophes: controls.cbReplaceApostrophes.value,
                    applyItalicStyle: controls.applyItalicStyleOpt.value,
                    applyExposantStyle: controls.applyExposantStyleOpt.value,
                    formatEspaces: controls.cbFormatEspaces.value
                },
                formatting: {
                    formatSiecles: controls.cbFormatSiecles.value,
                    formatOrdinaux: controls.cbFormatOrdinaux.value,
                    formatReferences: controls.cbFormatReferences.value,
                    formatNumbers: controls.cbFormatNumbers.value,
                    addSpaces: controls.cbAddSpaces.value,
                    useComma: controls.cbUseComma.value,
                    excludeYears: controls.cbExcludeYears.value
                },
                layout: {
                    enableStyleAfter: controls.cbEnableStyleAfter.value,
                    triggerStyles: [],
                    targetStyle: (controls.targetStyleDropdown.selection)
                        ? controls.targetStyleDropdown.selection.text : null,
                    applyMasterToLastPage: controls.cbApplyMasterToLastPage.value,
                    masterName: (controls.masterDropdown.selection)
                        ? controls.masterDropdown.selection.text : null
                }
            };

            // Collect trigger style names
            for (var i = 0; i < controls.triggerCheckboxes.length; i++) {
                if (controls.triggerCheckboxes[i].value) {
                    config.layout.triggerStyles.push(controls.triggerCheckboxes[i].text);
                }
            }

            return config;
        }

        /**
         * Applies loaded config data to dialog controls
         * @param {Object} configData - Parsed config data
         * @param {Object} controls - Object with references to all dialog controls
         * @param {Array} characterStyles - Character style names array
         * @param {Array} availableProfiles - Available profile descriptors
         */
        function applyToDialog(configData, controls, characterStyles, availableProfiles) {
            if (!configData) return;

            // Language profile
            if (configData.languageProfile && availableProfiles && availableProfiles.length > 0) {
                for (var pi = 0; pi < availableProfiles.length; pi++) {
                    if (availableProfiles[pi].id === configData.languageProfile) {
                        controls.profileDropdown.selection = pi;
                        LanguageProfile.load(configData.languageProfile);
                        break;
                    }
                }
            }

            // Styles — find index in character styles dropdown
            if (configData.styles) {
                var s = configData.styles;
                if (s.noteStyle) {
                    var idx = findStyleIndexByName(characterStyles, s.noteStyle);
                    if (idx >= 0) {
                        controls.noteStyleOpt.dropdown.selection = idx;
                        controls.noteStyleOpt.checkbox.value = true;
                        controls.noteStyleOpt.dropdown.enabled = true;
                    }
                }
                if (s.italicStyle) {
                    var idx = findStyleIndexByName(characterStyles, s.italicStyle);
                    if (idx >= 0) {
                        controls.cbItalicStyle.dropdown.selection = idx;
                        controls.cbItalicStyle.checkbox.value = true;
                        controls.cbItalicStyle.dropdown.enabled = true;
                    }
                }
                if (s.smallCapsStyle) {
                    var idx = findStyleIndexByName(characterStyles, s.smallCapsStyle);
                    if (idx >= 0) {
                        controls.romainsStyleOpt.dropdown.selection = idx;
                        controls.romainsStyleOpt.checkbox.value = true;
                        controls.romainsStyleOpt.dropdown.enabled = true;
                    }
                }
                if (s.capitalsStyle) {
                    var idx = findStyleIndexByName(characterStyles, s.capitalsStyle);
                    if (idx >= 0) {
                        controls.romainsMajStyleOpt.dropdown.selection = idx;
                        controls.romainsMajStyleOpt.checkbox.value = true;
                        controls.romainsMajStyleOpt.dropdown.enabled = true;
                    }
                }
                if (s.superscriptStyle) {
                    var idx = findStyleIndexByName(characterStyles, s.superscriptStyle);
                    if (idx >= 0) {
                        controls.exposantOrdinalStyleOpt.dropdown.selection = idx;
                        controls.exposantOrdinalStyleOpt.checkbox.value = true;
                        controls.exposantOrdinalStyleOpt.dropdown.enabled = true;
                    }
                }
            }

            // Corrections
            if (configData.corrections) {
                var c = configData.corrections;
                if (typeof c.removeSpacesBeforePunctuation === 'boolean') controls.cbRemoveSpacesBeforePunctuation.value = c.removeSpacesBeforePunctuation;
                if (typeof c.fixDoubleSpaces === 'boolean') controls.cbFixSpaces.value = c.fixDoubleSpaces;
                if (typeof c.fixTypoSpaces === 'boolean') {
                    controls.fixTypoSpacesOpt.checkbox.value = c.fixTypoSpaces;
                    controls.fixTypoSpacesOpt.dropdown.enabled = c.fixTypoSpaces;
                }
                if (typeof c.fixTypoSpacesType === 'number' && c.fixTypoSpacesType < controls.fixTypoSpacesOpt.dropdown.items.length) {
                    controls.fixTypoSpacesOpt.dropdown.selection = c.fixTypoSpacesType;
                }
                if (typeof c.fixDashIncises === 'boolean') {
                    controls.fixDashIncisesOpt.checkbox.value = c.fixDashIncises;
                    controls.fixDashIncisesOpt.dropdown.enabled = c.fixDashIncises;
                }
                if (typeof c.fixDashIncisesType === 'number' && c.fixDashIncisesType < controls.fixDashIncisesOpt.dropdown.items.length) {
                    controls.fixDashIncisesOpt.dropdown.selection = c.fixDashIncisesType;
                }
                if (typeof c.removeDoubleReturns === 'boolean') controls.cbDoubleReturns.value = c.removeDoubleReturns;
                if (typeof c.removeSpacesStartParagraph === 'boolean') controls.cbRemoveSpacesStartParagraph.value = c.removeSpacesStartParagraph;
                if (typeof c.removeSpacesEndParagraph === 'boolean') controls.cbRemoveSpacesEndParagraph.value = c.removeSpacesEndParagraph;
                if (typeof c.removeTabs === 'boolean') controls.cbRemoveTabs.value = c.removeTabs;
                if (typeof c.moveNotes === 'boolean') controls.cbMoveNotes.value = c.moveNotes;
                if (typeof c.applyNoteStyle === 'boolean') controls.applyNoteStyleOpt.value = c.applyNoteStyle;
                if (typeof c.replaceDashes === 'boolean') controls.cbDashes.value = c.replaceDashes;
                if (typeof c.fixIsolatedHyphens === 'boolean') controls.cbFixIsolatedHyphens.value = c.fixIsolatedHyphens;
                if (typeof c.fixValueRanges === 'boolean') controls.cbFixValueRanges.value = c.fixValueRanges;
                if (typeof c.convertEllipsis === 'boolean') controls.cbEllipsis.value = c.convertEllipsis;
                if (typeof c.replaceApostrophes === 'boolean') controls.cbReplaceApostrophes.value = c.replaceApostrophes;
                if (typeof c.applyItalicStyle === 'boolean') controls.applyItalicStyleOpt.value = c.applyItalicStyle;
                if (typeof c.applyExposantStyle === 'boolean') controls.applyExposantStyleOpt.value = c.applyExposantStyle;
                if (typeof c.formatEspaces === 'boolean') controls.cbFormatEspaces.value = c.formatEspaces;
            }

            // Formatting
            if (configData.formatting) {
                var f = configData.formatting;
                if (typeof f.formatSiecles === 'boolean') controls.cbFormatSiecles.value = f.formatSiecles;
                if (typeof f.formatOrdinaux === 'boolean') controls.cbFormatOrdinaux.value = f.formatOrdinaux;
                if (typeof f.formatReferences === 'boolean') controls.cbFormatReferences.value = f.formatReferences;
                if (typeof f.formatNumbers === 'boolean') {
                    controls.cbFormatNumbers.value = f.formatNumbers;
                    controls.numberSettingsPanel.enabled = f.formatNumbers;
                }
                if (typeof f.addSpaces === 'boolean') controls.cbAddSpaces.value = f.addSpaces;
                if (typeof f.useComma === 'boolean') controls.cbUseComma.value = f.useComma;
                if (typeof f.excludeYears === 'boolean') controls.cbExcludeYears.value = f.excludeYears;
            }

            // Layout
            if (configData.layout) {
                var l = configData.layout;
                if (typeof l.enableStyleAfter === 'boolean') controls.cbEnableStyleAfter.value = l.enableStyleAfter;
                if (typeof l.applyMasterToLastPage === 'boolean') {
                    controls.cbApplyMasterToLastPage.value = l.applyMasterToLastPage;
                    controls.masterDropdown.enabled = l.applyMasterToLastPage;
                }

                // Trigger styles
                if (l.triggerStyles && controls.triggerCheckboxes) {
                    // First uncheck all
                    for (var ti = 0; ti < controls.triggerCheckboxes.length; ti++) {
                        controls.triggerCheckboxes[ti].value = false;
                    }
                    // Then check the ones in config
                    for (var ts = 0; ts < l.triggerStyles.length; ts++) {
                        for (var tc = 0; tc < controls.triggerCheckboxes.length; tc++) {
                            if (controls.triggerCheckboxes[tc].text === l.triggerStyles[ts]) {
                                controls.triggerCheckboxes[tc].value = true;
                                break;
                            }
                        }
                    }
                }

                // Target style
                if (l.targetStyle && controls.targetStyleDropdown) {
                    for (var tsi = 0; tsi < controls.targetStyleDropdown.items.length; tsi++) {
                        if (controls.targetStyleDropdown.items[tsi].text === l.targetStyle) {
                            controls.targetStyleDropdown.selection = tsi;
                            break;
                        }
                    }
                }

                // Master
                if (l.masterName && controls.masterDropdown) {
                    for (var mi = 0; mi < controls.masterDropdown.items.length; mi++) {
                        if (controls.masterDropdown.items[mi].text === l.masterName) {
                            controls.masterDropdown.selection = mi;
                            break;
                        }
                    }
                }
            }
        }

        return {
            autoLoad: autoLoad,
            save: save,
            load: load,
            collectFromDialog: collectFromDialog,
            applyToDialog: applyToDialog
        };
    })();

    /**
     * Gestionnaire d'erreurs personnalisé
     * @private
     */
    var ErrorHandler = {
        /**
         * Affiche un message d'erreur avec détails et trace
         * @param {Error} error - L'erreur qui s'est produite
         * @param {string} context - Contexte dans lequel l'erreur s'est produite
         * @param {boolean} isFatal - Si true, arrête l'exécution du script
         */
        handleError: function(error, context, isFatal) {
            var message = I18n.__("errorInContext");

            if (context) {
                message += I18n.__("errorContextIn") + context;
            }

            message += " : " + (error.message || I18n.__("errorUnknown"));

            if (error.line) {
                message += I18n.__("errorLine") + error.line + ")";
            }
            
            try {
                if (typeof console !== "undefined" && console && console.log) {
                    console.log(message);
                    if (error.stack) {
                        console.log("Stack trace: " + error.stack);
                    }
                }
            } catch (e) {
                // Ignorer les erreurs de journalisation
            }
            
            if (isFatal) {
                alert(message);
                throw new Error(I18n.__("errorScriptHalted"));
            }
            
            return message;
        },
        
        /**
         * Vérifie qu'un objet n'est pas undefined ou null
         * @param {*} obj - L'objet à vérifier
         * @param {string} name - Nom de l'objet pour le message d'erreur
         * @param {boolean} throwError - Si true, lance une erreur si l'objet est null
         * @returns {boolean} True si l'objet est défini, false sinon
         */
        ensureDefined: function(obj, name, throwError) {
            if (obj === undefined || obj === null) {
                var message = I18n.__("errorObjectUndefined", name);
                
                if (throwError) {
                    throw new Error(message);
                }
                
                try {
                    if (typeof console !== "undefined" && console && console.log) {
                        console.log(message);
                    }
                } catch (e) {
                    // Ignorer les erreurs de journalisation
                }
                
                return false;
            }
            
            return true;
        }
    };
    
    /**
     * Fonctions utilitaires pour les opérations du script
     * @private
     */
    var Utilities = {
        /**
         * Vérifie qu'un document est ouvert avant de continuer
         * @returns {boolean} Vrai si des documents sont disponibles
         */
        validateDocumentOpen: function() {
            try {
                if (!ErrorHandler.ensureDefined(app, "app", false)) {
                    alert(I18n.__("errorInDesignUnavailable"));
                    return false;
                }

                if (!ErrorHandler.ensureDefined(app.documents, "app.documents", false)) {
                    alert(I18n.__("errorDocumentsUnavailable"));
                    return false;
                }

                if (app.documents.length === 0) {
                    alert(I18n.__("errorNoDocumentOpen"));
                    return false;
                }
                
                return true;
            } catch (error) {
                ErrorHandler.handleError(error, "validateDocumentOpen", true);
                return false;
            }
        },
        
        /**
         * Récupère les styles de caractère disponibles dans le document
         * @param {Document} doc - Document InDesign
         * @returns {Object} Objet contenant les noms de style et les index
         */
        getCharacterStyles: function(doc) {
            var result = {
                styles: [],
                superscriptIndex: 0,
                italicIndex: 0
            };
            
            try {
                if (!ErrorHandler.ensureDefined(doc, "document", true)) {
                    return result;
                }
                
                if (!ErrorHandler.ensureDefined(doc.characterStyles, "document.characterStyles", true)) {
                    // Utiliser des valeurs par défaut
                    result.styles = [I18n.__("defaultStyle"), CONFIG.DEFAULT_STYLES.SUPERSCRIPT, CONFIG.DEFAULT_STYLES.ITALIC];
                    return result;
                }
                
                for (var i = 0; i < doc.characterStyles.length; i++) {
                    try {
                        var style = doc.characterStyles[i];
                        
                        if (!ErrorHandler.ensureDefined(style, "style at index " + i, false)) {
                            continue;
                        }
                        
                        if (!ErrorHandler.ensureDefined(style.name, "style.name at index " + i, false)) {
                            continue;
                        }
                        
                        var styleName = style.name;
                        result.styles.push(styleName);
                        
                        // Trouver les index de style par défaut
                        if (styleName === CONFIG.DEFAULT_STYLES.SUPERSCRIPT) {
                            result.superscriptIndex = i;
                        }
                        if (styleName === CONFIG.DEFAULT_STYLES.ITALIC) {
                            result.italicIndex = i;
                        }
                    } catch (styleError) {
                        ErrorHandler.handleError(styleError, "getCharacterStyles loop at index " + i, false);
                        // Continuer avec le style suivant
                    }
                }
                
                // S'assurer qu'au moins un style existe
                if (result.styles.length === 0) {
                    result.styles.push(I18n.__("defaultStyle"));
                }
            } catch (error) {
                ErrorHandler.handleError(error, "getCharacterStyles", false);
                // Ajouter des styles par défaut en cas d'erreur
                result.styles = [I18n.__("defaultStyle"), CONFIG.DEFAULT_STYLES.SUPERSCRIPT, CONFIG.DEFAULT_STYLES.ITALIC];
            }
            
            return result;
        },
        
        /**
         * Récupère tous les styles de paragraphe, y compris ceux dans les groupes
         * @param {Document} doc - Document InDesign
         * @returns {Object} Objet contenant les styles de paragraphe et l'index du style "First paragraph"
         */
        getParagraphStyles: function(doc) {
            var result = {
                styles: [],
                firstParagraphIndex: -1
            };
            
            try {
                if (!ErrorHandler.ensureDefined(doc, "document", true)) {
                    return result;
                }
                
                // Fonction récursive pour collecter les styles dans les groupes
                function collectStyles(group) {
                    var styles = [];
                    
                    try {
                        if (!ErrorHandler.ensureDefined(group, "groupe de styles", false)) {
                            return styles;
                        }
                        
                        if (ErrorHandler.ensureDefined(group.paragraphStyles, "group.paragraphStyles", false)) {
                            for (var i = 0; i < group.paragraphStyles.length; i++) {
                                try {
                                    var style = group.paragraphStyles[i];
                                    
                                    if (!ErrorHandler.ensureDefined(style, "style at index " + i, false)) {
                                        continue;
                                    }
                                    
                                    if (!ErrorHandler.ensureDefined(style.name, "style.name at index " + i, false)) {
                                        continue;
                                    }
                                    
                                    // Ignorer les styles avec crochets
                                    if (!style.name.match(/^\[/)) {
                                        styles.push(style);
                                    }
                                } catch (styleError) {
                                    ErrorHandler.handleError(styleError, "getParagraphStyles loop at index " + i, false);
                                    // Continuer avec le style suivant
                                }
                            }
                        }
                        
                        if (ErrorHandler.ensureDefined(group.paragraphStyleGroups, "group.paragraphStyleGroups", false)) {
                            for (var j = 0; j < group.paragraphStyleGroups.length; j++) {
                                try {
                                    var subgroup = group.paragraphStyleGroups[j];
                                    
                                    if (!ErrorHandler.ensureDefined(subgroup, "subgroup at index " + j, false)) {
                                        continue;
                                    }
                                    
                                    styles = styles.concat(collectStyles(subgroup));
                                } catch (groupError) {
                                    ErrorHandler.handleError(groupError, "group loop at index " + j, false);
                                    // Continuer avec le groupe suivant
                                }
                            }
                        }
                    } catch (error) {
                        ErrorHandler.handleError(error, "collectStyles", false);
                    }
                    
                    return styles;
                }
                
                // Collecter tous les styles
                var allStyles = collectStyles(doc);
                result.styles = allStyles;
                
                // Trouver l'index du style "First paragraph" s'il existe
                for (var k = 0; k < allStyles.length; k++) {
                    if (allStyles[k].name === CONFIG.DEFAULT_STYLES.FIRST_PARAGRAPH) {
                        result.firstParagraphIndex = k;
                        break;
                    }
                }
            } catch (error) {
                ErrorHandler.handleError(error, "getParagraphStyles", false);
            }
            
            return result;
        },
        
        /**
         * Récupère ou crée un style de caractère dans le document
         * @param {Document} doc - Document InDesign
         * @param {string} name - Nom du style
         * @param {Object} properties - Propriétés du style à définir si créé
         * @returns {CharacterStyle} Le style récupéré ou créé
         */
        getOrCreateStyle: function(doc, name, properties) {
            var style = null;
            
            try {
                if (!ErrorHandler.ensureDefined(doc, "document", true)) {
                    return null;
                }
                
                if (!ErrorHandler.ensureDefined(name, "nom du style", true)) {
                    return null;
                }
                
                if (!ErrorHandler.ensureDefined(properties, "style properties", false)) {
                    properties = {}; // Créer un objet vide par défaut
                }
                
                if (!ErrorHandler.ensureDefined(doc.characterStyles, "document.characterStyles", true)) {
                    return null;
                }
                
                // Essayer de récupérer le style existant
                try {
                    style = doc.characterStyles.itemByName(name);
                    
                    // Vérifier si le style existe
                    if (!style || !style.isValid) {
                        throw new Error("Style not found");
                    }
                } catch (lookupError) {
                    // Créer le style s'il n'existe pas
                    try {
                        style = doc.characterStyles.add({name: name});
                        
                        // Appliquer les propriétés
                        for (var prop in properties) {
                            if (Object.prototype.hasOwnProperty.call(properties, prop)) {
                                style[prop] = properties[prop];
                            }
                        }
                    } catch (createError) {
                        ErrorHandler.handleError(createError, "creating style " + name, false);
                        return null;
                    }
                }
            } catch (error) {
                ErrorHandler.handleError(error, "getOrCreateStyle", false);
                return null;
            }
            
            return style;
        },
        
        /**
         * Vérifie si un paragraphe est vide
         * @param {Paragraph} paragraph - Paragraphe à vérifier
         * @returns {boolean} True si le paragraphe est vide
         */
        isEmptyParagraph: function(paragraph) {
            try {
                if (!ErrorHandler.ensureDefined(paragraph, "paragraph", true)) {
                    return false;
                }
                
                if (!ErrorHandler.ensureDefined(paragraph.contents, "paragraph.contents", true)) {
                    return false;
                }
                
                var contents = paragraph.contents;
                return contents.replace(/[\r\n\s\u200B\uFEFF]/g, "") === "";
            } catch (error) {
                ErrorHandler.handleError(error, "isEmptyParagraph", false);
                return false;
            }
        },
        
        /**
         * Vérifie si un style est dans une liste de styles
         * @param {ParagraphStyle} style - Style à vérifier
         * @param {Array} styleList - Liste des styles
         * @returns {boolean} True si le style est dans la liste
         */
        isStyleInList: function(style, styleList) {
            try {
                if (!ErrorHandler.ensureDefined(style, "style", true)) {
                    return false;
                }
                
                if (!ErrorHandler.ensureDefined(styleList, "styleList", true)) {
                    return false;
                }
                
                if (Object.prototype.toString.call(styleList) !== "[object Array]") {
                    return false;
                }
                
                for (var i = 0; i < styleList.length; i++) {
                    if (style.id === styleList[i].id) {
                        return true;
                    }
                }
                
                return false;
            } catch (error) {
                ErrorHandler.handleError(error, "isStyleInList", false);
                return false;
            }
        },
        
        /**
         * Réinitialise toutes les préférences de recherche
         */
        resetPreferences: function() {
            try {
                if (!ErrorHandler.ensureDefined(app, "app", true)) return;
                
                // Réinitialiser les préférences GREP
                if (ErrorHandler.ensureDefined(app.findGrepPreferences, "app.findGrepPreferences", false)) {
                    app.findGrepPreferences = NothingEnum.NOTHING;
                }
                
                if (ErrorHandler.ensureDefined(app.changeGrepPreferences, "app.changeGrepPreferences", false)) {
                    app.changeGrepPreferences = NothingEnum.NOTHING;
                }
                
                // Réinitialiser les préférences de texte
                if (ErrorHandler.ensureDefined(app.findTextPreferences, "app.findTextPreferences", false)) {
                    app.findTextPreferences = NothingEnum.NOTHING;
                }
                
                if (ErrorHandler.ensureDefined(app.changeTextPreferences, "app.changeTextPreferences", false)) {
                    app.changeTextPreferences = NothingEnum.NOTHING;
                }
                
                // Activer l'inclusion des notes de bas de page par défaut
                if (ErrorHandler.ensureDefined(app.findChangeGrepOptions, "app.findChangeGrepOptions", false)) {
                    app.findChangeGrepOptions.includeFootnotes = true;
                }
                
                if (ErrorHandler.ensureDefined(app.findChangeTextOptions, "app.findChangeTextOptions", false)) {
                    app.findChangeTextOptions.includeFootnotes = true;
                }
            } catch (error) {
                ErrorHandler.handleError(error, "resetPreferences", false);
            }
        }
    };
    
    /**
     * Fonctions de correction typographique
     * @private
     */
    var Corrections = {
        /**
         * Supprime les espaces avant la ponctuation simple et les notes de bas de page
         * @param {Document} doc - Document InDesign
         */
        removeSpacesBeforePunctuation: function(doc) {
            try {
                if (!ErrorHandler.ensureDefined(doc, "document", true)) return;
                if (!ErrorHandler.ensureDefined(app, "app", true)) return;
                if (!ErrorHandler.ensureDefined(app.findGrepPreferences, "app.findGrepPreferences", true)) return;
                if (!ErrorHandler.ensureDefined(app.changeGrepPreferences, "app.changeGrepPreferences", true)) return;
                
                // Supprimer les espaces avant les notes de bas de page
                Utilities.resetPreferences();
                app.findGrepPreferences.findWhat = CONFIG.REGEX.SPACE_BEFORE_FOOTNOTE;
                app.changeGrepPreferences.changeTo = "";
                doc.changeGrep();
                
                // Supprimer les espaces avant les points
                Utilities.resetPreferences();
                app.findGrepPreferences.findWhat = CONFIG.REGEX.SPACE_BEFORE_POINT;
                app.changeGrepPreferences.changeTo = "";
                doc.changeGrep();
                
                // Supprimer les espaces avant les virgules
                Utilities.resetPreferences();
                app.findGrepPreferences.findWhat = CONFIG.REGEX.SPACE_BEFORE_COMMA;
                app.changeGrepPreferences.changeTo = "";
                doc.changeGrep();
            } catch (error) {
                ErrorHandler.handleError(error, "removeSpacesBeforePunctuation", false);
            }
        },
        
        /**
         * Supprime les espaces en début de paragraphe
         * @param {Document} doc - Document InDesign
         */
        removeSpacesStartParagraph: function(doc) {
            try {
                if (!ErrorHandler.ensureDefined(doc, "document", true)) return;
                if (!ErrorHandler.ensureDefined(app, "app", true)) return;
                if (!ErrorHandler.ensureDefined(app.findGrepPreferences, "app.findGrepPreferences", true)) return;
                if (!ErrorHandler.ensureDefined(app.changeGrepPreferences, "app.changeGrepPreferences", true)) return;
                
                Utilities.resetPreferences();
                app.findGrepPreferences.findWhat = "^\\s+"; // Cherche un ou plusieurs espaces en début de paragraphe
                app.changeGrepPreferences.changeTo = "";
                doc.changeGrep();
            } catch (error) {
                ErrorHandler.handleError(error, "removeSpacesStartParagraph", false);
            }
        },
        
        /**
         * Supprime les espaces en fin de paragraphe
         * @param {Document} doc - Document InDesign
         */
        removeSpacesEndParagraph: function(doc) {
            try {
                if (!ErrorHandler.ensureDefined(doc, "document", true)) return;
                if (!ErrorHandler.ensureDefined(app, "app", true)) return;
                if (!ErrorHandler.ensureDefined(app.findGrepPreferences, "app.findGrepPreferences", true)) return;
                if (!ErrorHandler.ensureDefined(app.changeGrepPreferences, "app.changeGrepPreferences", true)) return;
                
                Utilities.resetPreferences();
                app.findGrepPreferences.findWhat = "\\s+$"; // Cherche un ou plusieurs espaces en fin de paragraphe
                app.changeGrepPreferences.changeTo = "";
                doc.changeGrep();
            } catch (error) {
                ErrorHandler.handleError(error, "removeSpacesEndParagraph", false);
            }
        },
        
        /**
         * Supprime les tabulations, sauf celles qui suivent une référence de note dans les notes de bas de page
         * @param {Document} doc - Document InDesign
         */
        removeTabs: function(doc) {
            try {
                if (!ErrorHandler.ensureDefined(doc, "document", true)) return;
                if (!ErrorHandler.ensureDefined(app, "app", true)) return;
                if (!ErrorHandler.ensureDefined(app.findGrepPreferences, "app.findGrepPreferences", true)) return;
                if (!ErrorHandler.ensureDefined(app.changeGrepPreferences, "app.changeGrepPreferences", true)) return;
                
                // Approche alternative: traiter les tabulations par contexte
                
                // 1. D'abord, supprimer les tabulations régulières (pas après ~F) partout
                Utilities.resetPreferences();
                app.findChangeGrepOptions.includeFootnotes = true; // Inclure les notes
                
                // Rechercher les tabulations qui ne sont pas précédées par ~F
                app.findGrepPreferences.findWhat = "(?<!~F)\\t";
                app.changeGrepPreferences.changeTo = "";
                doc.changeGrep();
                
                // 2. Ensuite, supprimer les tabulations après ~F, mais uniquement dans le texte principal
                Utilities.resetPreferences();
                app.findChangeGrepOptions.includeFootnotes = false; // Exclure les notes
                
                // Rechercher ~F suivi d'une tabulation dans le texte principal uniquement
                app.findGrepPreferences.findWhat = "~F\\t";
                app.changeGrepPreferences.changeTo = "~F"; // Garder la référence de note mais supprimer la tabulation
                doc.changeGrep();
                
            } catch (error) {
                ErrorHandler.handleError(error, "removeTabs", false);
            }
        },
        
        /**
         * Déplace les références de notes de bas de page avant la ponctuation
         * @param {Document} doc - Document InDesign
         */
        moveNotes: function(doc) {
            try {
                if (!ErrorHandler.ensureDefined(doc, "document", true)) return;
                if (!ErrorHandler.ensureDefined(app, "app", true)) return;
                if (!ErrorHandler.ensureDefined(app.findGrepPreferences, "app.findGrepPreferences", true)) return;
                if (!ErrorHandler.ensureDefined(app.changeGrepPreferences, "app.changeGrepPreferences", true)) return;
                if (!ErrorHandler.ensureDefined(app.findChangeGrepOptions, "app.findChangeGrepOptions", true)) return;
                if (!ErrorHandler.ensureDefined(LocationOptions, "LocationOptions", true)) return;
                
                var LO_BEFORE = LocationOptions.BEFORE;
                
                app.findGrepPreferences = null;
                app.changeGrepPreferences = null;
                app.findChangeGrepOptions.includeFootnotes = false;
                app.findGrepPreferences.findWhat = CONFIG.REGEX.FOOTNOTE_PATTERN;
                
                var foundItems = doc.findGrep();
                
                if (!ErrorHandler.ensureDefined(foundItems, "foundItems", false)) {
                    return;
                }
                
                var itemCount = foundItems.length;
                
                // Traiter les éléments en ordre inverse pour éviter les problèmes de décalage d'index
                for (var i = itemCount - 1; i >= 0; i--) {
                    try {
                        var item = foundItems[i];
                        
                        if (!ErrorHandler.ensureDefined(item, "item at index " + i, false)) {
                            continue;
                        }
                        
                        if (!ErrorHandler.ensureDefined(item.texts, "item.texts at index " + i, false)) {
                            continue;
                        }
                        
                        if (item.texts.length === 0) {
                            continue;
                        }
                        
                        var t = item.texts[0].characters;
                        
                        if (!ErrorHandler.ensureDefined(t, "characters at index " + i, false)) {
                            continue;
                        }
                        
                        t[-1].move(LO_BEFORE, t[0]);
                    } catch (itemError) {
                        ErrorHandler.handleError(itemError, "processing item " + i, false);
                        // Continuer avec l'élément suivant
                    }
                }
            } catch (error) {
                ErrorHandler.handleError(error, "moveNotes", false);
            }
        },
        
        /**
         * Corrige les doubles espaces dans tout le document
         * @param {Document} doc - Document InDesign
         */
        fixDoubleSpaces: function(doc) {
            try {
                if (!ErrorHandler.ensureDefined(doc, "document", true)) return;
                if (!ErrorHandler.ensureDefined(app, "app", true)) return;
                if (!ErrorHandler.ensureDefined(app.findGrepPreferences, "app.findGrepPreferences", true)) return;
                if (!ErrorHandler.ensureDefined(app.changeGrepPreferences, "app.changeGrepPreferences", true)) return;
                
                Utilities.resetPreferences();
                
                app.findGrepPreferences.findWhat = CONFIG.REGEX.DOUBLE_SPACES;
                app.changeGrepPreferences.changeTo = " ";
                doc.changeGrep();
            } catch (error) {
                ErrorHandler.handleError(error, "fixDoubleSpaces", false);
            }
        },
        
        /**
         * Applique un style de caractère aux références de notes de bas de page
         * @param {Document} doc - Document InDesign
         * @param {string} styleName - Nom du style de caractère
         */
        applyNoteStyle: function(doc, styleName) {
            try {
                if (!ErrorHandler.ensureDefined(doc, "document", true)) return;
                if (!ErrorHandler.ensureDefined(styleName, "styleName", true)) return;
                if (!ErrorHandler.ensureDefined(app, "app", true)) return;
                if (!ErrorHandler.ensureDefined(app.findGrepPreferences, "app.findGrepPreferences", true)) return;
                
                // Récupérer ou créer le style
                var style = Utilities.getOrCreateStyle(doc, styleName, { position: Position.SUPERSCRIPT });
                
                if (!ErrorHandler.ensureDefined(style, "style", false)) {
                    return;
                }
                
                if (!ErrorHandler.ensureDefined(doc.stories, "document.stories", true)) {
                    return;
                }
                
                // Traiter chaque histoire
                var stories = doc.stories;
                for (var i = 0; i < stories.length; i++) {
                    try {
                        var story = stories[i];
                        
                        if (!ErrorHandler.ensureDefined(story, "story at index " + i, false)) {
                            continue;
                        }
                        
                        // Rechercher les références de notes de bas de page seulement dans le texte principal
                        // et non dans les notes elles-mêmes
                        
                        // Assurons-nous que les paramètres de recherche sont correctement configurés
                        app.findGrepPreferences = NothingEnum.NOTHING;
                        app.changeGrepPreferences = NothingEnum.NOTHING;
                        
                        // Désactiver explicitement la recherche dans les notes de bas de page
                        app.findChangeGrepOptions.includeFootnotes = false;
                        
                        // Configurer la recherche
                        app.findGrepPreferences.findWhat = "~F";
                        
                        // Effectuer la recherche uniquement dans cette histoire
                        var founds = story.findGrep();
                        
                        if (!ErrorHandler.ensureDefined(founds, "founds dans story " + i, false)) {
                            continue;
                        }
                        
                        // Appliquer le style aux références trouvées
                        for (var j = 0; j < founds.length; j++) {
                            try {
                                var found = founds[j];
                                
                                if (!ErrorHandler.ensureDefined(found, "found at index " + j, false)) {
                                    continue;
                                }
                                
                                found.appliedCharacterStyle = style;
                            } catch (foundError) {
                                ErrorHandler.handleError(foundError, "applying style to found " + j, false);
                                // Continuer avec l'élément suivant
                            }
                        }
                    } catch (storyError) {
                        ErrorHandler.handleError(storyError, "traitement de l'histoire " + i, false);
                        // Continuer avec l'histoire suivante
                    }
                }
                
                // Réinitialiser les options de recherche à la fin
                app.findChangeGrepOptions.includeFootnotes = true;
            } catch (error) {
                ErrorHandler.handleError(error, "applyNoteStyle", false);
            }
        },
        
        /**
         * Corrige les espaces typographiques autour de la ponctuation et des guillemets
         * @param {Document} doc - Document InDesign
         * @param {string} spaceType - Code d'espace spécial InDesign
         */
        fixTypoSpaces: function(doc, spaceType, profile) {
            try {
                if (!ErrorHandler.ensureDefined(doc, "document", true)) return;
                if (!ErrorHandler.ensureDefined(app, "app", true)) return;
                if (!ErrorHandler.ensureDefined(app.findGrepPreferences, "app.findGrepPreferences", true)) return;
                if (!ErrorHandler.ensureDefined(app.changeGrepPreferences, "app.changeGrepPreferences", true)) return;

                // Read per-punctuation space types from language profile, fallback to spaceType
                var spaceOpenQuote = spaceType;
                var spaceCloseQuote = spaceType;
                var spaceSemicolon = spaceType;
                var spaceColon = spaceType;
                var spaceExclamation = spaceType;
                var spaceQuestion = spaceType;

                if (profile) {
                    spaceOpenQuote = LanguageProfile.get("punctuation.spaceInsideOpenQuote", spaceType);
                    spaceCloseQuote = LanguageProfile.get("punctuation.spaceInsideCloseQuote", spaceType);
                    spaceSemicolon = LanguageProfile.get("punctuation.spaceBeforeSemicolon", spaceType);
                    spaceColon = LanguageProfile.get("punctuation.spaceBeforeColon", spaceType);
                    spaceExclamation = LanguageProfile.get("punctuation.spaceBeforeExclamation", spaceType);
                    spaceQuestion = LanguageProfile.get("punctuation.spaceBeforeQuestion", spaceType);
                }

                // Skip entirely if no spaces are needed before any punctuation
                var hasAnySpace = spaceOpenQuote || spaceCloseQuote || spaceSemicolon || spaceColon || spaceExclamation || spaceQuestion;
                if (!hasAnySpace) return;

                Utilities.resetPreferences();

                // Space after opening quote
                if (spaceOpenQuote) {
                    app.findGrepPreferences.findWhat = CONFIG.REGEX.SPACE_AFTER_OPENING_QUOTE;
                    app.changeGrepPreferences.changeTo = spaceOpenQuote;
                    doc.changeGrep();
                    Utilities.resetPreferences();
                }

                // Space before closing quote
                if (spaceCloseQuote) {
                    app.findGrepPreferences.findWhat = CONFIG.REGEX.SPACE_BEFORE_CLOSING_QUOTE;
                    app.changeGrepPreferences.changeTo = spaceCloseQuote;
                    doc.changeGrep();
                    Utilities.resetPreferences();
                }

                // Build per-punctuation replacement pairs
                // Each punctuation mark may have a different space type
                var punctPairs = [];
                if (spaceSemicolon) punctPairs.push({ char: ";", space: spaceSemicolon });
                if (spaceColon) punctPairs.push({ char: ":", space: spaceColon });
                if (spaceExclamation) punctPairs.push({ char: "!", space: spaceExclamation });
                if (spaceQuestion) punctPairs.push({ char: "\\?", space: spaceQuestion });

                if (punctPairs.length > 0) {
                    // Group punctuation marks that share the same space type
                    var spaceGroups = {};
                    for (var i = 0; i < punctPairs.length; i++) {
                        var key = punctPairs[i].space;
                        if (!spaceGroups[key]) spaceGroups[key] = [];
                        spaceGroups[key].push(punctPairs[i].char);
                    }

                    for (var sp in spaceGroups) {
                        if (spaceGroups.hasOwnProperty(sp)) {
                            var chars = spaceGroups[sp].join("");
                            // Replace existing wrong-type spaces before these characters
                            Utilities.resetPreferences();
                            app.findGrepPreferences.findWhat = "[ \\t\\u00A0\\u2000-\\u200A\\u202F\\u205F\\u3000](?=[" + chars + "])";
                            app.changeGrepPreferences.changeTo = sp;
                            doc.changeGrep();

                            // Insert space when character is directly adjacent
                            Utilities.resetPreferences();
                            app.findGrepPreferences.findWhat = "([^\\s])(?=[" + chars + "])";
                            app.changeGrepPreferences.changeTo = "$1" + sp;
                            doc.changeGrep();
                        }
                    }
                }
            } catch (error) {
                ErrorHandler.handleError(error, "fixTypoSpaces", false);
            }
        },
        
        /**
         * Remplace les tirets cadratin par des tirets demi-cadratin
         * @param {Document} doc - Document InDesign
         */
        replaceDashes: function(doc, profile) {
            try {
                if (!ErrorHandler.ensureDefined(doc, "document", true)) return;
                if (!ErrorHandler.ensureDefined(app, "app", true)) return;
                if (!ErrorHandler.ensureDefined(app.findGrepPreferences, "app.findGrepPreferences", true)) return;
                if (!ErrorHandler.ensureDefined(app.changeGrepPreferences, "app.changeGrepPreferences", true)) return;

                // Check if profile says to skip this operation
                if (profile && LanguageProfile.get("dashes.replaceCadratinWithDemiCadratin") === false) {
                    return;
                }

                Utilities.resetPreferences();
                app.findGrepPreferences.findWhat = "\u2014"; // em dash —
                app.changeGrepPreferences.changeTo = "\u2013"; // en dash –
                doc.changeGrep();
            } catch (error) {
                ErrorHandler.handleError(error, "replaceDashes", false);
            }
        },
        
        /**
         * Corrige les tirets isolés en les transformant en tirets demi-cadratin
         * Version sans dictionnaire qui cible seulement les tirets vraiment isolés
         * @param {Document} doc - Document InDesign
         */
        fixIsolatedHyphens: function(doc) {
            try {
                if (!ErrorHandler.ensureDefined(doc, "document", true)) return;
                if (!ErrorHandler.ensureDefined(app, "app", true)) return;
                if (!ErrorHandler.ensureDefined(app.findGrepPreferences, "app.findGrepPreferences", true)) return;
                if (!ErrorHandler.ensureDefined(app.changeGrepPreferences, "app.changeGrepPreferences", true)) return;
                
                // Le tiret demi-cadratin dans Unicode
                var ENDASH = "\u2013"; // –
                
                // Réinitialiser les préférences de recherche
                Utilities.resetPreferences();
                
                // CAS 1: Tirets isolés entre deux espaces normaux
                // Simplement: espace tiret espace
                app.findGrepPreferences.findWhat = " - ";
                app.changeGrepPreferences.changeTo = " " + ENDASH + " ";
                doc.changeGrep();
                
                // CAS 2: Tirets en début de paragraphe
                // Simplement: début de paragraphe, tiret, espace
                Utilities.resetPreferences();
                app.findGrepPreferences.findWhat = "^- ";
                app.changeGrepPreferences.changeTo = ENDASH + " ";
                doc.changeGrep();
                
                // CAS 3: Tirets après tabulation (pour les listes)
                Utilities.resetPreferences();
                app.findGrepPreferences.findWhat = "\\t- ";
                app.changeGrepPreferences.changeTo = "\\t" + ENDASH + " ";
                doc.changeGrep();
                
                // Réinitialiser les préférences à la fin
                Utilities.resetPreferences();
            } catch (error) {
                ErrorHandler.handleError(error, "fixIsolatedHyphens", false);
            }
        },
        
        /**
         * Remplace les tirets simples par des tirets demi-cadratin dans les intervalles de valeurs
         * Version de base ultra-simplifiée qui devrait fonctionner dans tous les cas
         * @param {Document} doc - Document InDesign
         */
        fixValueRanges: function(doc) {
            try {
                if (!ErrorHandler.ensureDefined(doc, "document", true)) return;
                if (!ErrorHandler.ensureDefined(app, "app", true)) return;
                if (!ErrorHandler.ensureDefined(app.findGrepPreferences, "app.findGrepPreferences", true)) return;
                if (!ErrorHandler.ensureDefined(app.changeGrepPreferences, "app.changeGrepPreferences", true)) return;
                
                // Le tiret demi-cadratin dans Unicode
                var ENDASH = "\u2013"; // –
                
                // Cas 1: Années (1990-2000)
                Utilities.resetPreferences();
                app.findGrepPreferences.findWhat = "(\\d{4})-(\\d{4})";
                app.changeGrepPreferences.changeTo = "$1" + ENDASH + "$2";
                doc.changeGrep();
                
                // Cas 2: Pages (p. 10-20)
                Utilities.resetPreferences();
                app.findGrepPreferences.findWhat = "(p\\.) (\\d+)-(\\d+)";
                app.changeGrepPreferences.changeTo = "$1 $2" + ENDASH + "$3";
                doc.changeGrep();
                
                // Cas 3: Pages (pages 10-20)
                Utilities.resetPreferences();
                app.findGrepPreferences.findWhat = "(pages) (\\d+)-(\\d+)";
                app.changeGrepPreferences.changeTo = "$1 $2" + ENDASH + "$3";
                doc.changeGrep();
                
                // Cas 4: Heures (10h-12h)
                Utilities.resetPreferences();
                app.findGrepPreferences.findWhat = "(\\d+h)-(\\d+h)";
                app.changeGrepPreferences.changeTo = "$1" + ENDASH + "$2";
                doc.changeGrep();
                
                // Cas 5: Heures avec minutes (10h30-12h45)
                Utilities.resetPreferences();
                app.findGrepPreferences.findWhat = "(\\d+h\\d+)-(\\d+h\\d+)";
                app.changeGrepPreferences.changeTo = "$1" + ENDASH + "$2";
                doc.changeGrep();
                
                // Cas 6: Tableaux (tableau 1-3)
                Utilities.resetPreferences();
                app.findGrepPreferences.findWhat = "(tableau) (\\d+)-(\\d+)";
                app.changeGrepPreferences.changeTo = "$1 $2" + ENDASH + "$3";
                doc.changeGrep();
                
                // Cas 7: Figures (fig. 1-3)
                Utilities.resetPreferences();
                app.findGrepPreferences.findWhat = "(fig\\.) (\\d+)-(\\d+)";
                app.changeGrepPreferences.changeTo = "$1 $2" + ENDASH + "$3";
                doc.changeGrep();
                
                // Cas 8: De A à Z
                Utilities.resetPreferences();
                app.findGrepPreferences.findWhat = "([A-Z])-([A-Z])";
                app.changeGrepPreferences.changeTo = "$1" + ENDASH + "$2";
                doc.changeGrep();
                
                // Réinitialiser les préférences à la fin
                Utilities.resetPreferences();
            } catch (error) {
                ErrorHandler.handleError(error, "fixValueRanges", false);
            }
        },
        
        /**
         * Applique un style au texte avec formatage italique
         * @param {Document} doc - Document InDesign
         * @param {string} styleName - Nom du style de caractère
         */
        applyItalicStyle: function(doc, styleName) {
            try {
                if (!ErrorHandler.ensureDefined(doc, "document", true)) return;
                if (!ErrorHandler.ensureDefined(styleName, "styleName", true)) return;
                if (!ErrorHandler.ensureDefined(app, "app", true)) return;
                if (!ErrorHandler.ensureDefined(app.findTextPreferences, "app.findTextPreferences", true)) return;
                if (!ErrorHandler.ensureDefined(app.changeTextPreferences, "app.changeTextPreferences", true)) return;
                
                // Récupérer ou créer le style
                var style = Utilities.getOrCreateStyle(doc, styleName, { fontStyle: "Italic" });
                
                if (!ErrorHandler.ensureDefined(style, "style", false)) {
                    return;
                }
                
                // Appliquer le style au texte en italique
                Utilities.resetPreferences();
                app.findTextPreferences.fontStyle = "Italic";
                app.changeTextPreferences.appliedCharacterStyle = style;
                doc.changeText();
            } catch (error) {
                ErrorHandler.handleError(error, "applyItalicStyle", false);
            }
        },
        
        /**
         * Applique un style au texte avec formatage exposant
         * @param {Document} doc - Document InDesign
         * @param {string} styleName - Nom du style de caractère
         */
        applyExposantStyle: function(doc, styleName) {
            try {
                if (!ErrorHandler.ensureDefined(doc, "document", true)) return;
                if (!ErrorHandler.ensureDefined(styleName, "styleName", true)) return;
                if (!ErrorHandler.ensureDefined(app, "app", true)) return;
                if (!ErrorHandler.ensureDefined(app.findTextPreferences, "app.findTextPreferences", true)) return;
                if (!ErrorHandler.ensureDefined(app.changeTextPreferences, "app.changeTextPreferences", true)) return;
                
                // Récupérer ou créer le style
                var style = Utilities.getOrCreateStyle(doc, styleName, { position: Position.SUPERSCRIPT });
                
                if (!ErrorHandler.ensureDefined(style, "style", false)) {
                    return;
                }
                
                // Appliquer le style au texte en exposant
                Utilities.resetPreferences();
                app.findTextPreferences.position = Position.SUPERSCRIPT;
                app.changeTextPreferences.appliedCharacterStyle = style;
                doc.changeText();
            } catch (error) {
                ErrorHandler.handleError(error, "applyExposantStyle", false);
            }
        },
        
        /**
         * Supprime les doubles retours à la ligne
         * @param {Document} doc - Document InDesign
         */
        removeDoubleReturns: function(doc) {
            try {
                if (!ErrorHandler.ensureDefined(doc, "document", true)) return;
                if (!ErrorHandler.ensureDefined(app, "app", true)) return;
                if (!ErrorHandler.ensureDefined(app.findGrepPreferences, "app.findGrepPreferences", true)) return;
                if (!ErrorHandler.ensureDefined(app.changeGrepPreferences, "app.changeGrepPreferences", true)) return;
                
                Utilities.resetPreferences();
                
                app.findGrepPreferences.findWhat = CONFIG.REGEX.DOUBLE_RETURNS;
                app.changeGrepPreferences.changeTo = "\\r";
                doc.changeGrep();
            } catch (error) {
                ErrorHandler.handleError(error, "removeDoubleReturns", false);
            }
        },
        
        /**
         * Convertit trois points consécutifs en caractère points de suspension
         * @param {Document} doc - Document InDesign
         */
        convertEllipsis: function(doc) {
            try {
                if (!ErrorHandler.ensureDefined(doc, "document", true)) return;
                if (!ErrorHandler.ensureDefined(app, "app", true)) return;
                if (!ErrorHandler.ensureDefined(app.findGrepPreferences, "app.findGrepPreferences", true)) return;
                if (!ErrorHandler.ensureDefined(app.changeGrepPreferences, "app.changeGrepPreferences", true)) return;
                
                Utilities.resetPreferences();
                
                app.findGrepPreferences.findWhat = CONFIG.REGEX.TRIPLE_DOTS;
                app.changeGrepPreferences.changeTo = "…"; // U+2026 HORIZONTAL ELLIPSIS
                doc.changeGrep();
            } catch (error) {
                ErrorHandler.handleError(error, "convertEllipsis", false);
            }
        },
        
        /**
         * Remplace les apostrophes droites par des apostrophes typographiques
         * @param {Document} doc - Document InDesign
         * @returns {number} Nombre total d'apostrophes remplacées
         */
        replaceApostrophes: function(doc) {
            try {
                if (!ErrorHandler.ensureDefined(doc, "document", true)) return 0;
                
                var totalReplaced = 0;
                
                // 1. Sauvegarder les options de recherche actuelles
                var savedOptions = {};
                try {
                    savedOptions = {
                        caseSensitive: app.findChangeTextOptions.caseSensitive,
                        wholeWord: app.findChangeTextOptions.wholeWord,
                        includeFootnotes: app.findChangeTextOptions.includeFootnotes,
                        includeHiddenLayers: app.findChangeTextOptions.includeHiddenLayers,
                        includeMasterPages: app.findChangeTextOptions.includeMasterPages,
                        includeLockedLayersForFind: app.findChangeTextOptions.includeLockedLayersForFind,
                        includeLockedStoriesForFind: app.findChangeTextOptions.includeLockedStoriesForFind
                    };
                } catch (e) {
                    // Continuer même si certaines options ne peuvent pas être sauvegardées
                }
                
                // 2. Réinitialiser à la fois les préférences GREP et Text
                app.findGrepPreferences = NothingEnum.NOTHING;
                app.changeGrepPreferences = NothingEnum.NOTHING;
                app.findTextPreferences = NothingEnum.NOTHING;
                app.changeTextPreferences = NothingEnum.NOTHING;
                
                // 3. Configurer explicitement les options de recherche
                app.findChangeTextOptions.caseSensitive = false;
                app.findChangeTextOptions.wholeWord = false;
                app.findChangeTextOptions.includeFootnotes = true;
                app.findChangeTextOptions.includeHiddenLayers = true;
                app.findChangeTextOptions.includeMasterPages = true;
                app.findChangeTextOptions.includeLockedLayersForFind = true;
                app.findChangeTextOptions.includeLockedStoriesForFind = true;
                
                try {
                    // 4. Rechercher uniquement l'apostrophe droite (le cas le plus courant dans InDesign)
                    app.findTextPreferences.findWhat = "'";
                    app.changeTextPreferences.changeTo = "\u2019"; // apostrophe typographique
                    
                    // 5. Effectuer le remplacement global
                    var foundItems = doc.changeText();
                    totalReplaced = foundItems.length;
                    
                    // Version sans alerte pour production
                    // Si vous préférez conserver l'alerte, décommentez la ligne ci-dessous
                    try {
                        if (typeof console !== "undefined" && console && console.log) {
                            console.log("Apostrophes replaced: " + totalReplaced);
                        }
                    } catch (e) {
                        // Ignorer les erreurs de journalisation
                    }
                } catch (searchError) {
                    ErrorHandler.handleError(searchError, "replaceApostrophes - search operation", false);
                }
                
                // 6. Restaurer les options originales
                try {
                    for (var opt in savedOptions) {
                        if (Object.prototype.hasOwnProperty.call(savedOptions, opt)) {
                            app.findChangeTextOptions[opt] = savedOptions[opt];
                        }
                    }
                } catch (e) {
                    // Continuer même si la restauration échoue
                }
                
                // 7. Réinitialiser à nouveau les deux types de préférences
                app.findGrepPreferences = NothingEnum.NOTHING;
                app.changeGrepPreferences = NothingEnum.NOTHING;
                app.findTextPreferences = NothingEnum.NOTHING;
                app.changeTextPreferences = NothingEnum.NOTHING;
                
                return totalReplaced;
            } catch (error) {
                ErrorHandler.handleError(error, "replaceApostrophes", false);
                // Alerte seulement en cas d'erreur grave
                alert(I18n.__("errorReplaceApostrophes", error.message));
                return 0;
            }
        },
        
        /**
         * Applique un style de paragraphe après les paragraphes déclencheurs
         * @param {Document} doc - Document InDesign
         * @param {Array} triggerStyles - Styles déclencheurs
         * @param {ParagraphStyle} targetStyle - Style à appliquer
         */
        applyStyleAfterTriggers: function(doc, triggerStyles, targetStyle) {
            try {
                if (!ErrorHandler.ensureDefined(doc, "document", true)) return;
                if (!ErrorHandler.ensureDefined(triggerStyles, "triggerStyles", true)) return;
                if (!ErrorHandler.ensureDefined(targetStyle, "targetStyle", true)) return;
                
                // Vérifier que les collections d'objets sont accessibles
                if (!ErrorHandler.ensureDefined(doc.stories, "document.stories", true)) return;
                
                // Parcourir les articles du document
                var stories = doc.stories;
                for (var s = 0; s < stories.length; s++) {
                    try {
                        var story = stories[s];
                        
                        if (!ErrorHandler.ensureDefined(story, "story at index " + s, false)) {
                            continue;
                        }
                        
                        if (!ErrorHandler.ensureDefined(story.paragraphs, "story.paragraphs at index " + s, false)) {
                            continue;
                        }
                        
                        var paras = story.paragraphs;
                        var justLeftTriggerBlock = false;
                        
                        // Parcourir les paragraphes de l'article
                        for (var p = 0; p < paras.length; p++) {
                            try {
                                var para = paras[p];
                                
                                if (!ErrorHandler.ensureDefined(para, "paragraph at index " + p, false)) {
                                    continue;
                                }
                                
                                if (!ErrorHandler.ensureDefined(para.appliedParagraphStyle, "para.appliedParagraphStyle at index " + p, false)) {
                                    continue;
                                }
                                
                                var style = para.appliedParagraphStyle;
                                var isTriggerStyle = false;
                                
                                // Vérifier si le style actuel est un déclencheur
                                for (var t = 0; t < triggerStyles.length; t++) {
                                    if (style.id === triggerStyles[t].id) {
                                        isTriggerStyle = true;
                                        break;
                                    }
                                }
                                
                                if (isTriggerStyle) {
                                    justLeftTriggerBlock = true;
                                } else if (justLeftTriggerBlock) {
                                    // Si on vient de quitter un bloc de déclencheurs et que le paragraphe n'est pas vide
                                    if (!Utilities.isEmptyParagraph(para)) {
                                        // Appliquer le style cible
                                        para.appliedParagraphStyle = targetStyle;
                                        justLeftTriggerBlock = false; // Réinitialiser pour ne pas affecter plusieurs paragraphes consécutifs
                                    }
                                }
                            } catch (paraError) {
                                ErrorHandler.handleError(paraError, "traitement du paragraphe " + p, false);
                                // Continuer avec le paragraphe suivant
                            }
                        }
                    } catch (storyError) {
                        ErrorHandler.handleError(storyError, "traitement de l'article " + s, false);
                        // Continuer avec l'article suivant
                    }
                }
            } catch (error) {
                ErrorHandler.handleError(error, "applyStyleAfterTriggers", false);
            }
        },
        
        /**
         * Corrige les espaces autour des incises avec tirets demi-cadratins
         * @param {Document} doc - Document InDesign
         * @param {string} spaceType - Code d'espace spécial InDesign
         */
        fixDashIncises: function(doc, spaceType, profile) {
            try {
                if (!ErrorHandler.ensureDefined(doc, "document", true)) return;
                if (!ErrorHandler.ensureDefined(app, "app", true)) return;
                if (!ErrorHandler.ensureDefined(app.findGrepPreferences, "app.findGrepPreferences", true)) return;
                if (!ErrorHandler.ensureDefined(app.changeGrepPreferences, "app.changeGrepPreferences", true)) return;

                // Tiret demi-cadratin
                var ENDASH = "\u2013"; // –

                // Read incise space from profile if available
                var inciseSpaceConfig = spaceType;
                if (profile) {
                    inciseSpaceConfig = LanguageProfile.get("dashes.inciseSpace", spaceType);
                }

                // If profile says no space (empty string), skip this correction
                if (!inciseSpaceConfig) return;

                // Determine the actual Unicode character to insert
                var insecableChar;
                if (inciseSpaceConfig === "~<") {
                    insecableChar = "\u202F"; // Fine non-breaking space
                } else if (inciseSpaceConfig === "~S") {
                    insecableChar = "\u00A0"; // Non-breaking space
                } else if (inciseSpaceConfig === " ") {
                    insecableChar = " "; // Normal (breakable) space — for FR-CH, EN-UK, DE, etc.
                } else {
                    insecableChar = (spaceType === "~<") ? "\u202F" : "\u00A0";
                }
                
                // Réinitialiser les préférences
                Utilities.resetPreferences();
                app.findChangeGrepOptions.includeFootnotes = true;
                
                // 1. Gestion des incises (paires de tirets)
                // 1.1 Trouver les paires de tirets (tiret + espace + texte + espace + tiret)
                app.findGrepPreferences.findWhat = ENDASH + " ([^" + ENDASH + "\\r\\n]*?) " + ENDASH;
                app.changeGrepPreferences.changeTo = ENDASH + insecableChar + "$1" + insecableChar + ENDASH;
                doc.changeGrep();
                
                // 1.2 Traiter les tirets de dialogue en début de paragraphe
                Utilities.resetPreferences();
                app.findChangeGrepOptions.includeFootnotes = true;
                app.findGrepPreferences.findWhat = "^" + ENDASH + " ";
                app.changeGrepPreferences.changeTo = ENDASH + insecableChar;
                doc.changeGrep();
                
                // 1.3 Traiter les tirets ouvrants isolés (non suivis d'un autre tiret dans le paragraphe)
                Utilities.resetPreferences();
                app.findChangeGrepOptions.includeFootnotes = true;
                app.findGrepPreferences.findWhat = ENDASH + " ";
                
                var tousLesTirets = doc.findGrep();
                for (var i = 0; i < tousLesTirets.length; i++) {
                    var item = tousLesTirets[i];
                    var paragraphText = item.paragraphs[0].contents;
                    var itemPosition = paragraphText.indexOf(item.contents);
                    var positionApresTiret = itemPosition + item.contents.length;
                    
                    // Si ce tiret est le dernier du paragraphe, c'est un tiret ouvrant isolé
                    if (paragraphText.indexOf(ENDASH, positionApresTiret) === -1) {
                        // Utiliser findGrep/changeGrep localement sur cette occurrence
                        Utilities.resetPreferences();
                        app.findGrepPreferences.findWhat = ENDASH + " ";
                        app.changeGrepPreferences.changeTo = ENDASH + insecableChar;
                        item.changeGrep();
                    }
                }
                
                // 1.4 Traiter les tirets fermants isolés (non précédés d'un autre tiret dans le paragraphe)
                Utilities.resetPreferences();
                app.findChangeGrepOptions.includeFootnotes = true;
                app.findGrepPreferences.findWhat = " " + ENDASH;
                
                var tiretsFermants = doc.findGrep();
                for (var j = 0; j < tiretsFermants.length; j++) {
                    var fermant = tiretsFermants[j];
                    var paraText = fermant.paragraphs[0].contents;
                    var fermantPosition = paraText.indexOf(fermant.contents);
                    
                    // Si ce tiret est le premier du paragraphe, c'est un tiret fermant isolé
                    var precedingText = paraText.substring(0, fermantPosition);
                    if (precedingText.indexOf(ENDASH) === -1) {
                        // Utiliser findGrep/changeGrep localement sur cette occurrence
                        Utilities.resetPreferences();
                        app.findGrepPreferences.findWhat = " " + ENDASH;
                        app.changeGrepPreferences.changeTo = insecableChar + ENDASH;
                        fermant.changeGrep();
                    }
                }
                
                // Réinitialiser les préférences à la fin
                Utilities.resetPreferences();
            } catch (error) {
                ErrorHandler.handleError(error, "fixDashIncises", false);
            }
        },
        
        /**
         * Formate les nombres en respectant les conventions typographiques françaises
         * @param {Document} doc - Document InDesign
         * @param {boolean} addSpaces - Si true, ajoute des espaces insécables entre les milliers
         * @param {boolean} useComma - Si true, remplace les points décimaux par des virgules
         * @param {boolean} excludeYears - Si true, exclut les nombres entre 0 et 2050
         */
        formatNumbers: function(doc, addSpaces, useComma, excludeYears, profile) {
            try {
                if (!ErrorHandler.ensureDefined(doc, "document", true)) return;
                if (!ErrorHandler.ensureDefined(app, "app", true)) return;
                if (!ErrorHandler.ensureDefined(app.findGrepPreferences, "app.findGrepPreferences", true)) return;
                if (!ErrorHandler.ensureDefined(app.changeGrepPreferences, "app.changeGrepPreferences", true)) return;

                // Définir le tiret demi-cadratin
                var ENDASH = "\u2013"; // –

                // Read thousands separator from profile, default to fine non-breaking space
                var SEPARATEUR_MILLIERS = "~<";
                if (profile) {
                    SEPARATEUR_MILLIERS = LanguageProfile.get("numbers.thousandsSeparator", "~<");
                }
                
                if (addSpaces) {
                    // CORRECTION: Protéger d'abord les intervalles d'années avec tiret demi-cadratin
                    if (excludeYears) {
                        Utilities.resetPreferences();
                        
                        // 1. Protéger les intervalles d'années avec tiret demi-cadratin déjà formatées
                        app.findGrepPreferences.findWhat = "([0-1]\\d{3})" + ENDASH + "([0-1]\\d{3}|20[0-4]\\d|2050)";
                        app.changeGrepPreferences.changeTo = "YEAR_ENDASH_$1_ENDASH_$2";  
                        doc.changeGrep();
                        
                        // 2. Protéger les années entre parenthèses avec tiret demi-cadratin
                        app.findGrepPreferences.findWhat = "\\(([0-1]\\d{3})" + ENDASH + "([0-1]\\d{3}|20[0-4]\\d|2050)\\)";
                        app.changeGrepPreferences.changeTo = "(YEAR_ENDASH_$1_ENDASH_$2)";  
                        doc.changeGrep();
                        
                        // Première passe : protéger les plages d'années connectées par des tirets ou traits d'union
                        Utilities.resetPreferences();
                        app.findGrepPreferences.findWhat = "\\b([0-1]\\d{3})[-–—]([0-1]\\d{3}|20[0-4]\\d|2050)\\b";
                        app.changeGrepPreferences.changeTo = "YEAR_MARK_$1_YEAR_MARK-YEAR_MARK_$2_YEAR_MARK";
                        doc.changeGrep();
                        
                        // Deuxième passe : protéger les années individuelles restantes
                        Utilities.resetPreferences();
                        app.findGrepPreferences.findWhat = "\\b([0-1]\\d{3})\\b|\\b(20[0-4]\\d)\\b|\\b(2050)\\b";
                        app.changeGrepPreferences.changeTo = "YEAR_MARK_$1$2$3_YEAR_MARK";
                        doc.changeGrep();
                        
                        // Troisième passe : protéger les années dans les parenthèses (comme dans votre exemple)
                        Utilities.resetPreferences();
                        app.findGrepPreferences.findWhat = "\\(([0-1]\\d{3})[-–—]([0-1]\\d{3}|20[0-4]\\d|2050)\\)";
                        app.changeGrepPreferences.changeTo = "(YEAR_MARK_$1_YEAR_MARK-YEAR_MARK_$2_YEAR_MARK)";
                        doc.changeGrep();
                    }
                    
                    // Protéger nombres décimaux avant tout traitement
                    Utilities.resetPreferences();
                    if (useComma) {
                        app.findGrepPreferences.findWhat = "\\b(\\d+),(\\d+)\\b";
                    } else {
                        app.findGrepPreferences.findWhat = "\\b(\\d+)\\.(\\d+)\\b";
                    }
                    app.changeGrepPreferences.changeTo = "###$1###$2###";
                    doc.changeGrep();
                    
                    // Traiter d'abord les grands nombres sans séparateurs (4+ chiffres)
                    // mais qui ne sont pas protégés (années ou décimaux)
                    Utilities.resetPreferences();
                    app.findGrepPreferences.findWhat = "\\b(\\d{4,})\\b";
                    
                    var largeNumbers = doc.findGrep();
                    for (var ln = 0; ln < largeNumbers.length; ln++) {
                        try {
                            var numStr = largeNumbers[ln].contents;
                            
                            // Vérifier que ce n'est pas un nombre déjà protégé
                            if (numStr.indexOf("YEAR_MARK_") !== -1 || numStr.indexOf("YEAR_ENDASH_") !== -1 || numStr.indexOf("###") !== -1) {
                                continue;
                            }
                            
                            // Formater en préservant la méthode d'origine
                            var formattedNum = "";
                            var len = numStr.length;
                            var mod = len % 3;
                            
                            if (mod > 0) {
                                formattedNum = numStr.substring(0, mod);
                            } else {
                                formattedNum = numStr.substring(0, 3);
                                mod = 3;
                            }
                            
                            for (var j = mod; j < len; j += 3) {
                                formattedNum += " " + numStr.substring(j, j + 3);
                            }
                            
                            largeNumbers[ln].contents = formattedNum;
                        } catch (e) {
                            continue;
                        }
                    }
                    
                    // Traiter les nombres avec un séparateur suivi de 4+ chiffres (ex: 5 000000)
                    Utilities.resetPreferences();
                    app.findGrepPreferences.findWhat = "\\b(\\d+)[ \\t\\u00A0\\u202F~<~>~|~=](\\d{4,})\\b";
                    
                    var singleSepNumbers = doc.findGrep();
                    for (var i = 0; i < singleSepNumbers.length; i++) {
                        try {
                            var originalText = singleSepNumbers[i].contents;
                            
                            // Vérifier que ce n'est pas un nombre déjà protégé
                            if (originalText.indexOf("YEAR_MARK_") !== -1 || originalText.indexOf("YEAR_ENDASH_") !== -1 || originalText.indexOf("###") !== -1) {
                                continue;
                            }
                            
                            var matches = originalText.match(/(\d+)[ \t\u00A0\u202F~<~>~|~=](\d+)/);
                            
                            if (matches && matches.length === 3) {
                                var fullNumber = matches[1] + matches[2];
                                var tempKey = "SPECIAL_NUMBER_" + Math.random().toString(36).substring(2, 10) + "_";
                                
                                singleSepNumbers[i].contents = tempKey;
                                
                                var formattedNumber = "";
                                var len = fullNumber.length;
                                var mod = len % 3;
                                
                                if (mod > 0) {
                                    formattedNumber = fullNumber.substring(0, mod);
                                } else {
                                    formattedNumber = fullNumber.substring(0, 3);
                                    mod = 3;
                                }
                                
                                for (var j = mod; j < len; j += 3) {
                                    formattedNumber += " " + fullNumber.substring(j, j + 3);
                                }
                                
                                Utilities.resetPreferences();
                                app.findGrepPreferences.findWhat = tempKey;
                                app.changeGrepPreferences.changeTo = formattedNumber;
                                doc.changeGrep();
                            }
                        } catch (e) {
                            continue;
                        }
                    }
                    
                    // Traiter les nombres avec apostrophe (ex: 10'000)
                    Utilities.resetPreferences();
                    app.findGrepPreferences.findWhat = "\\b(\\d+)'(\\d{3,})\\b";
                    
                    var apostropheSepNumbers = doc.findGrep();
                    for (var i = 0; i < apostropheSepNumbers.length; i++) {
                        try {
                            var originalText = apostropheSepNumbers[i].contents;
                            
                            // Vérifier que ce n'est pas un nombre déjà protégé
                            if (originalText.indexOf("YEAR_MARK_") !== -1 || originalText.indexOf("YEAR_ENDASH_") !== -1 || originalText.indexOf("###") !== -1) {
                                continue;
                            }
                            
                            var parts = originalText.split("'");
                            
                            if (parts.length === 2) {
                                var fullNumber = parts[0] + parts[1];
                                var tempKey = "SPECIAL_NUMBER_" + Math.random().toString(36).substring(2, 10) + "_";
                                
                                apostropheSepNumbers[i].contents = tempKey;
                                
                                var formattedNumber = "";
                                var len = fullNumber.length;
                                var mod = len % 3;
                                
                                if (mod > 0) {
                                    formattedNumber = fullNumber.substring(0, mod);
                                } else {
                                    formattedNumber = fullNumber.substring(0, 3);
                                    mod = 3;
                                }
                                
                                for (var j = mod; j < len; j += 3) {
                                    formattedNumber += " " + fullNumber.substring(j, j + 3);
                                }
                                
                                Utilities.resetPreferences();
                                app.findGrepPreferences.findWhat = tempKey;
                                app.changeGrepPreferences.changeTo = formattedNumber;
                                doc.changeGrep();
                            }
                        } catch (e) {
                            continue;
                        }
                    }
                    
                    // Normaliser les séparateurs
                    // Traiter apostrophes
                    Utilities.resetPreferences();
                    app.findGrepPreferences.findWhat = "(\\d+)'(\\d{3})(?!'|\\d)"; 
                    app.changeGrepPreferences.changeTo = "$1" + SEPARATEUR_MILLIERS + "$2";
                    while (doc.changeGrep().length > 0) {}
                    
                    // Traiter points
                    Utilities.resetPreferences();
                    app.findGrepPreferences.findWhat = "(\\d+)\\.(\\d{3})(?![\\d,.])";
                    app.changeGrepPreferences.changeTo = "$1" + SEPARATEUR_MILLIERS + "$2";
                    while (doc.changeGrep().length > 0) {}
                    
                    // Normaliser les différents types d'espaces
                    var espacesPatterns = [
                        " (\\d{3})",                // Espace normal
                        "~>(\\d{3})",               // Espace fine
                        "~<(\\d{3})",               // Espace fine insécable
                        "~|(\\d{3})",               // Espace insécable alt
                        "~=(\\d{3})",               // Espace quart
                        "[\\u00A0\\u202F](\\d{3})" // Codes Unicode pour espaces insécables
                    ];
                    
                    for (var i = 0; i < espacesPatterns.length; i++) {
                        if (espacesPatterns[i].indexOf(SEPARATEUR_MILLIERS) === -1) {
                            Utilities.resetPreferences();
                            app.findGrepPreferences.findWhat = "(\\d+)" + espacesPatterns[i];
                            app.changeGrepPreferences.changeTo = "$1" + SEPARATEUR_MILLIERS + "$2";
                            while (doc.changeGrep().length > 0) {}
                        }
                    }
                    
                    // Traiter combinaisons mixtes
                    var mixedPatterns = [
                        "(\\d+)'(\\d{3})[\\s\\u00A0\\u202F~<~>~|~=](\\d{3})",  // Apostrophe puis espace
                        "(\\d+)[\\s\\u00A0\\u202F~<~>~|~=](\\d{3})'(\\d{3})",  // Espace puis apostrophe
                        "(\\d+)\\.(\\d{3})[\\s\\u00A0\\u202F~<~>~|~=](\\d{3})" // Point puis espace
                    ];
                    
                    for (var i = 0; i < mixedPatterns.length; i++) {
                        Utilities.resetPreferences();
                        app.findGrepPreferences.findWhat = mixedPatterns[i];
                        app.changeGrepPreferences.changeTo = "$1" + SEPARATEUR_MILLIERS + "$2" + SEPARATEUR_MILLIERS + "$3";
                        while (doc.changeGrep().length > 0) {}
                    }
                }
                
                // Remplacer points décimaux par virgules
                if (useComma) {
                    Utilities.resetPreferences();
                    app.findGrepPreferences.findWhat = "(\\d+)(?<!\\.)\\.((?!\\.)\\d+)";
                    app.changeGrepPreferences.changeTo = "$1,$2";
                    doc.changeGrep();
                }
                
                if (addSpaces) {
                    // Traiter nombres avec décimales qui ont été protégés
                    for (var nbDigits = 15; nbDigits >= 4; nbDigits--) {
                        var groupLength = nbDigits - 3;
                        
                        Utilities.resetPreferences();
                        var pattern = "###(\\d{" + groupLength + "})(\\d{3})###";
                        app.findGrepPreferences.findWhat = pattern;
                        app.changeGrepPreferences.changeTo = "###$1" + SEPARATEUR_MILLIERS + "$2###";
                        doc.changeGrep();
                    }
                    
                    // Restaurer nombres décimaux
                    Utilities.resetPreferences();
                    if (useComma) {
                        app.findGrepPreferences.findWhat = "###([^#]+)###([^#]+)###";
                        app.changeGrepPreferences.changeTo = "$1,$2";
                    } else {
                        app.findGrepPreferences.findWhat = "###([^#]+)###([^#]+)###";
                        app.changeGrepPreferences.changeTo = "$1.$2";
                    }
                    doc.changeGrep();
                    
                    // CORRECTION: Restaurer les différents types d'années protégées
                    if (excludeYears) {
                        // 1. D'abord restaurer les années avec tiret demi-cadratin
                        Utilities.resetPreferences();
                        app.findGrepPreferences.findWhat = "YEAR_ENDASH_(\\d+)_ENDASH_(\\d+)";
                        app.changeGrepPreferences.changeTo = "$1" + ENDASH + "$2";  // Restaure avec tiret demi-cadratin
                        doc.changeGrep();
                        
                        // 2. Restaurer les plages d'années avec tiret simple
                        Utilities.resetPreferences();
                        app.findGrepPreferences.findWhat = "YEAR_MARK_(\\d+)_YEAR_MARK-YEAR_MARK_(\\d+)_YEAR_MARK";
                        app.changeGrepPreferences.changeTo = "$1-$2";
                        doc.changeGrep();
                        
                        // 3. Restaurer les années individuelles
                        Utilities.resetPreferences();
                        app.findGrepPreferences.findWhat = "YEAR_MARK_(\\d+)_YEAR_MARK";
                        app.changeGrepPreferences.changeTo = "$1";
                        doc.changeGrep();
                    }
                    
                    // Traiter très grands nombres
                    for (var pass = 0; pass < 5; pass++) {
                        Utilities.resetPreferences();
                        app.findGrepPreferences.findWhat = "(\\d{1,3})" + SEPARATEUR_MILLIERS + "(\\d{3}" + SEPARATEUR_MILLIERS + "\\d{3})";
                        app.changeGrepPreferences.changeTo = "$1" + SEPARATEUR_MILLIERS + "$2";
                        
                        var foundItems = doc.findGrep();
                        if (foundItems.length === 0) break;
                        
                        for (var f = 0; f < foundItems.length; f++) {
                            var num = foundItems[f].contents;
                            if (num.match(/^\d{4,}/)) {
                                var parts = num.split(SEPARATEUR_MILLIERS);
                                var firstPart = parts[0];
                                
                                if (firstPart.length > 3) {
                                    var newFirstPart = "";
                                    var remaining = firstPart.length % 3;
                                    
                                    if (remaining > 0) {
                                        newFirstPart += firstPart.substring(0, remaining) + SEPARATEUR_MILLIERS;
                                    }
                                    
                                    for (var i = remaining; i < firstPart.length; i += 3) {
                                        newFirstPart += firstPart.substring(i, i + 3);
                                        if (i + 3 < firstPart.length) {
                                            newFirstPart += SEPARATEUR_MILLIERS;
                                        }
                                    }
                                    
                                    var newNum = newFirstPart;
                                    for (var i = 1; i < parts.length; i++) {
                                        newNum += SEPARATEUR_MILLIERS + parts[i];
                                    }
                                    
                                    foundItems[f].contents = newNum;
                                }
                            }
                        }
                    }
                }
                
                // Nettoyage final
                Utilities.resetPreferences();
                app.findGrepPreferences.findWhat = SEPARATEUR_MILLIERS + "{2,}";
                app.changeGrepPreferences.changeTo = SEPARATEUR_MILLIERS;
                doc.changeGrep();
                
                // CORRECTION: Inclure les nouveaux marqueurs dans le nettoyage
                Utilities.resetPreferences();
                app.findGrepPreferences.findWhat = "###|YEAR_MARK_|_YEAR_MARK|YEAR_ENDASH_|_ENDASH_|SPECIAL_NUMBER_[a-z0-9]+_";
                app.changeGrepPreferences.changeTo = "";
                doc.changeGrep();
                
                // Convertir espaces normaux en espaces fines insécables
                Utilities.resetPreferences();
                app.findGrepPreferences.findWhat = "(\\d+) (\\d{3})";
                app.changeGrepPreferences.changeTo = "$1" + SEPARATEUR_MILLIERS + "$2";
                while (doc.changeGrep().length > 0) {}
                
            } catch (error) {
                ErrorHandler.handleError(error, "formatNumbers", false);
            }
        }
    };
    
    /**
     * Module de barre de progression pour SuperScript
     * @private
     */
    var ProgressBar = {
        /**
         * Référence à la fenêtre de progression
         * @private
         */
        progressWin: null,
        
        /**
         * Crée une barre de progression
         * @param {String} title - Titre de la boîte de dialogue
         * @param {Number} maxValue - Valeur maximale pour la barre de progression
         */
        create: function(title, maxValue) {
            try {
                this.progressWin = new Window("palette", title);
                this.progressWin.progressBar = this.progressWin.add("progressbar", undefined, 0, maxValue);
                this.progressWin.progressBar.preferredSize.width = 300;
                this.progressWin.status = this.progressWin.add("statictext", undefined, "");
                this.progressWin.status.preferredSize.width = 300;
                
                // Centrer la fenêtre
                this.progressWin.center();
                this.progressWin.show();
            } catch (e) {
                ErrorHandler.handleError(e, "creating progress bar", false);
                // Continuer sans barre de progression
                this.progressWin = null;
            }
        },
        
        /**
         * Met à jour la barre de progression
         * @param {Number} value - Valeur actuelle de la progression
         * @param {String} statusText - Texte de statut à afficher
         */
        update: function(value, statusText) {
            if (!this.progressWin) return;
            
            try {
                this.progressWin.progressBar.value = value;
                this.progressWin.status.text = statusText;
                this.progressWin.update();
            } catch (e) {
                ErrorHandler.handleError(e, "updating progress bar", false);
            }
        },
        
        /**
         * Ferme la barre de progression
         */
        close: function() {
            if (this.progressWin) {
                try {
                    this.progressWin.close();
                    this.progressWin = null;
                } catch (e) {
                    ErrorHandler.handleError(e, "fermeture de la barre de progression", false);
                }
            }
        }
    };
    
    /**
     * Constructeur d'interface utilisateur pour le dialogue du script
     * @private
     */
    var UIBuilder = {
      /**
       * Crée et affiche le dialogue du script
       * @param {Array} characterStyles - Tableau des styles de caractère disponibles
       * @param {number} noteStyleIndex - Index du style de note par défaut
       * @param {number} italicStyleIndex - Index du style italique par défaut
       * @returns {Object} Résultat du dialogue avec les options de l'utilisateur
       */
      createDialog: function(characterStyles, noteStyleIndex, italicStyleIndex) {
        try {
        if (!ErrorHandler.ensureDefined(characterStyles, "characterStyles", true)) {
          characterStyles = [I18n.__("defaultStyle")];
        }
        
        // Création du dialogue principal
        var dialog = new Window("dialog", CONFIG.SCRIPT_TITLE);
        dialog.orientation = "column";
        dialog.alignChildren = "fill";
        dialog.preferredSize.width = 400;
        
        // Bannière supérieure avec attribution
        var topBanner = dialog.add("group");
        topBanner.orientation = "row";
        topBanner.alignment = "right";
        var attribution = topBanner.add("statictext", undefined, "entremonde / Spectral lab");

        // Language profile selector
        var profileGroup = dialog.add("group");
        profileGroup.orientation = "row";
        profileGroup.alignment = "fill";
        profileGroup.add("statictext", undefined, I18n.__("languageProfileLabel"));
        var profileDropdown = profileGroup.add("dropdownlist");
        profileDropdown.preferredSize.width = 200;

        var availableProfiles = LanguageProfile.getAvailableProfiles();
        var defaultProfileId = LanguageProfile.getDefaultProfileId();
        var defaultProfileIndex = 0;

        if (availableProfiles.length > 0) {
            for (var pi = 0; pi < availableProfiles.length; pi++) {
                var displayLabel = (I18n.getLanguage() === 'fr')
                    ? availableProfiles[pi].label
                    : availableProfiles[pi].labelEN;
                profileDropdown.add("item", displayLabel);
                if (availableProfiles[pi].id === defaultProfileId) {
                    defaultProfileIndex = pi;
                }
            }
            profileDropdown.selection = defaultProfileIndex;
        } else {
            profileDropdown.add("item", I18n.__("languageProfileNone"));
            profileDropdown.selection = 0;
        }

        // Load the default profile
        if (availableProfiles.length > 0) {
            LanguageProfile.load(availableProfiles[defaultProfileIndex].id);
        }

        // Configuration bar: Save / Load buttons + auto-detected indicator
        var configBar = dialog.add("group");
        configBar.orientation = "row";
        configBar.alignment = "fill";
        var saveConfigBtn = configBar.add("button", undefined, I18n.__("saveConfigButton"));
        saveConfigBtn.preferredSize.width = 100;
        var loadConfigBtn = configBar.add("button", undefined, I18n.__("loadConfigButton"));
        loadConfigBtn.preferredSize.width = 100;
        var configStatusText = configBar.add("statictext", undefined, I18n.__("configNotDetected"));
        configStatusText.preferredSize.width = 150;

        // dialogControls will be populated after all controls are created
        var dialogControls = {};

        // Function to update UI state based on loaded profile
        // Called after all controls are created and when profile changes
        var uiProfileControls = {}; // Will be populated as controls are created

        function updateUIForProfile() {
            var profile = LanguageProfile.getProfile();
            if (!profile) return;

            // Disable fixTypoSpaces if no spaces before punctuation
            var hasAnyPunctSpace = profile.punctuation &&
                (profile.punctuation.spaceBeforeSemicolon ||
                 profile.punctuation.spaceBeforeColon ||
                 profile.punctuation.spaceBeforeExclamation ||
                 profile.punctuation.spaceBeforeQuestion ||
                 profile.punctuation.spaceInsideOpenQuote ||
                 profile.punctuation.spaceInsideCloseQuote);

            if (uiProfileControls.fixTypoSpaces) {
                uiProfileControls.fixTypoSpaces.enabled = !!hasAnyPunctSpace;
                if (!hasAnyPunctSpace) uiProfileControls.fixTypoSpaces.checkbox.value = false;
            }

            // Disable SieclesModule when centuries.enabled === false
            var centuriesEnabled = profile.centuries && profile.centuries.enabled !== false;
            if (uiProfileControls.formatSiecles) {
                uiProfileControls.formatSiecles.enabled = centuriesEnabled;
                if (!centuriesEnabled) uiProfileControls.formatSiecles.value = false;
            }
            if (uiProfileControls.formatOrdinaux) {
                uiProfileControls.formatOrdinaux.enabled = centuriesEnabled;
                if (!centuriesEnabled) uiProfileControls.formatOrdinaux.value = false;
            }
            if (uiProfileControls.formatReferences) {
                uiProfileControls.formatReferences.enabled = centuriesEnabled;
                if (!centuriesEnabled) uiProfileControls.formatReferences.value = false;
            }

            // Update number formatting defaults from profile
            if (uiProfileControls.useComma && profile.numbers) {
                uiProfileControls.useComma.value = !!profile.numbers.replacePointWithComma;
            }
            if (uiProfileControls.addSpaces && profile.numbers) {
                uiProfileControls.addSpaces.value = !!profile.numbers.addThousandsSpaces;
            }
        }

        // Wire up profile dropdown onChange
        profileDropdown.onChange = function() {
            if (availableProfiles.length > 0 && profileDropdown.selection) {
                var selectedId = availableProfiles[profileDropdown.selection.index].id;
                LanguageProfile.load(selectedId);
                updateUIForProfile();
            }
        };

        // Création des onglets
        var tpanel = dialog.add("tabbedpanel");
        tpanel.alignChildren = "fill";
        
        // Nouvel onglet des corrections générales
        var tabCorrections = tpanel.add("tab", undefined, I18n.__("tabCorrections"));
        tabCorrections.orientation = "column";
        tabCorrections.alignChildren = "left";

        var tabSpaces = tpanel.add("tab", undefined, I18n.__("tabSpaces"));
        tabSpaces.orientation = "column";
        tabSpaces.alignChildren = "left";

        var tabStyle = tpanel.add("tab", undefined, I18n.__("tabStyles"));
        tabStyle.orientation = "column";
        tabStyle.alignChildren = "left";

        var tabOther = tpanel.add("tab", undefined, I18n.__("tabFormatting"));
        tabOther.orientation = "column";
        tabOther.alignChildren = "left";

        var tabStyles = tpanel.add("tab", undefined, I18n.__("tabPageLayout"));
        tabStyles.orientation = "column";
        tabStyles.alignChildren = "left";
        
        // Sélection de l'onglet par défaut
        tpanel.selection = tabCorrections;
        
        // Fonction utilitaire pour ajouter une case à cocher
        function addCheckboxOption(parent, label, checked) {
          var group = parent.add("group");
          group.orientation = "row";
          group.alignChildren = "left";
          var checkbox = group.add("checkbox", undefined, label);
          checkbox.value = checked;
          return checkbox;
        }
        
        // Fonction utilitaire pour ajouter une option avec liste déroulante
        function addDropdownOption(parent, label, items, checked) {
          var group = parent.add("group");
          group.orientation = "row";
          group.alignChildren = "left";
          var checkbox = group.add("checkbox", undefined, label);
          checkbox.value = checked;
          
          var dropdown = group.add("dropdownlist", undefined);
          for (var i = 0; i < items.length; i++) {
            var item = items[i];
            var itemLabel = item;
            
            if (typeof item === "object" && item !== null) {
              itemLabel = item.labelKey ? I18n.__(item.labelKey) : (item.label || item);
            }
            
            dropdown.add("item", itemLabel);
          }
          dropdown.selection = 0;
          dropdown.preferredSize.width = 150;
          dropdown.enabled = checkbox.value;
          
          checkbox.onClick = function() {
            dropdown.enabled = checkbox.value;
          };
          
          return { checkbox: checkbox, dropdown: dropdown };
        }
        
        // Ajout des options dans l'onglet Corrections
        var cbMoveNotes = addCheckboxOption(tabCorrections, I18n.__("moveNotesLabel"), true);
        var cbEllipsis = addCheckboxOption(tabCorrections, I18n.__("convertEllipsisLabel"), true);
        var cbReplaceApostrophes = addCheckboxOption(tabCorrections, I18n.__("replaceApostrophesLabel"), true);
        var cbDashes = addCheckboxOption(tabCorrections, I18n.__("replaceDashesLabel"), true);
        var cbFixIsolatedHyphens = addCheckboxOption(tabCorrections, I18n.__("fixIsolatedHyphensLabel"), true);
        var cbFixValueRanges = addCheckboxOption(tabCorrections, I18n.__("fixValueRangesLabel"), true);
        
        // Ajout des options dans l'onglet Espaces et retours
        var fixTypoSpacesOpt = addDropdownOption(tabSpaces, I18n.__("fixTypoSpacesLabel"), CONFIG.SPACE_TYPES, true);
        var fixDashIncisesOpt = addDropdownOption(tabSpaces, I18n.__("fixDashIncisesLabel"), CONFIG.SPACE_TYPES, false);
          fixDashIncisesOpt.dropdown.selection = 1;
        var cbFixSpaces = addCheckboxOption(tabSpaces, I18n.__("fixDoubleSpacesLabel"), true);
        var cbDoubleReturns = addCheckboxOption(tabSpaces, I18n.__("removeDoubleReturnsLabel"), true);
        var cbRemoveSpacesBeforePunctuation = addCheckboxOption(tabSpaces, I18n.__("removeSpacesBeforePunctuationLabel"), true);
        var cbRemoveSpacesStartParagraph = addCheckboxOption(tabSpaces, I18n.__("removeSpacesStartParagraphLabel"), true);
        var cbRemoveSpacesEndParagraph = addCheckboxOption(tabSpaces, I18n.__("removeSpacesEndParagraphLabel"), true);
        var cbRemoveTabs = addCheckboxOption(tabSpaces, I18n.__("removeTabsLabel"), true);
        var cbFormatEspaces = addCheckboxOption(tabSpaces, I18n.__("formatEspacesLabel"), true);
        
        // Ajout des options de style dans l'onglet Styles
        // Section pour la définition des styles
        var styleDefinitionPanel = tabStyle.add("panel", undefined, I18n.__("styleDefinitionPanel"));
        styleDefinitionPanel.orientation = "column";
        styleDefinitionPanel.alignChildren = "left";

        var noteStyleOpt = addDropdownOption(styleDefinitionPanel, I18n.__("noteStyleLabel"), characterStyles, true);
        var cbItalicStyle = addDropdownOption(styleDefinitionPanel, I18n.__("italicStyleLabel"), characterStyles, true);
        var romainsStyleOpt = addDropdownOption(styleDefinitionPanel, I18n.__("smallCapsStyleLabel"), characterStyles, true);
        var romainsMajStyleOpt = addDropdownOption(styleDefinitionPanel, I18n.__("capitalStyleLabel"), characterStyles, true);
        var exposantOrdinalStyleOpt = addDropdownOption(styleDefinitionPanel, I18n.__("superscriptStyleLabel"), characterStyles, true);
        
        // Sélection des styles par défaut
        if (noteStyleIndex >= 0 && noteStyleIndex < noteStyleOpt.dropdown.items.length) {
          noteStyleOpt.dropdown.selection = noteStyleIndex;
        }
        
        if (italicStyleIndex >= 0 && italicStyleIndex < cbItalicStyle.dropdown.items.length) {
          cbItalicStyle.dropdown.selection = italicStyleIndex;
        }
        
        // Rechercher les styles par défaut pour SieclesModule
        var defaultIndicesSiecles = SieclesModule.trouverStylesParDefaut(characterStyles);
        
        // Sélectionner les styles par défaut si trouvés
        if (defaultIndicesSiecles.petitesCapitales >= 0) {
          romainsStyleOpt.dropdown.selection = defaultIndicesSiecles.petitesCapitales;
          romainsStyleOpt.checkbox.value = true;
        } else {
          romainsStyleOpt.checkbox.value = false;
          romainsStyleOpt.dropdown.enabled = false;
        }
        
        if (defaultIndicesSiecles.capitales >= 0) {
          romainsMajStyleOpt.dropdown.selection = defaultIndicesSiecles.capitales;
          romainsMajStyleOpt.checkbox.value = true;
        } else {
          romainsMajStyleOpt.checkbox.value = false;
          romainsMajStyleOpt.dropdown.enabled = false;
        }
        
        if (defaultIndicesSiecles.exposant >= 0) {
          exposantOrdinalStyleOpt.dropdown.selection = defaultIndicesSiecles.exposant;
          exposantOrdinalStyleOpt.checkbox.value = true;
        } else if (noteStyleIndex >= 0) {
          // Utiliser le style d'exposant pour les notes si disponible
          exposantOrdinalStyleOpt.dropdown.selection = noteStyleIndex;
          exposantOrdinalStyleOpt.checkbox.value = true;
        } else {
          exposantOrdinalStyleOpt.checkbox.value = false;
          exposantOrdinalStyleOpt.dropdown.enabled = false;
        }
        
        // Garder des références aux cases à cocher des fonctionnalités
        var cbFormatSiecles = null;
        var cbFormatOrdinaux = null;
        var cbFormatReferences = null;
        var applyNoteStyleOpt = null;
        var applyItalicStyleOpt = null;
        var applyExposantStyleOpt = null;
        
        // Fonction qui met à jour l'état des fonctionnalités en fonction des styles disponibles
        function updateFeatureDependencies() {
          // --- Dépendances pour "Formater les siècles (XIVe siècle)" ---
          // Nécessite: Petites capitales + Exposant
          if (cbFormatSiecles) {
            var canFormatSiecles = romainsStyleOpt.checkbox.value && exposantOrdinalStyleOpt.checkbox.value;
            cbFormatSiecles.enabled = canFormatSiecles;
            
            if (!canFormatSiecles) {
              cbFormatSiecles.value = false;
            }
          }
          
          // --- Dépendances pour "Formater les expressions ordinales (IIe Internationale)" ---
          // Nécessite: Capitales + Exposant
          if (cbFormatOrdinaux) {
            var canFormatOrdinaux = romainsMajStyleOpt.checkbox.value && exposantOrdinalStyleOpt.checkbox.value;
            cbFormatOrdinaux.enabled = canFormatOrdinaux;
            
            if (!canFormatOrdinaux) {
              cbFormatOrdinaux.value = false;
            }
          }
          
          // --- Dépendances pour "Formater titres d'œuvres et noms propres (Tome III, Louis XIV)" ---
          // Nécessite: Capitales + Exposant (pour les cas comme "Louis Ier")
          if (cbFormatReferences) {
            var canFormatReferences = romainsMajStyleOpt.checkbox.value && exposantOrdinalStyleOpt.checkbox.value;
            cbFormatReferences.enabled = canFormatReferences;
            
            if (!canFormatReferences) {
              cbFormatReferences.value = false;
            }
          }
          
          // --- Dépendance pour "Appliquer un style aux références de notes" ---
          // Nécessite: Style d'appel de notes disponible
          if (applyNoteStyleOpt) {
            var canApplyNoteStyle = noteStyleOpt.checkbox.value;
            applyNoteStyleOpt.enabled = canApplyNoteStyle;
            
            if (!canApplyNoteStyle) {
              applyNoteStyleOpt.value = false;
            }
          }
          
          // --- Dépendance pour "Appliquer un style au texte en italique" ---
          // Nécessite: Style italique disponible
          if (applyItalicStyleOpt) {
            var canApplyItalicStyle = cbItalicStyle.checkbox.value;
            applyItalicStyleOpt.enabled = canApplyItalicStyle;
            
            if (!canApplyItalicStyle) {
              applyItalicStyleOpt.value = false;
            }
          }
          
          // --- Dépendance pour "Appliquer un style au texte en exposant" ---
          // Nécessite: Style d'exposant disponible
          if (applyExposantStyleOpt) {
            var canApplyExposantStyle = exposantOrdinalStyleOpt.checkbox.value;
            applyExposantStyleOpt.enabled = canApplyExposantStyle;
            
            if (!canApplyExposantStyle) {
              applyExposantStyleOpt.value = false;
            }
          }
        }
        
        // Ajouter des écouteurs d'événements pour toutes les cases à cocher de style
        romainsStyleOpt.checkbox.onClick = function() {
          romainsStyleOpt.dropdown.enabled = romainsStyleOpt.checkbox.value;
          updateFeatureDependencies();
        };
        
        romainsMajStyleOpt.checkbox.onClick = function() {
          romainsMajStyleOpt.dropdown.enabled = romainsMajStyleOpt.checkbox.value;
          updateFeatureDependencies();
        };
        
        exposantOrdinalStyleOpt.checkbox.onClick = function() {
          exposantOrdinalStyleOpt.dropdown.enabled = exposantOrdinalStyleOpt.checkbox.value;
          updateFeatureDependencies();
        };
        
        noteStyleOpt.checkbox.onClick = function() {
          noteStyleOpt.dropdown.enabled = noteStyleOpt.checkbox.value;
          updateFeatureDependencies();
        };
        
        cbItalicStyle.checkbox.onClick = function() {
          cbItalicStyle.dropdown.enabled = cbItalicStyle.checkbox.value;
          updateFeatureDependencies();
        };

        // Ajouter un peu d'espace entre les sections
        tabStyle.add("statictext", undefined, "");
        
        // Section pour l'application des styles
        var styleApplicationPanel = tabStyle.add("panel", undefined, I18n.__("styleApplicationPanel"));
        styleApplicationPanel.orientation = "column";
        styleApplicationPanel.alignChildren = "left";

        applyNoteStyleOpt = addCheckboxOption(styleApplicationPanel, I18n.__("applyNoteStyleLabel"), true);
        applyItalicStyleOpt = addCheckboxOption(styleApplicationPanel, I18n.__("applyItalicStyleLabel"), true);
        applyExposantStyleOpt = addCheckboxOption(styleApplicationPanel, I18n.__("applyExposantStyleLabel"), true);
        // Appliquer immédiatement les dépendances
        updateFeatureDependencies();
        
        // Ajout des options dans l'onglet Formatages (sans les options déplacées)
        // Options du module SieclesModule (restent dans l'onglet Formatages)
        var cbFormatSiecles = addCheckboxOption(tabOther, I18n.__("formatSieclesLabel"), true);
        var cbFormatOrdinaux = addCheckboxOption(tabOther, I18n.__("formatOrdinauxLabel"), true);
        var cbFormatReferences = addCheckboxOption(tabOther, I18n.__("formatReferencesLabel"), true);
        // Appliquer immédiatement les dépendances
        updateFeatureDependencies();
        
        // Ajout des options pour le formatage des nombres
        var cbFormatNumbers = addCheckboxOption(tabOther, I18n.__("formatNumbersLabel"), true);
        var numberSettingsPanel = tabOther.add("panel", undefined, I18n.__("numberSettingsPanel"));
        numberSettingsPanel.orientation = "column";
        numberSettingsPanel.alignChildren = "left";
        numberSettingsPanel.enabled = cbFormatNumbers.value;
        
        var cbAddSpaces = addCheckboxOption(numberSettingsPanel, I18n.__("addSpacesLabel"), true);
        var cbExcludeYears = addCheckboxOption(numberSettingsPanel, I18n.__("excludeYearsLabel"), true);
        var cbUseComma = addCheckboxOption(numberSettingsPanel, I18n.__("useCommaLabel"), true);
        
        // Activer/désactiver le panneau d'options selon l'état de la case à cocher principale
        cbFormatNumbers.onClick = function() {
          numberSettingsPanel.enabled = cbFormatNumbers.value;
        };

        // Populate profile control references for dynamic profile switching
        uiProfileControls.fixTypoSpaces = fixTypoSpacesOpt;
        uiProfileControls.formatSiecles = cbFormatSiecles;
        uiProfileControls.formatOrdinaux = cbFormatOrdinaux;
        uiProfileControls.formatReferences = cbFormatReferences;
        uiProfileControls.useComma = cbUseComma;
        uiProfileControls.addSpaces = cbAddSpaces;

        // Apply initial profile-based state
        updateUIForProfile();

        // Ajout des options dans le nouvel onglet Styles de paragraphe
        var cbEnableStyleAfter = addCheckboxOption(tabStyles, I18n.__("enableStyleAfterLabel"), true);
        
        // Récupérer les styles de paragraphe du document actif
        var doc = app.activeDocument;
        var allParaStyles = [];
        var paraStyleNames = [];
        
        try {
          // Fonction récursive pour collecter les styles
          function collectParagraphStyles(group) {
            var styles = [];
            
            try {
              if (group && group.paragraphStyles) {
                for (var i = 0; i < group.paragraphStyles.length; i++) {
                  var style = group.paragraphStyles[i];
                  if (style && style.name && !style.name.match(/^\[/)) {
                    styles.push(style);
                  }
                }
              }
              
              if (group && group.paragraphStyleGroups) {
                for (var j = 0; j < group.paragraphStyleGroups.length; j++) {
                  var subgroup = group.paragraphStyleGroups[j];
                  if (subgroup) {
                    styles = styles.concat(collectParagraphStyles(subgroup));
                  }
                }
              }
            } catch (e) {
              // Ignorer les erreurs silencieusement
            }
            
            return styles;
          }
          
          allParaStyles = collectParagraphStyles(doc);
          
          for (var i = 0; i < allParaStyles.length; i++) {
            paraStyleNames.push(allParaStyles[i].name);
          }
        } catch (e) {
          paraStyleNames = [I18n.__("noStylesAvailable")];
          allParaStyles = [];
        }

        if (paraStyleNames.length === 0) {
          paraStyleNames = [I18n.__("noStylesAvailable")];
        }

        var styleGroupPanel = tabStyles.add("panel", undefined, I18n.__("triggerStylesPanel"));
        styleGroupPanel.orientation = "column";
        styleGroupPanel.alignChildren = "left";
        styleGroupPanel.maximumSize.height = 200;
        styleGroupPanel.minimumSize.height = 150;
        
        // Ajouter un groupe déroulant pour les styles
        var scrollGroup = styleGroupPanel.add("group");
        scrollGroup.orientation = "column";
        scrollGroup.alignChildren = "left";
        scrollGroup.maximumSize.height = 180;
        
        // Créer les cases à cocher pour les styles déclencheurs
        var triggerCheckboxes = [];
        for (var i = 0; i < paraStyleNames.length; i++) {
          var cb = scrollGroup.add("checkbox", undefined, paraStyleNames[i]);
          triggerCheckboxes.push(cb);
          
          // Par défaut, tous les styles sont cochés sauf "Body text" et "First paragraph"
          if (paraStyleNames[i] === "Body text" || paraStyleNames[i] === "First paragraph") {
            cb.value = false;
          } else {
            cb.value = true;
          }
        }
        
        // Option pour le style cible
        var targetStyleGroup = tabStyles.add("group");
        targetStyleGroup.orientation = "row";
        targetStyleGroup.alignChildren = "center";
        targetStyleGroup.add("statictext", undefined, I18n.__("targetStyleLabel"));
        var targetStyleDropdown = targetStyleGroup.add("dropdownlist", undefined, paraStyleNames);
        targetStyleDropdown.preferredSize.width = 150;
        
        // Sélectionner "First paragraph" par défaut s'il existe
        var firstParaIndex = -1;
        for (var i = 0; i < paraStyleNames.length; i++) {
          if (paraStyleNames[i] === "First paragraph") {
            firstParaIndex = i;
            break;
          }
        }
        
        if (firstParaIndex !== -1) {
          targetStyleDropdown.selection = firstParaIndex;
        } else if (paraStyleNames.length > 0) {
          targetStyleDropdown.selection = 0;
        }
        
        // Option pour appliquer un gabarit à la dernière page
        var cbApplyMasterToLastPage = addCheckboxOption(tabStyles, I18n.__("applyMasterToLastPageLabel"), true);
        
        // Récupérer les gabarits du document
        var masterNames = [];
        var allMasters = [];
        
        try {
          if (ErrorHandler.ensureDefined(doc.masterSpreads, "doc.masterSpreads", false)) {
            for (var i = 0; i < doc.masterSpreads.length; i++) {
              var master = doc.masterSpreads[i];
              if (ErrorHandler.ensureDefined(master, "master at index " + i, false) && 
                ErrorHandler.ensureDefined(master.name, "master.name at index " + i, false)) {
                masterNames.push(master.name);
                allMasters.push(master);
              }
            }
          }
        } catch (e) {
          masterNames = [I18n.__("noMastersAvailable")];
          allMasters = [];
        }

        if (masterNames.length === 0) {
          masterNames = [I18n.__("noMastersAvailable")];
          cbApplyMasterToLastPage.enabled = false;
        }
        
        // Menu déroulant pour sélectionner le gabarit
        var masterGroup = tabStyles.add("group");
        masterGroup.orientation = "row";
        masterGroup.alignChildren = "center";
        masterGroup.add("statictext", undefined, I18n.__("masterLabel"));
        var masterDropdown = masterGroup.add("dropdownlist", undefined, masterNames);
        masterDropdown.preferredSize.width = 150;
        masterDropdown.enabled = cbApplyMasterToLastPage.value;
        
        // Activer/désactiver le menu déroulant selon la case à cocher
        cbApplyMasterToLastPage.onClick = function() {
          masterDropdown.enabled = cbApplyMasterToLastPage.value;
        };
        
        // Sélectionner le premier gabarit par défaut
        if (masterNames.length > 0) {
          masterDropdown.selection = 0;
        }
        
        // Populate dialogControls with references to all dialog controls
        dialogControls.profileDropdown = profileDropdown;
        dialogControls.noteStyleOpt = noteStyleOpt;
        dialogControls.cbItalicStyle = cbItalicStyle;
        dialogControls.romainsStyleOpt = romainsStyleOpt;
        dialogControls.romainsMajStyleOpt = romainsMajStyleOpt;
        dialogControls.exposantOrdinalStyleOpt = exposantOrdinalStyleOpt;
        dialogControls.cbRemoveSpacesBeforePunctuation = cbRemoveSpacesBeforePunctuation;
        dialogControls.cbFixSpaces = cbFixSpaces;
        dialogControls.fixTypoSpacesOpt = fixTypoSpacesOpt;
        dialogControls.fixDashIncisesOpt = fixDashIncisesOpt;
        dialogControls.cbDoubleReturns = cbDoubleReturns;
        dialogControls.cbRemoveSpacesStartParagraph = cbRemoveSpacesStartParagraph;
        dialogControls.cbRemoveSpacesEndParagraph = cbRemoveSpacesEndParagraph;
        dialogControls.cbRemoveTabs = cbRemoveTabs;
        dialogControls.cbMoveNotes = cbMoveNotes;
        dialogControls.applyNoteStyleOpt = applyNoteStyleOpt;
        dialogControls.cbDashes = cbDashes;
        dialogControls.cbFixIsolatedHyphens = cbFixIsolatedHyphens;
        dialogControls.cbFixValueRanges = cbFixValueRanges;
        dialogControls.cbEllipsis = cbEllipsis;
        dialogControls.cbReplaceApostrophes = cbReplaceApostrophes;
        dialogControls.applyItalicStyleOpt = applyItalicStyleOpt;
        dialogControls.applyExposantStyleOpt = applyExposantStyleOpt;
        dialogControls.cbFormatEspaces = cbFormatEspaces;
        dialogControls.cbFormatSiecles = cbFormatSiecles;
        dialogControls.cbFormatOrdinaux = cbFormatOrdinaux;
        dialogControls.cbFormatReferences = cbFormatReferences;
        dialogControls.cbFormatNumbers = cbFormatNumbers;
        dialogControls.numberSettingsPanel = numberSettingsPanel;
        dialogControls.cbAddSpaces = cbAddSpaces;
        dialogControls.cbUseComma = cbUseComma;
        dialogControls.cbExcludeYears = cbExcludeYears;
        dialogControls.cbEnableStyleAfter = cbEnableStyleAfter;
        dialogControls.triggerCheckboxes = triggerCheckboxes;
        dialogControls.targetStyleDropdown = targetStyleDropdown;
        dialogControls.cbApplyMasterToLastPage = cbApplyMasterToLastPage;
        dialogControls.masterDropdown = masterDropdown;

        // Helper to get current language profile ID from dropdown
        function getSelectedProfileId() {
            if (availableProfiles.length > 0 && profileDropdown.selection) {
                return availableProfiles[profileDropdown.selection.index].id;
            }
            return null;
        }

        // Auto-load configuration if found near the document
        var autoLoadedConfig = ConfigManager.autoLoad();
        if (autoLoadedConfig) {
            configStatusText.text = I18n.__("configDetected");
            ConfigManager.applyToDialog(autoLoadedConfig, dialogControls, characterStyles, availableProfiles);
            updateUIForProfile();
            updateFeatureDependencies();
        }

        // Wire Save button
        saveConfigBtn.onClick = function() {
            dialogControls.languageProfileId = getSelectedProfileId();
            var configData = ConfigManager.collectFromDialog(dialogControls);
            ConfigManager.save(configData);
        };

        // Wire Load button
        loadConfigBtn.onClick = function() {
            var configData = ConfigManager.load();
            if (configData) {
                ConfigManager.applyToDialog(configData, dialogControls, characterStyles, availableProfiles);
                updateUIForProfile();
                updateFeatureDependencies();
            }
        };

        // Boutons d'action
        var buttonGroup = dialog.add("group");
        buttonGroup.orientation = "row";
        buttonGroup.alignment = "right";
        
        // Ajouter un bouton d'aide
        var helpButton = buttonGroup.add("button", undefined, "?");
        helpButton.preferredSize.width = 25; // Mini bouton
        helpButton.preferredSize.height = 25;
        helpButton.helpTip = I18n.__("helpTooltip");

        var cancelButton = buttonGroup.add("button", undefined, I18n.__("cancelButton"), {name: "cancel"});
        var okButton = buttonGroup.add("button", undefined, I18n.__("applyButton"), {name: "ok"});
        
        // Fonction pour afficher la fenêtre d'aide
        helpButton.onClick = function() {
          // Créer une nouvelle boîte de dialogue pour l'aide
          var helpDialog = new Window("dialog", I18n.__("helpDialogTitle"));
          helpDialog.orientation = "column";
          helpDialog.alignChildren = "fill";
          helpDialog.preferredSize.width = 500;
          helpDialog.preferredSize.height = 400;

          var helpHeaderGroup = helpDialog.add("group");
          helpHeaderGroup.alignment = "center";
          helpHeaderGroup.add("statictext", undefined, I18n.__("helpDialogHeader"));

          var helpText = helpDialog.add("edittext", undefined, "", {multiline: true, readonly: true, scrollable: true});
          helpText.preferredSize.height = 300;
          helpText.text = I18n.__("helpContent");

          var closeButton = helpDialog.add("button", undefined, I18n.__("closeButton"), {name: "ok"});
          closeButton.alignment = "center";
          
          // Afficher la boîte de dialogue d'aide
          helpDialog.show();
        };
        
        // Événement du bouton Annuler
        cancelButton.onClick = function() {
          dialog.close(2);
        };
        
        // Afficher le dialogue et renvoyer les résultats
        if (dialog.show() == 1) {
          try {
            // Collecter les styles déclencheurs sélectionnés
            var selectedTriggerStyles = [];
            if (allParaStyles.length > 0) {
              for (var i = 0; i < triggerCheckboxes.length; i++) {
                if (triggerCheckboxes[i].value) {
                  selectedTriggerStyles.push(allParaStyles[i]);
                }
              }
            }
            
            // Récupérer les options du module SieclesModule
            var sieclesOptions = {
              formaterSiecles: cbFormatSiecles.value,
              formaterOrdinaux: cbFormatOrdinaux.value,
              formaterReferences: cbFormatReferences.value,
              formaterEspaces: cbFormatEspaces.value,
              romainsStyle: romainsStyleOpt.checkbox.value ? 
                doc.characterStyles[romainsStyleOpt.dropdown.selection.index] : null,
              romainsMajStyle: romainsMajStyleOpt.checkbox.value ? 
                doc.characterStyles[romainsMajStyleOpt.dropdown.selection.index] : null,
              exposantStyle: exposantOrdinalStyleOpt.checkbox.value ? 
                doc.characterStyles[exposantOrdinalStyleOpt.dropdown.selection.index] : null
            };
            
            // Renvoyer les sélections de l'utilisateur
            return {
              removeSpacesBeforePunctuation: cbRemoveSpacesBeforePunctuation.value,
              moveNotes: cbMoveNotes.value,
              applyNoteStyle: applyNoteStyleOpt.value && noteStyleOpt.checkbox.value,
              noteStyleName: applyNoteStyleOpt.value && noteStyleOpt.checkbox.value && noteStyleOpt.dropdown.selection ? 
                noteStyleOpt.dropdown.selection.text : null,
              fixDoubleSpaces: cbFixSpaces.value,
              fixTypoSpaces: fixTypoSpacesOpt.checkbox.value,
              spaceType: fixTypoSpacesOpt.checkbox.value && fixTypoSpacesOpt.dropdown.selection ? 
                CONFIG.SPACE_TYPES[fixTypoSpacesOpt.dropdown.selection.index].value : null,
              fixDashIncises: fixDashIncisesOpt.checkbox.value,
              dashIncisesSpaceType: fixDashIncisesOpt.checkbox.value && fixDashIncisesOpt.dropdown.selection ? 
                  CONFIG.SPACE_TYPES[fixDashIncisesOpt.dropdown.selection.index].value : null,
              replaceDashes: cbDashes.value,
              applyItalicStyle: applyItalicStyleOpt.value && cbItalicStyle.checkbox.value,
              italicStyleName: applyItalicStyleOpt.value && cbItalicStyle.checkbox.value && cbItalicStyle.dropdown.selection ? 
                cbItalicStyle.dropdown.selection.text : null,
              applyExposantStyle: applyExposantStyleOpt.value && exposantOrdinalStyleOpt.checkbox.value,
              exposantStyleName: applyExposantStyleOpt.value && exposantOrdinalStyleOpt.checkbox.value && exposantOrdinalStyleOpt.dropdown.selection ? 
                exposantOrdinalStyleOpt.dropdown.selection.text : null,
              removeDoubleReturns: cbDoubleReturns.value,
              convertEllipsis: cbEllipsis.value,
              replaceApostrophes: cbReplaceApostrophes.value,
              fixIsolatedHyphens: cbFixIsolatedHyphens.value,
              fixValueRanges: cbFixValueRanges.value,
              removeSpacesStartParagraph: cbRemoveSpacesStartParagraph.value,
              removeSpacesEndParagraph: cbRemoveSpacesEndParagraph.value,
              removeTabs: cbRemoveTabs.value,
              
              // Nouvelles options pour les styles
              enableStyleAfter: cbEnableStyleAfter.value,
              triggerStyles: selectedTriggerStyles,
              targetStyle: targetStyleDropdown.selection && allParaStyles.length > 0 ? 
                allParaStyles[targetStyleDropdown.selection.index] : null,
                
              applyMasterToLastPage: cbApplyMasterToLastPage.value,
              selectedMaster: masterDropdown.selection && allMasters.length > 0 ? 
                allMasters[masterDropdown.selection.index] : null,
              
              // Options pour le module SieclesModule
              sieclesOptions: sieclesOptions,
              
              // Options pour le formatage des nombres
              formatNumbers: cbFormatNumbers.value,
              addSpaces: cbAddSpaces.value,
              excludeYears: cbExcludeYears.value,
              useComma: cbUseComma.value,

              // Language profile
              languageProfileId: (availableProfiles.length > 0 && profileDropdown.selection)
                ? availableProfiles[profileDropdown.selection.index].id
                : null
            };
          } catch (resultError) {
            ErrorHandler.handleError(resultError, "dialog results", true);
            return null;
          }
        }
        
        return null;
        } catch (error) {
          ErrorHandler.handleError(error, "createDialog", true);
          return null;
        }
      }
    };
    
    /**
     * Processeur principal pour appliquer les corrections
     * @private
     */
    var Processor = {
      /**
        * Applique les corrections sélectionnées à un document
        * @param {Document} doc - Document InDesign
        * @param {Object} options - Options de correction
        */
      applyCorrections: function(doc, options) {
          try {
              if (!ErrorHandler.ensureDefined(doc, "document", true)) return;
              if (!ErrorHandler.ensureDefined(options, "options", true)) return;
              if (!ErrorHandler.ensureDefined(app, "app", true)) return;
              if (!ErrorHandler.ensureDefined(app.scriptPreferences, "app.scriptPreferences", false)) {
                  // Continuer sans modifier le redessinage si non disponible
              } else {
                  // Désactiver le redessinage pour de meilleures performances
                  app.scriptPreferences.enableRedraw = false;
              }
              
              // Compter le nombre d'étapes activées pour la barre de progression
              var totalSteps = 0;
              if (options.removeSpacesBeforePunctuation) totalSteps++;
              if (options.fixDoubleSpaces) totalSteps++;
              if (options.fixTypoSpaces && options.spaceType) totalSteps++;
              if (options.fixDashIncises && options.dashIncisesSpaceType) totalSteps++;
              if (options.removeDoubleReturns) totalSteps++;
              if (options.removeSpacesStartParagraph) totalSteps++;
              if (options.removeSpacesEndParagraph) totalSteps++;
              if (options.removeTabs) totalSteps++;
              if (options.moveNotes) totalSteps++;
              if (options.applyNoteStyle && options.noteStyleName) totalSteps++;
              if (options.replaceDashes) totalSteps++;
              if (options.fixIsolatedHyphens) totalSteps++;
              if (options.fixValueRanges) totalSteps++;
              if (options.applyItalicStyle && options.italicStyleName) totalSteps++;
              if (options.applyExposantStyle && options.exposantStyleName) totalSteps++;
              if (options.convertEllipsis) totalSteps++;
              if (options.replaceApostrophes) totalSteps++;
              if (options.enableStyleAfter && options.triggerStyles && options.triggerStyles.length > 0 && options.targetStyle) totalSteps++;
              if (options.applyMasterToLastPage && options.selectedMaster) totalSteps++;
              if (options.sieclesOptions && (options.sieclesOptions.formaterSiecles || options.sieclesOptions.formaterOrdinaux || options.sieclesOptions.formaterReferences || options.sieclesOptions.formaterEspaces)) totalSteps++;
              if (options.formatNumbers) totalSteps++;

              // Créer la barre de progression
              ProgressBar.create(I18n.__("progressTitle"), totalSteps);

              try {
                  // Compteur de progression
                  var progress = 0;
                  
                  // Corrections d'espaces et retours
                  if (options.removeSpacesBeforePunctuation) {
                      ProgressBar.update(++progress, I18n.__("progressRemoveSpacesBeforePunctuation"));
                      Corrections.removeSpacesBeforePunctuation(doc);
                  }
                  
                  if (options.fixDoubleSpaces) {
                      ProgressBar.update(++progress, I18n.__("progressFixDoubleSpaces"));
                      Corrections.fixDoubleSpaces(doc);
                  }
                  
                  if (options.fixTypoSpaces && options.spaceType) {
                      ProgressBar.update(++progress, I18n.__("progressFixTypoSpaces"));
                      Corrections.fixTypoSpaces(doc, options.spaceType, options.languageProfileId);
                  }

                  if (options.fixDashIncises && options.dashIncisesSpaceType) {
                      ProgressBar.update(++progress, I18n.__("progressFixDashIncises"));
                      Corrections.fixDashIncises(doc, options.dashIncisesSpaceType, options.languageProfileId);
                  }
                  
                  if (options.removeDoubleReturns) {
                      ProgressBar.update(++progress, I18n.__("progressRemoveDoubleReturns"));
                      Corrections.removeDoubleReturns(doc);
                  }
                  
                  if (options.removeSpacesStartParagraph) {
                      ProgressBar.update(++progress, I18n.__("progressRemoveSpacesStart"));
                      Corrections.removeSpacesStartParagraph(doc);
                  }
                  
                  if (options.removeSpacesEndParagraph) {
                      ProgressBar.update(++progress, I18n.__("progressRemoveSpacesEnd"));
                      Corrections.removeSpacesEndParagraph(doc);
                  }
                  
                  if (options.removeTabs) {
                      ProgressBar.update(++progress, I18n.__("progressRemoveTabs"));
                      Corrections.removeTabs(doc);
                  }
                  
                  // Autres corrections
                  if (options.moveNotes) {
                      ProgressBar.update(++progress, I18n.__("progressMoveNotes"));
                      Corrections.moveNotes(doc);
                  }
                  
                  if (options.applyNoteStyle && options.noteStyleName) {
                      ProgressBar.update(++progress, I18n.__("progressApplyNoteStyle"));
                      Corrections.applyNoteStyle(doc, options.noteStyleName);
                  }
                  
                  if (options.replaceDashes) {
                      ProgressBar.update(++progress, I18n.__("progressReplaceDashes"));
                      Corrections.replaceDashes(doc, options.languageProfileId);
                  }
                  
                  if (options.fixIsolatedHyphens) {
                      ProgressBar.update(++progress, I18n.__("progressFixIsolatedHyphens"));
                      Corrections.fixIsolatedHyphens(doc);
                  }
                  
                  if (options.fixValueRanges) {
                      ProgressBar.update(++progress, I18n.__("progressFixValueRanges"));
                      Corrections.fixValueRanges(doc);
                  }
                  
                  if (options.applyItalicStyle && options.italicStyleName) {
                      ProgressBar.update(++progress, I18n.__("progressApplyItalicStyle"));
                      Corrections.applyItalicStyle(doc, options.italicStyleName);
                  }
                  
                  if (options.applyExposantStyle && options.exposantStyleName) {
                      ProgressBar.update(++progress, I18n.__("progressApplyExposantStyle"));
                      Corrections.applyExposantStyle(doc, options.exposantStyleName);
                  }
                  
                  if (options.convertEllipsis) {
                      ProgressBar.update(++progress, I18n.__("progressConvertEllipsis"));
                      Corrections.convertEllipsis(doc);
                  }
                  
                  if (options.replaceApostrophes) {
                      ProgressBar.update(++progress, I18n.__("progressReplaceApostrophes"));
                      Corrections.replaceApostrophes(doc);
                  }
                  
                  // Nouvelle correction : application de style après déclencheurs
                  if (options.enableStyleAfter && 
                      options.triggerStyles && 
                      options.triggerStyles.length > 0 && 
                      options.targetStyle) {
                      ProgressBar.update(++progress, I18n.__("progressApplyConditionalStyles"));
                      Corrections.applyStyleAfterTriggers(doc, options.triggerStyles, options.targetStyle);
                  }
                  
                  // Application du gabarit à la dernière page
                  if (options.applyMasterToLastPage && options.selectedMaster) {
                      ProgressBar.update(++progress, I18n.__("progressApplyMasterToLastPage"));
                      // Stocker le nom du gabarit plutôt que l'objet gabarit lui-même
                      var masterName = options.selectedMaster.name;
                      
                      // Appliquer en utilisant le nom du gabarit
                      applyMasterToLastPageStandalone(doc, masterName);
                  }
                  
                  // Intégration du module SieclesModule
                  if (options.sieclesOptions && 
                      (options.sieclesOptions.formaterSiecles || 
                       options.sieclesOptions.formaterOrdinaux || 
                       options.sieclesOptions.formaterReferences || 
                       options.sieclesOptions.formaterEspaces)) {
                      
                      ProgressBar.update(++progress, I18n.__("progressFormatSiecles"));
                      
                      // Appliquer les corrections du module SieclesModule
                      SieclesModule.processDocument(doc, options.sieclesOptions);
                  }
                  
                  // Traitement du formatage des nombres
                  if (options.formatNumbers) {
                      ProgressBar.update(++progress, I18n.__("progressFormatNumbers"));
                      Corrections.formatNumbers(doc, options.addSpaces, options.useComma, options.excludeYears, options.languageProfileId);
                  }
                  
                  // Finalisation
                  ProgressBar.update(totalSteps, I18n.__("progressComplete"));
                  
              } catch (correctionsError) {
                  ErrorHandler.handleError(correctionsError, "applying corrections", false);
              } finally {
                  // Fermer la barre de progression
                  ProgressBar.close();
                  
                  // Réinitialiser les préférences
                  Utilities.resetPreferences();
                  
                  // Réactiver le redessinage
                  if (ErrorHandler.ensureDefined(app, "app", false) && 
                      ErrorHandler.ensureDefined(app.scriptPreferences, "app.scriptPreferences", false)) {
                      app.scriptPreferences.enableRedraw = true;
                  }
              }
          } catch (error) {
              ErrorHandler.handleError(error, "applyCorrections", true);
          }
      },
      
      /**
        * Traite les documents en fonction des options de l'utilisateur
        * @param {Object} options - Options sélectionnées par l'utilisateur
        */
      processDocuments: function(options) {
          try {
              if (!ErrorHandler.ensureDefined(options, "options", true)) return;
              if (!ErrorHandler.ensureDefined(app, "app", true)) return;
              if (!ErrorHandler.ensureDefined(app.activeDocument, "app.activeDocument", true)) return;

              // Load the selected language profile
              if (options.languageProfileId) {
                  LanguageProfile.load(options.languageProfileId);
              }

              Processor.applyCorrections(app.activeDocument, options);
              alert(I18n.__("successCorrectionsApplied"));
          } catch (error) {
              ErrorHandler.handleError(error, "processDocuments", true);
          }
      }
    };
    
    /**
     * Module SieclesModule - Intégration isolée du script de formatage des siècles
     * @private
     */
    var SieclesModule = {
        /**
         * Configuration spécifique au formatage des siècles
         * Copiée directement du script original
         */
        CONFIG: {
            // Style auto-detection names (universal, not language-specific)
            STYLES_PETITES_CAPITALES: ["Small caps", "Small cap", "Small capitals", "Small capital", "Petites capitales", "Petites caps"],
            STYLES_CAPITALES: ["Large Capitals", "Capital", "Capitals"],
            STYLES_EXPOSANT: ["Superscript", "Exposant", "Superior"]
        },
        
        /**
         * Trouve les index des styles par défaut pour le formatage des siècles
         * @param {Array} characterStyles - Tableau des styles de caractère disponibles
         * @returns {Object} Indices des styles par défaut
         */
        trouverStylesParDefaut: function(characterStyles) {
            var indices = {
                petitesCapitales: -1,  // Initialiser à -1 pour indiquer "non trouvé"
                capitales: -1,
                exposant: -1
            };
            
            // Rechercher les styles par défaut
            for (var i = 0; i < characterStyles.length; i++) {
                var styleName = characterStyles[i].toLowerCase();
                
                // Recherche pour les petites capitales (correspondance exacte)
                for (var sc = 0; sc < this.CONFIG.STYLES_PETITES_CAPITALES.length; sc++) {
                    if (styleName === this.CONFIG.STYLES_PETITES_CAPITALES[sc].toLowerCase()) {
                        indices.petitesCapitales = i;
                        break;
                    }
                }
                
                // Recherche pour les capitales (correspondance exacte)
                for (var ac = 0; ac < this.CONFIG.STYLES_CAPITALES.length; ac++) {
                    if (styleName === this.CONFIG.STYLES_CAPITALES[ac].toLowerCase()) {
                        indices.capitales = i;
                        break;
                    }
                }
                
                // Recherche pour les exposants (correspondance exacte)
                for (var ex = 0; ex < this.CONFIG.STYLES_EXPOSANT.length; ex++) {
                    if (styleName === this.CONFIG.STYLES_EXPOSANT[ex].toLowerCase()) {
                        indices.exposant = i;
                        break;
                    }
                }
            }
            
            return indices;
        },
        
        /**
         * Récupère les options sélectionnées par l'utilisateur
         * @param {Object} controls - Contrôles créés dans l'interface
         * @param {Object} styles - Styles communs récupérés du script principal
         * @returns {Object} Options sélectionnées
         */
        getOptions: function(controls, styles) {
            return {
                formaterSiecles: controls.formatSiecles.value,
                formaterOrdinaux: controls.formatOrdinaux.value,
                formaterReferences: controls.formatReferences.value,
                formaterEspaces: controls.formatEspaces.value,
                romainsStyle: controls.romainsStyle.checkbox.value ? 
                    app.activeDocument.characterStyles[controls.romainsStyle.dropdown.selection.index] : null,
                romainsMajStyle: controls.romainsMajStyle.checkbox.value ? 
                    app.activeDocument.characterStyles[controls.romainsMajStyle.dropdown.selection.index] : null,
                // Utiliser le style d'exposant spécifique pour les ordinaux
                exposantStyle: controls.exposantStyle.checkbox.value ? 
                    app.activeDocument.characterStyles[controls.exposantStyle.dropdown.selection.index] : null
            };
        },
        
        /**
         * Applique les corrections de formatage selon les options choisies
         * @param {Document} doc - Document InDesign
         * @param {Object} options - Options sélectionnées
         */
        /**
         * Gets a CONFIG data array from LanguageProfile if available, otherwise from hardcoded CONFIG
         * @param {String} configKey - Key in SieclesModule.CONFIG (e.g., "MOTS_ORDINAUX")
         * @param {String} profilePath - Dot path in LanguageProfile (e.g., "data.motsOrdinaux")
         * @return {Array} The data array
         */
        getConfigData: function(configKey, profilePath) {
            if (LanguageProfile.getCurrentId()) {
                var profileData = LanguageProfile.getList(profilePath);
                if (profileData.length > 0) return profileData;
            }
            return [];
        },

        processDocument: function(doc, options) {
          try {
            // Vérifier si au moins une option est activée
            if (!options.formaterSiecles && !options.formaterOrdinaux &&
                !options.formaterReferences && !options.formaterEspaces) {
                return;
            }

            // Initialiser les variables nécessaires pour le script isolé
            this.initializeIsolatedEnvironment();
            
            // Appliquer les corrections selon les options choisies
            if (options.formaterOrdinaux) {
                this.formaterOrdinaux(doc, options.romainsMajStyle, options.exposantStyle);
            }
            
            if (options.formaterSiecles) {
                this.formaterSiecles(doc, options.romainsStyle, options.exposantStyle);
            }
            
            if (options.formaterReferences) {
                this.formaterReferences(doc, options.romainsMajStyle, options.exposantStyle);
            }
            
            if (options.formaterEspaces) {
                this.formaterEspaces(doc);
            }
            
            // Réinitialiser l'environnement après utilisation
            this.resetIsolatedEnvironment();
          } catch (error) {
              alert(I18n.__("errorFormatSiecles", error));
          }
        },
        
        /**
         * Initialise l'environnement isolé pour le script de formatage des siècles
         * Crée des fonctions locales nécessaires au fonctionnement du script
         */
        initializeIsolatedEnvironment: function() {
            // Créer une référence locale au module pour utilisation dans les fonctions
            var self = this;
            
            // Recréer les fonctions utilitaires nécessaires au script
            this.Utilities = {
                estEnItalique: function(texte) {
                    try {
                        return texte.fontStyle && texte.fontStyle.indexOf("Italic") !== -1;
                    } catch (e) {
                        return false;
                    }
                },
                
                estEnGras: function(texte) {
                    try {
                        return texte.fontStyle && texte.fontStyle.indexOf("Bold") !== -1;
                    } catch (e) {
                        return false;
                    }
                },
                
                appliquerStylePolice: function(caractere, estItalique, estGras) {
                    try {
                        if (estItalique && estGras) {
                            caractere.fontStyle = "Bold Italic";
                        } else if (estItalique) {
                            caractere.fontStyle = "Italic";
                        } else if (estGras) {
                            caractere.fontStyle = "Bold";
                        }
                    } catch (e) {
                        // Ignorer l'erreur si l'application du style échoue
                    }
                },
                
                appliquerGrepPartout: function(doc, findWhat, changeParams, specificAction) {
                    app.findGrepPreferences.findWhat = findWhat;
                    
                    // Si changeParams contient des paramètres spécifiques (comme un style)
                    if (changeParams) {
                        for (var prop in changeParams) {
                            app.changeGrepPreferences[prop] = changeParams[prop];
                        }
                    }
                    
                    // Parcourir toutes les histoires normales
                    for (var s = 0; s < doc.stories.length; s++) {
                        // Si specificAction est fournie (pour un traitement spécial), l'utiliser
                        if (specificAction) {
                            var results = doc.stories[s].findGrep();
                            for (var r = 0; r < results.length; r++) {
                                specificAction(results[r]);
                            }
                        } else {
                            // Sinon, utiliser la méthode standard changeGrep()
                            doc.stories[s].changeGrep();
                        }
                        
                        // Traiter les notes de bas de page de cette histoire
                        try {
                            // Vérifier si l'histoire contient des notes de bas de page
                            if (doc.stories[s].footnotes && doc.stories[s].footnotes.length > 0) {
                                for (var fn = 0; fn < doc.stories[s].footnotes.length; fn++) {
                                    var footnote = doc.stories[s].footnotes[fn];
                                    if (specificAction) {
                                        var fnResults = footnote.texts[0].findGrep();
                                        for (var fr = 0; fr < fnResults.length; fr++) {
                                            specificAction(fnResults[fr]);
                                        }
                                    } else {
                                        footnote.texts[0].changeGrep();
                                    }
                                }
                            }
                        } catch (e) {
                            // Ignorer les erreurs
                        }
                    }
                    
                    // Réinitialiser les préférences
                    app.findGrepPreferences = app.changeGrepPreferences = null;
                },
                
                preparerRegex: function(liste) {
                    var regex = "";
                    for (var i = 0; i < liste.length; i++) {
                        if (i > 0) {
                            regex += "|";
                        }
                        regex += liste[i];
                    }
                    return regex;
                },
                
                preparerRegexAmbigus: function(liste) {
                    var regex = "";
                    for (var i = 0; i < liste.length; i++) {
                        if (i > 0) {
                            regex += "|";
                        }
                        regex += "\\b" + liste[i] + "\\b";
                    }
                    return regex;
                }
            };
            
            // Sauvegarde des préférences actuelles pour restauration ultérieure
            this.savedPreferences = {
                findGrepPreferences: app.findGrepPreferences,
                changeGrepPreferences: app.changeGrepPreferences,
                findTextPreferences: app.findTextPreferences,
                changeTextPreferences: app.changeTextPreferences
            };
        },
        
        /**
         * Réinitialise l'environnement après exécution du script isolé
         */
        resetIsolatedEnvironment: function() {
            // Réinitialiser les préférences GREP
            app.findGrepPreferences = app.changeGrepPreferences = null;
            app.findTextPreferences = app.changeTextPreferences = null;
        },
        
        /**
         * Formate les siècles (ex: XIVe siècle)
         */
        formaterSiecles: function(doc, romainsStyle, exposantStyle) {
          var self = this;
          try {
            // Vérifier que les styles sont définis
            if (!romainsStyle || !exposantStyle) {
              alert(I18n.__("errorRequiredStyles"));
              return;
            }
            
            // Espace insécable
            var ESPACE_INSECABLE = String.fromCharCode(0x00A0);
            
            // Préparer les expressions régulières
            var motsClefsRegex = this.Utilities.preparerRegex(this.getConfigData("MOTS_ORDINAUX", "data.motsOrdinaux"));
            var ambigusRegex = this.Utilities.preparerRegexAmbigus(this.getConfigData("MOTS_AMBIGUS", "data.motsAmbigus"));
            
            // Appliquer les styles à tout le document
            app.findGrepPreferences = app.changeGrepPreferences = null;
            
            // === TRAITEMENT SPÉCIAL POUR LE Ier SIÈCLE ===
            // Rechercher "Ie siècle" ou "ie siècle" et le remplacer par "Ier siècle"
            var modelesPremierSiecle = [
                // Modèles originaux
                "(?<=\\s|^|:)([Ii])e(?=\\s+(?i:siècle))",
                "(?<=\\s|^|:)([Ii])e(?=[\\.,;:\\s!\\?])(?!\\s+(?i:" + motsClefsRegex + "))",
                "\\bie(?=\\s+(?i:siècle))",
                "(?<=\\s|^|:)([Ii])e(?=[\\)\\]\\}\\>])(?!\\s+(?i:" + motsClefsRegex + "))"
            ];
            
            // Traitement explicite pour "ie siècle"
            app.findGrepPreferences = app.changeGrepPreferences = null;
            app.findGrepPreferences.findWhat = "\\bie\\s+siècle";
            
            function traiterIeSiecle(resultat) {
              if (resultat && resultat.contents) {
                // Forcer le résultat à "ier"
                var estItaliqueLocal = self.Utilities.estEnItalique(resultat);
                var estGrasLocal = self.Utilities.estEnGras(resultat);
                
                // Récupérer la position de "ie"
                var texteOriginal = resultat.contents;
                var positionIe = texteOriginal.indexOf("ie");
                
                if (positionIe !== -1) {
                  // Remplacer "ie" par "ier"
                  resultat.contents = texteOriginal.substring(0, positionIe) + "ier" + texteOriginal.substring(positionIe + 2);
                  
                  // Appliquer les styles caractère par caractère
                  if (positionIe < resultat.characters.length) {
                    // Style pour "i"
                    resultat.characters[positionIe].appliedCharacterStyle = romainsStyle;
                    self.Utilities.appliquerStylePolice(resultat.characters[positionIe], estItaliqueLocal, estGrasLocal);
                    
                    // Style pour "er"
                    if (positionIe + 1 < resultat.characters.length) {
                      resultat.characters[positionIe + 1].appliedCharacterStyle = exposantStyle;
                      self.Utilities.appliquerStylePolice(resultat.characters[positionIe + 1], estItaliqueLocal, estGrasLocal);
                    }
                    if (positionIe + 2 < resultat.characters.length) {
                      resultat.characters[positionIe + 2].appliedCharacterStyle = exposantStyle;
                      self.Utilities.appliquerStylePolice(resultat.characters[positionIe + 2], estItaliqueLocal, estGrasLocal);
                    }
                  }
                }
              }
            }
            
            self.Utilities.appliquerGrepPartout(doc, "\\bie\\s+siècle", null, traiterIeSiecle);
            
            // Traitement explicite pour "Ie siècle" (I majuscule)
            app.findGrepPreferences = app.changeGrepPreferences = null;
            app.findGrepPreferences.findWhat = "\\bIe\\s+siècle";
            
            function traiterIeCapSiecle(resultat) {
              if (resultat && resultat.contents) {
                // Forcer le résultat à "Ier"
                var estItaliqueLocal = self.Utilities.estEnItalique(resultat);
                var estGrasLocal = self.Utilities.estEnGras(resultat);
                
                // Récupérer la position de "Ie"
                var texteOriginal = resultat.contents;
                var positionIe = texteOriginal.indexOf("Ie");
                
                if (positionIe !== -1) {
                  // Remplacer "Ie" par "Ier"
                  resultat.contents = texteOriginal.substring(0, positionIe) + "Ier" + texteOriginal.substring(positionIe + 2);
                  
                  // Appliquer les styles caractère par caractère
                  if (positionIe < resultat.characters.length) {
                    // Style pour "I"
                    resultat.characters[positionIe].appliedCharacterStyle = romainsStyle;
                    self.Utilities.appliquerStylePolice(resultat.characters[positionIe], estItaliqueLocal, estGrasLocal);
                    
                    // Style pour "er"
                    if (positionIe + 1 < resultat.characters.length) {
                      resultat.characters[positionIe + 1].appliedCharacterStyle = exposantStyle;
                      self.Utilities.appliquerStylePolice(resultat.characters[positionIe + 1], estItaliqueLocal, estGrasLocal);
                    }
                    if (positionIe + 2 < resultat.characters.length) {
                      resultat.characters[positionIe + 2].appliedCharacterStyle = exposantStyle;
                      self.Utilities.appliquerStylePolice(resultat.characters[positionIe + 2], estItaliqueLocal, estGrasLocal);
                    }
                  }
                }
              }
            }
            
            self.Utilities.appliquerGrepPartout(doc, "\\bIe\\s+siècle", null, traiterIeCapSiecle);
            
            // Fonction spéciale pour traiter le premier siècle
            function traiterPremierSiecle(resultat) {
                // Obtenir le caractère initial (I ou i)
                var premierChar = resultat.characters[0].contents;
                
                // Vérifier si le texte est en italique ou en gras
                var estItaliqueLocal = self.Utilities.estEnItalique(resultat);
                var estGrasLocal = self.Utilities.estEnGras(resultat);
                
                // Remplacer le contenu par "Ier" ou "ier"
                resultat.contents = premierChar + "er";
                
                // S'assurer qu'on a bien les caractères attendus
                if (resultat.characters.length >= 3) {
                  // Appliquer le style pour "I" ou "i"
                  resultat.characters[0].appliedCharacterStyle = romainsStyle;
                  self.Utilities.appliquerStylePolice(resultat.characters[0], estItaliqueLocal, estGrasLocal);
                  
                  // Appliquer le style pour "er"
                  resultat.characters[1].appliedCharacterStyle = exposantStyle;
                  resultat.characters[2].appliedCharacterStyle = exposantStyle;
                  
                  self.Utilities.appliquerStylePolice(resultat.characters[1], estItaliqueLocal, estGrasLocal);
                  self.Utilities.appliquerStylePolice(resultat.characters[2], estItaliqueLocal, estGrasLocal);
                } else {
                  // Si le nombre de caractères n'est pas celui attendu, journal d'erreur
                  try {
                    if (typeof console !== "undefined" && console && console.log) {
                      console.log("Attention: Nombre incorrect de caractères pour le premier siècle: " + resultat.contents);
                    }
                  } catch (e) {
                    // Ignorer les erreurs de journalisation
                  }
                  
                  // Tentons quand même d'appliquer les styles aux caractères disponibles
                  if (resultat.characters.length > 0) {
                    resultat.characters[0].appliedCharacterStyle = romainsStyle;
                    self.Utilities.appliquerStylePolice(resultat.characters[0], estItaliqueLocal, estGrasLocal);
                    
                    // Appliquer le style pour "er" s'ils existent
                    for (var i = 1; i < resultat.characters.length; i++) {
                      resultat.characters[i].appliedCharacterStyle = exposantStyle;
                      self.Utilities.appliquerStylePolice(resultat.characters[i], estItaliqueLocal, estGrasLocal);
                    }
                  }
                }
              }
              
            for (var ps = 0; ps < modelesPremierSiecle.length; ps++) {
              this.Utilities.appliquerGrepPartout(doc, modelesPremierSiecle[ps], null, traiterPremierSiecle);
            }

            // === TRAITEMENT POUR LE Ier SIÈCLE DÉJÀ FORMATÉ ===
            // Rechercher "Ier siècle" ou "ier siècle" déjà présent
            var modelesPremierSiecleExistant = [
                "(?<=\\s)([Ii])er(?=\\s+(?i:siècle))",
                "(?<=\\s)([Ii])er(?=[\\.,;:\\s!\\?])(?!\\s+(?i:" + motsClefsRegex + "))",
                "(?<=\\s)([Ii])er(?=[\\)\\]\\}\\>])(?!\\s+(?i:" + motsClefsRegex + "))"
            ];
            
            function traiterPremierSiecleExistant(resultat) {
              var estItaliqueLocal = self.Utilities.estEnItalique(resultat);
              var estGrasLocal = self.Utilities.estEnGras(resultat);

              // Appliquer le style pour la première lettre (I/i)
              resultat.characters[0].appliedCharacterStyle = romainsStyle;
              self.Utilities.appliquerStylePolice(resultat.characters[0], estItaliqueLocal, estGrasLocal);

              // Appliquer le style d'exposant aux caractères "er" (indices 1 et 2)
              if (resultat.characters.length >= 3) {
                resultat.characters[1].appliedCharacterStyle = exposantStyle;
                resultat.characters[2].appliedCharacterStyle = exposantStyle;

                self.Utilities.appliquerStylePolice(resultat.characters[1], estItaliqueLocal, estGrasLocal);
                self.Utilities.appliquerStylePolice(resultat.characters[2], estItaliqueLocal, estGrasLocal);
              }
            }

            for (var pse = 0; pse < modelesPremierSiecleExistant.length; pse++) {
              this.Utilities.appliquerGrepPartout(doc, modelesPremierSiecleExistant[pse], null, traiterPremierSiecleExistant);
            }
            
            // === FORMATAGE DES SIÈCLES - MAJUSCULES ===
            var modeleSiecleMaj = "(?<=\\s)([IVX]{1,5})e(?=[\\.,;:\\s!\\?\\)\\]\\}\\>])(?!\\s+(?i:" + motsClefsRegex + "))";

            function traiterSiecleMaj(resultat) {
              var texteOriginal = resultat.contents;

              // Vérifier si le texte est en italique ou en gras
              var estItaliqueLocal = self.Utilities.estEnItalique(resultat);
              var estGrasLocal = self.Utilities.estEnGras(resultat);

              // Séparer le chiffre romain du "e"
              var chiffreRomain = texteOriginal.slice(0, -1).toLowerCase();
              var e = texteOriginal.slice(-1);

              // Remplacer le texte
              resultat.contents = chiffreRomain + e;

              // Appliquer les styles séparément
              for (var c = 0; c < resultat.characters.length - 1; c++) {
                resultat.characters[c].appliedCharacterStyle = romainsStyle;
                self.Utilities.appliquerStylePolice(resultat.characters[c], estItaliqueLocal, estGrasLocal);
              }

              // Appliquer le style d'exposant au dernier caractère
              resultat.characters[resultat.characters.length - 1].appliedCharacterStyle = exposantStyle;
              self.Utilities.appliquerStylePolice(resultat.characters[resultat.characters.length - 1], estItaliqueLocal, estGrasLocal);
            }

            this.Utilities.appliquerGrepPartout(doc, modeleSiecleMaj, null, traiterSiecleMaj);
            
            // === FORMATAGE DES SIÈCLES - MINUSCULES ===
            var modeleSiecleMin = "(?<=\\s)(?!(?i:" + ambigusRegex + "))([ivx]{1,5})e(?=[\\.,;:\\s!\\?\\)\\]\\}\\>])(?!\\s+(?i:" + motsClefsRegex + "))";

            function traiterSiecleMin(resultat) {
              var texteOriginal = resultat.contents;

              // Vérifier si le texte est en italique ou en gras
              var estItaliqueLocal = self.Utilities.estEnItalique(resultat);
              var estGrasLocal = self.Utilities.estEnGras(resultat);

              // Séparer le chiffre romain du "e"
              var chiffreRomain = texteOriginal.slice(0, -1);
              var e = texteOriginal.slice(-1);

              // Remplacer le texte
              resultat.contents = chiffreRomain + e;

              // Appliquer les styles séparément
              for (var c = 0; c < resultat.characters.length - 1; c++) {
                resultat.characters[c].appliedCharacterStyle = romainsStyle;
                self.Utilities.appliquerStylePolice(resultat.characters[c], estItaliqueLocal, estGrasLocal);
              }

              // Appliquer le style d'exposant au dernier caractère
              resultat.characters[resultat.characters.length - 1].appliedCharacterStyle = exposantStyle;
              self.Utilities.appliquerStylePolice(resultat.characters[resultat.characters.length - 1], estItaliqueLocal, estGrasLocal);
            }

            this.Utilities.appliquerGrepPartout(doc, modeleSiecleMin, null, traiterSiecleMin);
            
            // === TRAITEMENT EXPLICITE POUR "SIÈCLE" ET "SIÈCLES" ===
            var modeleSiecleExplicite = "(?<=\\s)([IVXivx]{1,5})e(?=\\s+(?i:siècles?))";
            
            function traiterSiecleExplicite(resultat) {
              var texteOriginal = resultat.contents;

              // Vérifier si le texte est en italique ou en gras
              var estItaliqueLocal = self.Utilities.estEnItalique(resultat);
              var estGrasLocal = self.Utilities.estEnGras(resultat);

              // Séparer le chiffre romain du "e"
              var chiffreRomain = texteOriginal.slice(0, -1).toLowerCase();
              var e = texteOriginal.slice(-1);

              // Remplacer le texte
              resultat.contents = chiffreRomain + e;

              // Appliquer les styles séparément
              for (var c = 0; c < resultat.characters.length - 1; c++) {
                resultat.characters[c].appliedCharacterStyle = romainsStyle;
                self.Utilities.appliquerStylePolice(resultat.characters[c], estItaliqueLocal, estGrasLocal);
              }

              // Appliquer le style d'exposant au dernier caractère
              resultat.characters[resultat.characters.length - 1].appliedCharacterStyle = exposantStyle;
              self.Utilities.appliquerStylePolice(resultat.characters[resultat.characters.length - 1], estItaliqueLocal, estGrasLocal);
            }

            this.Utilities.appliquerGrepPartout(doc, modeleSiecleExplicite, null, traiterSiecleExplicite);
          } catch (error) {
            alert(I18n.__("errorFormatSiecles", error));
          }
        },
        
        /**
         * Formate les expressions ordinales (ex: IIIe régiment, Ire République)
         */
        formaterOrdinaux: function(doc, romainsMajStyle, exposantStyle) {
            try {
                // Créer une référence locale au module
                var self = this;
                
                // Espace insécable
                var ESPACE_INSECABLE = String.fromCharCode(0x00A0);
                
                // Préparer les expressions régulières
                var motsClefsRegex = this.Utilities.preparerRegex(this.getConfigData("MOTS_ORDINAUX", "data.motsOrdinaux"));
                var motsClefsRegexAvant = this.Utilities.preparerRegex(this.getConfigData("MOTS_AVANT_ORDINAUX", "data.motsAvantOrdinaux"));
                var ambigusRegex = this.Utilities.preparerRegexAmbigus(this.getConfigData("MOTS_AMBIGUS", "data.motsAmbigus"));
                
                // === CORRECTIONS PRÉLIMINAIRES ===
                // Remplacer les caractères problématiques avant tout traitement
                app.findGrepPreferences = app.changeGrepPreferences = null;
                
                // 1. Corriger les formes féminines avec accent
                app.findGrepPreferences.findWhat = "\\b([Ii])ère\\b";
                app.changeGrepPreferences.changeTo = "$1re";
                doc.changeGrep();

                // 2. Corriger les formes féminines sans accent
                app.findGrepPreferences = app.changeGrepPreferences = null;
                app.findGrepPreferences.findWhat = "\\b([Ii])ere\\b";
                app.changeGrepPreferences.changeTo = "$1re";
                doc.changeGrep();

                // 3. Corriger le cas "IeR"
                app.findGrepPreferences = app.changeGrepPreferences = null;
                app.findGrepPreferences.findWhat = "\\b([Ii])eR\\b";
                app.changeGrepPreferences.changeTo = "$1er";
                doc.changeGrep();
                
                // === TRAITEMENT DES ORDINAUX SPÉCIFIQUES ===
                // Rechercher "Ier" et "Ire" premier/première, mais pas les mots ambigus
                app.findGrepPreferences = app.changeGrepPreferences = null;
                
                // Expression qui exclut les mots ambigus
                app.findGrepPreferences.findWhat = "\\b(?!(?i:" + ambigusRegex + "))([Ii])(?:er|re)\\b";
                
                var resultatsIer = doc.findGrep();
                
                for (var i = 0; i < resultatsIer.length; i++) {
                    try {
                        var resultat = resultatsIer[i];
                        var estItaliqueLocal = self.Utilities.estEnItalique(resultat);
                        var estGrasLocal = self.Utilities.estEnGras(resultat);
                        
                        // Appliquer le style au "I"
                        resultat.characters[0].appliedCharacterStyle = romainsMajStyle;
                        self.Utilities.appliquerStylePolice(resultat.characters[0], estItaliqueLocal, estGrasLocal);
                        
                        // Appliquer le style d'exposant au reste
                        for (var j = 1; j < resultat.characters.length; j++) {
                            resultat.characters[j].appliedCharacterStyle = exposantStyle;
                            self.Utilities.appliquerStylePolice(resultat.characters[j], estItaliqueLocal, estGrasLocal);
                        }
                    } catch (e) {
                        // Ignorer les erreurs
                    }
                }
                
                // === CODE ORIGINAL POUR LES CAS GÉNÉRAUX (II, III, IV, etc.) ===
                // Cas 1: précédés par des mots spécifiques (majuscules + minuscules)
                var modelesOrdCase1 = [
                    "(?<=(?i:" + motsClefsRegexAvant + ")\\s)([IVX]{1,5})e(?=\\s)",
                    "(?<=(?i:" + motsClefsRegexAvant + ")\\s)(?!(?i:" + ambigusRegex + "))([ivx]{1,5})e(?=\\s)"
                ];

                // Fonction pour traiter les expressions ordinales
                function traiterExpressionOrdinale(resultat) {
                    try {
                        var texteOriginal = resultat.contents;
                        var estItaliqueLocal = self.Utilities.estEnItalique(resultat);
                        var estGrasLocal = self.Utilities.estEnGras(resultat);
                        
                        // Séparer le chiffre romain du "e"
                        var chiffreRomain = texteOriginal.slice(0, -1);
                        var e = texteOriginal.slice(-1);
                        
                        // Appliquer les styles séparément
                        for (var c = 0; c < resultat.characters.length - 1; c++) {
                            resultat.characters[c].appliedCharacterStyle = romainsMajStyle;
                            self.Utilities.appliquerStylePolice(resultat.characters[c], estItaliqueLocal, estGrasLocal);
                        }
                        
                        // Appliquer le style d'exposant au dernier caractère
                        resultat.characters[resultat.characters.length - 1].appliedCharacterStyle = exposantStyle;
                        self.Utilities.appliquerStylePolice(resultat.characters[resultat.characters.length - 1], estItaliqueLocal, estGrasLocal);
                    } catch (e) {
                        // Ignorer les erreurs
                    }
                }
                
                // Appliquer aux cas 1
                for (var o = 0; o < modelesOrdCase1.length; o++) {
                    this.Utilities.appliquerGrepPartout(doc, modelesOrdCase1[o], null, traiterExpressionOrdinale);
                }

                // Cas 2: suivis par des mots spécifiques (majuscules + minuscules)
                var modelesOrdCase2 = [
                    "(?<=\\s)([IVX]{1,5})e(?=\\s+(?i:" + motsClefsRegex + "))",
                    "(?<=\\s)(?!(?i:" + ambigusRegex + "))([ivx]{1,5})e(?=\\s+(?i:" + motsClefsRegex + "))"
                ];

                // Appliquer aux cas 2
                for (var p = 0; p < modelesOrdCase2.length; p++) {
                    this.Utilities.appliquerGrepPartout(doc, modelesOrdCase2[p], null, traiterExpressionOrdinale);
                }
                
                // === RECHERCHE DIRECTE DES FORMES II, III, etc. + e ===
                // Cette approche est plus directe et peut attraper les cas manqués
                // En excluant explicitement les mots ambigus
                app.findGrepPreferences = app.changeGrepPreferences = null;
                app.findGrepPreferences.findWhat = "\\b(?!(?i:" + ambigusRegex + "))([IVXivx][IVXivx]+)e\\b";
                
                var resultatsAutres = doc.findGrep();
                
                for (var i = 0; i < resultatsAutres.length; i++) {
                    try {
                        var resultat = resultatsAutres[i];
                        var estItaliqueLocal = self.Utilities.estEnItalique(resultat);
                        var estGrasLocal = self.Utilities.estEnGras(resultat);
                        
                        // Séparer le chiffre romain du "e"
                        var texteOriginal = resultat.contents;
                        var longueurChiffre = texteOriginal.length - 1;
                        
                        // Appliquer le style au chiffre romain
                        for (var c = 0; c < longueurChiffre; c++) {
                            resultat.characters[c].appliedCharacterStyle = romainsMajStyle;
                            self.Utilities.appliquerStylePolice(resultat.characters[c], estItaliqueLocal, estGrasLocal);
                        }
                        
                        // Appliquer le style au "e"
                        resultat.characters[longueurChiffre].appliedCharacterStyle = exposantStyle;
                        self.Utilities.appliquerStylePolice(resultat.characters[longueurChiffre], estItaliqueLocal, estGrasLocal);
                    } catch (e) {
                        // Ignorer les erreurs
                    }
                }
                
                // === ESPACES INSÉCABLES ===
                // Ajouter des espaces insécables après les ordinaux
                
                // Pour "Ier" et "Ire" suivis d'un mot commençant par une majuscule
                app.findGrepPreferences = app.changeGrepPreferences = null;
                app.findGrepPreferences.findWhat = "(\\b[Ii](?:er|re)) ([A-Z]\\w*)";
                app.changeGrepPreferences.changeTo = "$1" + ESPACE_INSECABLE + "$2";
                doc.changeGrep();

                // Pour les chiffres romains + e suivis d'un mot commençant par une majuscule
                app.findGrepPreferences = app.changeGrepPreferences = null;
                app.findGrepPreferences.findWhat = "(\\b[IVXivx]+e) ([A-Z]\\w*)";
                app.changeGrepPreferences.changeTo = "$1" + ESPACE_INSECABLE + "$2";
                doc.changeGrep();
            } catch (error) {
                alert(I18n.__("errorFormatOrdinaux", error));
            }
        },
        
        /**
         * Formate les parties d'œuvres et titres de personnes
         */
        formaterReferences: function(doc, romainsMajStyle, exposantStyle) {
          try {
              // Créer une référence locale au module
              var self = this;
              
              // Espace insécable
              var ESPACE_INSECABLE = String.fromCharCode(0x00A0);
              
              app.findGrepPreferences = app.changeGrepPreferences = null;
              
              // Combiner les mots d'œuvres et les titres de personnes
              var tousLesMots = this.getConfigData("MOTS_OEUVRES", "data.motsOeuvres").concat(this.getConfigData("TITRES_PERSONNES", "data.titresPersonnes"));
              
              // Joindre tous les mots en un seul pattern d'alternation
              var motsJoined = tousLesMots.join("|");

              // ÉTAPE 1: Remplacer les espaces normaux par des espaces insécables (un seul changeGrep)
              app.findGrepPreferences = app.changeGrepPreferences = null;
              app.findGrepPreferences.findWhat = "(?i)(" + motsJoined + ") ([ivxIVX][ivxIVX]*)";
              app.changeGrepPreferences.changeTo = "$1" + ESPACE_INSECABLE + "$2";
              doc.changeGrep();

              // ÉTAPE 2: Formater les chiffres romains trouvés (un seul findGrep)
              app.findGrepPreferences = app.changeGrepPreferences = null;
              app.findGrepPreferences.findWhat = "(?i)(" + motsJoined + ")" + ESPACE_INSECABLE + "([ivxIVX][ivxIVX]*)";
              var resultats = doc.findGrep();

              // Pour chaque occurrence
              for (var r = 0; r < resultats.length; r++) {
                  // Le texte complet (mot + espace insécable + chiffre romain)
                  var texteComplet = resultats[r].contents;

                  // Trouver la position de l'espace insécable pour séparer mot et chiffre
                  var posEspace = texteComplet.indexOf(ESPACE_INSECABLE);
                  var mot = texteComplet.substring(0, posEspace);
                  var chiffreRomain = texteComplet.substring(posEspace + 1);

                  // Conserver les attributs
                  var estItaliqueLocal = self.Utilities.estEnItalique(resultats[r]);
                  var estGrasLocal = self.Utilities.estEnGras(resultats[r]);

                  // Remplacer le texte en mettant le chiffre romain en majuscules
                  resultats[r].contents = mot + ESPACE_INSECABLE + chiffreRomain.toUpperCase();

                  // Traiter le chiffre romain séparément
                  var longueurChiffre = chiffreRomain.length;
                  var debutChiffre = mot.length + 1; // +1 pour l'espace insécable

                  // Appliquer le style au chiffre romain uniquement
                  for (var c = debutChiffre; c < debutChiffre + longueurChiffre; c++) {
                      if (c < resultats[r].characters.length) {
                          resultats[r].characters[c].appliedCharacterStyle = romainsMajStyle;

                          // Préserver l'italique et/ou le gras
                          self.Utilities.appliquerStylePolice(resultats[r].characters[c], estItaliqueLocal, estGrasLocal);
                      }
                  }
              }
              
              // === TRAITEMENT UNIQUEMENT DES "Ier" DÉJÀ FORMATÉS (IGNORE LES "Ie") ===
              // Créer une expression regex pour les noms avec premier
              var nomsPremierRegex = this.Utilities.preparerRegex(this.getConfigData("NOMS_PREMIER", "data.nomsPremier"));
              
              // Rechercher seulement les cas déjà formatés en "Ier" (pas "Ie")
              app.findGrepPreferences = app.changeGrepPreferences = null;
              app.findGrepPreferences.findWhat = "(?i)(" + nomsPremierRegex + ")[ " + ESPACE_INSECABLE + "]([Ii])er";
              
              var resultatsIer = doc.findGrep();

              for (var i = 0; i < resultatsIer.length; i++) {
                  try {
                      var texteOriginal = resultatsIer[i].contents;
                      var estItaliqueLocal = self.Utilities.estEnItalique(resultatsIer[i]);
                      var estGrasLocal = self.Utilities.estEnGras(resultatsIer[i]);
                      
                      // Trouver l'espace qui sépare le nom du "Ier"
                      var indexEspace = -1;
                      for (var j = 0; j < texteOriginal.length; j++) {
                          if (texteOriginal.charAt(j) === ' ' || texteOriginal.charAt(j) === ESPACE_INSECABLE) {
                              indexEspace = j;
                              break;
                          }
                      }
                      
                      if (indexEspace !== -1) {
                          // Position du "I" après l'espace
                          var positionI = indexEspace + 1;
                          
                          // Appliquer le style pour le "I"
                          if (positionI < resultatsIer[i].characters.length) {
                              resultatsIer[i].characters[positionI].appliedCharacterStyle = romainsMajStyle;
                              self.Utilities.appliquerStylePolice(resultatsIer[i].characters[positionI], estItaliqueLocal, estGrasLocal);
                              
                              // Appliquer le style exposant pour "er"
                              for (var k = positionI + 1; k < resultatsIer[i].characters.length; k++) {
                                  resultatsIer[i].characters[k].appliedCharacterStyle = exposantStyle;
                                  self.Utilities.appliquerStylePolice(resultatsIer[i].characters[k], estItaliqueLocal, estGrasLocal);
                              }
                          }
                      }
                  } catch (e) {
                      // Ignorer les erreurs
                  }
              }
              
              // === TRAITEMENT DES CHIFFRES ARABES (1er) ===
              try {
                  // Rechercher explicitement "1er" (déjà formaté)
                  app.findGrepPreferences = app.changeGrepPreferences = null;
                  app.findGrepPreferences.findWhat = "(\\s|^)(1er)";
                  
                  var resultats1er = doc.findGrep();
                  
                  for (var r = 0; r < resultats1er.length; r++) {
                      try {
                          var texteComplet = resultats1er[r].contents;
                          var estItaliqueLocal = self.Utilities.estEnItalique(resultats1er[r]);
                          var estGrasLocal = self.Utilities.estEnGras(resultats1er[r]);
                          
                          // Identifier où commence le "1er"
                          var debut = texteComplet.indexOf("1");
                          
                          // Appliquer le style d'exposant aux caractères "er"
                          if (debut !== -1 && debut + 1 < resultats1er[r].characters.length) {
                              resultats1er[r].characters[debut + 1].appliedCharacterStyle = exposantStyle;
                              if (debut + 2 < resultats1er[r].characters.length) {
                                  resultats1er[r].characters[debut + 2].appliedCharacterStyle = exposantStyle;
                              }
                              
                              self.Utilities.appliquerStylePolice(resultats1er[r].characters[debut + 1], estItaliqueLocal, estGrasLocal);
                              if (debut + 2 < resultats1er[r].characters.length) {
                                  self.Utilities.appliquerStylePolice(resultats1er[r].characters[debut + 2], estItaliqueLocal, estGrasLocal);
                              }
                          }
                      } catch (e) {
                          // Ignorer les erreurs
                      }
                  }
              } catch (e) {
                  alert(I18n.__("errorFormat1er", e.message));
              }
          } catch (error) {
              alert(I18n.__("errorFormatReferences", error.message));
          }
        },
        
        /**
         * Formate les espaces insécables dans les références
         */
        formaterEspaces: function(doc) {
          var self = this;
          
          try {
              // Espace insécable
              var ESPACE_INSECABLE = String.fromCharCode(0x00A0);
              
              app.findGrepPreferences = app.changeGrepPreferences = null;
              
              // Préparer des abréviations plus précises
              var abreviationsPrecises = [];
              
              // Traiter les abréviations avec points
              for (var i = 0; i < this.getConfigData("ABREVIATIONS_REFS", "data.abreviationsRefs").length; i++) {
                  // Si l'abréviation contient déjà un point d'échappement, elle est correcte
                  if (this.getConfigData("ABREVIATIONS_REFS", "data.abreviationsRefs")[i].indexOf("\\.") !== -1) {
                      abreviationsPrecises.push(this.getConfigData("ABREVIATIONS_REFS", "data.abreviationsRefs")[i]);
                  } else {
                      // Sinon, ajouter des délimiteurs de mot
                      abreviationsPrecises.push("\\b" + this.getConfigData("ABREVIATIONS_REFS", "data.abreviationsRefs")[i] + "\\b");
                  }
              }
              
              // Ajouter les références aux volumes avec délimiteurs de mot
              for (var i = 0; i < this.getConfigData("ABREVIATIONS_VOLUMES", "data.abreviationsVolumes").length; i++) {
                  if (this.getConfigData("ABREVIATIONS_VOLUMES", "data.abreviationsVolumes")[i].indexOf("\\.") !== -1) {
                      abreviationsPrecises.push(this.getConfigData("ABREVIATIONS_VOLUMES", "data.abreviationsVolumes")[i]);
                  } else {
                      abreviationsPrecises.push("\\b" + this.getConfigData("ABREVIATIONS_VOLUMES", "data.abreviationsVolumes")[i] + "\\b");
                  }
              }
              
              // Ajouter les références temporelles avec délimiteurs de mot
              for (var i = 0; i < this.getConfigData("ABREVIATIONS_TEMPORELLES", "data.abreviationsTemporelles").length; i++) {
                  if (this.getConfigData("ABREVIATIONS_TEMPORELLES", "data.abreviationsTemporelles")[i].indexOf("\\.") !== -1) {
                      abreviationsPrecises.push(this.getConfigData("ABREVIATIONS_TEMPORELLES", "data.abreviationsTemporelles")[i]);
                  } else {
                      abreviationsPrecises.push("\\b" + this.getConfigData("ABREVIATIONS_TEMPORELLES", "data.abreviationsTemporelles")[i] + "\\b");
                  }
              }
              
              // Ajouter les références aux numéros avec délimiteurs de mot
              for (var i = 0; i < this.getConfigData("ABREVIATIONS_NUMEROS", "data.abreviationsNumeros").length; i++) {
                  if (this.getConfigData("ABREVIATIONS_NUMEROS", "data.abreviationsNumeros")[i].indexOf("\\.") !== -1) {
                      abreviationsPrecises.push(this.getConfigData("ABREVIATIONS_NUMEROS", "data.abreviationsNumeros")[i]);
                  } else if (this.getConfigData("ABREVIATIONS_NUMEROS", "data.abreviationsNumeros")[i].indexOf("°") !== -1) {
                      // Cas spécial pour n° et n°s - pas de \b à la fin
                      abreviationsPrecises.push("\\b" + this.getConfigData("ABREVIATIONS_NUMEROS", "data.abreviationsNumeros")[i]);
                  } else {
                      abreviationsPrecises.push("\\b" + this.getConfigData("ABREVIATIONS_NUMEROS", "data.abreviationsNumeros")[i] + "\\b");
                  }
              }
              
              // Ajouter les références directionnelles avec délimiteurs de mot
              for (var i = 0; i < this.getConfigData("ABREVIATIONS_DIRECTION", "data.abreviationsDirection").length; i++) {
                  if (this.getConfigData("ABREVIATIONS_DIRECTION", "data.abreviationsDirection")[i].indexOf("\\.") !== -1) {
                      abreviationsPrecises.push(this.getConfigData("ABREVIATIONS_DIRECTION", "data.abreviationsDirection")[i]);
                  } else {
                      abreviationsPrecises.push("\\b" + this.getConfigData("ABREVIATIONS_DIRECTION", "data.abreviationsDirection")[i] + "\\b");
                  }
              }
              
              // Ajouter les titres et appellations avec délimiteurs de mot
              for (var i = 0; i < this.getConfigData("TITRES_APPELLATIONS", "data.titresAppellations").length; i++) {
                  if (this.getConfigData("TITRES_APPELLATIONS", "data.titresAppellations")[i].indexOf("\\.") !== -1) {
                      abreviationsPrecises.push(this.getConfigData("TITRES_APPELLATIONS", "data.titresAppellations")[i]);
                  } else {
                      abreviationsPrecises.push("\\b" + this.getConfigData("TITRES_APPELLATIONS", "data.titresAppellations")[i] + "\\b");
                  }
              }
              
              // Joindre toutes les abréviations en un seul pattern d'alternation
              var abrevJoined = abreviationsPrecises.join("|");

              // 1. Abréviations suivies d'un chiffre ou d'un nombre (un seul changeGrep)
              app.findGrepPreferences = app.changeGrepPreferences = null;
              app.findGrepPreferences.findWhat = "(?i)(" + abrevJoined + ")[ \\t" + ESPACE_INSECABLE + "]+([0-9][0-9.,\\-\u2013\u2014]*(?:[0-9]|[a-z]))";
              app.changeGrepPreferences.changeTo = "$1" + ESPACE_INSECABLE + "$2";
              doc.changeGrep();

              // 2. Traitement spécial pour les intervalles de pages
              var abreviationsRefs = [];
              for (var i = 0; i < this.getConfigData("ABREVIATIONS_REFS", "data.abreviationsRefs").length; i++) {
                  if (this.getConfigData("ABREVIATIONS_REFS", "data.abreviationsRefs")[i].indexOf("\\.") !== -1) {
                      abreviationsRefs.push(this.getConfigData("ABREVIATIONS_REFS", "data.abreviationsRefs")[i]);
                  } else {
                      abreviationsRefs.push("\\b" + this.getConfigData("ABREVIATIONS_REFS", "data.abreviationsRefs")[i] + "\\b");
                  }
              }
              var refsJoined = abreviationsRefs.join("|");

              app.findGrepPreferences = app.changeGrepPreferences = null;
              app.findGrepPreferences.findWhat = "(?i)(" + refsJoined + ")[ \\t" + ESPACE_INSECABLE + "]+([0-9]+)[ ]*([-\u2013\u2014])[ ]*([0-9]+)";
              app.changeGrepPreferences.changeTo = "$1" + ESPACE_INSECABLE + "$2$3$4";
              doc.changeGrep();

              // 3. Traitement spécial pour les numéros avec lettre (un seul changeGrep)
              app.findGrepPreferences = app.changeGrepPreferences = null;
              app.findGrepPreferences.findWhat = "(?i)(" + abrevJoined + ")[ \\t" + ESPACE_INSECABLE + "]+([0-9]+[a-z])\\b";
              app.changeGrepPreferences.changeTo = "$1" + ESPACE_INSECABLE + "$2";
              doc.changeGrep();

              // 4. Traitement des unités de mesure précédées de nombres (un seul changeGrep)
              var unitesJoined = this.getConfigData("UNITES_MESURE", "data.unitesMesure").join("|");
              app.findGrepPreferences = app.changeGrepPreferences = null;
              app.findGrepPreferences.findWhat = "([0-9]+[.,]?[0-9]*)[ ](\\b(?:" + unitesJoined + ")(?=\\s|[.,;:!?)]|$))";
              app.changeGrepPreferences.changeTo = "$1" + ESPACE_INSECABLE + "$2";
              doc.changeGrep();
          } catch (error) {
              alert(I18n.__("errorFormatEspaces", error.message));
          }
        }
    };
    
    /**
     * Fonction principale pour initialiser et exécuter le script
     * @private
     */
    function main() {
        try {
            // Vérifier si un document est ouvert
            if (!Utilities.validateDocumentOpen()) {
                return;
            }
            
            // Récupérer les styles de caractère
            if (!ErrorHandler.ensureDefined(app, "app", true)) return;
            if (!ErrorHandler.ensureDefined(app.activeDocument, "app.activeDocument", true)) return;
            
            var doc = app.activeDocument;
            var styleInfo = Utilities.getCharacterStyles(doc);
            
            // Créer et afficher le dialogue
            var options = UIBuilder.createDialog(
                styleInfo.styles, 
                styleInfo.superscriptIndex, 
                styleInfo.italicIndex
            );
            
            // Traiter si l'utilisateur a cliqué sur Appliquer
            if (options) {
                try {
                    // Vérifier les constantes nécessaires
                    if (!ErrorHandler.ensureDefined(ScriptLanguage, "ScriptLanguage", true)) return;
                    if (!ErrorHandler.ensureDefined(UndoModes, "UndoModes", true)) return;
                    
                    // Exécuter les corrections dans un bloc d'annulation
                    app.doScript(
                        function() {
                            Processor.processDocuments(options);
                        },
                        ScriptLanguage.JAVASCRIPT,
                        undefined,
                        UndoModes.FAST_ENTIRE_SCRIPT,
                        CONFIG.SCRIPT_TITLE
                    );
                } catch (scriptError) {
                    ErrorHandler.handleError(scriptError, "script execution in undo block", true);
                }
            }
        } catch (error) {
            ErrorHandler.handleError(error, I18n.__("errorMainFunction"), true);
        }
    }
    
    // Initialiser le script avec une gestion des erreurs globale
    try {
        main();
    } catch (fatalError) {
        try {
            alert(I18n.__("errorFatal", fatalError.message));

            if (typeof console !== "undefined" && console && console.log) {
                console.log(I18n.__("errorFatal", fatalError.message));

                if (fatalError.line) {
                    console.log(I18n.__("errorLine") + fatalError.line + ")");
                }

                if (fatalError.stack) {
                    console.log("Stack trace: " + fatalError.stack);
                }
            }
        } catch (e) {
            alert(I18n.__("errorUnrecoverable"));
        }
    }
    
    /**
     * Fonction autonome pour appliquer un gabarit à la dernière page
     */
    function applyMasterToLastPageStandalone(document, masterName) {
        try {
            if (!document) {
                alert(I18n.__("errorInvalidDocument"));
                return false;
            }

            if (!masterName) {
                alert(I18n.__("errorInvalidMasterName"));
                return false;
            }
            
            // Accéder à la dernière page 
            var pageCount = document.pages.length;
            if (pageCount <= 0) {
                return false;
            }
            
            var lastPage = document.pages.item(pageCount - 1);
            if (!lastPage || !lastPage.isValid) {
                return false;
            }
            
            // Trouver le gabarit correspondant DANS CE DOCUMENT
            var targetMaster = null;
            for (var i = 0; i < document.masterSpreads.length; i++) {
                if (document.masterSpreads[i].name === masterName) {
                    targetMaster = document.masterSpreads[i];
                    break;
                }
            }
            
            if (!targetMaster) {
                alert(I18n.__("errorMasterNotFound", masterName, document.name));
                return false;
            }

            lastPage.appliedMaster = targetMaster;
            return true;
        } catch (error) {
            alert(I18n.__("errorApplyMaster", error.message));
            return false;
        }
    }
})();