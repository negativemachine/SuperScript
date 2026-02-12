/**
 * PageLayout — Conditional paragraph styles & master page application for InDesign
 *
 * Extracted from SuperScript v2.0 — contains the "Page Layout" tab functionality:
 *   - Apply a paragraph style to the first paragraph after "trigger" styles
 *   - Apply a master page (gabarit) to the last facing page
 *
 * Bilingual UI (FR/EN) via I18n module, auto-detected from InDesign locale.
 * Save/load user preferences via ConfigManager (pagelayout.json).
 *
 * Architecture:
 *   safeJSON       — ES3-compatible JSON stringify/parse
 *   I18n           — Bilingual UI translations (FR/EN)
 *   ConfigManager  — Save/load/autoLoad user preferences to JSON
 *   ErrorHandler   — Error handling with context
 *   Utilities      — Document validation, paragraph style helpers
 *   Corrections    — applyStyleAfterTriggers, applyMasterToLastPage
 *   ProgressBar    — Non-blocking palette progress indicator
 *   UIBuilder      — Single-tab dialog with config bar
 *   Processor      — Orchestrates corrections on active document
 *
 * @version 1.0
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
            FIRST_PARAGRAPH: "First paragraph"
        },
        SCRIPT_TITLE: "PageLayout"
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
                'errorMainFunction': 'main function',

                // Dialog
                'dialogTitle': 'PageLayout',

                // Page layout tab
                'tabPageLayout': 'Page Layout',
                'enableStyleAfterLabel': 'Conditional style application',
                'triggerStylesPanel': 'Trigger styles',
                'targetStyleLabel': 'Style to apply to the following paragraph:',
                'applyMasterToLastPageLabel': 'Apply master to last facing page',
                'masterLabel': 'Master to apply:',
                'noStylesAvailable': 'No styles available',
                'noMastersAvailable': 'No masters available',
                'defaultStyle': '[None]',

                // Config
                'saveConfigTitle': 'Save configuration',
                'loadConfigTitle': 'Load configuration',
                'saveConfigButton': 'Save',
                'loadConfigButton': 'Load',
                'configDetected': 'Config detected',
                'configNotDetected': '',
                'configSaved': 'Config saved',
                'configLoaded': 'Config loaded',
                'errorSaveConfig': 'Error saving config: %s',
                'errorLoadConfig': 'Error loading config: %s',
                'errorParseConfig': 'Error parsing config: %s',
                'errorOpenConfig': 'Error: Cannot open config file',

                // Progress
                'progressTitle': 'PageLayout — Processing...',
                'progressApplyConditionalStyles': 'Applying conditional styles...',
                'progressApplyMasterToLastPage': 'Applying master to last page...',
                'progressComplete': 'Complete',

                // Buttons
                'cancelButton': 'Cancel',
                'applyButton': 'Apply',
                'helpTooltip': 'Help',
                'helpDialogTitle': 'PageLayout — Help',
                'helpDialogHeader': 'PageLayout v1.0',
                'closeButton': 'Close',
                'helpContent': 'PageLayout applies two page layout operations:\n\n'
                    + '1. CONDITIONAL STYLE APPLICATION\n'
                    + 'Select trigger styles (e.g. Heading 1, Heading 2). '
                    + 'The first non-empty paragraph after a trigger block will receive the target style '
                    + '(e.g. "First paragraph").\n\n'
                    + '2. MASTER PAGE APPLICATION\n'
                    + 'Applies the selected master page to the last page of the document.\n\n'
                    + 'Configuration can be saved/loaded as JSON for reuse.',

                // Success
                'successCorrectionsApplied': 'PageLayout corrections applied successfully.'
            },
            'fr': {
                // App-level
                'errorInDesignAccess': "Impossible d'acc\u00E9der \u00E0 l'application InDesign.",
                'errorUnrecoverable': "Une erreur irr\u00E9cup\u00E9rable s'est produite.",
                'errorFatal': 'Erreur fatale\u2009: %s',
                'errorScriptHalted': "Le script a \u00E9t\u00E9 arr\u00EAt\u00E9 en raison d'une erreur fatale",
                'errorObjectUndefined': "L'objet '%s' est ind\u00E9fini ou null",
                'errorInContext': 'Erreur',
                'errorContextIn': ' dans ',
                'errorLine': ' (ligne ',
                'errorInDesignUnavailable': "L'application InDesign n'est pas accessible",
                'errorDocumentsUnavailable': "La collection de documents n'est pas accessible",
                'errorNoDocumentOpen': 'Veuillez ouvrir un document avant de lancer ce script.',
                'errorInvalidDocument': 'Erreur\u2009: Document invalide',
                'errorInvalidMasterName': 'Erreur\u2009: Nom de gabarit invalide',
                'errorMasterNotFound': "Le gabarit '%s' n'a pas \u00E9t\u00E9 trouv\u00E9 dans le document %s",
                'errorApplyMaster': "Erreur lors de l'application du gabarit\u2009: %s",
                'errorMainFunction': 'fonction principale',

                // Dialog
                'dialogTitle': 'PageLayout',

                // Page layout tab
                'tabPageLayout': 'Mise en page',
                'enableStyleAfterLabel': 'Application conditionnelle de styles',
                'triggerStylesPanel': 'Styles d\u00E9clencheurs',
                'targetStyleLabel': 'Style \u00E0 appliquer au paragraphe suivant\u2009:',
                'applyMasterToLastPageLabel': 'Appliquer un gabarit \u00E0 la derni\u00E8re page en vis-\u00E0-vis',
                'masterLabel': 'Gabarit \u00E0 appliquer\u2009:',
                'noStylesAvailable': 'Aucun style disponible',
                'noMastersAvailable': 'Aucun gabarit disponible',
                'defaultStyle': '[Aucun]',

                // Config
                'saveConfigTitle': 'Enregistrer la configuration',
                'loadConfigTitle': 'Charger la configuration',
                'saveConfigButton': 'Enregistrer',
                'loadConfigButton': 'Charger',
                'configDetected': 'Config d\u00E9tect\u00E9e',
                'configNotDetected': '',
                'configSaved': 'Config enregistr\u00E9e',
                'configLoaded': 'Config charg\u00E9e',
                'errorSaveConfig': "Erreur lors de l'enregistrement\u2009: %s",
                'errorLoadConfig': 'Erreur lors du chargement\u2009: %s',
                'errorParseConfig': "Erreur lors de l'analyse\u2009: %s",
                'errorOpenConfig': 'Erreur\u2009: Impossible d\'ouvrir le fichier de configuration',

                // Progress
                'progressTitle': 'PageLayout \u2014 Traitement en cours...',
                'progressApplyConditionalStyles': 'Application des styles conditionnels...',
                'progressApplyMasterToLastPage': 'Application du gabarit \u00E0 la derni\u00E8re page...',
                'progressComplete': 'Termin\u00E9',

                // Buttons
                'cancelButton': 'Annuler',
                'applyButton': 'Appliquer',
                'helpTooltip': 'Aide',
                'helpDialogTitle': 'PageLayout \u2014 Aide',
                'helpDialogHeader': 'PageLayout v1.0',
                'closeButton': 'Fermer',
                'helpContent': "PageLayout applique deux op\u00E9rations de mise en page\u2009:\n\n"
                    + "1. APPLICATION CONDITIONNELLE DE STYLES\n"
                    + "S\u00E9lectionnez les styles d\u00E9clencheurs (ex.\u2009: Titre 1, Titre 2). "
                    + "Le premier paragraphe non vide apr\u00E8s un bloc d\u00E9clencheur recevra le style cible "
                    + "(ex.\u2009: \u00AB\u2009First paragraph\u2009\u00BB).\n\n"
                    + "2. APPLICATION DE GABARIT\n"
                    + "Applique le gabarit s\u00E9lectionn\u00E9 \u00E0 la derni\u00E8re page du document.\n\n"
                    + "La configuration peut \u00EAtre enregistr\u00E9e et recharg\u00E9e en JSON.",

                // Success
                'successCorrectionsApplied': 'Corrections PageLayout appliqu\u00E9es avec succ\u00E8s.'
            }
        };

        function detectInDesignLanguage() {
            try {
                if (typeof app !== 'undefined' && app.locale) {
                    var locale = app.locale.toString().toLowerCase();
                    if (locale.indexOf('french') !== -1 || locale.indexOf('fran') !== -1) {
                        return 'fr';
                    }
                }
            } catch (e) {}
            return 'en';
        }

        currentLanguage = detectInDesignLanguage();

        return {
            __: function(key) {
                var dict = translations[currentLanguage] || translations['en'];
                var text = dict[key] || translations['en'][key] || key;
                if (arguments.length > 1) {
                    for (var i = 1; i < arguments.length; i++) {
                        text = text.replace(/%s/, String(arguments[i]));
                    }
                }
                return text;
            },
            setLanguage: function(lang) {
                if (translations[lang]) currentLanguage = lang;
            },
            getLanguage: function() {
                return currentLanguage;
            }
        };
    })();

    // =========================================================================
    // ConfigManager — Save/Load user preferences
    // =========================================================================

    var ConfigManager = (function() {
        var CONFIG_FILENAME = "pagelayout.json";
        var CONFIG_VERSION = 1;

        /**
         * Searches recursively for config files
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
            } catch (e) {}
            return files;
        }

        /**
         * Automatically loads configuration from near the active document
         * @return {Object|null} Object with { data, inConfigFolder } or null
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

                var inConfigFolder = false;
                var configFile = files[0];
                for (var fi = 0; fi < files.length; fi++) {
                    var parentFolder = files[fi].parent;
                    if (parentFolder && parentFolder.name.toLowerCase() === "config") {
                        configFile = files[fi];
                        inConfigFolder = true;
                        break;
                    }
                }

                configFile.encoding = "UTF-8";

                if (configFile.open("r")) {
                    try {
                        var content = configFile.read();
                        configFile.close();
                        var configData = safeJSON.parse(content);
                        return { data: configData, inConfigFolder: inConfigFolder };
                    } catch (e) {
                        return null;
                    }
                }
            } catch (e) {}
            return null;
        }

        /**
         * Builds runtime options from saved config (for silent mode)
         */
        function buildOptionsFromConfig(configData, doc) {
            var l = configData.layout || {};

            // Helper to find a paragraph style by name (recursively in groups)
            function findParaStyle(name) {
                if (!name) return null;
                try {
                    var result = null;
                    function searchInGroup(group) {
                        for (var i = 0; i < group.paragraphStyles.length; i++) {
                            if (group.paragraphStyles[i].name === name) {
                                result = group.paragraphStyles[i];
                                return;
                            }
                        }
                        for (var j = 0; j < group.paragraphStyleGroups.length; j++) {
                            searchInGroup(group.paragraphStyleGroups[j]);
                            if (result) return;
                        }
                    }
                    searchInGroup(doc);
                } catch (e) {}
                return result;
            }

            // Helper to find a master spread by name
            function findMaster(name) {
                if (!name) return null;
                try {
                    for (var i = 0; i < doc.masterSpreads.length; i++) {
                        if (doc.masterSpreads[i].name === name) return doc.masterSpreads[i];
                    }
                } catch (e) {}
                return null;
            }

            // Resolve trigger styles
            var triggerStyles = [];
            if (l.enableStyleAfter && l.triggerStyles) {
                for (var ti = 0; ti < l.triggerStyles.length; ti++) {
                    var ts = findParaStyle(l.triggerStyles[ti]);
                    if (ts) triggerStyles.push(ts);
                }
            }

            return {
                enableStyleAfter: !!l.enableStyleAfter,
                triggerStyles: triggerStyles,
                targetStyle: findParaStyle(l.targetStyle),
                applyMasterToLastPage: !!l.applyMasterToLastPage,
                selectedMaster: findMaster(l.masterName)
            };
        }

        /**
         * Saves configuration to a user-selected file
         */
        function save(configData) {
            try {
                var defaultPath = "";
                if (typeof app !== 'undefined' && app.activeDocument && app.activeDocument.saved) {
                    defaultPath = app.activeDocument.filePath + "/";
                }
                var defaultFile = new File(defaultPath + "pagelayout");
                var saveFile = defaultFile.saveDlg(
                    I18n.__("saveConfigTitle"),
                    "JSON files:*.json"
                );
                if (!saveFile) return false;

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
         */
        function collectFromDialog(controls) {
            var config = {
                version: CONFIG_VERSION,
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

            // Collect trigger style names from checkboxes
            if (controls.triggerCheckboxes) {
                for (var i = 0; i < controls.triggerCheckboxes.length; i++) {
                    if (controls.triggerCheckboxes[i].value) {
                        config.layout.triggerStyles.push(controls.triggerCheckboxes[i].text);
                    }
                }
            }

            return config;
        }

        /**
         * Applies loaded config data to dialog controls
         */
        function applyToDialog(configData, controls) {
            if (!configData) return;

            var l = configData.layout;
            if (!l) return;

            if (typeof l.enableStyleAfter === 'boolean') controls.cbEnableStyleAfter.value = l.enableStyleAfter;
            if (typeof l.applyMasterToLastPage === 'boolean') {
                controls.cbApplyMasterToLastPage.value = l.applyMasterToLastPage;
                controls.masterDropdown.enabled = l.applyMasterToLastPage;
            }

            // Trigger styles — restore checkboxes
            if (l.triggerStyles && controls.triggerCheckboxes) {
                for (var tc = 0; tc < controls.triggerCheckboxes.length; tc++) {
                    controls.triggerCheckboxes[tc].value = false;
                }
                for (var ts = 0; ts < l.triggerStyles.length; ts++) {
                    for (var tc2 = 0; tc2 < controls.triggerCheckboxes.length; tc2++) {
                        if (controls.triggerCheckboxes[tc2].text === l.triggerStyles[ts]) {
                            controls.triggerCheckboxes[tc2].value = true;
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

        return {
            autoLoad: autoLoad,
            save: save,
            load: load,
            collectFromDialog: collectFromDialog,
            applyToDialog: applyToDialog,
            buildOptionsFromConfig: buildOptionsFromConfig
        };
    })();

    // =========================================================================
    // ErrorHandler
    // =========================================================================

    var ErrorHandler = {
        handleError: function(error, context, isFatal) {
            var message = I18n.__("errorInContext");

            if (context) {
                message += I18n.__("errorContextIn") + context;
            }

            message += " : " + (error.message || "Unknown error");

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
            } catch (e) {}

            if (isFatal) {
                alert(message);
                throw new Error(I18n.__("errorScriptHalted"));
            }

            return message;
        },

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
                } catch (e) {}
                return false;
            }
            return true;
        }
    };

    // =========================================================================
    // Utilities
    // =========================================================================

    var Utilities = {
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

                function collectStyles(group) {
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
                                    styles = styles.concat(collectStyles(subgroup));
                                }
                            }
                        }
                    } catch (e) {}
                    return styles;
                }

                var allStyles = collectStyles(doc);
                result.styles = allStyles;

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
         * Vérifie si un paragraphe est vide
         */
        isEmptyParagraph: function(paragraph) {
            try {
                if (!paragraph || !paragraph.contents) return false;
                var contents = paragraph.contents;
                return contents.replace(/[\r\n\s\u200B\uFEFF]/g, "") === "";
            } catch (error) {
                return false;
            }
        }
    };

    // =========================================================================
    // Corrections — Style and master page operations
    // =========================================================================

    var Corrections = {
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
                if (!ErrorHandler.ensureDefined(doc.stories, "document.stories", true)) return;

                var stories = doc.stories;
                for (var s = 0; s < stories.length; s++) {
                    try {
                        var story = stories[s];
                        if (!story || !story.paragraphs) continue;

                        var paras = story.paragraphs;
                        var justLeftTriggerBlock = false;

                        for (var p = 0; p < paras.length; p++) {
                            try {
                                var para = paras[p];
                                if (!para || !para.appliedParagraphStyle) continue;

                                var style = para.appliedParagraphStyle;
                                var isTriggerStyle = false;

                                for (var t = 0; t < triggerStyles.length; t++) {
                                    if (style.id === triggerStyles[t].id) {
                                        isTriggerStyle = true;
                                        break;
                                    }
                                }

                                if (isTriggerStyle) {
                                    justLeftTriggerBlock = true;
                                } else if (justLeftTriggerBlock) {
                                    if (!Utilities.isEmptyParagraph(para)) {
                                        para.appliedParagraphStyle = targetStyle;
                                        justLeftTriggerBlock = false;
                                    }
                                }
                            } catch (paraError) {
                                ErrorHandler.handleError(paraError, "paragraph " + p, false);
                            }
                        }
                    } catch (storyError) {
                        ErrorHandler.handleError(storyError, "story " + s, false);
                    }
                }
            } catch (error) {
                ErrorHandler.handleError(error, "applyStyleAfterTriggers", false);
            }
        },

        /**
         * Applique un gabarit à la dernière page du document
         * @param {Document} doc - Document InDesign
         * @param {string} masterName - Nom du gabarit
         * @returns {boolean} True si l'application a réussi
         */
        applyMasterToLastPage: function(doc, masterName) {
            try {
                if (!doc) {
                    alert(I18n.__("errorInvalidDocument"));
                    return false;
                }
                if (!masterName) {
                    alert(I18n.__("errorInvalidMasterName"));
                    return false;
                }

                var pageCount = doc.pages.length;
                if (pageCount <= 0) return false;

                var lastPage = doc.pages.item(pageCount - 1);
                if (!lastPage || !lastPage.isValid) return false;

                var targetMaster = null;
                for (var i = 0; i < doc.masterSpreads.length; i++) {
                    if (doc.masterSpreads[i].name === masterName) {
                        targetMaster = doc.masterSpreads[i];
                        break;
                    }
                }

                if (!targetMaster) {
                    alert(I18n.__("errorMasterNotFound", masterName, doc.name));
                    return false;
                }

                lastPage.appliedMaster = targetMaster;
                return true;
            } catch (error) {
                alert(I18n.__("errorApplyMaster", error.message));
                return false;
            }
        }
    };

    // =========================================================================
    // ProgressBar — Non-blocking palette
    // =========================================================================

    var ProgressBar = {
        progressWin: null,

        create: function(title, maxValue) {
            try {
                this.progressWin = new Window("palette", title);
                this.progressWin.progressBar = this.progressWin.add("progressbar", undefined, 0, maxValue);
                this.progressWin.progressBar.preferredSize.width = 300;
                this.progressWin.status = this.progressWin.add("statictext", undefined, "");
                this.progressWin.status.preferredSize.width = 300;
                this.progressWin.center();
                this.progressWin.show();
            } catch (e) {
                this.progressWin = null;
            }
        },

        update: function(value, statusText) {
            if (!this.progressWin) return;
            try {
                this.progressWin.progressBar.value = value;
                this.progressWin.status.text = statusText;
                this.progressWin.update();
            } catch (e) {}
        },

        close: function() {
            if (this.progressWin) {
                try {
                    this.progressWin.close();
                    this.progressWin = null;
                } catch (e) {}
            }
        }
    };

    // =========================================================================
    // UIBuilder — Single-tab dialog
    // =========================================================================

    var UIBuilder = {
        /**
         * Crée et affiche le dialogue du script
         * @param {Object} [preloadConfig] - Config to preload (reopen after UI language switch)
         * @returns {Object|null} Options de l'utilisateur ou null si annulé
         */
        createDialog: function(preloadConfig) {
            try {
                var doc = app.activeDocument;

                // Création du dialogue principal
                var dialog = new Window("dialog", CONFIG.SCRIPT_TITLE);
                dialog.orientation = "column";
                dialog.alignChildren = "fill";
                dialog.preferredSize.width = 400;

                // Bannière supérieure avec attribution + sélecteur de langue UI
                var topBanner = dialog.add("group");
                topBanner.orientation = "row";
                topBanner.alignment = "right";
                var attribution = topBanner.add("statictext", undefined, "entremonde / Spectral lab");
                topBanner.add("statictext", undefined, "  ");
                var langDropdown = topBanner.add("dropdownlist", undefined, ["En", "Fr"]);
                langDropdown.selection = I18n.getLanguage() === 'fr' ? 1 : 0;

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

                var dialogControls = {};

                // Wire up UI language dropdown
                langDropdown.onChange = function() {
                    var newLang = langDropdown.selection.index === 1 ? 'fr' : 'en';
                    if (newLang !== I18n.getLanguage()) {
                        dialog.reopenConfig = ConfigManager.collectFromDialog(dialogControls);
                        I18n.setLanguage(newLang);
                        dialog.close(3); // code 3 = reopen with new UI language
                    }
                };

                // Création de l'onglet unique
                var tpanel = dialog.add("tabbedpanel");
                tpanel.alignChildren = "fill";

                var tabStyles = tpanel.add("tab", undefined, I18n.__("tabPageLayout"));
                tabStyles.orientation = "column";
                tabStyles.alignChildren = "left";

                tpanel.selection = tabStyles;

                // Fonction utilitaire pour ajouter une case à cocher
                function addCheckboxOption(parent, label, checked) {
                    var group = parent.add("group");
                    group.orientation = "row";
                    group.alignChildren = "left";
                    var checkbox = group.add("checkbox", undefined, label);
                    checkbox.value = checked;
                    return checkbox;
                }

                // === APPLICATION CONDITIONNELLE DE STYLES ===
                var cbEnableStyleAfter = addCheckboxOption(tabStyles, I18n.__("enableStyleAfterLabel"), true);

                // Récupérer les styles de paragraphe du document actif
                var allParaStyles = [];
                var paraStyleNames = [];

                try {
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
                        } catch (e) {}
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
                styleGroupPanel.alignChildren = "fill";

                // 3-column layout for trigger style checkboxes
                var triggerCheckboxes = [];
                var colCount = 3;
                var rowCount = Math.ceil(paraStyleNames.length / colCount);
                for (var ri = 0; ri < rowCount; ri++) {
                    var row = styleGroupPanel.add("group");
                    row.orientation = "row";
                    row.alignChildren = "left";
                    row.alignment = ["fill", "top"];
                    for (var ci = 0; ci < colCount; ci++) {
                        var idx = ri * colCount + ci;
                        if (idx < paraStyleNames.length) {
                            var cb = row.add("checkbox", undefined, paraStyleNames[idx]);
                            cb.preferredSize.width = 150;
                            cb.value = (paraStyleNames[idx] !== "Body text" && paraStyleNames[idx] !== "First paragraph");
                            triggerCheckboxes.push(cb);
                        }
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
                for (var fpi = 0; fpi < paraStyleNames.length; fpi++) {
                    if (paraStyleNames[fpi] === "First paragraph") {
                        firstParaIndex = fpi;
                        break;
                    }
                }

                if (firstParaIndex !== -1) {
                    targetStyleDropdown.selection = firstParaIndex;
                } else if (paraStyleNames.length > 0) {
                    targetStyleDropdown.selection = 0;
                }

                // === APPLICATION DU GABARIT À LA DERNIÈRE PAGE ===
                var cbApplyMasterToLastPage = addCheckboxOption(tabStyles, I18n.__("applyMasterToLastPageLabel"), true);

                // Récupérer les gabarits du document
                var masterNames = [];
                var allMasters = [];

                try {
                    if (doc.masterSpreads) {
                        for (var mi = 0; mi < doc.masterSpreads.length; mi++) {
                            var master = doc.masterSpreads[mi];
                            if (master && master.name) {
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

                cbApplyMasterToLastPage.onClick = function() {
                    masterDropdown.enabled = cbApplyMasterToLastPage.value;
                };

                if (masterNames.length > 0) {
                    masterDropdown.selection = 0;
                }

                // Populate dialogControls
                dialogControls.cbEnableStyleAfter = cbEnableStyleAfter;
                dialogControls.triggerCheckboxes = triggerCheckboxes;
                dialogControls.targetStyleDropdown = targetStyleDropdown;
                dialogControls.cbApplyMasterToLastPage = cbApplyMasterToLastPage;
                dialogControls.masterDropdown = masterDropdown;

                // Load configuration: preloadConfig (reopen) takes priority over autoLoad
                var configToApply = preloadConfig;
                if (!configToApply) {
                    var autoResult = ConfigManager.autoLoad();
                    if (autoResult) configToApply = autoResult.data;
                }
                if (configToApply) {
                    if (!preloadConfig) configStatusText.text = I18n.__("configDetected");
                    ConfigManager.applyToDialog(configToApply, dialogControls);
                }

                // Wire Save button
                saveConfigBtn.onClick = function() {
                    var configData = ConfigManager.collectFromDialog(dialogControls);
                    if (ConfigManager.save(configData)) {
                        configStatusText.text = I18n.__("configSaved");
                    }
                };

                // Wire Load button
                loadConfigBtn.onClick = function() {
                    var configData = ConfigManager.load();
                    if (configData) {
                        ConfigManager.applyToDialog(configData, dialogControls);
                        configStatusText.text = I18n.__("configLoaded");
                    }
                };

                // Boutons d'action
                var buttonGroup = dialog.add("group");
                buttonGroup.orientation = "row";
                buttonGroup.alignment = "right";

                var helpButton = buttonGroup.add("button", undefined, "?");
                helpButton.preferredSize.width = 25;
                helpButton.preferredSize.height = 25;
                helpButton.helpTip = I18n.__("helpTooltip");

                var cancelButton = buttonGroup.add("button", undefined, I18n.__("cancelButton"), {name: "cancel"});
                var okButton = buttonGroup.add("button", undefined, I18n.__("applyButton"), {name: "ok"});

                helpButton.onClick = function() {
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

                    helpDialog.show();
                };

                cancelButton.onClick = function() {
                    dialog.close(2);
                };

                // Afficher le dialogue
                var dialogResult = dialog.show();

                // Code 3 = reopen with new UI language
                if (dialogResult == 3 && dialog.reopenConfig) {
                    return UIBuilder.createDialog(dialog.reopenConfig);
                }

                if (dialogResult == 1) {
                    try {
                        var selectedTriggerStyles = [];
                        for (var sti = 0; sti < triggerCheckboxes.length; sti++) {
                            if (triggerCheckboxes[sti].value && sti < allParaStyles.length) {
                                selectedTriggerStyles.push(allParaStyles[sti]);
                            }
                        }

                        return {
                            enableStyleAfter: cbEnableStyleAfter.value,
                            triggerStyles: selectedTriggerStyles,
                            targetStyle: targetStyleDropdown.selection && allParaStyles.length > 0 ?
                                allParaStyles[targetStyleDropdown.selection.index] : null,
                            applyMasterToLastPage: cbApplyMasterToLastPage.value,
                            selectedMaster: masterDropdown.selection && allMasters.length > 0 ?
                                allMasters[masterDropdown.selection.index] : null
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

    // =========================================================================
    // Processor — Orchestrates operations
    // =========================================================================

    var Processor = {
        applyCorrections: function(doc, options) {
            try {
                if (!ErrorHandler.ensureDefined(doc, "document", true)) return;
                if (!ErrorHandler.ensureDefined(options, "options", true)) return;

                var totalSteps = 0;
                if (options.enableStyleAfter && options.triggerStyles && options.triggerStyles.length > 0 && options.targetStyle) totalSteps++;
                if (options.applyMasterToLastPage && options.selectedMaster) totalSteps++;

                if (totalSteps === 0) return;

                if (!options.silentMode) {
                    ProgressBar.create(I18n.__("progressTitle"), totalSteps);
                }

                try {
                    var progress = 0;

                    if (options.enableStyleAfter &&
                        options.triggerStyles &&
                        options.triggerStyles.length > 0 &&
                        options.targetStyle) {
                        ProgressBar.update(++progress, I18n.__("progressApplyConditionalStyles"));
                        Corrections.applyStyleAfterTriggers(doc, options.triggerStyles, options.targetStyle);
                    }

                    if (options.applyMasterToLastPage && options.selectedMaster) {
                        ProgressBar.update(++progress, I18n.__("progressApplyMasterToLastPage"));
                        var masterName = options.selectedMaster.name;
                        Corrections.applyMasterToLastPage(doc, masterName);
                    }

                    ProgressBar.update(totalSteps, I18n.__("progressComplete"));

                } catch (correctionsError) {
                    ErrorHandler.handleError(correctionsError, "applying corrections", false);
                } finally {
                    ProgressBar.close();
                }
            } catch (error) {
                ErrorHandler.handleError(error, "applyCorrections", true);
            }
        },

        processDocuments: function(options) {
            try {
                if (!ErrorHandler.ensureDefined(options, "options", true)) return;
                if (!ErrorHandler.ensureDefined(app, "app", true)) return;
                if (!ErrorHandler.ensureDefined(app.activeDocument, "app.activeDocument", true)) return;

                Processor.applyCorrections(app.activeDocument, options);
                if (!options.silentMode) {
                    alert(I18n.__("successCorrectionsApplied"));
                }
            } catch (error) {
                ErrorHandler.handleError(error, "processDocuments", true);
            }
        }
    };

    // =========================================================================
    // main() — Entry point
    // =========================================================================

    function main() {
        try {
            if (!Utilities.validateDocumentOpen()) {
                return;
            }

            if (!ErrorHandler.ensureDefined(app, "app", true)) return;
            if (!ErrorHandler.ensureDefined(app.activeDocument, "app.activeDocument", true)) return;

            var doc = app.activeDocument;

            // Check if called from BookCreator via scriptArgs
            var callerArg = "";
            try { callerArg = app.scriptArgs.getValue("caller") || ""; } catch (e) {}
            try { app.scriptArgs.setValue("caller", ""); } catch (e) {}

            var isFromBookCreator = (callerArg === "BookCreator");

            // Priority 1: pagelayout.json in config/ subfolder
            var autoResult = ConfigManager.autoLoad();
            if (autoResult && autoResult.data && (autoResult.inConfigFolder || isFromBookCreator)) {
                try {
                    var silentOptions = ConfigManager.buildOptionsFromConfig(autoResult.data, doc);
                    silentOptions.silentMode = true;

                    if (!ErrorHandler.ensureDefined(ScriptLanguage, "ScriptLanguage", true)) return;
                    if (!ErrorHandler.ensureDefined(UndoModes, "UndoModes", true)) return;

                    app.doScript(
                        function() {
                            Processor.processDocuments(silentOptions);
                        },
                        ScriptLanguage.JAVASCRIPT,
                        undefined,
                        UndoModes.FAST_ENTIRE_SCRIPT,
                        CONFIG.SCRIPT_TITLE
                    );

                    doc.save();
                } catch (silentError) {
                    ErrorHandler.handleError(silentError, "silent mode execution", !isFromBookCreator);
                }
                return;
            }

            // If called from BookCreator but no config — exit silently
            if (isFromBookCreator) {
                return;
            }

            // Interactive mode
            var options = UIBuilder.createDialog();

            if (options) {
                try {
                    if (!ErrorHandler.ensureDefined(ScriptLanguage, "ScriptLanguage", true)) return;
                    if (!ErrorHandler.ensureDefined(UndoModes, "UndoModes", true)) return;

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

    // Initialiser le script
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
})();
