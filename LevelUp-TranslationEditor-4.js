(function () {
    var ROOT_ID = "lu_translate_panel";
    var STYLE_ID = "lu_translate_style";
    var Xrm = window.Xrm || (window.parent && window.parent.Xrm) || (window.top && window.top.Xrm);
    var formContext = Xrm && Xrm.Page;
    var state = {
        targetLcid: null,
        userLcid: null,
        orgLcid: null,
        languages: [],
        allLanguages: [],
        showAllLanguages: false,
        showLockedFields: false,
        scope: "fields",
        search: "",
        items: [],
        entityName: "",
        entityLogicalName: "",
        formId: "",
        loading: false,
        hasChanges: false,
        itemsCache: {},
        pendingChanges: {},
        solutionTracking: {
            enabled: false,
            solutionId: null,
            solutionName: null,
            solutions: [],
            componentsCache: {},
            loading: true
        }
    };

    if (!Xrm || !formContext || !formContext.data || !formContext.ui) {
        alert("No supported form context found.");
        return;
    }

    if (document.getElementById(ROOT_ID)) {
        removePanel();
        return;
    }

    state.entityName = formContext.data.entity.getEntityName();
    state.entityLogicalName = state.entityName;
    state.formId = formContext.ui.formSelector.getCurrentItem() ? formContext.ui.formSelector.getCurrentItem().getId() : "";
    state.userLcid = Xrm.Utility.getGlobalContext().userSettings.languageId;
    state.orgLcid = Xrm.Utility.getGlobalContext().organizationSettings.languageId;

    addStyles();
    renderShell();
    setStatus("Loading languages...");
    
    loadLanguages().then(function () {
        if (state.languages.length > 0) {
            state.targetLcid = state.languages[0].lcid;
            updateLanguageDisplay();
        }
        // Update checkbox state after loading languages
        var showAllLangsCheck = document.getElementById("lu_show_all_langs");
        if (showAllLangsCheck) {
            showAllLangsCheck.checked = state.showAllLanguages;
        }
        // Load solutions for solution tracking
        return loadSolutions().then(function() {
            updateSolutionDisplay();
            return loadTranslations();
        });
    }).then(function () {
        renderItems();
        setStatus("Ready");
    }, function (err) {
        setStatus("Error");
        alert("Load failed: " + getErrorMessage(err));
    });

    function removePanel() {
        var el = document.getElementById(ROOT_ID);
        if (el && el.parentNode) el.parentNode.removeChild(el);
    }

    function esc(v) {
        return String(v == null ? "" : v)
            .replace(/&/g, "&amp;")
            .replace(/</g, "&lt;")
            .replace(/>/g, "&gt;")
            .replace(/"/g, "&quot;");
    }

    function cssEsc(v) {
        return String(v || "").replace(/\\/g, "\\\\").replace(/"/g, "\\\"");
    }

    function getErrorMessage(err) {
        if (!err) return "Unknown error";
        if (typeof err === "string") return err;
        if (err.message) return err.message;
        try { return JSON.stringify(err); } catch (e) { return "Unknown error"; }
    }

    function labelText(labelObj, lcid, fallback) {
        if (!labelObj) return fallback || "";
        if (typeof labelObj === "string") return labelObj;
        
        if (labelObj.LocalizedLabels) {
            for (var i = 0; i < labelObj.LocalizedLabels.length; i++) {
                if (labelObj.LocalizedLabels[i].LanguageCode === lcid) {
                    return labelObj.LocalizedLabels[i].Label || "";
                }
            }
        }
        
        if (labelObj.UserLocalizedLabel && labelObj.UserLocalizedLabel.Label) {
            return labelObj.UserLocalizedLabel.Label;
        }
        
        return fallback || "";
    }

    // Helper functions for XML manipulation
    function getXmlAttr(el, names) {
        if (!el) return "";
        for (var i = 0; i < names.length; i++) {
            var v = el.getAttribute(names[i]);
            if (v != null && v !== "") return v;
        }
        return "";
    }

    function getDirectChildByTag(parent, tagName) {
        if (!parent) return null;
        var nodes = parent.childNodes;
        for (var i = 0; i < nodes.length; i++) {
            if (nodes[i].nodeType === 1 && String(nodes[i].tagName).toLowerCase() === String(tagName).toLowerCase()) {
                return nodes[i];
            }
        }
        return null;
    }

    function getDirectChildrenByTag(parent, tagName) {
        var result = [];
        if (!parent) return result;
        var nodes = parent.childNodes;
        for (var i = 0; i < nodes.length; i++) {
            if (nodes[i].nodeType === 1 && String(nodes[i].tagName).toLowerCase() === String(tagName).toLowerCase()) {
                result.push(nodes[i]);
            }
        }
        return result;
    }

    function getLocalizedDescriptionFromContainer(container, lcid, useFallback) {
        var labelsEl = getDirectChildByTag(container, "labels");
        if (!labelsEl) return "";
        var labels = getDirectChildrenByTag(labelsEl, "label");
        
        // Try to find exact match first
        var exactMatch = "";
        var foundExactMatch = false;
        for (var i = 0; i < labels.length; i++) {
            var code = parseInt(getXmlAttr(labels[i], ["languagecode", "languageCode"]), 10);
            if (code === lcid) {
                exactMatch = getXmlAttr(labels[i], ["description"]) || "";
                foundExactMatch = true;
                // If we have a non-empty match, return it immediately
                if (exactMatch && exactMatch.trim() !== "") {
                    return exactMatch;
                }
                // Otherwise continue to fallback if enabled
                break;
            }
        }
        
        // If we found an exact match but it was empty, or no exact match found, and useFallback is true
        if (useFallback && labels.length > 0) {
            // Try organization language first (skip if it's the same as lcid)
            if (state.orgLcid !== lcid) {
                for (var i = 0; i < labels.length; i++) {
                    var code = parseInt(getXmlAttr(labels[i], ["languagecode", "languageCode"]), 10);
                    if (code === state.orgLcid) {
                        var val = getXmlAttr(labels[i], ["description"]) || "";
                        if (val && val.trim() !== "") return val;
                    }
                }
            }
            
            // Try English (1033) as fallback (skip if it's the same as lcid or orgLcid)
            if (1033 !== lcid && 1033 !== state.orgLcid) {
                for (var i = 0; i < labels.length; i++) {
                    var code = parseInt(getXmlAttr(labels[i], ["languagecode", "languageCode"]), 10);
                    if (code === 1033) {
                        var val = getXmlAttr(labels[i], ["description"]) || "";
                        if (val && val.trim() !== "") return val;
                    }
                }
            }
            
            // If still nothing, return first available non-empty label
            for (var i = 0; i < labels.length; i++) {
                var val = getXmlAttr(labels[i], ["description"]) || "";
                if (val && val.trim() !== "") return val;
            }
        }
        
        // Return the exact match (even if empty) if we found one, otherwise empty string
        return exactMatch;
    }

    function ensureLabelsElement(container, xmlDoc) {
        var labelsEl = getDirectChildByTag(container, "labels");
        if (labelsEl) return labelsEl;

        labelsEl = xmlDoc.createElement("labels");
        if (container.firstChild) {
            container.insertBefore(labelsEl, container.firstChild);
        } else {
            container.appendChild(labelsEl);
        }
        return labelsEl;
    }

    function setLocalizedDescriptionOnContainer(container, lcid, text, xmlDoc) {
        var normalized = String(text == null ? "" : text);
        var labelsEl = ensureLabelsElement(container, xmlDoc);
        var labels = getDirectChildrenByTag(labelsEl, "label");
        var targetLabel = null;

        for (var i = 0; i < labels.length; i++) {
            var code = parseInt(getXmlAttr(labels[i], ["languagecode", "languageCode"]), 10);
            if (code === lcid) {
                targetLabel = labels[i];
                break;
            }
        }

        if (!normalized.trim()) {
            if (targetLabel) {
                labelsEl.removeChild(targetLabel);
                return true;
            }
            return false;
        }

        if (!targetLabel) {
            targetLabel = xmlDoc.createElement("label");
            targetLabel.setAttribute("languagecode", String(lcid));
            labelsEl.appendChild(targetLabel);
        }

        var current = getXmlAttr(targetLabel, ["description"]) || "";
        if (current === normalized) return false;

        targetLabel.setAttribute("description", normalized);
        return true;
    }

    function getFieldControlFromCell(cell) {
        if (!cell) return null;
        var controls = cell.getElementsByTagName("control");
        for (var i = 0; i < controls.length; i++) {
            var dataField = getXmlAttr(controls[i], ["datafieldname", "dataFieldName"]);
            if (dataField) return controls[i];
        }
        return null;
    }

    function buildStableRowKey(item, scope) {
        return [
            scope || "",
            item.type || "",
            item.propertyName || "",
            item.labelObjectId || "",
            item.id || "",
            item.logicalName || "",
            item.elementType || ""
        ].join("|");
    }

    function ensureRowKeys(items, scope) {
        for (var i = 0; i < items.length; i++) {
            if (!items[i].rowKey) {
                items[i].rowKey = buildStableRowKey(items[i], scope);
            }
        }
        return items;
    }

    function buildFormLabelNodeIndex(xmlDoc) {
        var index = {};

        function add(node, fallbackType) {
            if (!node) return;
            var objectId = getXmlAttr(node, ["labelid", "labelId", "id", "Id"]);
            if (!objectId) return;
            index[objectId] = {
                node: node,
                type: fallbackType || ""
            };
        }

        var tabs = xmlDoc.getElementsByTagName("tab");
        for (var i = 0; i < tabs.length; i++) add(tabs[i], "tab");

        var sections = xmlDoc.getElementsByTagName("section");
        for (var j = 0; j < sections.length; j++) add(sections[j], "section");

        var cells = xmlDoc.getElementsByTagName("cell");
        for (var k = 0; k < cells.length; k++) {
            var ctrl = getFieldControlFromCell(cells[k]);
            if (ctrl) add(cells[k], "cell");
        }

        return index;
    }

    function fetchJson(url, method, body, extraHeaders) {
        var opts = {
            method: method || "GET",
            headers: {
                "Accept": "application/json",
                "Content-Type": "application/json",
                "OData-Version": "4.0",
                "OData-MaxVersion": "4.0"
            },
            credentials: "same-origin"
        };
        
        // Add extra headers if provided (e.g., MSCRM.MergeLabels)
        if (extraHeaders) {
            for (var key in extraHeaders) {
                opts.headers[key] = extraHeaders[key];
            }
        }
        
        if (body) {
            opts.body = JSON.stringify(body);
        }
        
        return fetch(url, opts).then(function (r) {
            if (!r.ok) {
                return r.text().then(function (txt) {
                    var msg = txt || ("HTTP " + r.status);
                    try {
                        var obj = JSON.parse(txt);
                        if (obj && obj.error && obj.error.message) msg = obj.error.message;
                    } catch (e) {}
                    throw new Error(msg);
                });
            }
            
            // Handle 204 No Content or empty responses
            if (r.status === 204 || r.headers.get("content-length") === "0") {
                return {};
            }
            
            return r.json();
        });
    }

    function loadSolutions() {
        var base = Xrm.Utility.getGlobalContext().getClientUrl() + "/api/data/v9.2/";
        var url = base + "solutions?$select=solutionid,uniquename,friendlyname&$filter=ismanaged eq false and isvisible eq true&$orderby=friendlyname asc";
        
        state.solutionTracking.loading = true;
        
        return fetchJson(url).then(function (result) {
            var solutions = [];
            if (result && result.value) {
                for (var i = 0; i < result.value.length; i++) {
                    var sol = result.value[i];
                    solutions.push({
                        id: sol.solutionid,
                        uniqueName: sol.uniquename,
                        friendlyName: sol.friendlyname || sol.uniquename
                    });
                }
            }
            state.solutionTracking.solutions = solutions;
            state.solutionTracking.loading = false;
            console.log("Loaded", solutions.length, "unmanaged solutions");
            return solutions;
        }).catch(function (err) {
            console.warn("Failed to load solutions:", err);
            state.solutionTracking.solutions = [];
            state.solutionTracking.loading = false;
            return [];
        });
    }

    function loadLanguages() {
        var base = Xrm.Utility.getGlobalContext().getClientUrl() + "/api/data/v9.2/";
        
        // Use the RetrieveProvisionedLanguages unbound function to get enabled languages
        var provisionedUrl = base + "RetrieveProvisionedLanguages";
        
        return fetchJson(provisionedUrl).then(function (provisionedResult) {
            var enabledLcids = [];
            var enabledMap = {};
            
            // The response property is RetrieveProvisionedLanguages (array of LCIDs)
            if (provisionedResult && provisionedResult.RetrieveProvisionedLanguages) {
                enabledLcids = provisionedResult.RetrieveProvisionedLanguages;
                console.log("Provisioned languages (LCIDs):", enabledLcids);
                for (var i = 0; i < enabledLcids.length; i++) {
                    enabledMap[enabledLcids[i]] = true;
                }
            }
            
            var hasEnabledLanguages = enabledLcids.length > 0;
            
            // Get all language locales
            var localeUrl = base + "languagelocale?$select=localeid,name,language&$filter=statecode eq 0";
            return fetchJson(localeUrl).then(function (localeResult) {
                var allLangs = [];
                var provisionedLangs = [];
                
                if (localeResult && localeResult.value) {
                    for (var i = 0; i < localeResult.value.length; i++) {
                        var l = localeResult.value[i];
                        var isProvisioned = hasEnabledLanguages ? !!enabledMap[l.localeid] : false;
                        var lang = {
                            lcid: l.localeid,
                            name: l.name || l.language || String(l.localeid),
                            code: l.language || "",
                            provisioned: isProvisioned
                        };
                        allLangs.push(lang);
                        if (lang.provisioned) {
                            provisionedLangs.push(lang);
                        }
                    }
                }
                
                allLangs.sort(function (a, b) {
                    return String(a.name).localeCompare(String(b.name));
                });
                provisionedLangs.sort(function (a, b) {
                    return String(a.name).localeCompare(String(b.name));
                });
                
                console.log("Total languages:", allLangs.length);
                console.log("Provisioned languages:", provisionedLangs.length);
                console.log("Provisioned language details:", provisionedLangs.map(function(l) { return l.name + " (" + l.lcid + ")"; }));
                
                state.allLanguages = allLangs;
                // Only show provisioned languages by default
                if (hasEnabledLanguages && provisionedLangs.length > 0) {
                    state.languages = provisionedLangs;
                    state.showAllLanguages = false;
                } else {
                    // If no enabled languages detected, show all and warn user
                    state.languages = allLangs;
                    state.showAllLanguages = true;
                    console.warn("No provisioned languages detected, showing all languages");
                }
            });
        }).catch(function (err) {
            // Fallback if RetrieveProvisionedLanguages fails
            console.error("Failed to retrieve provisioned languages:", err);
            alert("Warning: Could not detect installed languages. Showing all languages.");
            
            var localeUrl = base + "languagelocale?$select=localeid,name,language&$filter=statecode eq 0";
            return fetchJson(localeUrl).then(function (localeResult) {
                var allLangs = [];
                
                if (localeResult && localeResult.value) {
                    for (var i = 0; i < localeResult.value.length; i++) {
                        var l = localeResult.value[i];
                        var lang = {
                            lcid: l.localeid,
                            name: l.name || l.language || String(l.localeid),
                            code: l.language || "",
                            provisioned: false
                        };
                        allLangs.push(lang);
                    }
                }
                
                allLangs.sort(function (a, b) {
                    return String(a.name).localeCompare(String(b.name));
                });
                
                state.allLanguages = allLangs;
                state.languages = allLangs;
                state.showAllLanguages = true;
            });
        });
    }

    function loadTranslations() {
        // Check if we have cached items for this scope
        var cacheKey = state.scope + "_" + state.targetLcid;
        if (state.itemsCache[cacheKey]) {
            state.items = state.itemsCache[cacheKey];
            
            // Restore pending changes if any  
            if (state.pendingChanges[cacheKey]) {
                for (var i = 0; i < state.items.length; i++) {
                    var rowKey = state.items[i].rowKey;
                    if (state.pendingChanges[cacheKey][rowKey]) {
                        state.items[i].newLabel = state.pendingChanges[cacheKey][rowKey].newLabel;
                        state.items[i].newDescription = state.pendingChanges[cacheKey][rowKey].newDescription;
                    }
                }
            }
            
            setStatus("Loaded from cache");
            return Promise.resolve();
        }
        
        setStatus("Loading translations...");
        
        var promise;
        if (state.scope === "fields") {
            promise = loadFieldTranslations().then(cacheItems);
        } else if (state.scope === "forms") {
            promise = loadFormTranslations().then(cacheItems);
        } else if (state.scope === "views") {
            promise = loadViewTranslations().then(cacheItems);
        } else if (state.scope === "entity") {
            promise = loadEntityTranslations().then(cacheItems);
        } else if (state.scope === "formlabels") {
            promise = loadFormLabelTranslations().then(cacheItems);
        } else {
            promise = Promise.resolve();
        }
        
        return promise.then(function() {
            setStatus("Loaded " + state.items.length + " item(s)");
        });
    }
    
    function cacheItems() {
        var cacheKey = state.scope + "_" + state.targetLcid;
        ensureRowKeys(state.items, state.scope);
        state.itemsCache[cacheKey] = state.items.slice(0);
    }

    function loadFieldTranslations() {
        var base = Xrm.Utility.getGlobalContext().getClientUrl() + "/api/data/v9.2/";
        var key = "EntityDefinitions(LogicalName=%27" + encodeURIComponent(state.entityLogicalName) + "%27)";
        // Fetch IsCustomizable and IsRenameable to detect locked fields
        var url = base + key + "/Attributes?$select=LogicalName,DisplayName,Description,MetadataId,IsCustomizable,IsRenameable";
        
        return fetchJson(url).then(function (result) {
            var items = [];
            if (result && result.value) {
                for (var i = 0; i < result.value.length; i++) {
                    var attr = result.value[i];
                    if (!attr.LogicalName || !attr.MetadataId) continue;
                    
                    var baseLabel = labelText(attr.DisplayName, state.orgLcid, attr.LogicalName);
                    var userLabel = labelText(attr.DisplayName, state.userLcid, attr.LogicalName);
                    var targetLabel = labelText(attr.DisplayName, state.targetLcid, "");
                    var userDesc = labelText(attr.Description, state.userLcid, "");
                    var targetDesc = labelText(attr.Description, state.targetLcid, "");
                    
                    // Check if field is customizable
                    var isCustomizable = attr.IsCustomizable && attr.IsCustomizable.Value !== false;
                    var isRenameable = attr.IsRenameable && attr.IsRenameable.Value !== false;
                    var canModify = isCustomizable && isRenameable;
                    
                    items.push({
                        id: attr.MetadataId,
                        logicalName: attr.LogicalName,
                        type: "field",
                        baseLabel: baseLabel,
                        userLabel: userLabel,
                        targetLabel: targetLabel,
                        userDescription: userDesc,
                        targetDescription: targetDesc,
                        newLabel: targetLabel,
                        newDescription: targetDesc,
                        isCustomizable: isCustomizable,
                        isRenameable: isRenameable,
                        canModify: canModify
                    });
                }
            }
            
            items.sort(function (a, b) {
                return String(a.userLabel || "").localeCompare(String(b.userLabel || ""));
            });
            
            state.items = items;
        });
    }

    function retrieveLocLabel(entityType, entityId, attributeName, lcid) {
        var base = Xrm.Utility.getGlobalContext().getClientUrl() + "/api/data/v9.2/";
        var id = String(entityId).replace(/[{}]/g, "").toLowerCase();
        
        // Use @odata.id format as documented by Microsoft
        var entitySet = entityType === "savedquery" ? "savedqueries" : 
                       entityType === "systemform" ? "systemforms" : entityType + "s";
        var entityMonikerParam = encodeURIComponent(JSON.stringify({
            "@odata.id": entitySet + "(" + id + ")"
        }));
        
        var url = base + "RetrieveLocLabels(EntityMoniker=@p1,AttributeName=@p2,IncludeUnpublished=@p3)"
            + "?@p1=" + entityMonikerParam
            + "&@p2='" + attributeName + "'"
            + "&@p3=true";  // Include unpublished to get latest changes
        
        return fetchJson(url).then(function(result) {
            if (result && result.Label && result.Label.LocalizedLabels) {
                for (var i = 0; i < result.Label.LocalizedLabels.length; i++) {
                    if (result.Label.LocalizedLabels[i].LanguageCode === lcid) {
                        return result.Label.LocalizedLabels[i].Label || "";
                    }
                }
            }
            return "";
        }).catch(function(err) {
            console.warn("Failed to retrieve '" + attributeName + "' translation:", err);
            return "";
        });
    }

    function loadFormTranslations() {
        var base = Xrm.Utility.getGlobalContext().getClientUrl() + "/api/data/v9.2/";
        var url = base + "systemforms?$select=formid,name,description,objecttypecode&$filter=objecttypecode eq '" + state.entityLogicalName + "' and type eq 2";
        
        return fetchJson(url).then(function (result) {
            var items = [];
            if (result && result.value) {
                for (var i = 0; i < result.value.length; i++) {
                    var form = result.value[i];
                    if (!form.formid) continue;
                    
                    items.push({
                        id: form.formid,
                        logicalName: "",
                        type: "form",
                        userLabel: form.name || "",
                        targetLabel: "",
                        userDescription: form.description || "",
                        targetDescription: "",
                        newLabel: "",
                        newDescription: ""
                    });
                }
            }
            
            // Load translations for each form
            var chain = Promise.resolve();
            items.forEach(function(item) {
                chain = chain.then(function() {
                    return retrieveLocLabel("systemform", item.id, "name", state.targetLcid).then(function(nameLabel) {
                        item.targetLabel = nameLabel;
                        item.newLabel = nameLabel;
                        
                        return retrieveLocLabel("systemform", item.id, "description", state.targetLcid).then(function(descLabel) {
                            item.targetDescription = descLabel;
                            item.newDescription = descLabel;
                        });
                    });
                });
            });
            
            return chain.then(function() {
                state.items = items;
            });
        });
    }

    function loadViewTranslations() {
        var base = Xrm.Utility.getGlobalContext().getClientUrl() + "/api/data/v9.2/";
        // Only load system views (querytype = 0 means system view)
        // Exclude personal views (querytype = 1) as they don't support localization
        var url = base + "savedqueries?$select=savedqueryid,name,description,returnedtypecode,querytype&$filter=returnedtypecode eq '" + state.entityLogicalName + "' and querytype eq 0";
        
        return fetchJson(url).then(function (result) {
            var items = [];
            if (result && result.value) {
                for (var i = 0; i < result.value.length; i++) {
                    var view = result.value[i];
                    if (!view.savedqueryid) continue;
                    
                    items.push({
                        id: view.savedqueryid,
                        logicalName: "",
                        type: "view",
                        userLabel: view.name || "",
                        targetLabel: "",
                        userDescription: view.description || "",
                        targetDescription: "",
                        newLabel: "",
                        newDescription: ""
                    });
                }
            }
            
            // Load translations for each view in parallel batches for efficiency
            var promises = [];
            items.forEach(function(item) {
                var namePromise = retrieveLocLabel("savedquery", item.id, "name", state.targetLcid).then(function(nameLabel) {
                    item.targetLabel = nameLabel;
                    item.newLabel = nameLabel;
                });
                var descPromise = retrieveLocLabel("savedquery", item.id, "description", state.targetLcid).then(function(descLabel) {
                    item.targetDescription = descLabel;
                    item.newDescription = descLabel;
                });
                promises.push(Promise.all([namePromise, descPromise]));
            });
            
            return Promise.all(promises).then(function() {
                items.sort(function (a, b) {
                    return String(a.userLabel || "").localeCompare(String(b.userLabel || ""));
                });
                state.items = items;
            });
        });
    }

    function loadEntityTranslations() {
        var base = Xrm.Utility.getGlobalContext().getClientUrl() + "/api/data/v9.2/";
        var key = "EntityDefinitions(LogicalName=%27" + encodeURIComponent(state.entityLogicalName) + "%27)";
        var url = base + key + "?$select=LogicalName,DisplayName,DisplayCollectionName,Description,MetadataId";
        
        return fetchJson(url).then(function (entity) {
            var items = [];
            
            if (entity && entity.MetadataId) {
                var userLabel = labelText(entity.DisplayName, state.userLcid, state.entityLogicalName);
                var targetLabel = labelText(entity.DisplayName, state.targetLcid, "");
                var userPluralLabel = labelText(entity.DisplayCollectionName, state.userLcid, "");
                var targetPluralLabel = labelText(entity.DisplayCollectionName, state.targetLcid, "");
                var userDesc = labelText(entity.Description, state.userLcid, "");
                var targetDesc = labelText(entity.Description, state.targetLcid, "");
                
                items.push({
                    id: entity.MetadataId,
                    logicalName: entity.LogicalName,
                    type: "entity-name",
                    userLabel: userLabel,
                    targetLabel: targetLabel,
                    userDescription: userDesc,
                    targetDescription: targetDesc,
                    newLabel: targetLabel,
                    newDescription: targetDesc,
                    propertyName: "DisplayName"
                });
                
                items.push({
                    id: entity.MetadataId,
                    logicalName: entity.LogicalName,
                    type: "entity-plural",
                    userLabel: userPluralLabel,
                    targetLabel: targetPluralLabel,
                    userDescription: "",
                    targetDescription: "",
                    newLabel: targetPluralLabel,
                    newDescription: "",
                    propertyName: "DisplayCollectionName"
                });
                
                items.push({
                    id: entity.MetadataId,
                    logicalName: entity.LogicalName,
                    type: "entity-description",
                    userLabel: userDesc,
                    targetLabel: targetDesc,
                    userDescription: "",
                    targetDescription: "",
                    newLabel: targetDesc,
                    newDescription: "",
                    propertyName: "Description"
                });
            }
            
            state.items = items;
        });
    }

    function loadFormLabelTranslations() {
        if (!state.formId) {
            alert("No form selected. Please open a form first.");
            state.items = [];
            return Promise.resolve();
        }
        
        var base = Xrm.Utility.getGlobalContext().getClientUrl() + "/api/data/v9.2/";
        
        // First, fetch ALL field metadata for fallback when FormXML labels are missing
        var metadataUrl = base + "EntityDefinitions(LogicalName='" + encodeURIComponent(state.entityLogicalName) + "')/Attributes?$select=LogicalName,DisplayName";
        return fetchJson(metadataUrl).then(function(metadataResult) {
            // Build a map of field metadata by logical name
            var fieldMetadata = {};
            if (metadataResult && metadataResult.value) {
                for (var m = 0; m < metadataResult.value.length; m++) {
                    var attr = metadataResult.value[m];
                    if (attr.LogicalName) {
                        fieldMetadata[attr.LogicalName.toLowerCase()] = attr.DisplayName;
                    }
                }
            }
            console.log("Loaded field metadata for", Object.keys(fieldMetadata).length, "attributes as fallback");
            
            return fieldMetadata;
        }).catch(function(err) {
            console.warn("Failed to load field metadata, continuing without fallback:", err);
            return {}; // Continue with empty metadata
        }).then(function(fieldMetadata) {
            // Now fetch the FormXML
            var url = base + "systemforms(" + state.formId + ")?$select=formxml";
            
            return fetchJson(url).then(function (result) {
                var items = [];
                
                if (!result || !result.formxml) {
                    state.items = items;
                    return;
                }
                
                try {
                    var parser = new DOMParser();
                    var xmlDoc = parser.parseFromString(result.formxml, "text/xml");
                    
                    // Check for parse errors
                    var parseError = xmlDoc.querySelector("parsererror");
                    if (parseError) {
                        throw new Error("XML parse error: " + parseError.textContent);
                    }
                    
                    // Parse tabs - use labelid or id
                    var tabs = xmlDoc.getElementsByTagName("tab");
                    for (var i = 0; i < tabs.length; i++) {
                        var tab = tabs[i];
                        var labelObjectId = getXmlAttr(tab, ["labelid", "labelId", "id", "Id"]);
                        if (!labelObjectId) labelObjectId = "tab_" + i;
                        var elementId = getXmlAttr(tab, ["id", "Id"]) || labelObjectId;
                        
                        var userLabel = getLocalizedDescriptionFromContainer(tab, state.userLcid, false);
                        var baseLabel = getLocalizedDescriptionFromContainer(tab, state.orgLcid, false);
                        var targetLabel = getLocalizedDescriptionFromContainer(tab, state.targetLcid, false);
                        
                        items.push({
                            id: elementId,
                            labelObjectId: labelObjectId,
                            logicalName: labelObjectId,
                            type: "formlabel",
                            baseLabel: baseLabel || "(Tab " + (i + 1) + ")",
                            userLabel: userLabel || baseLabel || "(Tab " + (i + 1) + ")",
                            targetLabel: targetLabel,
                            userDescription: "",
                            targetDescription: "",
                            newLabel: targetLabel,
                            newDescription: "",
                            elementType: "tab"
                        });
                    }
                    
                    // Parse sections - use labelid or id
                    var sections = xmlDoc.getElementsByTagName("section");
                    for (var i = 0; i < sections.length; i++) {
                        var section = sections[i];
                        var labelObjectId = getXmlAttr(section, ["labelid", "labelId", "id", "Id", "name"]);
                        if (!labelObjectId) labelObjectId = "section_" + i;
                        var elementId = getXmlAttr(section, ["id", "Id", "name"]) || labelObjectId;
                        
                        var userLabel = getLocalizedDescriptionFromContainer(section, state.userLcid, false);
                        var baseLabel = getLocalizedDescriptionFromContainer(section, state.orgLcid, false);
                        var targetLabel = getLocalizedDescriptionFromContainer(section, state.targetLcid, false);
                        
                        items.push({
                            id: elementId,
                            labelObjectId: labelObjectId,
                            logicalName: labelObjectId,
                            type: "formlabel",
                            baseLabel: baseLabel || "(Section " + (i + 1) + ")",
                            userLabel: userLabel || baseLabel ||  "(Section " + (i + 1) + ")",
                            targetLabel: targetLabel,
                            userDescription: "",
                            targetDescription: "",
                            newLabel: targetLabel,
                            newDescription: "",
                            elementType: "section"
                        });
                    }
                    
                    // Parse CELL labels (field labels) with field metadata fallback
                    var cells = xmlDoc.getElementsByTagName("cell");
                    for (var i = 0; i < cells.length; i++) {
                        var cell = cells[i];
                        
                        // Find the control with datafieldname inside this cell
                        var control = getFieldControlFromCell(cell);
                        if (!control) continue;
                        
                        var dataField = getXmlAttr(control, ["datafieldname", "dataFieldName"]);
                        if (!dataField) continue;
                        
                        // Get the label object ID from the cell
                        var labelObjectId = getXmlAttr(cell, ["labelid", "labelId", "id", "Id"]);
                        if (!labelObjectId) continue;
                        
                        var elementId = getXmlAttr(cell, ["id", "Id"]) || labelObjectId;
                        
                        // Extract labels from the CELL
                        var userLabel = getLocalizedDescriptionFromContainer(cell, state.userLcid, false);
                        var baseLabel = getLocalizedDescriptionFromContainer(cell, state.orgLcid, false);
                        var targetLabel = getLocalizedDescriptionFromContainer(cell, state.targetLcid, false);
                        
                        // FALLBACK: If labels are missing from FormXML, get them from field metadata
                        var metadata = fieldMetadata[dataField.toLowerCase()];
                        var usedFallback = false;
                        if (metadata) {
                            if (!baseLabel || baseLabel.trim() === "") {
                                baseLabel = labelText(metadata, state.orgLcid, "");
                                if (baseLabel) console.log("Field metadata fallback for baseLabel:", dataField, "->", baseLabel);
                            }
                            if (!targetLabel || targetLabel.trim() === "") {
                                targetLabel = labelText(metadata, state.targetLcid, "");
                                if (targetLabel) {
                                    console.log("Field metadata fallback for targetLabel:", dataField, "->", targetLabel);
                                    usedFallback = true; // Mark that target came from metadata
                                }
                            }
                            if (!userLabel || userLabel.trim() === "") {
                                userLabel = labelText(metadata, state.userLcid, dataField);
                            }
                        }
                        
                        items.push({
                            id: elementId,
                            labelObjectId: labelObjectId,
                            logicalName: dataField,
                            type: "formlabel",
                            baseLabel: baseLabel || dataField,
                            userLabel: userLabel || baseLabel || dataField,
                            targetLabel: targetLabel,
                            userDescription: "",
                            targetDescription: "",
                            newLabel: targetLabel,
                            newDescription: "",
                            elementType: "cell",
                            usedMetadataFallback: usedFallback // Track if this came from metadata
                        });
                    }
                    
                    items.sort(function (a, b) {
                        return String(a.userLabel || "").localeCompare(String(b.userLabel || ""));
                    });
                    
                    console.log("Loaded", items.length, "form labels with field metadata fallback");
                    
                } catch (e) {
                    console.error("Failed to parse FormXML:", e);
                    alert("Failed to parse form XML: " + e.message);
                }
                
                state.items = items;
            });
        });
    }

    function addStyles() {
        if (document.getElementById(STYLE_ID)) return;
        var css = "#" + ROOT_ID + "{position:fixed;top:64px;right:12px;width:700px;max-width:calc(100vw - 24px);height:calc(100vh - 76px);background:#ffffff;border:1px solid #d1d5db;border-radius:16px;box-shadow:0 20px 50px rgba(0,0,0,.2);z-index:2147483647;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Arial,sans-serif;display:flex;flex-direction:column;overflow:hidden;}#" + ROOT_ID + " *{box-sizing:border-box;}#" + ROOT_ID + " .hd{padding:20px 22px 22px;border-bottom:1px solid #e5e7eb;background:linear-gradient(180deg,#fafbfc 0%,#f5f7fa 100%);}#" + ROOT_ID + " .top{display:flex;align-items:flex-start;justify-content:space-between;gap:12px;margin-bottom:14px;}#" + ROOT_ID + " .title{font-size:18px;font-weight:700;line-height:1.3;color:#0d1421;letter-spacing:-.01em;}#" + ROOT_ID + " .sub{font-size:12px;color:#6b7280;margin-top:5px;line-height:1.5;}#" + ROOT_ID + " .close{border:1px solid #d1d5db;background:#fff;border-radius:8px;padding:7px 13px;font-size:12px;font-weight:600;cursor:pointer;appearance:none;-webkit-appearance:none;transition:all .15s ease;color:#374151;}#" + ROOT_ID + " .close:hover{background:#f9fafb;border-color:#9ca3af;}#" + ROOT_ID + " .search{width:100%;margin-bottom:16px;padding:10px 13px;border:1px solid #d1d5db;border-radius:10px;font-size:13px;outline:none;transition:border-color .15s ease;}#" + ROOT_ID + " .search:focus{border-color:#0f6cbd;box-shadow:0 0 0 3px rgba(15,108,189,.1);}#" + ROOT_ID + " .lu-collapse-toggle{display:flex;align-items:center;justify-content:space-between;cursor:pointer;padding:10px 0;border-bottom:1px solid #e5e7eb;margin-bottom:16px;user-select:none;}#" + ROOT_ID + " .lu-collapse-title{font-size:11px;font-weight:700;color:#374151;text-transform:uppercase;letter-spacing:.03em;}#" + ROOT_ID + " .lu-collapse-icon{font-size:16px;color:#6b7280;transition:transform .2s ease;line-height:1;}#" + ROOT_ID + " .lu-collapse-content{overflow:hidden;transition:max-height .3s ease,opacity .3s ease;}#" + ROOT_ID + " .lu-collapse-content.collapsed{max-height:0!important;opacity:0;pointer-events:none;}#" + ROOT_ID + " .lu-switch-row{display:flex;align-items:center;gap:12px;margin-bottom:18px;}#" + ROOT_ID + " .lu-switch-label{font-size:13px;color:#374151;font-weight:600;text-transform:uppercase;letter-spacing:.02em;font-size:11px;}#" + ROOT_ID + " .lu-segmented{display:inline-flex;align-items:stretch;gap:0;padding:3px;border:1px solid #d1dae6;background:#e8edf5;border-radius:10px;box-shadow:inset 0 1px 2px rgba(16,24,40,.06);}#" + ROOT_ID + " .lu-segment{appearance:none;-webkit-appearance:none;border:0 !important;outline:none;background:transparent;color:#4b5563;border-radius:7px;padding:8px 16px;font-size:12px;font-weight:600;line-height:1;cursor:pointer;transition:all .2s cubic-bezier(.4,0,.2,1);white-space:nowrap;position:relative;}#" + ROOT_ID + " .lu-segment:hover:not(.active){background:rgba(255,255,255,.5);}#" + ROOT_ID + " .lu-segment.active{background:#ffffff;color:#0f6cbd;box-shadow:0 2px 4px rgba(16,24,40,.1),0 1px 2px rgba(16,24,40,.06),inset 0 0 0 1px rgba(15,108,189,.1);}#" + ROOT_ID + " .filter-section{margin-bottom:18px;}#" + ROOT_ID + " .filter-section:last-child{margin-bottom:8px;}#" + ROOT_ID + " .filter-label{font-size:11px;color:#374151;font-weight:700;text-transform:uppercase;letter-spacing:.03em;margin-bottom:10px;}#" + ROOT_ID + " .lang-selector{width:100%;padding:10px 13px;border:1px solid #d1d5db;border-radius:10px;font-size:13px;outline:none;transition:border-color .15s ease;background:#fff;}#" + ROOT_ID + " .lang-selector:focus{border-color:#0f6cbd;box-shadow:0 0 0 3px rgba(15,108,189,.1);}#" + ROOT_ID + " .lang-selector:disabled{background:#f9fafb;color:#9ca3af;cursor:not-allowed;opacity:0.7;}#" + ROOT_ID + " .lang-info{margin-top:10px;padding:10px 13px;background:#f0f9ff;border:1px solid #bae6fd;border-radius:8px;font-size:11px;color:#0369a1;line-height:1.5;}#" + ROOT_ID + " .toolbar{display:flex;gap:8px;flex-wrap:wrap;padding:13px 18px;border-bottom:1px solid #e5e7eb;background:#fafbfc;}#" + ROOT_ID + " .btn{border:1px solid #d1d5db;background:#fff;border-radius:8px;padding:9px 15px;font-size:12px;font-weight:600;cursor:pointer;appearance:none;-webkit-appearance:none;transition:all .15s ease;color:#374151;}#" + ROOT_ID + " .btn:hover{background:#f3f4f6;}#" + ROOT_ID + " .btn.primary{background:#0f6cbd;color:#fff;border-color:#0f6cbd;}#" + ROOT_ID + " .btn.primary:hover{background:#0c5ba6;}#" + ROOT_ID + " .btn.primary:disabled{background:#e5e7eb;border-color:#d1d5db;color:#9ca3af;cursor:not-allowed;box-shadow:none;}#" + ROOT_ID + " .btn.primary:disabled:hover{background:#e5e7eb;}#" + ROOT_ID + " .status{padding:11px 18px;border-bottom:1px solid #e5e7eb;background:#f9fafb;font-size:12px;color:#374151;font-weight:500;}#" + ROOT_ID + " .list{flex:1;overflow:auto;background:#f5f7fa;padding:16px;}#" + ROOT_ID + " .item{background:#fff;border:1px solid #e5e7eb;border-radius:12px;padding:16px;margin-bottom:12px;box-shadow:0 1px 2px rgba(0,0,0,.04);transition:box-shadow .15s ease;}#" + ROOT_ID + " .item:hover{box-shadow:0 2px 8px rgba(0,0,0,.08);}#" + ROOT_ID + " .item-hd{margin-bottom:12px;}#" + ROOT_ID + " .lbl{font-size:14px;font-weight:600;color:#111827;line-height:1.4;margin-bottom:4px;}#" + ROOT_ID + " .meta{font-size:11px;color:#6b7280;margin-bottom:10px;word-break:break-word;font-family:Monaco,Consolas,monospace;background:#f9fafb;padding:3px 8px;border-radius:5px;display:inline-block;}#" + ROOT_ID + " .lu-chip-row{display:flex;gap:8px;flex-wrap:wrap;align-items:center;}#" + ROOT_ID + " .lu-chip{display:inline-flex;align-items:center;gap:6px;font-size:11px;font-weight:600;border-radius:6px;padding:5px 11px;white-space:nowrap;border:1px solid;line-height:1.3;text-transform:capitalize;}#" + ROOT_ID + " .lu-chip-dot{width:5px;height:5px;border-radius:999px;background:currentColor;display:inline-block;}#" + ROOT_ID + " .lu-chip-type{background:#e0e7ff;border-color:#c7d2fe;color:#3730a3;}#" + ROOT_ID + " .lu-chip-status.translated{background:#d1fae5;border-color:#6ee7b7;color:#065f46;}#" + ROOT_ID + " .lu-chip-status.missing{background:#fee2e2;border-color:#fecaca;color:#991b1b;}#" + ROOT_ID + " .trans-section{margin-top:12px;}#" + ROOT_ID + " .trans-label{font-size:11px;font-weight:700;color:#374151;text-transform:uppercase;letter-spacing:.02em;margin-bottom:6px;}#" + ROOT_ID + " .trans-text{padding:10px 12px;background:#f9fafb;border:1px solid #e5e7eb;border-radius:8px;font-size:13px;color:#111827;line-height:1.5;min-height:38px;margin-bottom:10px;}#" + ROOT_ID + " .input,#" + ROOT_ID + " .ta{width:100%;padding:10px 12px;border:1px solid #d1d5db;border-radius:8px;font-size:13px;background:#fff;color:#111827;transition:border-color .15s ease;}#" + ROOT_ID + " .input:focus,#" + ROOT_ID + " .ta:focus{outline:none;border-color:#0f6cbd;box-shadow:0 0 0 3px rgba(15,108,189,.1);}#" + ROOT_ID + " .ta{min-height:70px;resize:vertical;font-family:inherit;}#" + ROOT_ID + " .chk{display:flex;align-items:center;gap:8px;font-size:12px;color:#374151;padding:0;}#" + ROOT_ID + " .chk input{width:15px;height:15px;cursor:pointer;}#" + ROOT_ID + " .empty{text-align:center;color:#6b7280;padding:40px 20px;font-size:13px;}";
        var st = document.createElement("style");
        st.id = STYLE_ID;
        st.appendChild(document.createTextNode(css));
        document.head.appendChild(st);
    }

    function renderShell() {
        var host = document.createElement("div");
        host.id = ROOT_ID;
        host.innerHTML = "<div class=\"hd\"><div class=\"top\"><div><div class=\"title\">Translation Editor</div><div class=\"sub\" id=\"lu_subtitle\">" + esc(state.entityName) + "</div></div><button class=\"close\" type=\"button\" id=\"lu_close\">Close</button></div><input class=\"search\" id=\"lu_search\" type=\"text\" placeholder=\"Search translations...\"><div class=\"lu-collapse-toggle\" id=\"lu_collapse_toggle\"><span class=\"lu-collapse-title\">⚙️ FILTERS &amp; SETTINGS</span><span class=\"lu-collapse-icon\" id=\"lu_collapse_icon\">▼</span></div><div class=\"lu-collapse-content\" id=\"lu_collapse_content\"><div class=\"lu-switch-row\"><div class=\"lu-switch-label\">Translate</div><div class=\"lu-segmented\"><button class=\"lu-segment active\" data-scope=\"fields\" type=\"button\">Fields</button><button class=\"lu-segment\" data-scope=\"entity\" type=\"button\">Entity</button><button class=\"lu-segment\" data-scope=\"forms\" type=\"button\">Forms</button><button class=\"lu-segment\" data-scope=\"formlabels\" type=\"button\">Form Labels</button><button class=\"lu-segment\" data-scope=\"views\" type=\"button\">Views</button></div></div><div class=\"filter-section\"><div class=\"filter-label\">Target Language</div><select class=\"lang-selector\" id=\"lu_lang_select\"></select><label class=\"chk\" style=\"margin-top:10px;\"><input type=\"checkbox\" id=\"lu_show_all_langs\"> Show all languages (including not provisioned)</label><div class=\"lang-info\" id=\"lu_lang_info\">Loading languages...</div></div><div class=\"filter-section\"><div class=\"filter-label\">Display Options</div><label class=\"chk\"><input type=\"checkbox\" id=\"lu_show_locked\"> Show locked fields (non-customizable)</label></div><div class=\"filter-section\"><div class=\"filter-label\">Solution Tracking (ALM)</div><label class=\"chk\" style=\"margin-bottom:10px;\"><input type=\"checkbox\" id=\"lu_solution_tracking\"> Track changes in solution</label><select class=\"lang-selector\" id=\"lu_solution_select\" style=\"display:none;\"></select><div class=\"lang-info\" id=\"lu_solution_info\" style=\"display:none;\">Select a solution to automatically add modified components</div></div></div></div><div class=\"toolbar\"><button class=\"btn primary\" type=\"button\" id=\"lu_save\">Save Translations</button><button class=\"btn\" type=\"button\" id=\"lu_refresh\">Refresh</button></div><div class=\"status\" id=\"lu_status\">Loading...</div><div class=\"list\" id=\"lu_list\"></div>";
        document.body.appendChild(host);
        
        document.getElementById("lu_close").onclick = removePanel;
        document.getElementById("lu_search").oninput = function () {
            state.search = String(this.value || "").toLowerCase();
            renderItems();
        };
        
        var saveBtn = document.getElementById("lu_save");
        saveBtn.onclick = saveTranslations;
        saveBtn.disabled = true;
        
        document.getElementById("lu_refresh").onclick = function () {
            if (state.loading) return;
            state.loading = true;
            setStatus("Refreshing (clearing cache)...");
            
            // Clear ALL caches to force complete reload
            state.itemsCache = {};
            state.pendingChanges = {};
            
            loadTranslations().then(function () {
                state.loading = false;
                renderItems();
                setStatus("Refreshed - fetched latest from server (IncludeUnpublished=true)");
            }, function (err) {
                state.loading = false;
                setStatus("Refresh failed: " + getErrorMessage(err));
            });
        };
        
        var langSelect = document.getElementById("lu_lang_select");
        langSelect.onchange = function () {
            state.targetLcid = parseInt(this.value, 10);
            updateLanguageDisplay();
            loadTranslations().then(function () {
                renderItems();
                setStatus("Language changed");
            });
        };
        
        var showAllLangsCheck = document.getElementById("lu_show_all_langs");
        if (showAllLangsCheck) {
            showAllLangsCheck.onchange = function () {
                state.showAllLanguages = this.checked;
                state.languages = state.showAllLanguages ? state.allLanguages : state.allLanguages.filter(function (l) { return l.provisioned; });
                updateLanguageDisplay();
            };
        }
        
        var showLockedCheck = document.getElementById("lu_show_locked");
        if (showLockedCheck) {
            showLockedCheck.onchange = function () {
                state.showLockedFields = this.checked;
                renderItems(); // Re-render with new filter
            };
        }
        
        var solutionTrackingCheck = document.getElementById("lu_solution_tracking");
        if (solutionTrackingCheck) {
            solutionTrackingCheck.onchange = function () {
                console.log("Solution tracking checkbox changed:", this.checked);
                state.solutionTracking.enabled = this.checked;
                
                // Ensure the collapse section is expanded when enabling solution tracking
                if (this.checked) {
                    var collapseContent = document.getElementById("lu_collapse_content");
                    var collapseIcon = document.getElementById("lu_collapse_icon");
                    if (collapseContent && collapseContent.classList.contains("collapsed")) {
                        console.log("Expanding collapsed section to show solution dropdown");
                        collapseContent.classList.remove("collapsed");
                        collapseContent.style.maxHeight = collapseContent.scrollHeight + "px";
                        if (collapseIcon) collapseIcon.textContent = "▼";
                    }
                }
                
                console.log("Calling updateSolutionDisplay...");
                updateSolutionDisplay();
                checkForChanges(); // Re-validate save button based on solution selection
            };
        }
        
        var solutionSelect = document.getElementById("lu_solution_select");
        if (solutionSelect) {
            solutionSelect.onchange = function () {
                if (this.value === "") {
                    // Deselected - clear solution tracking
                    state.solutionTracking.solutionId = null;
                    state.solutionTracking.solutionName = null;
                    updateSolutionDisplay();
                    checkForChanges(); // Re-validate save button
                } else {
                    var selectedSolution = null;
                    for (var i = 0; i < state.solutionTracking.solutions.length; i++) {
                        if (state.solutionTracking.solutions[i].uniqueName === this.value) {
                            selectedSolution = state.solutionTracking.solutions[i];
                            break;
                        }
                    }
                    if (selectedSolution) {
                        state.solutionTracking.solutionId = selectedSolution.id;
                        state.solutionTracking.solutionName = selectedSolution.uniqueName;
                        updateSolutionDisplay();
                        checkForChanges(); // Re-validate save button
                    }
                }
            };
        }
        
        // Collapse/expand toggle
        var collapseToggle = document.getElementById("lu_collapse_toggle");
        var collapseContent = document.getElementById("lu_collapse_content");
        var collapseIcon = document.getElementById("lu_collapse_icon");
        
        if (collapseToggle && collapseContent && collapseIcon) {
            // Set initial max-height for smooth animation
            collapseContent.style.maxHeight = collapseContent.scrollHeight + "px";
            
            collapseToggle.onclick = function () {
                var isCollapsed = collapseContent.classList.contains("collapsed");
                
                if (isCollapsed) {
                    // Expand
                    collapseContent.classList.remove("collapsed");
                    collapseContent.style.maxHeight = collapseContent.scrollHeight + "px";
                    collapseIcon.textContent = "▼";
                } else {
                    // Collapse
                    collapseContent.style.maxHeight = collapseContent.scrollHeight + "px";
                    // Force reflow
                    collapseContent.offsetHeight;
                    collapseContent.classList.add("collapsed");
                    collapseIcon.textContent = "▶";
                }
            };
        }
        
        host.onclick = function (e) {
            var t = e.target || e.srcElement;
            if (!t) return;
            
            var scope = t.getAttribute("data-scope");
            if (scope) {
                state.scope = scope;
                var sbs = host.querySelectorAll(".lu-segment");
                for (var si = 0; si < sbs.length; si++) sbs[si].className = "lu-segment";
                t.className = "lu-segment active";
                loadTranslations().then(function () {
                    renderItems();
                    setStatus("Ready");
                });
                return;
            }
        };
    }

    function updateSolutionDisplay() {
        console.log("updateSolutionDisplay called");
        console.log("  enabled:", state.solutionTracking.enabled);
        console.log("  loading:", state.solutionTracking.loading);
        console.log("  solutions count:", state.solutionTracking.solutions.length);
        
        var solutionSelect = document.getElementById("lu_solution_select");
        var solutionInfo = document.getElementById("lu_solution_info");
        
        console.log("  solutionSelect found:", !!solutionSelect);
        console.log("  solutionInfo found:", !!solutionInfo);
        
        if (state.solutionTracking.enabled) {
            console.log("  Solution tracking ENABLED - showing dropdown");
            // Show dropdown and info
            if (solutionSelect) {
                solutionSelect.style.display = "block";
                console.log("  Dropdown display set to: block");
                
                if (state.solutionTracking.loading) {
                    // Loading state - disable dropdown and show loading message
                    solutionSelect.disabled = true;
                    solutionSelect.innerHTML = "<option value=\"\">⏳ Loading solutions...</option>";
                    console.log("  Dropdown in LOADING state");
                } else {
                    // Loaded state - enable dropdown and populate with solutions
                    solutionSelect.disabled = false;
                    solutionSelect.innerHTML = "<option value=\"\">-- Select Solution --</option>";
                    
                    for (var i = 0; i < state.solutionTracking.solutions.length; i++) {
                        var sol = state.solutionTracking.solutions[i];
                        var opt = document.createElement("option");
                        opt.value = sol.uniqueName;
                        opt.textContent = sol.friendlyName + " (" + sol.uniqueName + ")";
                        if (sol.uniqueName === state.solutionTracking.solutionName) {
                            opt.selected = true;
                        }
                        solutionSelect.appendChild(opt);
                    }
                    console.log("  Dropdown POPULATED with", state.solutionTracking.solutions.length, "solutions");
                }
            } else {
                console.warn("  ERROR: solutionSelect element not found!");
            }
            
            if (solutionInfo) {
                solutionInfo.style.display = "block";
                console.log("  Info box display set to: block");
                
                if (state.solutionTracking.loading) {
                    // Loading state - show info message
                    solutionInfo.textContent = "⏳ Loading available solutions...";
                    solutionInfo.style.background = "#e0e7ff";
                    solutionInfo.style.borderColor = "#c7d2fe";
                    solutionInfo.style.color = "#3730a3";
                } else if (state.solutionTracking.solutionName) {
                    // Solution selected - show success state
                    solutionInfo.textContent = "Modified components will be added to: " + state.solutionTracking.solutionName;
                    solutionInfo.style.background = "#d1fae5";
                    solutionInfo.style.borderColor = "#6ee7b7";
                    solutionInfo.style.color = "#065f46";
                } else {
                    // No solution selected - show warning state
                    solutionInfo.textContent = "⚠️ You must select a solution before saving";
                    solutionInfo.style.background = "#fef3c7";
                    solutionInfo.style.borderColor = "#fcd34d";
                    solutionInfo.style.color = "#92400e";
                }
            } else {
                console.warn("  ERROR: solutionInfo element not found!");
            }
            
            // Recalculate collapse section max-height after showing/hiding dropdown
            var collapseContent = document.getElementById("lu_collapse_content");
            if (collapseContent && !collapseContent.classList.contains("collapsed")) {
                console.log("  Recalculating collapse max-height after showing dropdown");
                collapseContent.style.maxHeight = collapseContent.scrollHeight + "px";
            }
        } else {
            console.log("  Solution tracking DISABLED - hiding dropdown");
            // Hide dropdown and info
            if (solutionSelect) {
                solutionSelect.style.display = "none";
                solutionSelect.disabled = false; // Reset disabled state
            }
            if (solutionInfo) solutionInfo.style.display = "none";
            
            // Recalculate collapse section max-height after hiding dropdown
            var collapseContent = document.getElementById("lu_collapse_content");
            if (collapseContent && !collapseContent.classList.contains("collapsed")) {
                console.log("  Recalculating collapse max-height after hiding dropdown");
                collapseContent.style.maxHeight = collapseContent.scrollHeight + "px";
            }
        }
    }

    function updateLanguageDisplay() {
        var langSelect = document.getElementById("lu_lang_select");
        if (langSelect) {
            langSelect.innerHTML = "";
            for (var i = 0; i < state.languages.length; i++) {
                var opt = document.createElement("option");
                opt.value = state.languages[i].lcid;
                opt.textContent = state.languages[i].name + " (" + state.languages[i].lcid + ")";
                if (state.languages[i].lcid === state.targetLcid) {
                    opt.selected = true;
                }
                langSelect.appendChild(opt);
            }
        }
        
        var langInfo = document.getElementById("lu_lang_info");
        if (langInfo) {
            var targetLang = null;
            for (var i = 0; i < state.allLanguages.length; i++) {
                if (state.allLanguages[i].lcid === state.targetLcid) {
                    targetLang = state.allLanguages[i];
                    break;
                }
            }
            
            if (targetLang) {
                var provisionedText = targetLang.provisioned ? "" : " [NOT PROVISIONED]";
                var provisionedCount = 0;
                for (var i = 0; i < state.allLanguages.length; i++) {
                    if (state.allLanguages[i].provisioned) provisionedCount++;
                }
                var langCountText = " | Showing: " + state.languages.length + " of " + state.allLanguages.length + " (" + provisionedCount + " enabled)";
                langInfo.textContent = "Translating to " + targetLang.name + provisionedText + langCountText + " | Your LCID: " + state.userLcid + " | Org LCID: " + state.orgLcid;
            } else {
                langInfo.textContent = "Select a target language";
            }
        }
        
        updateSubtitle();
    }

    function updateSubtitle() {
        var sub = document.getElementById("lu_subtitle");
        if (sub) {
            var scopeText = state.scope === "fields" ? "Fields" : 
                           state.scope === "entity" ? "Entity" :
                           state.scope === "forms" ? "Forms" : 
                           state.scope === "formlabels" ? "Form Labels" : "Views";
            
            // Count locked items for fields scope
            var totalItems = state.items.length;
            var lockedCount = 0;
            if (state.scope === "fields") {
                for (var i = 0; i < state.items.length; i++) {
                    if (state.items[i].canModify === false) {
                        lockedCount++;
                    }
                }
            }
            
            var itemText = "Items: " + totalItems;
            if (lockedCount > 0 && !state.showLockedFields) {
                var visibleCount = totalItems - lockedCount;
                itemText = "Items: " + visibleCount + " visible (" + lockedCount + " locked hidden)";
            } else if (lockedCount > 0) {
                itemText = "Items: " + totalItems + " (" + lockedCount + " locked)";
            }
            
            sub.textContent = state.entityName + " | " + scopeText + " | " + itemText;
        }
    }

    function setStatus(text) {
        var el = document.getElementById("lu_status");
        if (el) el.textContent = text;
    }

    function matchItem(item) {
        // Filter out locked fields unless explicitly shown
        if (!state.showLockedFields && item.type === "field" && item.canModify === false) {
            return false;
        }
        
        if (!state.search) return true;
        var hay = (item.userLabel + " " + item.logicalName + " " + item.targetLabel + " " + item.type).toLowerCase();
        return hay.indexOf(state.search) >= 0;
    }

    function renderItem(item) {
        var html = [];
        html.push("<div class=\"item\">");
        html.push("<div class=\"item-hd\">");
        html.push("<div class=\"lbl\">" + esc(item.userLabel || item.logicalName || "Unnamed") + "</div>");
        
        if (item.logicalName) {
            html.push("<div class=\"meta\">" + esc(item.logicalName) + "</div>");
        }
        
        html.push("<div class=\"lu-chip-row\">");
        
        var typeLabel = item.type === "field" ? "Field" :
                       item.type === "form" ? "Form" :
                       item.type === "view" ? "View" :
                       item.type === "entity-name" ? "Entity Name" :
                       item.type === "entity-plural" ? "Plural Name" :
                       item.type === "entity-description" ? "Description" :
                       item.type === "formlabel" ? "Form Label" : item.type;
        html.push("<span class=\"lu-chip lu-chip-type\">" + esc(typeLabel) + "</span>");
        
        // Show locked indicator for non-customizable fields
        if (item.type === "field" && item.canModify === false) {
            html.push("<span class=\"lu-chip lu-chip-locked\" title=\"This field cannot be customized (IsCustomizable or IsRenameable is false)\">🔒 Locked</span>");
        }
        
        var hasTranslation = item.targetLabel && item.targetLabel.trim() !== "";
        var statusClass = hasTranslation ? "translated" : "missing";
        var statusText = hasTranslation ? "Translated" : "Missing";
        html.push("<span class=\"lu-chip lu-chip-status " + statusClass + "\"><span class=\"lu-chip-dot\"></span>" + statusText + "</span>");
        
        html.push("</div>");
        html.push("</div>");
        
        html.push("<div class=\"trans-section\">");
        
        // Show base language reference for form labels and fields
        if ((item.type === "formlabel" || item.type === "field") && item.baseLabel) {
            html.push("<div class=\"trans-label\">Reference (" + state.orgLcid + ") - Base Language</div>");
            html.push("<div class=\"trans-text\">" + esc(item.baseLabel || "(empty)") + "</div>");
        }
        
        html.push("<div class=\"trans-label\">Current (" + state.userLcid + ")</div>");
        html.push("<div class=\"trans-text\">" + esc(item.userLabel || "(empty)") + "</div>");
        
        html.push("<div class=\"trans-label\">Translation (" + state.targetLcid + ")</div>");
        var disabledAttr = (item.type === "field" && item.canModify === false) ? " disabled title=\"Cannot modify: IsCustomizable or IsRenameable is false\"" : "";
        html.push("<input class=\"input\" type=\"text\" data-id=\"" + esc(item.rowKey || item.id) + "\" data-field=\"label\" value=\"" + esc(item.newLabel || "") + "\" placeholder=\"Enter translation...\"" + disabledAttr + ">");
        html.push("</div>");
        
        // Always show description for views, forms, and items that have descriptions
        if (item.type === "view" || item.type === "form" || item.userDescription || item.targetDescription) {
            html.push("<div class=\"trans-section\" style=\"margin-top:16px;\">");
            html.push("<div class=\"trans-label\">Description (" + state.userLcid + ")</div>");
            html.push("<div class=\"trans-text\">" + esc(item.userDescription || "(empty)") + "</div>");
            
            html.push("<div class=\"trans-label\">Description Translation (" + state.targetLcid + ")</div>");
            var descDisabledAttr = (item.type === "field" && item.canModify === false) ? " disabled title=\"Cannot modify: IsCustomizable or IsRenameable is false\"" : "";
            html.push("<textarea class=\"ta\" data-id=\"" + esc(item.rowKey || item.id) + "\" data-field=\"description\" placeholder=\"Enter description translation...\"" + descDisabledAttr + ">" + esc(item.newDescription || "") + "</textarea>");
            html.push("</div>");
        }
        
        html.push("</div>");
        return html.join("");
    }

    function renderItems() {
        var list = document.getElementById("lu_list");
        if (!list) return;
        
        var html = [];
        var shown = 0;
        
        for (var i = 0; i < state.items.length; i++) {
            if (!matchItem(state.items[i])) continue;
            shown++;
            html.push(renderItem(state.items[i]));
        }
        
        if (!shown) {
            html.push("<div class=\"empty\">No items found for the current filter.</div>");
        }
        
        list.innerHTML = html.join("");
        updateSubtitle();
        attachChangeListeners();
    }

    function attachChangeListeners() {
        var inputs = document.querySelectorAll("#" + ROOT_ID + " .input, #" + ROOT_ID + " .ta");
        for (var i = 0; i < inputs.length; i++) {
            inputs[i].removeEventListener("input", onInputChange);
            inputs[i].addEventListener("input", onInputChange);
        }
    }

    function onInputChange(e) {
        var el = e.target || e.srcElement;
        
        // Ignore changes from disabled fields
        if (el.disabled) {
            return;
        }
        
        var rowKey = el.getAttribute("data-id");
        var field = el.getAttribute("data-field");
        var value = el.value;
        
        for (var i = 0; i < state.items.length; i++) {
            var itemKey = state.items[i].rowKey || String(state.items[i].id);
            if (itemKey === rowKey) {
                // Double-check: don't allow changes to non-customizable fields
                if (state.items[i].canModify === false) {
                    console.warn("Attempted to modify locked field:", state.items[i].logicalName);
                    return;
                }
                
                if (field === "label") {
                    state.items[i].newLabel = value;
                } else if (field === "description") {
                    state.items[i].newDescription = value;
                }
                
                // Save pending changes
                var cacheKey = state.scope + "_" + state.targetLcid;
                if (!state.pendingChanges[cacheKey]) {
                    state.pendingChanges[cacheKey] = {};
                }
                state.pendingChanges[cacheKey][rowKey] = {
                    newLabel: state.items[i].newLabel,
                    newDescription: state.items[i].newDescription
                };
                
                break;
            }
        }
        
        checkForChanges();
    }

    function checkForChanges() {
        var changed = false;
        
        // Check all scopes for changes
        for (var key in state.pendingChanges) {
            if (Object.keys(state.pendingChanges[key]).length > 0) {
                for (var itemId in state.pendingChanges[key]) {
                    changed = true;
                    break;
                }
            }
            if (changed) break;
        }
        
        state.hasChanges = changed;
        
        // Validate solution tracking: if enabled, a solution must be selected
        var canSave = changed;
        if (changed && state.solutionTracking.enabled && !state.solutionTracking.solutionName) {
            canSave = false;
        }
        
        var saveBtn = document.getElementById("lu_save");
        if (saveBtn) saveBtn.disabled = !canSave;
    }

    function getComponentType(scope, item) {
        // Component types from Microsoft.Crm.Sdk.SolutionComponentType
        // Entity = 1, Attribute = 2, SavedQuery = 26, SystemForm = 60
        if (scope === "fields") return 2; // Attribute
        if (scope === "entity") return 1; // Entity
        if (scope === "views") return 26; // SavedQuery
        if (scope === "forms" || scope === "formlabels") return 60; // SystemForm
        return null;
    }

    function isComponentInSolution(componentId, componentType, solutionName) {
        var base = Xrm.Utility.getGlobalContext().getClientUrl() + "/api/data/v9.2/";
        
        // Check cache first
        var cacheKey = solutionName + "_" + componentType + "_" + componentId;
        if (state.solutionTracking.componentsCache[cacheKey] !== undefined) {
            return Promise.resolve(state.solutionTracking.componentsCache[cacheKey]);
        }
        
        // Query solutioncomponent entity
        var filter = "objectid eq " + componentId + " and componenttype eq " + componentType;
        var url = base + "solutioncomponents?$select=solutioncomponentid&$filter=" + encodeURIComponent(filter) + "&$top=1";
        url += "&$expand=solutionid($select=uniquename)";
        
        return fetchJson(url).then(function(result) {
            var exists = false;
            if (result && result.value && result.value.length > 0) {
                for (var i = 0; i < result.value.length; i++) {
                    if (result.value[i].solutionid && result.value[i].solutionid.uniquename === solutionName) {
                        exists = true;
                        break;
                    }
                }
            }
            // Cache result
            state.solutionTracking.componentsCache[cacheKey] = exists;
            return exists;
        }).catch(function(err) {
            console.warn("Failed to check component in solution:", err);
            return false; // Assume not present on error
        });
    }

    function addToSolution(componentId, componentType, solutionName) {
        var base = Xrm.Utility.getGlobalContext().getClientUrl() + "/api/data/v9.2/";
        
        // Option A: Check first, then add
        return isComponentInSolution(componentId, componentType, solutionName).then(function(exists) {
            if (exists) {
                console.log("Component already in solution:", componentType, componentId);
                return { added: false, reason: "already_exists" };
            }
            
            // Add component to solution
            var payload = {
                ComponentId: componentId,
                ComponentType: componentType,
                SolutionUniqueName: solutionName,
                AddRequiredComponents: false
            };
            
            return fetchJson(base + "AddSolutionComponent", "POST", payload).then(function() {
                console.log("Added component to solution:", componentType, componentId);
                // Update cache
                var cacheKey = solutionName + "_" + componentType + "_" + componentId;
                state.solutionTracking.componentsCache[cacheKey] = true;
                return { added: true, reason: "success" };
            }).catch(function(err) {
                console.error("Failed to add component to solution:", err);
                return { added: false, reason: "error", error: err };
            });
        });
    }

    function saveTranslations() {
        if (state.loading) return;
        
        // Validate target language is provisioned
        var targetLang = null;
        for (var i = 0; i < state.allLanguages.length; i++) {
            if (state.allLanguages[i].lcid === state.targetLcid) {
                targetLang = state.allLanguages[i];
                break;
            }
        }
        
        if (!targetLang || !targetLang.provisioned) {
            alert("Warning: The selected target language (" + state.targetLcid + ") is not provisioned/enabled in this environment.\n\nSaving translations for non-provisioned languages may not work correctly.\n\nPlease enable the language in Settings > Languages first.");
            return;
        }
        
        // Validate solution tracking
        if (state.solutionTracking.enabled && !state.solutionTracking.solutionName) {
            alert("Please select a solution before saving.\n\nYou have enabled 'Track changes in solution' but haven't selected a solution yet.");
            return;
        }
        
        state.loading = true;
        
        console.log("=== SAVE TRANSLATIONS START ===");
        console.log("Current scope:", state.scope, "LCID:", state.targetLcid);
        console.log("Pending changes keys:", Object.keys(state.pendingChanges));
        console.log("ItemsCache keys:", Object.keys(state.itemsCache));
        
        Xrm.Utility.showProgressIndicator("Saving translations...");
        setStatus("Saving translations...");
        
        var saveBtn = document.getElementById("lu_save");
        if (saveBtn) saveBtn.disabled = true;
        
        // Collect all changes from all scopes
        var allPromises = [];
        var totalUpdates = 0;
        
        for (var cacheKey in state.pendingChanges) {
            console.log("Processing cacheKey:", cacheKey, "changes:", Object.keys(state.pendingChanges[cacheKey]).length);
            
            if (Object.keys(state.pendingChanges[cacheKey]).length === 0) continue;
            
            var parts = cacheKey.split("_");
            var scope = parts[0];
            var lcid = parseInt(parts[1], 10);
            
            console.log("  Scope:", scope, "LCID:", lcid);
            
            var cachedItems = state.itemsCache[cacheKey];
            console.log("  Cached items:", cachedItems ? cachedItems.length : "none");
            
            // If no cached items but this is the current scope/lcid, use state.items
            if (!cachedItems && scope === state.scope && lcid === state.targetLcid) {
                console.log("  Using state.items as fallback (", state.items.length, "items)");
                cachedItems = state.items;
            }
            
            if (!cachedItems) {
                console.log("  No items found, skipping");
                continue;
            }
            
            var updates = [];
            for (var i = 0; i < cachedItems.length; i++) {
                var item = cachedItems[i];
                var itemKey = item.rowKey || String(item.id);
                var pendingChange = state.pendingChanges[cacheKey][itemKey];
                if (pendingChange) {
                    item.newLabel = pendingChange.newLabel;
                    item.newDescription = pendingChange.newDescription;
                    
                    console.log("  Item", itemKey, "- newLabel:", item.newLabel, "targetLabel:", item.targetLabel);
                    
                    // For views/forms, targetLabel might be empty since we can't easily retrieve it
                    // So we check if newLabel has been filled in by the user (non-empty and different from original)
                    var hasLabelChange = item.newLabel && item.newLabel.trim() !== "" && item.newLabel !== item.targetLabel;
                    var hasDescChange = item.newDescription !== item.targetDescription;
                    
                    if (hasLabelChange || hasDescChange) {
                        console.log("    -> Added to updates");
                        updates.push(item);
                    } else {
                        console.log("    -> Skipped (no change or empty)");
                    }
                }
            }
            
            console.log("  Total updates for this scope:", updates.length);
            
            if (updates.length > 0) {
                totalUpdates += updates.length;
                allPromises.push({ scope: scope, updates: updates, lcid: lcid });
            }
        }
        
        console.log("Total jobs to execute:", allPromises.length);
        console.log("Total updates across all scopes:", totalUpdates);
        
        if (allPromises.length === 0) {
            console.log("No changes to save, aborting");
            state.loading = false;
            Xrm.Utility.closeProgressIndicator();
            setStatus("No changes to save");
            if (saveBtn) saveBtn.disabled = false;
            return;
        }
        
        // Execute saves sequentially to avoid race condition
        var chain = Promise.resolve();
        var solutionStats = { added: 0, skipped: 0, failed: 0 };
        
        for (var j = 0; j < allPromises.length; j++) {
            (function(job) {
                chain = chain.then(function() {
                    console.log("Executing save for scope:", job.scope, "LCID:", job.lcid, "updates:", job.updates.length);
                    return saveScope(job.scope, job.updates, job.lcid).then(function(result) {
                        // Aggregate solution stats
                        if (result.addResult) {
                            solutionStats.added += result.addResult.added;
                            solutionStats.skipped += result.addResult.skipped;
                            solutionStats.failed += result.addResult.failed;
                        }
                    });
                });
            })(allPromises[j]);
        }
        
        chain.then(function() {
            console.log("All saves completed successfully");
            console.log("Solution stats:", solutionStats);
            state.loading = false;
            Xrm.Utility.closeProgressIndicator();
            
            // Clear ALL caches before reloading to ensure fresh data
            state.pendingChanges = {};
            state.itemsCache = {};
            
            // Build status message
            var statusMsg = "Saved " + totalUpdates + " translation(s)";
            if (state.solutionTracking.enabled && state.solutionTracking.solutionName) {
                statusMsg += " | Solution: " + solutionStats.added + " added, " + solutionStats.skipped + " already present";
                if (solutionStats.failed > 0) {
                    statusMsg += ", " + solutionStats.failed + " failed";
                }
            }
            statusMsg += " - reloading...";
            
            setStatus(statusMsg);
            
            return loadTranslations().then(function () {
                renderItems();
                state.hasChanges = false;
                
                var finalMsg = "Saved successfully! Refresh CE (Ctrl+F5) to see changes there.";
                if (state.solutionTracking.enabled && state.solutionTracking.solutionName) {
                    finalMsg += " | Components added to solution: " + state.solutionTracking.solutionName;
                }
                setStatus(finalMsg);
                
                if (saveBtn) saveBtn.disabled = true;
                
                // Show alert with instructions
                var alertMsg = "Translations saved and published!\n\nNote: Dynamics 365 CE caches metadata heavily.\n\nTo see changes in CE:\n1. Press Ctrl+F5 (hard refresh)\n2. Or clear browser cache\n3. Or close and reopen CE\n\nChanges are immediately visible in Power Apps (powerapps.com).";
                if (state.solutionTracking.enabled && state.solutionTracking.solutionName) {
                    alertMsg += "\n\nSolution Tracking:\n- Added " + solutionStats.added + " component(s)\n- Skipped " + solutionStats.skipped + " (already present)";
                    if (solutionStats.failed > 0) {
                        alertMsg += "\n- Failed " + solutionStats.failed;
                    }
                }
                alert(alertMsg);
            });
        }).catch(function(err) {
            state.loading = false;
            Xrm.Utility.closeProgressIndicator();
            setStatus("Save failed: " + getErrorMessage(err));
            alert("Save failed: " + getErrorMessage(err));
            if (saveBtn) saveBtn.disabled = false;
        });
    }
    
    function saveScope(scope, updates, lcid) {
        console.log("saveScope called - scope:", scope, "lcid:", lcid, "updates:", updates.length);
        
        var promise;
        if (scope === "fields") {
            promise = saveFieldTranslations(updates, lcid);
        } else if (scope === "entity") {
            promise = saveEntityTranslations(updates, lcid);
        } else if (scope === "forms") {
            promise = saveFormTranslations(updates, lcid);
        } else if (scope === "views") {
            promise = saveViewTranslations(updates, lcid);
        } else if (scope === "formlabels") {
            promise = saveFormLabelTranslations(updates, lcid);
        } else {
            promise = Promise.resolve();
        }
        
        return promise.then(function(result) {
            // Add components to solution if tracking is enabled
            if (state.solutionTracking.enabled && state.solutionTracking.solutionName) {
                console.log("Solution tracking enabled, adding components to:", state.solutionTracking.solutionName);
                return addComponentsToSolution(scope, updates).then(function(addResult) {
                    return { saveResult: result, addResult: addResult };
                });
            }
            return { saveResult: result, addResult: null };
        }).catch(function(err) {
            console.log("saveScope error for scope:", scope, "lcid:", lcid, "error:", err);
            throw err;
        });
    }
    
    function addComponentsToSolution(scope, updates) {
        if (!state.solutionTracking.enabled || !state.solutionTracking.solutionName) {
            return Promise.resolve({ added: 0, skipped: 0, failed: 0 });
        }
        
        var componentType = getComponentType(scope, null);
        if (!componentType) {
            console.warn("Unknown component type for scope:", scope);
            return Promise.resolve({ added: 0, skipped: 0, failed: 0 });
        }
        
        var stats = { added: 0, skipped: 0, failed: 0 };
        var chain = Promise.resolve();
        
        // Collect unique component IDs
        var componentIds = {};
        for (var i = 0; i < updates.length; i++) {
            var item = updates[i];
            var compId;
            
            if (scope === "fields" || scope === "entity") {
                // Use MetadataId for attributes and entities
                compId = item.id;
            } else if (scope === "views") {
                // Use savedqueryid for views
                compId = String(item.id).replace(/[{}]/g, "").toLowerCase();
            } else if (scope === "forms" || scope === "formlabels") {
                // Use formid for forms
                compId = state.formId ? String(state.formId).replace(/[{}]/g, "").toLowerCase() : String(item.id).replace(/[{}]/g, "").toLowerCase();
            }
            
            if (compId && !componentIds[compId]) {
                componentIds[compId] = true;
            }
        }
        
        // Add each unique component to solution  
        var compIdList = Object.keys(componentIds);
        console.log("Adding", compIdList.length, "unique component(s) to solution");
        
        for (var j = 0; j < compIdList.length; j++) {
            (function(componentId) {
                chain = chain.then(function() {
                    return addToSolution(componentId, componentType, state.solutionTracking.solutionName).then(function(result) {
                        if (result.added) {
                            stats.added++;
                        } else if (result.reason === "already_exists") {
                            stats.skipped++;
                        } else {
                            stats.failed++;
                        }
                    });
                });
            })(compIdList[j]);
        }
        
        return chain.then(function() {
            console.log("Solution add stats:", stats);
            return stats;
        });
    }

    // Helper to merge localized labels without losing other languages
    function mergeLocalizedLabels(existingLabels, lcid, newText) {
        var result = [];
        var normalized = String(newText == null ? "" : newText);
        
        // Preserve all existing labels except the one we're updating
        if (existingLabels) {
            for (var i = 0; i < existingLabels.length; i++) {
                if (existingLabels[i].LanguageCode !== lcid) {
                    result.push({
                        "@odata.type": "Microsoft.Dynamics.CRM.LocalizedLabel",
                        "Label": existingLabels[i].Label,
                        "LanguageCode": existingLabels[i].LanguageCode
                    });
                }
            }
        }
        
        // Add the new/updated label
        if (normalized.trim() !== "") {
            result.push({
                "@odata.type": "Microsoft.Dynamics.CRM.LocalizedLabel",
                "Label": normalized,
                "LanguageCode": lcid
            });
        }
        
        return result;
    }

    function getAttributeMetadataType(item) {
        if (!item.attributeTypeName) return null;
        var typeName = String(item.attributeTypeName).toLowerCase();
        
        if (typeName.indexOf("boolean") >= 0) return "BooleanAttributeMetadata";
        if (typeName.indexOf("datetime") >= 0) return "DateTimeAttributeMetadata";
        if (typeName.indexOf("decimal") >= 0) return "DecimalAttributeMetadata";
        if (typeName.indexOf("double") >= 0) return "DoubleAttributeMetadata";
        if (typeName.indexOf("integer") >= 0) return "IntegerAttributeMetadata";
        if (typeName.indexOf("bigint") >= 0) return "BigIntAttributeMetadata";
        if (typeName.indexOf("money") >= 0) return "MoneyAttributeMetadata";
        if (typeName.indexOf("picklist") >= 0) return "PicklistAttributeMetadata";
        if (typeName.indexOf("state") >= 0) return "StateAttributeMetadata";
        if (typeName.indexOf("status") >= 0) return "StatusAttributeMetadata";
        if (typeName.indexOf("memo") >= 0) return "MemoAttributeMetadata";
        if (typeName.indexOf("lookup") >= 0) return "LookupAttributeMetadata";
        if (typeName.indexOf("customer") >= 0) return "LookupAttributeMetadata";
        if (typeName.indexOf("owner") >= 0) return "LookupAttributeMetadata";
        if (typeName.indexOf("multiselectpicklist") >= 0) return "MultiSelectPicklistAttributeMetadata";
        if (typeName.indexOf("uniqueidentifier") >= 0) return "UniqueIdentifierAttributeMetadata";
        
        return "StringAttributeMetadata";
    }

    function saveFieldTranslations(updates, lcid) {
        var base = Xrm.Utility.getGlobalContext().getClientUrl() + "/api/data/v9.2/";
        var entityPath = "EntityDefinitions(LogicalName='" + state.entityLogicalName + "')";
        var chain = Promise.resolve();
        var saved = 0;
        
        console.log("saveFieldTranslations (Minimal PUT + UserLocalizedLabel) - lcid:", lcid, "updates:", updates.length);
        
        for (var i = 0; i < updates.length; i++) {
            (function (item) {
                chain = chain.then(function () {
                    // Skip non-customizable fields
                    if (item.canModify === false) {
                        console.log("  ⚠️ Skipping locked field (not customizable):", item.logicalName);
                        return;
                    }
                    
                    // Fetch current labels to merge
                    var readUrl = base + entityPath + "/Attributes(" + item.id + ")?$select=DisplayName,Description,LogicalName,AttributeType";
                    
                    console.log("  Processing field:", item.logicalName, "MetadataId:", item.id);
                    console.log("    newLabel:", item.newLabel);
                    console.log("    targetLabel:", item.targetLabel);
                    
                    return fetchJson(readUrl).then(function (currentAttr) {
                        var payload = {
                            "@odata.type": currentAttr["@odata.type"] || "Microsoft.Dynamics.CRM.StringAttributeMetadata",
                            "LogicalName": currentAttr.LogicalName,
                            "AttributeType": currentAttr.AttributeType
                        };
                        
                        var hasChanges = false;
                        
                        // Update DisplayName if changed
                        if (item.newLabel !== item.targetLabel) {
                            var existingLabels = currentAttr.DisplayName && currentAttr.DisplayName.LocalizedLabels;
                            var mergedLabels = mergeLocalizedLabels(existingLabels, lcid, item.newLabel);
                            
                            payload.DisplayName = {
                                "@odata.type": "Microsoft.Dynamics.CRM.Label",
                                "LocalizedLabels": mergedLabels
                            };
                            
                            // CRITICAL: Also set UserLocalizedLabel for the target language
                            payload.DisplayName.UserLocalizedLabel = {
                                "@odata.type": "Microsoft.Dynamics.CRM.LocalizedLabel",
                                "Label": item.newLabel,
                                "LanguageCode": lcid
                            };
                            
                            hasChanges = true;
                            console.log("    DisplayName updated with", mergedLabels.length, "labels");
                        }
                        
                        // Update Description if changed
                        if (item.newDescription !== item.targetDescription) {
                            var existingDescs = currentAttr.Description && currentAttr.Description.LocalizedLabels;
                            var mergedDescs = mergeLocalizedLabels(existingDescs, lcid, item.newDescription);
                            
                            payload.Description = {
                                "@odata.type": "Microsoft.Dynamics.CRM.Label",
                                "LocalizedLabels": mergedDescs
                            };
                            
                            // CRITICAL: Also set UserLocalizedLabel for the target language
                            payload.Description.UserLocalizedLabel = {
                                "@odata.type": "Microsoft.Dynamics.CRM.LocalizedLabel",
                                "Label": item.newDescription,
                                "LanguageCode": lcid
                            };
                            
                            hasChanges = true;
                            console.log("    Description updated with", mergedDescs.length, "labels");
                        }
                        
                        if (!hasChanges) {
                            console.log("  No changes for field, skipping");
                            return;
                        }
                        
                        var putUrl = base + entityPath + "/Attributes(" + item.id + ")";
                        console.log("    PUT URL:", putUrl);
                        console.log("    Payload:", JSON.stringify(payload));
                        
                        // PUT with minimal metadata and MSCRM.MergeLabels header
                        return fetchJson(putUrl, "PUT", payload, { "MSCRM.MergeLabels": "true" }).then(function () {
                            saved++;
                            console.log("  ✓ Saved field:", item.logicalName);
                        }).catch(function(err) {
                            console.error("  ✗ Failed to save field:", item.logicalName, err);
                            throw err;
                        });
                    });
                });
            })(updates[i]);
        }
        
        return chain.then(function () {
            console.log("Saved " + saved + " field translation(s), now publishing...");
            
            // Publish the entity to make changes visible
            var publishPayload = {
                ParameterXml: "<importexportxml><entities><entity>" + state.entityLogicalName + "</entity></entities></importexportxml>"
            };
            
            return fetchJson(base + "PublishXml", "POST", publishPayload).then(function () {
                console.log("Published " + state.entityLogicalName + " field translations successfully.");
                return saved;
            }).catch(function (publishErr) {
                console.warn("Publish failed. Changes saved but may require manual publish:", publishErr);
                return saved;
            });
        });
    }

    function saveEntityTranslations(updates, lcid) {
        var base = Xrm.Utility.getGlobalContext().getClientUrl() + "/api/data/v9.2/";
        var entityPath = "EntityDefinitions(LogicalName='" + state.entityLogicalName + "')";
        var chain = Promise.resolve();
        var saved = 0;
        
        console.log("saveEntityTranslations (HAR-route) - lcid:", lcid, "updates:", updates.length);
        
        for (var i = 0; i < updates.length; i++) {
            (function (item) {
                chain = chain.then(function () {
                    // Fetch current entity labels
                    var readUrl = base + entityPath + "?$select=DisplayName,DisplayCollectionName,Description";
                    
                    return fetchJson(readUrl).then(function(currentEntity) {
                        var payload = {
                            "@odata.type": "Microsoft.Dynamics.CRM.EntityMetadata"
                        };
                        
                        if (item.propertyName === "DisplayName") {
                            var existingLabels = currentEntity.DisplayName && currentEntity.DisplayName.LocalizedLabels;
                            payload.DisplayName = {
                                "@odata.type": "Microsoft.Dynamics.CRM.Label",
                                "LocalizedLabels": mergeLocalizedLabels(existingLabels, lcid, item.newLabel)
                            };
                        } else if (item.propertyName === "DisplayCollectionName") {
                            var existingLabels = currentEntity.DisplayCollectionName && currentEntity.DisplayCollectionName.LocalizedLabels;
                            payload.DisplayCollectionName = {
                                "@odata.type": "Microsoft.Dynamics.CRM.Label",
                                "LocalizedLabels": mergeLocalizedLabels(existingLabels, lcid, item.newLabel)
                            };
                        } else if (item.propertyName === "Description") {
                            var existingLabels = currentEntity.Description && currentEntity.Description.LocalizedLabels;
                            payload.Description = {
                                "@odata.type": "Microsoft.Dynamics.CRM.Label",
                                "LocalizedLabels": mergeLocalizedLabels(existingLabels, lcid, item.newLabel)
                            };
                        }
                        
                        // Use UpdateEntity via Organization Request for better reliability
                        var requestPayload = {
                            Entity: payload,
                            MergeLabels: true
                        };
                        requestPayload.Entity.MetadataId = item.id;
                        requestPayload.Entity.LogicalName = state.entityLogicalName;
                        
                        var updateUrl = base + "UpdateEntity";
                        return fetchJson(updateUrl, "POST", requestPayload).then(function () {
                            saved++;
                            console.log("  Saved entity property:", item.propertyName);
                        }).catch(function(err) {
                            console.error("  Failed to save entity property:", item.propertyName, err);
                            // Fallback to direct PUT if UpdateEntity fails
                            return fetchJson(base + entityPath, "PUT", payload, { "MSCRM.MergeLabels": "true" }).then(function () {
                                saved++;
                                console.log("  Saved entity property (fallback PUT):", item.propertyName);
                            });
                        });
                    });
                });
            })(updates[i]);
        }
        
        return chain.then(function () {
            console.log("Saved " + saved + " entity translation(s), now publishing...");
            
            // Publish the entity to make changes visible
            var publishPayload = {
                ParameterXml: "<importexportxml><entities><entity>" + state.entityLogicalName + "</entity></entities></importexportxml>"
            };
            
            return fetchJson(base + "PublishXml", "POST", publishPayload).then(function () {
                console.log("Published " + state.entityLogicalName + " entity translations successfully.");
                return saved;
            }).catch(function (publishErr) {
                console.warn("Publish failed. Changes saved but may require manual publish:", publishErr);
                // Try PublishAllXml as fallback
                return fetchJson(base + "PublishAllXml", "POST", {}).then(function() {
                    console.log("PublishAllXml succeeded");
                    return saved;
                }).catch(function(err2) {
                    console.warn("PublishAllXml also failed:", err2);
                    return saved;
                });
            });
        });
    }

    function saveFormTranslations(updates, lcid) {
        var base = Xrm.Utility.getGlobalContext().getClientUrl() + "/api/data/v9.2/";
        var chain = Promise.resolve();
        var saved = 0;
        
        console.log("saveFormTranslations - lcid:", lcid, "updates:", updates.length);
        
        for (var i = 0; i < updates.length; i++) {
            (function (item) {
                chain = chain.then(function () {
                    var id = String(item.id).replace(/[{}]/g, "").toLowerCase();
                    var promises = [];
                    
                    // Update name using SetLocLabels for proper localization
                    if (item.newLabel !== item.targetLabel) {
                        var namePayload = {
                            EntityMoniker: {
                                "@odata.id": base + "systemforms(" + id + ")"
                            },
                            AttributeName: "name",
                            Labels: [
                                {
                                    "@odata.type": "Microsoft.Dynamics.CRM.LocalizedLabel",
                                    Label: item.newLabel || "",
                                    LanguageCode: lcid
                                }
                            ]
                        };
                        console.log("  SetLocLabels form name payload:", JSON.stringify(namePayload));
                        promises.push(fetchJson(base + "SetLocLabels", "POST", namePayload));
                    }
                    
                    // Update description using SetLocLabels
                    if (item.newDescription !== item.targetDescription) {
                        var descPayload = {
                            EntityMoniker: {
                                "@odata.id": base + "systemforms(" + id + ")"
                            },
                            AttributeName: "description",
                            Labels: [
                                {
                                    "@odata.type": "Microsoft.Dynamics.CRM.LocalizedLabel",
                                    Label: item.newDescription || "",
                                    LanguageCode: lcid
                                }
                            ]
                        };
                        console.log("  SetLocLabels form description payload:", JSON.stringify(descPayload));
                        promises.push(fetchJson(base + "SetLocLabels", "POST", descPayload));
                    }
                    
                    return Promise.all(promises).then(function () {
                        saved++;
                    });
                });
            })(updates[i]);
        }
        
        return chain.then(function () {
            console.log("Saved " + saved + " form translation(s)");
            return saved;
        });
    }

    function saveViewTranslations(updates, lcid) {
        var base = Xrm.Utility.getGlobalContext().getClientUrl() + "/api/data/v9.2/";
        var chain = Promise.resolve();
        var saved = 0;
        
        console.log("saveViewTranslations - lcid:", lcid, "updates:", updates.length);
        
        for (var i = 0; i < updates.length; i++) {
            (function (item) {
                chain = chain.then(function () {
                    var id = String(item.id).replace(/[{}]/g, "").toLowerCase();
                    var promises = [];
                    
                    // Update name using SetLocLabels for proper localization
                    if (item.newLabel !== item.targetLabel) {
                        var namePayload = {
                            EntityMoniker: {
                                "@odata.id": base + "savedqueries(" + id + ")"
                            },
                            AttributeName: "name",
                            Labels: [
                                {
                                    "@odata.type": "Microsoft.Dynamics.CRM.LocalizedLabel",
                                    Label: item.newLabel || "",
                                    LanguageCode: lcid
                                }
                            ]
                        };
                        console.log("  SetLocLabels name payload:", JSON.stringify(namePayload));
                        promises.push(fetchJson(base + "SetLocLabels", "POST", namePayload));
                    }
                    
                    // Update description using SetLocLabels
                    if (item.newDescription !== item.targetDescription) {
                        var descPayload = {
                            EntityMoniker: {
                                "@odata.id": base + "savedqueries(" + id + ")"
                            },
                            AttributeName: "description",
                            Labels: [
                                {
                                    "@odata.type": "Microsoft.Dynamics.CRM.LocalizedLabel",
                                    Label: item.newDescription || "",
                                    LanguageCode: lcid
                                }
                            ]
                        };
                        console.log("  SetLocLabels description payload:", JSON.stringify(descPayload));
                        promises.push(fetchJson(base + "SetLocLabels", "POST", descPayload));
                    }
                    
                    return Promise.all(promises).then(function () {
                        saved++;
                    });
                });
            })(updates[i]);
        }
        
        return chain.then(function () {
            console.log("Saved " + saved + " view translation(s), now publishing...");
            
            // Publish the entity (not just savedquery) to refresh metadata cache
            var publishPayload = {
                ParameterXml: "<importexportxml><entities><entity>" + state.entityLogicalName + "</entity></entities></importexportxml>"
            };
            
            return fetchJson(base + "PublishXml", "POST", publishPayload).then(function () {
                console.log("Published " + state.entityLogicalName + " successfully.");
                console.log("Note: You may need to refresh CE (Ctrl+F5) or clear browser cache to see changes.");
                return saved;
            }).catch(function (publishErr) {
                console.warn("Publish failed. Changes saved but may require manual publish:", publishErr);
                // Try PublishAllXml as fallback
                return fetchJson(base + "PublishAllXml", "POST", {}).then(function() {
                    console.log("PublishAllXml succeeded");
                    return saved;
                }).catch(function(err2) {
                    console.warn("PublishAllXml also failed:", err2);
                    return saved;
                });
            });
        });
    }

    function saveFormLabelTranslations(updates, lcid) {
        if (!state.formId) {
            console.error("No form selected");
            return Promise.reject("No form selected");
        }
        
        console.log("saveFormLabelTranslations - lcid:", lcid, "updates:", updates.length);
        
        // Separate updates into FormXML updates and field metadata updates
        var formXmlUpdates = [];
        var fieldMetadataUpdates = [];
        
        for (var i = 0; i < updates.length; i++) {
            if (updates[i].elementType === "cell" && updates[i].usedMetadataFallback) {
                // This label came from field metadata fallback, so save it to field metadata
                fieldMetadataUpdates.push(updates[i]);
                console.log("Will save to field metadata:", updates[i].logicalName);
            } else {
                // This label is in FormXML, so save it to FormXML
                formXmlUpdates.push(updates[i]);
            }
        }
        
        console.log("FormXML updates:", formXmlUpdates.length, "Field metadata updates:", fieldMetadataUpdates.length);
        
        // Save field metadata updates first
        var chain = Promise.resolve();
        if (fieldMetadataUpdates.length > 0) {
            chain = saveFieldMetadataLabels(fieldMetadataUpdates, lcid);
        }
        
        // Then save FormXML updates
        return chain.then(function() {
            if (formXmlUpdates.length === 0) {
                console.log("No FormXML updates to save");
                return updates.length; // Return total count
            }
            
            return saveFormXMLLabels(formXmlUpdates, lcid);
        });
    }
    
    function saveFieldMetadataLabels(updates, lcid) {
        console.log("saveFieldMetadataLabels (HAR-route) - lcid:", lcid, "updates:", updates.length);
        var base = Xrm.Utility.getGlobalContext().getClientUrl() + "/api/data/v9.2/";
        var entityPath = "EntityDefinitions(LogicalName='" + state.entityLogicalName + "')";
        var chain = Promise.resolve();
        var saved = 0;
        
        // Group updates by field to avoid multiple API calls for the same field
        var fieldUpdates = {};
        for (var i = 0; i < updates.length; i++) {
            var fieldName = updates[i].logicalName;
            if (!fieldUpdates[fieldName]) {
                fieldUpdates[fieldName] = updates[i];
            }
        }
        
        // Update each field using HAR-route
        for (var fieldName in fieldUpdates) {
            (function(item) {
                chain = chain.then(function() {
                    // First, get MetadataId by LogicalName
                    var lookupUrl = base + entityPath + "/Attributes(LogicalName='" + item.logicalName + "')?$select=MetadataId,DisplayName,Description";
                    
                    return fetchJson(lookupUrl).then(function(currentAttr) {
                        if (!currentAttr || !currentAttr.MetadataId) {
                            console.warn("Could not find attribute:", item.logicalName);
                            return;
                        }
                        
                        var payload = {};
                        
                        // Merge DisplayName labels
                        var existingLabels = currentAttr.DisplayName && currentAttr.DisplayName.LocalizedLabels;
                        payload.DisplayName = {
                            "@odata.type": "Microsoft.Dynamics.CRM.Label",
                            "LocalizedLabels": mergeLocalizedLabels(existingLabels, lcid, item.newLabel)
                        };
                        
                        console.log("  Field metadata DisplayName merged", payload.DisplayName.LocalizedLabels.length, "labels for:", item.logicalName);
                        
                        var putUrl = base + entityPath + "/Attributes(" + currentAttr.MetadataId + ")";
                        // MSCRM.MergeLabels: true to preserve other language labels
                        return fetchJson(putUrl, "PUT", payload, { "MSCRM.MergeLabels": "true" }).then(function() {
                            saved++;
                            console.log("Saved field metadata for:", item.logicalName);
                        });
                    }).catch(function(err) {
                        console.error("Failed to update field metadata for", item.logicalName, ":", err);
                        // Continue with other updates even if one fails
                    });
                });
            })(fieldUpdates[fieldName]);
        }
        
        return chain.then(function() {
            console.log("Saved", saved, "field metadata label(s), now publishing...");
            
            // Publish the entity to make changes visible
            var publishPayload = {
                ParameterXml: "<importexportxml><entities><entity>" + state.entityLogicalName + "</entity></entities></importexportxml>"
            };
            
            return fetchJson(base + "PublishXml", "POST", publishPayload).then(function () {
                console.log("Published field metadata label changes successfully.");
                return saved;
            }).catch(function (publishErr) {
                console.warn("Publish failed. Changes saved but may require manual publish:", publishErr);
                // Try PublishAllXml as fallback
                return fetchJson(base + "PublishAllXml", "POST", {}).then(function() {
                    console.log("PublishAllXml succeeded");
                    return saved;
                }).catch(function(err2) {
                    console.warn("PublishAllXml also failed:", err2);
                    return saved;
                });
            });
        });
    }
    
    function saveFormXMLLabels(updates, lcid) {
        console.log("saveFormXMLLabels - lcid:", lcid, "updates:", updates.length);
        
        var base = Xrm.Utility.getGlobalContext().getClientUrl() + "/api/data/v9.2/";
        var url = base + "systemforms(" + state.formId + ")?$select=formxml";
        
        return fetchJson(url).then(function (result) {
            if (!result || !result.formxml) {
                throw new Error("Failed to load form XML");
            }
            
            var parser = new DOMParser();
            var xmlDoc = parser.parseFromString(result.formxml, "text/xml");
            
            // Check for parse errors
            var parseError = xmlDoc.querySelector("parsererror");
            if (parseError) {
                throw new Error("XML parse error: " + parseError.textContent);
            }
            
            // Build index of all label containers (tabs, sections, cells)
            var index = buildFormLabelNodeIndex(xmlDoc);
            var modified = false;
            
            console.log("Built node index with", Object.keys(index).length, "entries");
            console.log("Processing", updates.length, "form label updates");
            
            // Process each update using DOM manipulation
            for (var i = 0; i < updates.length; i++) {
                var item = updates[i];
                
                // Find the node using labelObjectId or fallback to id
                var lookupKey = item.labelObjectId || item.id;
                var hit = index[lookupKey];
                
                if (!hit) {
                    console.warn("Node not found in index for key:", lookupKey, "elementType:", item.elementType);
                    continue;
                }
                
                console.log("Updating", item.elementType, "key:", lookupKey, "label:", item.newLabel);
                
                // Update the label in the DOM
                setLocalizedDescriptionOnContainer(hit.node, lcid, item.newLabel || "", xmlDoc);
                modified = true;
            }
            
            if (!modified) {
                console.log("No FormXML changes to save");
                return 0;
            }
            
            // Serialize back to XML string
            var serializer = new XMLSerializer();
            var newFormXml = serializer.serializeToString(xmlDoc);
            
            console.log("Saving form XML, original:", result.formxml.length, "new:", newFormXml.length);
            
            // Save the updated form XML
            var saveUrl = base + "systemforms(" + state.formId + ")";
            var payload = { formxml: newFormXml };
            
            return fetchJson(saveUrl, "PATCH", payload).then(function () {
                console.log("Form XML saved, now publishing...");
                
                // Publish the customizations
                var publishUrl = base + "PublishXml";
                var publishPayload = {
                    ParameterXml: "<importexportxml><entities><entity>" + state.entityLogicalName + "</entity></entities></importexportxml>"
                };
                
                return fetchJson(publishUrl, "POST", publishPayload).then(function () {
                    console.log("Customizations published successfully");
                    return updates.length;
                }).catch(function (publishErr) {
                    console.warn("Publish warning:", publishErr);
                    return updates.length;
                });
            });
        });
    }
})();
