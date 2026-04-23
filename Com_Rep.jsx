(function (thisObj) {
    var getAepBaseName = function () {
        var f = app.project && app.project.file ? app.project.file : null;
        if (!f) return "Untitled";
        return decodeURI(f.name).replace(/\.aep$/i, "");
    };

    var trim = function (s) {
        return (s || "").replace(/^\s+|\s+$/g, "");
    };

    var getCompsInFolder = function (folderItem) {
        var comps = [];
        var walk = function (folder) {
            for (var i = 1; i <= app.project.numItems; i++) {
                var item = app.project.item(i);
                if (item.parentFolder !== folder) continue;
                if (item instanceof CompItem) comps.push(item);
                else if (item instanceof FolderItem) walk(item);
            }
        };
        walk(folderItem);
        return comps;
    };

    var getFoldersInFolder = function (folderItem) {
        var folders = [];
        var walk = function (folder) {
            for (var i = 1; i <= app.project.numItems; i++) {
                var item = app.project.item(i);
                if (item.parentFolder !== folder) continue;
                if (item instanceof FolderItem) {
                    folders.push(item);
                    walk(item);
                }
            }
        };
        walk(folderItem);
        return folders;
    };

    var getOrCreateChildFolder = function (parentFolder, name) {
        for (var i = 1; i <= app.project.numItems; i++) {
            var item = app.project.item(i);
            if (item instanceof FolderItem && item.parentFolder === parentFolder && item.name === name) {
                return item;
            }
        }
        var newFolder = app.project.items.addFolder(name);
        newFolder.parentFolder = parentFolder;
        return newFolder;
    };

    var isDescendantFolder = function (candidate, ancestor) {
        var cur = candidate;
        while (cur && cur !== app.project.rootFolder) {
            if (cur.id === ancestor.id) return true;
            cur = cur.parentFolder;
        }
        return false;
    };

    var win = (thisObj instanceof Panel) ? thisObj : new Window("palette", "Com_Rep", undefined);
    win.orientation = "column";
    win.alignChildren = ["fill", "top"];
    win.spacing = 12;
    win.margins = 16;

    // 1) プロジェクト・フォルダ名
    var pNameGroup = win.add("panel", undefined, "1. プロジェクト・フォルダ名");
    pNameGroup.orientation = "column";
    pNameGroup.alignChildren = ["left", "top"];
    pNameGroup.margins = 12;

    var modeGroup = pNameGroup.add("group");
    modeGroup.orientation = "row";
    modeGroup.alignChildren = ["left", "center"];
    modeGroup.alignment = ["fill", "top"];
    var rbAuto = modeGroup.add("radiobutton", undefined, "ファイル名を使う（自動）");
    var rbCustom = modeGroup.add("radiobutton", undefined, "自分で打つ（カスタム）");
    rbAuto.alignment = ["left", "center"];
    rbCustom.alignment = ["left", "center"];

    var updateModeGroupLayout = function () {
        var w = pNameGroup.size && pNameGroup.size.width ? pNameGroup.size.width : win.size.width;
        modeGroup.orientation = (w < 320) ? "column" : "row";
        modeGroup.layout.layout(true);
    };

    var editProjName = pNameGroup.add("edittext", undefined, getAepBaseName());
    editProjName.alignment = ["fill", "center"];

    rbAuto.value = true;
    editProjName.enabled = false;

    var customNameCache = "";
    var syncNameMode = function () {
        if (rbAuto.value) {
            // auto に戻った時点でキャッシュをクリア
            customNameCache = "";
            editProjName.text = getAepBaseName();
            editProjName.enabled = false;
        } else {
            editProjName.enabled = true;
            editProjName.text = customNameCache ? customNameCache : getAepBaseName();
            editProjName.active = true;
        }
    };
    rbAuto.onClick = syncNameMode;
    rbCustom.onClick = syncNameMode;

    // コピー元（マスター）選択
    var compGroup = win.add("panel", undefined, "コピー元のデザイン（雛形）を選択");
    compGroup.orientation = "column";
    compGroup.alignChildren = ["fill", "top"];
    compGroup.margins = 12;

    var selectTypeGroup = compGroup.add("group");
    selectTypeGroup.orientation = "row";
    selectTypeGroup.alignChildren = ["left", "center"];
    var rbSelectComp = selectTypeGroup.add("radiobutton", undefined, "ファイル（コンポ）");
    var rbSelectFolder = selectTypeGroup.add("radiobutton", undefined, "フォルダ");
    rbSelectComp.value = true;

    var compList = compGroup.add("dropdownlist", undefined, []);
    compList.alignment = "fill";

    var encodeCompGroup = compGroup.add("group");
    encodeCompGroup.orientation = "row";
    encodeCompGroup.alignChildren = ["left", "center"];
    encodeCompGroup.add("statictext", undefined, "エンコード用コンポ:");
    var encodeCompList = encodeCompGroup.add("dropdownlist", undefined, []);
    encodeCompList.minimumSize.width = 180;

    var saveGroup = compGroup.add("group");
    saveGroup.orientation = "row";
    saveGroup.alignChildren = ["left", "center"];
    saveGroup.add("statictext", undefined, "保存先階層:");
    var saveFolderList = saveGroup.add("dropdownlist", undefined, []);
    saveFolderList.minimumSize.width = 180;

    // ドロップダウン項目の直接参照を保持（CompItem / FolderItem）
    var masterTemplateRefs = [];
    var saveFolderRefs = [];

    var updateCompList = function () {
        compList.removeAll();
        masterTemplateRefs = [];
        if (!app.project) return;
        for (var i = 1; i <= app.project.numItems; i++) {
            var item = app.project.item(i);
            if (rbSelectComp.value && item instanceof CompItem) {
                compList.add("item", item.name);
                masterTemplateRefs.push(item);
            } else if (rbSelectFolder.value && item instanceof FolderItem && item !== app.project.rootFolder) {
                compList.add("item", item.name);
                masterTemplateRefs.push(item);
            }
        }
        if (compList.items.length > 0) compList.selection = 0;
        updateEncodeCompList();
        updateSaveFolderList();
    };

    var updateSaveFolderList = function () {
        var prevFolder = (saveFolderList.selection && saveFolderRefs[saveFolderList.selection.index])
            ? saveFolderRefs[saveFolderList.selection.index]
            : null;
        saveFolderList.removeAll();
        saveFolderRefs = [];
        if (!app.project) return;

        saveFolderList.add("item", "[Root] プロジェクト直下");
        saveFolderRefs.push(app.project.rootFolder);

        for (var i = 1; i <= app.project.numItems; i++) {
            var item = app.project.item(i);
            if (item instanceof FolderItem && item !== app.project.rootFolder) {
                saveFolderList.add("item", item.name);
                saveFolderRefs.push(item);
            }
        }

        if (saveFolderList.items.length > 0) {
            var restored = false;
            if (prevFolder) {
                for (var j = 0; j < saveFolderRefs.length; j++) {
                    if (saveFolderRefs[j].id === prevFolder.id) {
                        saveFolderList.selection = j;
                        restored = true;
                        break;
                    }
                }
            }
            if (!restored) saveFolderList.selection = 0;
        }
    };

    var updateEncodeCompList = function () {
        encodeCompList.removeAll();
        encodeCompList.enabled = false;
        if (!compList.selection) return;

        var selectedTemplate = masterTemplateRefs[compList.selection.index];
        if (selectedTemplate instanceof CompItem) {
            encodeCompList.add("item", selectedTemplate.name);
            encodeCompList.selection = 0;
            encodeCompList.enabled = false;
            return;
        }

        if (selectedTemplate instanceof FolderItem) {
            var compsInFolder = getCompsInFolder(selectedTemplate);
            for (var i = 0; i < compsInFolder.length; i++) {
                encodeCompList.add("item", compsInFolder[i].name);
            }
            if (encodeCompList.items.length > 0) {
                encodeCompList.selection = 0;
                encodeCompList.enabled = true;
            }
        }
    };
    updateCompList();
    compList.onChange = updateEncodeCompList;
    rbSelectComp.onClick = updateCompList;
    rbSelectFolder.onClick = updateCompList;

    var btnRefresh = compGroup.add("button", undefined, "リスト更新");
    btnRefresh.onClick = updateCompList;

    // 生成リスト
    var listGroup = win.add("panel", undefined, "生成名リスト");
    listGroup.orientation = "column";
    listGroup.alignChildren = ["fill", "top"];
    listGroup.margins = 12;

    var checkList = listGroup.add("checkbox", undefined, "一括生成モード（改行区切りリストを使う）");
    checkList.value = true;

    var editWorks = listGroup.add("edittext", undefined, "Intro\rMotion_A\rMotion_B\rVFX_Shot\rOutro", { multiline: true });
    editWorks.preferredSize.height = 110;
    editWorks.alignment = "fill";

    checkList.onClick = function () {
        editWorks.enabled = checkList.value;
    };

    // 4) 実行
    var btnRun = win.add("button", undefined, "フォルダとリストを生成");

    btnRun.onClick = function () {
        if (!app.project) {
            alert("プロジェクトが開かれていません。");
            return;
        }
        if (!compList.selection) {
            alert("コピー元のコンポ／フォルダを選んでください。");
            return;
        }

        var folderName = trim(editProjName.text);
        if (!folderName) {
            alert("フォルダ名を入力してください。");
            return;
        }

        var works = [];
        if (checkList.value) {
            var lines = editWorks.text.split(/\r\n|\n|\r/);
            for (var i = 0; i < lines.length; i++) {
                var t = trim(lines[i]);
                if (t) works.push(t);
            }
            if (works.length === 0) {
                alert("生成名リストが空です。");
                return;
            }
        } else {
            works.push(folderName);
        }

        var proj = app.project;
        var selIdx = compList.selection.index;
        var selectedTemplate = masterTemplateRefs[selIdx];
        var masterComp = null;
        var encodeSourceComp = null;
        var sourceFolders = null;
        var sourceComps = null;

        if (selectedTemplate instanceof CompItem) {
            masterComp = selectedTemplate;
            encodeSourceComp = selectedTemplate;
        } else if (selectedTemplate instanceof FolderItem) {
            sourceComps = getCompsInFolder(selectedTemplate);
            sourceFolders = getFoldersInFolder(selectedTemplate);
            var compsInFolder = sourceComps;
            if (compsInFolder.length === 0) {
                alert("選択フォルダ内にコンポが見つかりません。");
                return;
            }
            if (!encodeCompList.selection) {
                alert("エンコード用コンポを選択してください。");
                return;
            }
            encodeSourceComp = compsInFolder[encodeCompList.selection.index];
        }

        if (!(selectedTemplate instanceof FolderItem) && !(masterComp instanceof CompItem)) {
            alert("選択した雛形コンポが見つかりません。");
            return;
        }

        // 生成前プレビュー確認
        var saveParentFolder = (saveFolderList.selection && saveFolderRefs[saveFolderList.selection.index])
            ? saveFolderRefs[saveFolderList.selection.index]
            : proj.rootFolder;
        var effectiveSaveParent = saveParentFolder;
        if (selectedTemplate instanceof FolderItem
            && effectiveSaveParent.id !== selectedTemplate.id
            && isDescendantFolder(effectiveSaveParent, selectedTemplate)) {
            effectiveSaveParent = selectedTemplate.parentFolder || proj.rootFolder;
        }
        var targetFolderName = folderName;
        if (selectedTemplate instanceof FolderItem && effectiveSaveParent.id === selectedTemplate.parentFolder.id && folderName === selectedTemplate.name) {
            targetFolderName = folderName + "_out";
        }
        var preview = "以下を生成します。\n\n";
        preview += "フォルダ: " + folderName + "\n";
        preview += "保存先: " + (effectiveSaveParent === proj.rootFolder ? "[Root] プロジェクト直下" : effectiveSaveParent.name) + "\n";
        preview += "雛形: " + (selectedTemplate instanceof FolderItem ? selectedTemplate.name + " (Folder)" : masterComp.name) + "\n";
        if (encodeSourceComp) preview += "リネーム対象: " + encodeSourceComp.name + "\n";
        preview += "件数: " + works.length + "\n\n";
        for (var k = 0; k < works.length; k++) {
            var previewName = checkList.value ? (folderName + "_" + works[k]) : folderName;
            preview += previewName + "\n";
        }

        if (!confirm(preview)) return;

        app.beginUndoGroup("Create Portfolio Folder and List");
        var targetFolder = getOrCreateChildFolder(effectiveSaveParent, targetFolderName);

        for (var n = 0; n < works.length; n++) {
            var generatedName = checkList.value ? (folderName + "_" + works[n]) : folderName;

            if (selectedTemplate instanceof CompItem) {
                var newComp = masterComp.duplicate();
                newComp.name = generatedName;
                newComp.parentFolder = targetFolder;

                // 注：以下の処理は、雛形コンポが最低1層を持つことを前提としています。
                // 雛形コンポの1層目のレイヤー名が「作品名」に変更される仕様に依存しています。
                if (newComp.numLayers > 0) {
                    newComp.layer(1).name = generatedName;
                }
                continue;
            }

            // フォルダ選択時: フォルダ内を丸ごと複製し、指定1コンポのみリネーム
            var setFolder = targetFolder;
            if (checkList.value) {
                setFolder = getOrCreateChildFolder(targetFolder, generatedName);
            }

            var folderMap = {};
            folderMap[selectedTemplate.id] = setFolder;

            for (var f = 0; f < sourceFolders.length; f++) {
                var srcFolder = sourceFolders[f];
                var dupFolder = getOrCreateChildFolder(folderMap[srcFolder.parentFolder.id] || setFolder, srcFolder.name);
                folderMap[srcFolder.id] = dupFolder;
            }

            for (var c = 0; c < sourceComps.length; c++) {
                var srcComp = sourceComps[c];
                var dupComp = srcComp.duplicate();
                var compParent = folderMap[srcComp.parentFolder.id] || setFolder;
                dupComp.parentFolder = compParent;
                if (encodeSourceComp && srcComp.id === encodeSourceComp.id) {
                    dupComp.name = generatedName;
                } else {
                    dupComp.name = srcComp.name;
                }
            }
        }

        app.endUndoGroup();
        alert("フォルダ「" + folderName + "」内に " + works.length + " 個のコンポを生成しました。");
    };

    win.onResizing = win.onResize = function () {
        updateModeGroupLayout();
        pNameGroup.layout.layout(true);
        this.layout.resize();
    };
    updateModeGroupLayout();

    if (win instanceof Window) {
        win.center();
        win.show();
    } else {
        win.layout.layout(true);
    }
})(this);
