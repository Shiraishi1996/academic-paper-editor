const I18N = {
    ja: {
        app_name:"論文エディタ", saved:"保存済み", saving:"保存中...", publisher_preset:"出版社書式プリセット",
        generic:"汎用", new_project:"新規作成", save_project:"プロジェクト保存", open_project:"プロジェクトを開く", backup:"バックアップ", restore:"復元", export_format:"出力形式...", export_docx:"Word・画像・数式埋込 (.docx)",
        export_pdf:"PDF印刷表示", export_latex_zip:"LaTeX + 画像 (.zip)", export_images_zip:"画像のみ一括 (.zip)",
        structure_assets:"構成 / 素材", edit:"編集", preview:"プレビュー", start_writing:"論文の執筆を始める",
        start_desc:"既存原稿をアップロードすると、タイトル・段落・画像・Word数式を構造化して読み込みます。",
        initial_upload:"ここにファイルをドロップして読み込む", supported_desc:"Word (.docx), PDF, テキスト等に対応。アップロードしたファイル名は論文タイトルには使用しません。",
        start_blank:"空から始める", outline:"構成", assets:"素材", references:"文献", section_paragraphs:"セクション / 段落", add:"追加",
        upload_material:"資料を解析して追加", upload:"解析して挿入", material_drop:"ここに資料をドロップして追加", material_supported:"Wordは段落・画像・表・数式まで解析して現在位置へ挿入します。PDF・テキスト・画像にも対応します。", drag_hint:"💡 初回アップロードと同じ解析器を使用します。Word数式と埋込画像も保持します。", bibtex_data:"BibTeXデータ",
        manuscript_editor:"原稿編集", paper_title:"論文タイトル", authors:"著者名", affiliations:"所属", abstract:"概要 (Abstract)", keywords:"キーワード（カンマ区切り）", translation:"翻訳", academic_style_on:"日本語本文は常体・である調に自動整形",
        to_english:"英語へ", to_japanese:"日本語へ", translate_paragraph:"選択段落を翻訳", translate_all:"全文を翻訳",
        translation_local:"ローカル（API不要）", translation_online:"オンライン（任意）", translation_online_free:"オンライン翻訳（APIキー不要）", translation_online_help:"APIキー不要のオンライン翻訳を使用します。翻訳時には対象テキストがオンライン翻訳サービスへ送信されます。サービス側の混雑や制限により一時的に失敗する場合があります。", translation_model:"ローカルモデル", model_fugumt:"FuguMT（日英専用・推奨）", model_m2m100:"M2M100 418M（大容量）", model_opus:"OPUS-MT（軽量・旧方式）", glossary:"用語辞書", glossary_label:"学術用語辞書（1行に「日本語 = English」）", glossary_placeholder:"例: 可変浮力 = variable buoyancy", glossary_help:"数式、引用番号、DOI/URL、Figure/Table/Equation番号は翻訳時に保護されます。", translation_local_help:"既定は日英専用FuguMTです。品質をさらに比較したい場合はM2M100を選べます。初回のみモデルを取得し、その後はオフライン利用できます。", local_model_loading:"ローカル翻訳モデルを読み込み中（初回はモデル取得を含みます）...",
        section_heading_optional:"セクション見出し（不要なら空欄）", next_paragraph:"次に段落", editor_help:"左の構成欄から段落・画像・数式を個別に選択できます。数式はWord OMMLを保持します。",
        paragraph_text:"段落テキスト...", live_preview:"リアルタイムプレビュー", manual_backup:"手動バックアップ", backup_note:"メモ...", cancel:"キャンセル", save:"保存", restore_history:"履歴から復元", close:"閉じる",
        paragraph:"段落", image:"画像", table:"表", equation:"数式", block:"ブロック", no_heading:"（見出しなし）", empty_paragraph:"空の段落", new_section:"新規セクション",
        confirm_delete_section:"このセクションを削除しますか？", confirm_delete_block:"このブロックを削除しますか？", processing_import:"段落・画像・数式を解析しています...",
        import_summary:(s,b,m)=>`${s}セクション / ${b}ブロック / 数式${m}個を抽出しました`, parse_error:"解析エラー", choose_file:"ファイルを選択してください", processing:"処理中...", upload_done:"アップロード完了", error:"エラー",
        insert_here:"現在位置へ挿入", save_image:"画像保存", insert_paragraphs:"構造を再解析して挿入", view_content:"内容を見る", confirm_delete_asset:"素材ファイルを削除しますか？参照画像が表示できなくなる場合があります。",
        preset_note:"最終投稿時は対象ジャーナル固有テンプレートを優先してください。", backup_saved:"保存しました", restore_button:"復元", confirm_restore:"現在の内容が置き換わります。よろしいですか？", restored:"復元しました",
        translating:"翻訳中...", translated:"翻訳しました", translate_failed:"翻訳エラー", translate_skip:"画像・数式ブロックは翻訳対象外です。", full_translate_confirm:"現在の原稿全体を翻訳結果で置き換えます。続行しますか？",
        new_project_confirm:"現在の作業内容を閉じて、新しいプロジェクトを作成しますか？必要なら先に「プロジェクト保存」をしてください。", new_project_created:"新しいプロジェクトを作成しました", saving_project:"プロジェクトを保存中...", project_saved:"プロジェクトを保存しました", project_save_error:"プロジェクト保存エラー", opening_project:"プロジェクトを開いています...", project_opened:"プロジェクトを開きました", project_open_error:"プロジェクト読込エラー", material_imported:(s,b,m,i)=>`素材を追加しました：${s}セクション / ${b}ブロック / 数式${m}個 / 画像${i}個`
    },
    en: {
        app_name:"Paper Editor", saved:"Saved", saving:"Saving...", publisher_preset:"Publisher format preset",
        generic:"Generic", new_project:"New Project", save_project:"Save Project", open_project:"Open Project", backup:"Backup", restore:"Restore", export_format:"Export...", export_docx:"Word with images & equations (.docx)",
        export_pdf:"Print/PDF view", export_latex_zip:"LaTeX + images (.zip)", export_images_zip:"Images only (.zip)",
        structure_assets:"Outline / Assets", edit:"Edit", preview:"Preview", start_writing:"Start writing your paper",
        start_desc:"Upload an existing manuscript to structure its title, paragraphs, images, and Word equations.",
        initial_upload:"Drop a file here to import", supported_desc:"Supports Word (.docx), PDF, and text. The uploaded filename is never used as the manuscript title.",
        start_blank:"Start blank", outline:"Outline", assets:"Assets", references:"References", section_paragraphs:"Sections / Paragraphs", add:"Add",
        upload_material:"Parse and add material", upload:"Parse & insert", material_drop:"Drop material here to add", material_supported:"Word files are parsed into paragraphs, images, tables, and equations and inserted at the current position. PDF, text, and images are also supported.", drag_hint:"💡 Uses the same parser as the initial import. Embedded images and Word equations are preserved.", bibtex_data:"BibTeX data",
        manuscript_editor:"Manuscript Editor", paper_title:"Paper title", authors:"Authors", affiliations:"Affiliations", abstract:"Abstract", keywords:"Keywords (comma-separated)", translation:"Translation", academic_style_on:"Japanese manuscript text is normalized to academic plain style",
        to_english:"To English", to_japanese:"To Japanese", translate_paragraph:"Translate selected paragraph", translate_all:"Translate whole paper",
        translation_local:"Local (no API)", translation_online:"Online (optional)", translation_online_free:"Online translation (no API key)", translation_online_help:"Uses online translation without an API key. Text selected for translation is sent to the online translation service. Temporary failures can occur if the service is busy or rate-limited.", translation_model:"Local model", model_fugumt:"FuguMT (Japanese/English, recommended)", model_m2m100:"M2M100 418M (large)", model_opus:"OPUS-MT (lightweight legacy)", glossary:"Glossary", glossary_label:"Academic glossary (one `Japanese = English` pair per line)", glossary_placeholder:"Example: 可変浮力 = variable buoyancy", glossary_help:"Equations, citation numbers, DOI/URLs, and Figure/Table/Equation numbers are protected during translation.", translation_local_help:"FuguMT is the default Japanese/English-specific model. M2M100 is available for comparison. Models download only on first use and then run offline from cache.", local_model_loading:"Loading local translation model (first use may download model files)...",
        section_heading_optional:"Section heading (leave blank if unnecessary)", next_paragraph:"Next paragraph", editor_help:"Select paragraphs, images, and equations individually from the outline. Word OMML equations are preserved.",
        paragraph_text:"Paragraph text...", live_preview:"Live Preview", manual_backup:"Manual Backup", backup_note:"Note...", cancel:"Cancel", save:"Save", restore_history:"Restore from history", close:"Close",
        paragraph:"Paragraph", image:"Image", table:"Table", equation:"Equation", block:"Block", no_heading:"(No heading)", empty_paragraph:"Empty paragraph", new_section:"New section",
        confirm_delete_section:"Delete this section?", confirm_delete_block:"Delete this block?", processing_import:"Parsing paragraphs, images, and equations...",
        import_summary:(s,b,m)=>`${s} sections / ${b} blocks / ${m} equations extracted`, parse_error:"Parse error", choose_file:"Choose a file first", processing:"Processing...", upload_done:"Upload complete", error:"Error",
        insert_here:"Insert here", save_image:"Save image", insert_paragraphs:"Re-parse structure & insert", view_content:"View content", confirm_delete_asset:"Delete this asset? Referenced images may stop displaying.",
        preset_note:"For final submission, use the target journal's official template as the authority.", backup_saved:"Saved", restore_button:"Restore", confirm_restore:"This will replace the current content. Continue?", restored:"Restored",
        translating:"Translating...", translated:"Translation complete", translate_failed:"Translation error", translate_skip:"Image and equation blocks are not translated.", full_translate_confirm:"Replace the whole manuscript with the translated version?",
        new_project_confirm:"Close the current workspace and create a new project? Save the project first if you want to keep it.", new_project_created:"New project created", saving_project:"Saving project...", project_saved:"Project saved", project_save_error:"Project save error", opening_project:"Opening project...", project_opened:"Project opened", project_open_error:"Project open error", material_imported:(s,b,m,i)=>`Material added: ${s} sections / ${b} blocks / ${m} equations / ${i} images`
    }
};
let uiLanguage = localStorage.getItem("paperEditorLanguage") || "ja";
function t(key, ...args) {
    const v = (I18N[uiLanguage] || I18N.ja)[key] ?? I18N.ja[key] ?? key;
    return typeof v === "function" ? v(...args) : v;
}
function applyI18n() {
    document.documentElement.lang = uiLanguage;
    document.querySelectorAll("[data-i18n]").forEach(el => { const k=el.getAttribute("data-i18n"); el.textContent=t(k); });
    document.querySelectorAll("[data-i18n-placeholder]").forEach(el => { el.placeholder=t(el.getAttribute("data-i18n-placeholder")); });
    document.querySelectorAll("[data-i18n-title]").forEach(el => { el.title=t(el.getAttribute("data-i18n-title")); });
    const sel=document.getElementById("uiLanguage"); if(sel) sel.value=uiLanguage;
    const mobilePreset=document.getElementById("mobilePublisherPreset"); if(mobilePreset) mobilePreset.value=paperData?.format_preset || "generic";
    updateBlockMeta(); renderSectionList(); changePublisherPreset(false);
}
function setUILanguage(lang) {
    uiLanguage = lang === "en" ? "en" : "ja";
    localStorage.setItem("paperEditorLanguage", uiLanguage);
    applyI18n(); renderAssetList(); renderPreview();
}

async function apiErrorMessage(response, fallback="Request failed") {
    try {
        const data = await response.clone().json();
        let msg = data.error || fallback;
        if (data.detail) msg += `: ${data.detail}`;
        if (data.error_id) msg += ` [${data.error_id}]`;
        return msg;
    } catch (_) {
        return `${fallback} (HTTP ${response.status})`;
    }
}


function toggleMobileActions(force) {
    const sheet = document.getElementById("mobileActionSheet");
    const backdrop = document.getElementById("mobileSheetBackdrop");
    if (!sheet || !backdrop) return;
    const open = typeof force === "boolean" ? force : !sheet.classList.contains("open");
    sheet.classList.toggle("open", open);
    backdrop.classList.toggle("open", open);
    document.body.classList.toggle("mobile-sheet-open", open);
}

async function newProject(confirmFirst = true) {
    if (confirmFirst && !confirm(t("new_project_confirm"))) return;
    const status = document.getElementById("saveStatus");
    if (status) status.textContent = t("processing");
    try {
        const r = await fetch("/api/project/new", { method: "POST" });
        const data = await r.json();
        if (!r.ok) throw new Error(data.error || "New project failed");
        paperData = normalizeClientData(data.paper);
        activeSectionIndex = 0;
        activeBlockIndex = 0;
        assetsData = [];
        document.getElementById("bibInput").value = "";
        applyPaperToForm();
        renderAssetList();
        closeWelcomeScreen();
        if (window.innerWidth <= 900) switchMainPanel("editor-panel");
        if (status) status.textContent = t("saved");
    } catch (e) {
        if (status) status.textContent = t("error");
        alert(t("error") + ": " + e.message);
    }
}

async function saveProject() {
    syncFormToPaper();
    const status = document.getElementById("saveStatus");
    if (status) status.textContent = t("saving_project");
    try {
        const r = await fetch("/api/project/save", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ paper: paperData, ui_language: uiLanguage })
        });
        if (!r.ok) {
            let msg = "Save failed";
            try { msg = (await r.json()).error || msg; } catch (_) {}
            throw new Error(msg);
        }
        const blob = await r.blob();
        const cd = r.headers.get("content-disposition") || "";
        let filename = "untitled_project.paperproj.zip";
        const utf8 = cd.match(/filename\*=UTF-8''([^;]+)/i);
        const plain = cd.match(/filename="?([^";]+)"?/i);
        if (utf8) filename = decodeURIComponent(utf8[1]);
        else if (plain) filename = plain[1];
        const url = URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url; a.download = filename; document.body.appendChild(a); a.click(); a.remove();
        setTimeout(() => URL.revokeObjectURL(url), 1000);
        if (status) status.textContent = t("saved");
    } catch (e) {
        if (status) status.textContent = t("project_save_error");
        alert(t("project_save_error") + ": " + e.message);
    }
}

function chooseProjectFile() {
    const input = document.getElementById("projectOpenInput");
    if (input) input.click();
}

async function openProjectFile(inputEl) {
    const file = inputEl?.files?.[0];
    if (!file) return;
    const status = document.getElementById("saveStatus");
    if (status) status.textContent = t("opening_project");
    try {
        const fd = new FormData();
        fd.append("file", file);
        const r = await fetch("/api/project/open", { method: "POST", body: fd });
        const data = await r.json();
        if (!r.ok) throw new Error(data.error || "Open failed");
        paperData = normalizeClientData(data.paper);
        activeSectionIndex = 0;
        activeBlockIndex = 0;
        if (data.ui_language === "ja" || data.ui_language === "en") {
            uiLanguage = data.ui_language;
            localStorage.setItem("paperEditorLanguage", uiLanguage);
        }
        assetsData = Array.isArray(data.assets) ? data.assets : [];
        document.getElementById("bibInput").value = paperData.bib_data || "";
        applyI18n();
        applyPaperToForm();
        updateBibData(false);
        renderAssetList();
        closeWelcomeScreen();
        if (window.innerWidth <= 900) switchMainPanel("editor-panel");
        if (status) status.textContent = t("saved");
    } catch (e) {
        if (status) status.textContent = t("project_open_error");
        alert(t("project_open_error") + ": " + e.message);
    } finally {
        if (inputEl) inputEl.value = "";
    }
}

let paperData = { title: "", authors: "", affiliations: "", abstract: "", keywords: "", bib_data: "", format_preset: "generic", translation_provider: "online", translation_model: "", translation_glossary: "", sections: [] };
let activeSectionIndex = 0;
let activeBlockIndex = 0;
let autoSaveTimer = null;
let assetsData = [];
let isTinyMceReady = false;
let suppressEditorEvents = false;

function uid(prefix = "id") {
    return `${prefix}_${Date.now()}_${Math.random().toString(16).slice(2, 8)}`;
}

function makeBlock(type = "paragraph", html = "<p></p>", plainText = "") {
    return { id: uid("blk"), type, html, plain_text: plainText };
}

function stripHtml(html) {
    const div = document.createElement("div");
    div.innerHTML = html || "";
    return (div.textContent || div.innerText || "").trim();
}

function legacyContentToBlocks(content) {
    const wrapper = document.createElement("div");
    wrapper.innerHTML = content || "";
    const children = Array.from(wrapper.children);
    if (!children.length) {
        const text = (wrapper.textContent || "").trim();
        return [makeBlock("paragraph", text ? `<p>${escapeHtml(text)}</p>` : "<p></p>", text)];
    }
    return children.map(el => {
        const html = el.outerHTML;
        const hasImage = !!el.querySelector("img") || el.tagName.toLowerCase() === "img";
        const type = el.tagName.toLowerCase() === "table" ? "table" : (hasImage && !stripHtml(html) ? "image" : "paragraph");
        return makeBlock(type, html, stripHtml(html));
    });
}

function normalizeClientData(data) {
    data = data || {};
    data.title = data.title || "";
    data.authors = data.authors || "";
    data.affiliations = data.affiliations || "";
    data.abstract = data.abstract || "";
    data.keywords = data.keywords || "";
    data.bib_data = data.bib_data || "";
    data.format_preset = data.format_preset || "generic";
    data.translation_provider = "online";
    data.translation_model = "";
    data.translation_glossary = data.translation_glossary || "";
    data.sections = Array.isArray(data.sections) ? data.sections : [];
    if (!data.sections.length) {
        data.sections = [{ id: uid("sec"), title: "", blocks: [makeBlock()], content: "<p></p>" }];
    }
    data.sections.forEach(sec => {
        sec.id = sec.id || uid("sec");
        sec.title = sec.title || "";
        if (!Array.isArray(sec.blocks) || !sec.blocks.length) sec.blocks = legacyContentToBlocks(sec.content || "");
        sec.blocks.forEach(blk => {
            blk.id = blk.id || uid("blk");
            blk.type = blk.type || "paragraph";
            blk.html = blk.html || "";
            blk.plain_text = blk.plain_text ?? stripHtml(blk.html);
        });
        syncSectionContent(sec);
    });
    return data;
}

function syncSectionContent(sec) {
    if (!sec) return;
    sec.content = (sec.blocks || []).map(b => b.html || "").join("\n");
}

function getActiveSection() {
    return paperData.sections[activeSectionIndex] || null;
}

function getActiveBlock() {
    const sec = getActiveSection();
    return sec && sec.blocks ? sec.blocks[activeBlockIndex] : null;
}

document.addEventListener("DOMContentLoaded", () => {
    applyI18n();
    initTinyMCE();
    loadPaperData();
    loadAssetsData();
    if (window.innerWidth <= 900) switchMainPanel("sidebar-panel");
});

function switchMainPanel(panelId) {
    if (window.innerWidth > 900) return;
    document.querySelectorAll(".panel").forEach(p => p.classList.remove("active-panel"));
    const target = document.getElementById(panelId);
    if (target) target.classList.add("active-panel");
    document.querySelectorAll(".mobile-nav button").forEach(b => b.classList.remove("active"));
    const nav = document.getElementById("nav-" + panelId.split("-")[0]);
    if (nav) nav.classList.add("active");
    if (panelId === "editor-panel" && isTinyMceReady) {
        setTimeout(() => tinymce.get("currentSectionContent")?.execCommand("mceRepaint"), 100);
    }
}

function initTinyMCE() {
    const mobileEditor = window.innerWidth <= 700;
    tinymce.init({
        selector: "#currentSectionContent",
        plugins: "image link lists table hr code",
        toolbar: mobileEditor ? "undo redo | bold italic | bullist numlist | link image | table code" : "undo redo | blocks | bold italic | alignleft aligncenter alignright | bullist numlist | link image table | code",
        toolbar_mode: "sliding",
        menubar: false,
        height: mobileEditor ? 360 : 430,
        branding: false,
        promotion: false,
        convert_urls: false,
        relative_urls: false,
        remove_script_host: false,
        extended_valid_elements: "span[class|contenteditable|data-latex|data-omml-b64|data-math-display],div[class],table[class],td,th,tr,tbody",
        content_style: "img{max-width:100%;height:auto;display:block;margin:10px auto;} table{border-collapse:collapse;width:100%;} td,th{border:1px solid #ccc;padding:4px;} .math-inline,.math-display{font-family:Cambria Math,serif;background:#eef2ff;border:1px solid #c7d2fe;border-radius:4px;padding:1px 4px;color:#312e81;} .math-display{display:inline-block;padding:5px 8px;}",
        setup: function(editor) {
            editor.on("init", function() {
                isTinyMceReady = true;
                loadActiveBlockToEditor();
            });
            editor.on("Change KeyUp Input Undo Redo", function() {
                if (suppressEditorEvents) return;
                const block = getActiveBlock();
                const sec = getActiveSection();
                if (!block || !sec) return;
                block.html = editor.getContent();
                block.plain_text = stripHtml(block.html);
                const probe = document.createElement("div");
                probe.innerHTML = block.html;
                if (probe.querySelector("img") && !stripHtml(block.html)) block.type = "image";
                else if (probe.querySelector("table")) block.type = "table";
                else block.type = "paragraph";
                syncSectionContent(sec);
                updateBlockMeta();
                renderSectionList();
                renderPreview();
                triggerAutoSave();
            });
        },
        images_upload_handler: function(blobInfo) {
            return new Promise((resolve, reject) => {
                const formData = new FormData();
                formData.append("file", blobInfo.blob(), blobInfo.filename());
                fetch("/api/upload", { method: "POST", body: formData })
                    .then(res => res.json())
                    .then(data => {
                        if (data.status === "success") {
                            loadAssetsData();
                            resolve("/uploads/" + data.asset.id);
                        } else reject(data.error || "Upload failed");
                    })
                    .catch(err => reject("Error: " + err));
            });
        }
    });
}

function loadPaperData() {
    fetch("/api/paper").then(r => r.json()).then(data => {
        paperData = normalizeClientData(data);
        activeSectionIndex = Math.min(activeSectionIndex, paperData.sections.length - 1);
        activeBlockIndex = 0;

        if (!paperData.title && paperData.sections.length === 1 && paperData.sections[0].id === "sec_init") {
            document.getElementById("welcomeScreen")?.classList.add("active");
        }

        document.getElementById("paperTitle").value = paperData.title;
        document.getElementById("paperAuthors").value = paperData.authors;
        document.getElementById("paperAffiliations").value = paperData.affiliations;
        document.getElementById("paperAbstract").value = paperData.abstract;
        if(document.getElementById("paperKeywords")) document.getElementById("paperKeywords").value = paperData.keywords || "";
        document.getElementById("bibInput").value = paperData.bib_data;
        const presetSel = document.getElementById("publisherPreset");
        if (presetSel) presetSel.value = paperData.format_preset || "generic";
        updateBibData(false);
        renderSectionList();
        renderPreview();
        loadActiveBlockToEditor();
    });
}

function handleInitialImport(inputEl) {
    const file = inputEl.files[0];
    if (!file) return;
    const formData = new FormData();
    formData.append("file", file);
    const status = document.getElementById("welcomeStatus");
    if (status) status.innerText = t("processing_import");

    fetch("/api/import_document", { method: "POST", body: formData })
        .then(async r => {
            const data = await r.json();
            if (!r.ok) throw new Error(data.error || "Import failed");
            return data;
        })
        .then(data => {
            paperData = normalizeClientData(data.paper);
            document.getElementById("paperTitle").value = paperData.title;
            document.getElementById("paperAuthors").value = paperData.authors || "";
            document.getElementById("paperAffiliations").value = paperData.affiliations || "";
            document.getElementById("paperAbstract").value = paperData.abstract || "";
    if(document.getElementById("paperKeywords")) document.getElementById("paperKeywords").value = paperData.keywords || "";
    const tg=document.getElementById("translationGlossary"); if(tg) tg.value=paperData.translation_glossary || "";
            activeSectionIndex = 0;
            activeBlockIndex = 0;
            renderSectionList();
            loadActiveBlockToEditor();
            renderPreview();
            closeWelcomeScreen();
            loadAssetsData();
            const count = paperData.sections.reduce((n, s) => n + (s.blocks?.length || 0), 0);
            const mathCount = paperData.sections.reduce((n, s) => n + (s.blocks || []).reduce((m,b) => m + ((b.html || "").match(/data-omml-b64=/g) || []).length, 0), 0);
            if (status) status.innerText = t("import_summary", paperData.sections.length, count, mathCount);
            if (window.innerWidth <= 900) switchMainPanel("editor-panel");
        })
        .catch(err => {
            if (status) status.innerText = t("parse_error") + ": " + err.message;
        });
}

function closeWelcomeScreen() {
    document.getElementById("welcomeScreen")?.classList.remove("active");
}

function loadActiveBlockToEditor() {
    const sec = getActiveSection();
    if (!sec) return;
    if (!sec.blocks?.length) sec.blocks = [makeBlock()];
    activeBlockIndex = Math.min(activeBlockIndex, sec.blocks.length - 1);
    const block = getActiveBlock();
    document.getElementById("currentSectionTitle").value = sec.title || "";
    updateBlockMeta();
    if (isTinyMceReady && tinymce.get("currentSectionContent")) {
        suppressEditorEvents = true;
        tinymce.get("currentSectionContent").setContent(block?.html || "<p></p>");
        setTimeout(() => suppressEditorEvents = false, 0);
    }
}

function updateBlockMeta() {
    const sec = getActiveSection();
    const block = getActiveBlock();
    const el = document.getElementById("currentBlockMeta");
    if (!el || !sec || !block) return;
    const labels = { paragraph: t("paragraph"), image: t("image"), table: t("table"), equation: t("equation") };
    el.textContent = `${labels[block.type] || t("block")} ${activeBlockIndex + 1} / ${sec.blocks.length}`;
}

function loadAssetsData() {
    return fetch("/api/assets").then(r => r.json()).then(data => {
        assetsData = data || [];
        renderAssetList();
        return assetsData;
    });
}

function renderSectionList() {
    const list = document.getElementById("sectionList");
    if (!list) return;
    list.innerHTML = "";

    paperData.sections.forEach((sec, sIdx) => {
        const group = document.createElement("div");
        group.className = "section-group" + (sIdx === activeSectionIndex ? " active" : "");

        const header = document.createElement("div");
        header.className = "section-header-row";
        header.onclick = () => selectSection(sIdx);

        const title = document.createElement("span");
        title.className = "section-title-text";
        title.textContent = sec.title || t("no_heading");
        header.appendChild(title);

        const actions = document.createElement("div");
        actions.className = "section-actions";
        [
            ["↑", () => moveSection(sIdx, -1), "btn-secondary"],
            ["↓", () => moveSection(sIdx, 1), "btn-secondary"],
            ["×", () => deleteSection(sIdx), "btn-danger"]
        ].forEach(([label, fn, cls]) => {
            const b = document.createElement("button");
            b.className = `btn ${cls} btn-sm`;
            b.textContent = label;
            b.onclick = e => { e.stopPropagation(); fn(); };
            actions.appendChild(b);
        });
        header.appendChild(actions);
        group.appendChild(header);

        if (sIdx === activeSectionIndex) {
            const blockList = document.createElement("div");
            blockList.className = "paragraph-list";
            (sec.blocks || []).forEach((blk, bIdx) => {
                const row = document.createElement("div");
                row.className = "paragraph-item" + (bIdx === activeBlockIndex ? " active" : "");
                row.onclick = () => selectBlock(bIdx);

                const icon = blk.type === "image" ? "🖼" : (blk.type === "table" ? "▦" : (blk.type === "equation" ? "∑" : "¶"));
                const preview = blk.type === "image" ? t("image") : (blk.type === "equation" ? ((blk.html.match(/data-latex="([^"]*)"/) || [])[1] || stripHtml(blk.html)).slice(0, 50) : (stripHtml(blk.html).slice(0, 50) || t("empty_paragraph")));
                const text = document.createElement("span");
                text.className = "paragraph-preview";
                text.textContent = `${icon} ${bIdx + 1}. ${preview}`;
                row.appendChild(text);

                const miniActions = document.createElement("div");
                miniActions.className = "paragraph-actions";
                [["↑", -1], ["↓", 1]].forEach(([label, dir]) => {
                    const b = document.createElement("button");
                    b.className = "mini-btn";
                    b.textContent = label;
                    b.onclick = e => { e.stopPropagation(); moveBlock(bIdx, dir); };
                    miniActions.appendChild(b);
                });
                const del = document.createElement("button");
                del.className = "mini-btn danger";
                del.textContent = "×";
                del.onclick = e => { e.stopPropagation(); deleteBlock(bIdx); };
                miniActions.appendChild(del);
                row.appendChild(miniActions);
                blockList.appendChild(row);
            });
            group.appendChild(blockList);
        }

        list.appendChild(group);
    });
}

function selectSection(idx) {
    activeSectionIndex = idx;
    activeBlockIndex = 0;
    renderSectionList();
    loadActiveBlockToEditor();
    switchMainPanel("editor-panel");
}

function selectBlock(idx) {
    activeBlockIndex = idx;
    renderSectionList();
    loadActiveBlockToEditor();
    switchMainPanel("editor-panel");
}

function addSection(title = null) {
    paperData.sections.push({
        id: uid("sec"),
        title: title || "",
        blocks: [makeBlock()],
        content: "<p></p>"
    });
    activeSectionIndex = paperData.sections.length - 1;
    activeBlockIndex = 0;
    renderSectionList();
    loadActiveBlockToEditor();
    triggerAutoSave();
    switchMainPanel("editor-panel");
}

function deleteSection(idx) {
    if (!confirm(t("confirm_delete_section"))) return;
    paperData.sections.splice(idx, 1);
    if (!paperData.sections.length) addSection("");
    activeSectionIndex = Math.max(0, Math.min(activeSectionIndex, paperData.sections.length - 1));
    activeBlockIndex = 0;
    renderSectionList();
    loadActiveBlockToEditor();
    renderPreview();
    triggerAutoSave();
}

function moveSection(idx, dir) {
    const t = idx + dir;
    if (t < 0 || t >= paperData.sections.length) return;
    [paperData.sections[idx], paperData.sections[t]] = [paperData.sections[t], paperData.sections[idx]];
    if (activeSectionIndex === idx) activeSectionIndex = t;
    else if (activeSectionIndex === t) activeSectionIndex = idx;
    renderSectionList();
    renderPreview();
    triggerAutoSave();
}

function addParagraph(afterCurrent = true) {
    const sec = getActiveSection();
    if (!sec) return;
    const insertAt = afterCurrent ? activeBlockIndex + 1 : sec.blocks.length;
    sec.blocks.splice(insertAt, 0, makeBlock());
    activeBlockIndex = insertAt;
    syncSectionContent(sec);
    renderSectionList();
    loadActiveBlockToEditor();
    renderPreview();
    triggerAutoSave();
}

function addImageBlockFromAsset(asset) {
    const sec = getActiveSection();
    if (!sec || !asset) return;
    const html = `<div class="figure-block"><img src="/uploads/${escapeAttr(asset.id)}" alt="${escapeAttr(asset.original_name || "image")}" class="doc-image"></div>`;
    const insertAt = activeBlockIndex + 1;
    sec.blocks.splice(insertAt, 0, makeBlock("image", html, ""));
    activeBlockIndex = insertAt;
    syncSectionContent(sec);
    renderSectionList();
    loadActiveBlockToEditor();
    renderPreview();
    triggerAutoSave();
    switchMainPanel("editor-panel");
}

function deleteBlock(idx) {
    const sec = getActiveSection();
    if (!sec || !confirm(t("confirm_delete_block"))) return;
    sec.blocks.splice(idx, 1);
    if (!sec.blocks.length) sec.blocks.push(makeBlock());
    activeBlockIndex = Math.max(0, Math.min(activeBlockIndex, sec.blocks.length - 1));
    syncSectionContent(sec);
    renderSectionList();
    loadActiveBlockToEditor();
    renderPreview();
    triggerAutoSave();
}

function moveBlock(idx, dir) {
    const sec = getActiveSection();
    if (!sec) return;
    const t = idx + dir;
    if (t < 0 || t >= sec.blocks.length) return;
    [sec.blocks[idx], sec.blocks[t]] = [sec.blocks[t], sec.blocks[idx]];
    if (activeBlockIndex === idx) activeBlockIndex = t;
    else if (activeBlockIndex === t) activeBlockIndex = idx;
    syncSectionContent(sec);
    renderSectionList();
    loadActiveBlockToEditor();
    renderPreview();
    triggerAutoSave();
}

function updateCurrentSectionTitle() {
    const sec = getActiveSection();
    if (!sec) return;
    sec.title = document.getElementById("currentSectionTitle").value;
    renderSectionList();
    renderPreview();
    triggerAutoSave();
}

function triggerAutoSave() {
    const status = document.getElementById("saveStatus");
    if (status) status.innerText = t("saving");
    paperData.title = document.getElementById("paperTitle")?.value || "";
    paperData.authors = document.getElementById("paperAuthors")?.value || "";
    paperData.affiliations = document.getElementById("paperAffiliations")?.value || "";
    paperData.abstract = document.getElementById("paperAbstract")?.value || "";
    paperData.keywords = document.getElementById("paperKeywords")?.value || "";
    paperData.translation_provider = "online";
    paperData.translation_model = "";
    paperData.translation_glossary = document.getElementById("translationGlossary")?.value || "";
    paperData.format_preset = document.getElementById("publisherPreset")?.value || paperData.format_preset || "generic";
    paperData.sections.forEach(syncSectionContent);
    renderPreview();
    clearTimeout(autoSaveTimer);
    autoSaveTimer = setTimeout(() => {
        fetch("/api/paper", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify(paperData)
        }).then(async r => {
            if (!r.ok) throw new Error(await apiErrorMessage(r, "Save failed"));
            const data = await r.json();
            // The server stores the manuscript in academic Japanese style. Keep the local structure intact
            // while editing; a reload/export always uses the normalized server version.
            if (data?.paper) paperData = normalizeClientData(data.paper);
            if (status) status.innerText = t("saved");
        }).catch(err => {
            if (status) status.innerText = `${t("error")}: ${err.message}`;
            console.error("Auto-save failed", err);
        });
    }, 700);
}

// --- Assets Logic ---
async function uploadAsset() {
    const fileInput = document.getElementById("fileInput");
    if (!fileInput?.files?.length) return alert(t("choose_file"));
    const file = fileInput.files[0];
    const formData = new FormData();
    formData.append("file", file);
    const status = document.getElementById("uploadStatus");
    if (status) status.innerText = t("processing_import");
    try {
        syncFormToPaper();
        const r = await fetch("/api/import_material", { method:"POST", body:formData });
        const data = await r.json();
        if (!r.ok) throw new Error(data.error || "Material import failed");
        insertImportedSections(data.sections || []);
        await loadAssetsData();
        const q = data.summary || {};
        if (status) status.innerText = t("material_imported", q.sections || 0, q.blocks || 0, q.equations || 0, q.images || 0);
        fileInput.value = "";
        if (window.innerWidth <= 900) switchMainPanel("editor-panel");
    } catch (e) {
        if (status) status.innerText = t("parse_error") + ": " + e.message;
    }
}

function insertImportedSections(sections) {
    sections = Array.isArray(sections) ? sections.map(s => ({...s, blocks:Array.isArray(s.blocks)?s.blocks:[]})) : [];
    if (!sections.length) return;
    const ed = tinymce.get("currentSectionContent");
    const activeSec = getActiveSection();
    const activeBlk = getActiveBlock();
    if (ed && activeSec && activeBlk) {
        activeBlk.html = ed.getContent(); activeBlk.plain_text = stripHtml(activeBlk.html); syncSectionContent(activeSec);
    }
    let secInsertAt = activeSectionIndex + 1;
    const first = sections[0];
    // A title-less leading fragment belongs exactly at the current block position.
    if (activeSec && !(first.title || "").trim()) {
        const incoming = (first.blocks || []).filter(Boolean);
        const isEmptyCurrent = activeBlk && !stripHtml(activeBlk.html || "") && !/data-omml-b64=|<img|<table/i.test(activeBlk.html || "");
        if (isEmptyCurrent && incoming.length) {
            activeSec.blocks.splice(activeBlockIndex, 1, ...incoming);
            activeBlockIndex = activeBlockIndex + incoming.length - 1;
        } else if (incoming.length) {
            activeSec.blocks.splice(activeBlockIndex + 1, 0, ...incoming);
            activeBlockIndex += incoming.length;
        }
        syncSectionContent(activeSec);
        sections = sections.slice(1);
    }
    if (sections.length) {
        paperData.sections.splice(secInsertAt, 0, ...sections);
    }
    renderSectionList(); loadActiveBlockToEditor(); renderPreview(); triggerAutoSave();
}

async function reparseAssetAndInsert(assetId) {
    const status = document.getElementById("uploadStatus");
    if (status) status.innerText = t("processing_import");
    try {
        const r = await fetch(`/api/assets/${encodeURIComponent(assetId)}/parse`, {method:"POST"});
        const data = await r.json();
        if (!r.ok) throw new Error(data.error || "Parse failed");
        insertImportedSections(data.sections || []);
        await loadAssetsData();
        const q=data.summary||{};
        if(status) status.innerText=t("material_imported",q.sections||0,q.blocks||0,q.equations||0,q.images||0);
    } catch(e) { if(status) status.innerText=t("parse_error")+": "+e.message; }
}

function handleDragStart(e, textToInsert) {
    e.dataTransfer.setData("text/html", textToInsert);
    e.dataTransfer.setData("text/plain", stripHtml(textToInsert));
}

function insertAsset(assetId) {
    const asset = assetsData.find(a => a.id === assetId);
    if (!asset) return;
    if (asset.is_image) {
        addImageBlockFromAsset(asset);
    } else {
        const sec = getActiveSection();
        if (!sec) return;
        const lines = (asset.extracted_text || "").split(/\n+/).map(s => s.trim()).filter(Boolean);
        const newBlocks = lines.slice(0, 1000).map(line => makeBlock("paragraph", `<p>${escapeHtml(line)}</p>`, line));
        const insertAt = activeBlockIndex + 1;
        sec.blocks.splice(insertAt, 0, ...(newBlocks.length ? newBlocks : [makeBlock()]));
        activeBlockIndex = insertAt;
        syncSectionContent(sec);
        renderSectionList();
        loadActiveBlockToEditor();
        renderPreview();
        triggerAutoSave();
        switchMainPanel("editor-panel");
    }
}

function renderAssetList() {
    const list = document.getElementById("assetList");
    if (!list) return;
    list.innerHTML = "";
    assetsData.forEach(asset => {
        const item = document.createElement("div");
        item.className = "asset-card";

        const head = document.createElement("div");
        head.className = "asset-card-head";
        const name = document.createElement("div");
        name.className = "asset-name";
        name.title = asset.original_name || asset.id;
        name.textContent = asset.original_name || asset.id;
        head.appendChild(name);
        const del = document.createElement("button");
        del.className = "btn btn-danger btn-sm";
        del.textContent = "×";
        del.onclick = () => deleteAsset(asset.id);
        head.appendChild(del);
        item.appendChild(head);

        if (asset.is_image) {
            const img = document.createElement("img");
            img.src = `/uploads/${asset.id}`;
            img.alt = asset.original_name || "image";
            img.className = "asset-thumb";
            img.loading = "lazy";
            item.appendChild(img);

            const actions = document.createElement("div");
            actions.className = "action-buttons";
            const insert = document.createElement("button");
            insert.className = "btn btn-primary btn-sm";
            insert.textContent = t("insert_here");
            insert.onclick = () => insertAsset(asset.id);
            actions.appendChild(insert);
            const dl = document.createElement("a");
            dl.className = "btn btn-secondary btn-sm";
            dl.href = `/api/assets/${encodeURIComponent(asset.id)}/download`;
            dl.textContent = t("save_image");
            actions.appendChild(dl);
            item.appendChild(actions);

            item.draggable = true;
            item.ondragstart = e => handleDragStart(e, `<img src="/uploads/${asset.id}" alt="${escapeAttr(asset.original_name || "image")}">`);
        } else {
            const actions = document.createElement("div");
            actions.className = "action-buttons";
            const insert = document.createElement("button");
            insert.className = "btn btn-primary btn-sm";
            insert.textContent = t("insert_paragraphs");
            insert.onclick = () => reparseAssetAndInsert(asset.id);
            actions.appendChild(insert);
            const toggle = document.createElement("button");
            toggle.className = "btn btn-secondary btn-sm";
            toggle.textContent = t("view_content");
            toggle.onclick = () => toggleAssetText(asset.id);
            actions.appendChild(toggle);
            item.appendChild(actions);

            const preview = document.createElement("div");
            preview.className = "asset-text-preview";
            preview.id = `txt_${asset.id}`;
            preview.style.display = "none";
            preview.textContent = asset.extracted_text || "";
            item.appendChild(preview);
        }
        list.appendChild(item);
    });
}

function toggleAssetText(id) {
    const el = document.getElementById("txt_" + id);
    if (el) el.style.display = el.style.display === "block" ? "none" : "block";
}

function deleteAsset(id) {
    if (!confirm(t("confirm_delete_asset"))) return;
    fetch("/api/assets/" + encodeURIComponent(id), { method: "DELETE" }).then(() => loadAssetsData());
}

function updateBibData(shouldSave = true) {
    paperData.bib_data = document.getElementById("bibInput")?.value || "";
    const keyRegex = /@\w+\s*{\s*([^,\s]+)/g;
    const badges = document.getElementById("citeKeyBadges");
    if (badges) badges.innerHTML = "";
    let match;
    while ((match = keyRegex.exec(paperData.bib_data)) !== null) {
        const k = match[1];
        const b = document.createElement("span");
        b.className = "cite-badge";
        b.innerText = `[@${k}]`;
        b.onclick = () => {
            tinymce.get("currentSectionContent")?.insertContent(`[@${k}]`);
            switchMainPanel("editor-panel");
        };
        badges?.appendChild(b);
    }
    if (shouldSave) triggerAutoSave();
}

function renderPreview() {
    const title = document.getElementById("prevTitle");
    if (!title) return;
    title.innerText = paperData.title || "";
    const sheet = document.getElementById("previewSheet");
    if (sheet) {
        sheet.className = `paper-sheet preset-${paperData.format_preset || "generic"}`;
        sheet.dataset.preset = paperData.format_preset || "generic";
    }
    document.getElementById("prevAuthors").innerText = paperData.authors || "";
    document.getElementById("prevAffiliations").innerText = paperData.affiliations || "";
    const box = document.getElementById("prevAbstractBox");
    if (paperData.abstract) {
        box.style.display = "block";
        document.getElementById("prevAbstract").innerText = paperData.abstract;
    } else box.style.display = "none";

    const sc = document.getElementById("prevSections");
    sc.innerHTML = "";
    let cIdx = 1, cMap = {};
    paperData.sections.forEach(sec => {
        const sectionWrap = document.createElement("section");
        if (sec.title) {
            const h = document.createElement("div");
            h.className = "preview-section-title";
            h.textContent = sec.title;
            sectionWrap.appendChild(h);
        }

        (sec.blocks || []).forEach((blk, idx) => {
            let htmlText = blk.html || "";
            htmlText = htmlText.replace(/\[@([^\]]+)\]/g, (m, k) => {
                if (!cMap[k]) cMap[k] = cIdx++;
                return `<strong>[${cMap[k]}]</strong>`;
            });
            const d = document.createElement("div");
            d.className = `preview-block preview-block-${blk.type || "paragraph"}`;
            d.dataset.blockIndex = idx;
            d.innerHTML = htmlText;
            sectionWrap.appendChild(d);
        });
        sc.appendChild(sectionWrap);
    });
    if (window.MathJax && MathJax.typesetPromise) {
        try { MathJax.typesetPromise([sc]); } catch(e) {}
    }
}

function escapeHtml(str) {
    return String(str ?? "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&#039;");
}
function escapeAttr(str) { return escapeHtml(str); }


function changePublisherPreset(shouldSave = true, sourceValue = null) {
    const sel = document.getElementById("publisherPreset");
    const mobileSel = document.getElementById("mobilePublisherPreset");
    const value = sourceValue || sel?.value || mobileSel?.value || paperData.format_preset || "generic";
    paperData.format_preset = value;
    if (sel && sel.value !== value) sel.value = value;
    if (mobileSel && mobileSel.value !== value) mobileSel.value = value;
    const note = document.getElementById("presetNote");
    const labels = { generic: uiLanguage === "ja" ? "汎用学術書式" : "Generic academic", ieee: "IEEE", elsevier: "Elsevier", springer_nature: "Springer Nature", mdpi: "MDPI", wiley: "Wiley", taylor_francis: "Taylor & Francis" };
    if (note) note.textContent = `${labels[paperData.format_preset] || "Publisher preset"} — ${t("preset_note")}`;
    renderPreview();
    if (shouldSave) triggerAutoSave();
}

function syncFormToPaper() {
    paperData.title = document.getElementById("paperTitle")?.value || "";
    paperData.authors = document.getElementById("paperAuthors")?.value || "";
    paperData.affiliations = document.getElementById("paperAffiliations")?.value || "";
    paperData.abstract = document.getElementById("paperAbstract")?.value || "";
    paperData.keywords = document.getElementById("paperKeywords")?.value || "";
    paperData.translation_provider = "online";
    paperData.translation_model = "";
    paperData.translation_glossary = document.getElementById("translationGlossary")?.value || "";
    const sec = getActiveSection(), blk = getActiveBlock(), ed = tinymce.get("currentSectionContent");
    if (sec && blk && ed) { blk.html = ed.getContent(); blk.plain_text = stripHtml(blk.html); syncSectionContent(sec); }
}
function applyPaperToForm() {
    document.getElementById("paperTitle").value = paperData.title || "";
    document.getElementById("paperAuthors").value = paperData.authors || "";
    document.getElementById("paperAffiliations").value = paperData.affiliations || "";
    document.getElementById("paperAbstract").value = paperData.abstract || "";
    if(document.getElementById("paperKeywords")) document.getElementById("paperKeywords").value = paperData.keywords || "";
    const tg=document.getElementById("translationGlossary"); if(tg) tg.value=paperData.translation_glossary || "";
    const preset=document.getElementById("publisherPreset"); if(preset) preset.value=paperData.format_preset || "generic";
    renderSectionList(); loadActiveBlockToEditor(); renderPreview();
}
function updateTranslationSettings() {
    paperData.translation_provider = "online";
    paperData.translation_model = "";
    paperData.translation_glossary = document.getElementById("translationGlossary")?.value || "";
    triggerAutoSave();
}
function toggleTranslationGlossary() {
    const panel=document.getElementById("translationGlossaryPanel");
    if(panel) panel.style.display = panel.style.display === "none" ? "block" : "none";
}
function currentTranslationOptions() {
    return {
        provider: "online",
        local_model: "",
        glossary: document.getElementById("translationGlossary")?.value || paperData.translation_glossary || ""
    };
}

async function translateCurrentBlock() {
    const status=document.getElementById("translateStatus");
    const blk=getActiveBlock(); if(!blk) return;
    if(["image","equation"].includes(blk.type)) { status.textContent=t("translate_skip"); return; }
    syncFormToPaper();
    status.textContent=t("translating");
    try {
        const target=document.getElementById("translateTarget")?.value || "en";
        const opts=currentTranslationOptions();
        paperData.translation_provider=opts.provider; paperData.translation_model=opts.local_model; paperData.translation_glossary=opts.glossary;
        
        const r=await fetch("/api/translate",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({html:blk.html,target,provider:opts.provider,local_model:opts.local_model,glossary:opts.glossary})});
        const data=await r.json(); if(!r.ok) throw new Error(data.error || "Translation failed");
        blk.html=data.html; blk.plain_text=stripHtml(blk.html); syncSectionContent(getActiveSection());
        loadActiveBlockToEditor(); renderSectionList(); renderPreview(); triggerAutoSave(); status.textContent=t("translated");
    } catch(e) { status.textContent=t("translate_failed")+": "+e.message; }
}
async function translateWholePaper() {
    if(!confirm(t("full_translate_confirm"))) return;
    const status=document.getElementById("translateStatus"); syncFormToPaper(); status.textContent=t("translating");
    try {
        const target=document.getElementById("translateTarget")?.value || "en";
        const opts=currentTranslationOptions();
        paperData.translation_provider=opts.provider; paperData.translation_model=opts.local_model; paperData.translation_glossary=opts.glossary;
        
        const r=await fetch("/api/translate",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({paper:paperData,target,skip_references:true,provider:opts.provider,local_model:opts.local_model,glossary:opts.glossary})});
        const data=await r.json(); if(!r.ok) throw new Error(data.error || "Translation failed");
        paperData=normalizeClientData(data.paper); activeSectionIndex=0; activeBlockIndex=0; applyPaperToForm(); triggerAutoSave(); status.textContent=t("translated");
    } catch(e) { status.textContent=t("translate_failed")+": "+e.message; }
}

function switchSidebarTab(t) {
    document.querySelectorAll(".tab-btn").forEach(b => b.classList.remove("active"));
    document.querySelectorAll(".tab-content").forEach(c => c.style.display = "none");
    const btn = document.querySelector(`.tab-btn[onclick="switchSidebarTab('${t}')"]`);
    if (btn) btn.classList.add("active");
    const tab = document.getElementById(t + "Tab");
    if (tab) tab.style.display = "block";
}

function openModal(id) { document.getElementById(id)?.classList.add("active"); }
function closeModal(id) { document.getElementById(id)?.classList.remove("active"); }
function submitSnapshot() {
    fetch("/api/snapshot", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ note: document.getElementById("snapshotNote").value })
    }).then(() => {
        closeModal("snapshotModal");
        alert(t("backup_saved"));
        document.getElementById("snapshotNote").value = "";
    });
}
function openRestoreModal() {
    fetch("/api/snapshots").then(r => r.json()).then(snaps => {
        const l = document.getElementById("snapshotList");
        l.innerHTML = "";
        snaps.forEach(s => {
            const d = document.createElement("div");
            d.className = "list-item";
            d.innerHTML = `<div><strong>${escapeHtml(s.timestamp)}</strong> - ${escapeHtml(s.note)}</div><button class="btn btn-primary btn-sm">${t("restore_button")}</button>`;
            d.querySelector("button").onclick = () => restoreSnapshot(s.filename);
            l.appendChild(d);
        });
        openModal("restoreModal");
    });
}
function restoreSnapshot(fname) {
    if (!confirm(t("confirm_restore"))) return;
    fetch("/api/snapshot/restore", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ filename: fname })
    }).then(r => r.json()).then(data => {
        paperData = normalizeClientData(data.data);
        activeSectionIndex = 0;
        activeBlockIndex = 0;
        loadPaperData();
        closeModal("restoreModal");
        alert(t("restored"));
    });
}
async function handleExport(sel) {
    const val = sel.value;
    if (!val) return;
    sel.disabled = true;
    try {
        syncFormToPaper();
        paperData.sections.forEach(syncSectionContent);
        const r = await fetch(`/api/export/${val}`, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ paper: paperData })
        });
        if (!r.ok) throw new Error(await apiErrorMessage(r, "Export failed"));
        const blob = await r.blob();
        const cd = r.headers.get("Content-Disposition") || "";
        const m = cd.match(/filename\*?=(?:UTF-8''|")?([^";]+)/i);
        let filename = m ? decodeURIComponent(m[1].replace(/"/g, "")) : `paper_${paperData.format_preset || "generic"}.${val === "pdf_print" ? "html" : val}`;
        const url = URL.createObjectURL(blob);
        if (val === "pdf_print" || val === "html") {
            const win = window.open(url, "_blank");
            if (!win) window.location.href = url;
            setTimeout(() => URL.revokeObjectURL(url), 60000);
        } else {
            const a = document.createElement("a");
            a.href = url; a.download = filename; a.style.display = "none";
            document.body.appendChild(a); a.click(); a.remove();
            setTimeout(() => URL.revokeObjectURL(url), 5000);
        }
    } catch (e) {
        alert((uiLanguage === "ja" ? "エクスポートに失敗しました: " : "Export failed: ") + e.message);
    } finally {
        sel.disabled = false;
        sel.value = "";
    }
}
