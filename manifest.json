<!DOCTYPE html>
<html lang="fr">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0, viewport-fit=cover">
    <title>Suivi des Lettres de Change - Tunisie (TND)</title>
    <link rel="manifest" href="manifest.json">
    <meta name="theme-color" content="#2c7da0">
    <meta name="apple-mobile-web-app-capable" content="yes">
    <meta name="apple-mobile-web-app-status-bar-style" content="black-translucent">
    <meta name="apple-mobile-web-app-title" content="Lettres Change">
    <link rel="apple-touch-icon" href="icons/icon-192x192.png">
    <style>
        * { box-sizing: border-box; font-family: 'Segoe UI', Roboto, Arial, sans-serif; }
        body { background: #e9f0f5; margin: 0; padding: 20px; }
        .container { max-width: 1400px; margin: auto; background: white; border-radius: 24px; box-shadow: 0 10px 30px rgba(0,0,0,0.1); padding: 20px; }
        .client-header { background: #f0f6fa; padding: 15px 20px; border-radius: 20px; margin-bottom: 20px; display: flex; align-items: center; gap: 20px; flex-wrap: wrap; border: 1px solid #cbd5e1; }
        .logo-area { max-width: 150px; max-height: 80px; }
        .logo-area img { max-width: 100%; max-height: 80px; object-fit: contain; }
        .client-info { flex: 1; }
        .client-info h2 { margin: 0 0 5px; color: #1e4a6e; }
        .client-info p { margin: 0; font-size: 14px; }
        h1 { margin-top: 0; color: #1e4a6e; }
        .tabs { display: flex; gap: 10px; margin-bottom: 20px; border-bottom: 2px solid #cbd5e1; flex-wrap: wrap; }
        .tab-btn { background: none; border: none; padding: 10px 20px; font-size: 16px; cursor: pointer; border-radius: 20px 20px 0 0; }
        .tab-btn.active { background: #2c7da0; color: white; }
        .tab-content { display: none; }
        .tab-content.active { display: block; }
        .dashboard { display: grid; grid-template-columns: repeat(auto-fit, minmax(220px, 1fr)); gap: 15px; margin-bottom: 25px; }
        .card { background: #f8fafc; border-radius: 20px; padding: 15px; border-left: 5px solid #2c7da0; }
        .card h3 { margin: 0 0 8px 0; font-size: 14px; }
        .card .number { font-size: 28px; font-weight: bold; color: #0f3b4c; }
        .filters { background: #f1f5f9; padding: 15px; border-radius: 20px; margin-bottom: 20px; display: flex; flex-wrap: wrap; gap: 12px; align-items: flex-end; }
        .filters div { display: flex; flex-direction: column; gap: 4px; }
        input, select, button { padding: 8px 12px; border-radius: 12px; border: 1px solid #cbd5e1; }
        button { background: #2c7da0; color: white; border: none; cursor: pointer; font-weight: bold; }
        button:hover { background: #1f5e7a; }
        .btn-small { padding: 4px 8px; font-size: 12px; margin-left: 5px; }
        .btn-edit { background: #f39c12; }
        .btn-delete { background: #e74c3c; }
        .btn-export { background: #27ae60; }
        .btn-warning { background: #e67e22; }
        .btn-save { background: #2980b9; }
        .btn-import { background: #8e44ad; }
        table { width: 100%; border-collapse: collapse; }
        th, td { text-align: left; padding: 8px; border-bottom: 1px solid #e2e8f0; }
        th { background: #e6f0f5; }
        .modal { display: none; position: fixed; top:0; left:0; width:100%; height:100%; background: rgba(0,0,0,0.5); justify-content: center; align-items: center; z-index: 1000; }
        .modal-content { background: white; padding: 20px; border-radius: 24px; width: 90%; max-width: 500px; }
        .form-group { margin-bottom: 12px; display: flex; flex-direction: column; }
        .inline-add { display: flex; gap: 8px; align-items: center; margin-top: 5px; }
        .inline-add button { background: #27ae60; }
        .list-item { display: flex; align-items: center; justify-content: space-between; padding: 6px 0; border-bottom: 1px solid #e2e8f0; }
        .list-item span { flex: 1; }
        .list-actions button { margin-left: 5px; }
        .date-input { width: 130px; }
        small { font-size: 10px; color: #666; }
        .report-filters { background: #f8fafc; padding: 15px; border-radius: 20px; margin-bottom: 20px; display: flex; flex-wrap: wrap; gap: 15px; align-items: flex-end; }
        .report-table { overflow-x: auto; margin-top: 20px; }
        .export-buttons { margin: 15px 0; display: flex; gap: 10px; flex-wrap: wrap; }
        .status-section { background: #fff; border: 1px solid #cbd5e1; border-radius: 20px; padding: 15px; margin-top: 25px; }
        .status-section h3 { margin-top: 0; }
        .status-buttons { display: flex; gap: 10px; flex-wrap: wrap; margin-bottom: 15px; }
        .status-filters { margin-bottom: 15px; display: flex; align-items: flex-end; gap: 15px; flex-wrap: wrap; }
    </style>
    <script src="https://cdn.sheetjs.com/xlsx-0.20.2/package/dist/xlsx.full.min.js"></script>
</head>
<body>
<div class="container">
    <div id="clientHeader" class="client-header">
        <div id="logoContainer" class="logo-area"></div>
        <div class="client-info">
            <h2 id="clientNom">Ma Société</h2>
            <p id="clientAdresse">Adresse : ...</p>
            <p id="clientActivite">Activité : ...</p>
        </div>
    </div>

    <div class="tabs">
        <button class="tab-btn active" data-tab="tab1">📋 Lettres de change</button>
        <button class="tab-btn" data-tab="tab2">📇 Gestion fournisseurs & banques</button>
        <button class="tab-btn" data-tab="tab3">📊 Rapports & États</button>
        <button class="tab-btn" data-tab="tab4">⚙️ Personnalisation</button>
    </div>

    <!-- TAB 1 -->
    <div id="tab1" class="tab-content active">
        <div class="dashboard">
            <div class="card"><h3>📅 Ce mois</h3><div class="number" id="moisCount">0</div><div class="sub">Montant: <span id="moisMontant">0</span> TND</div></div>
            <div class="card"><h3>📆 Cette année</h3><div class="number" id="anneeCount">0</div><div class="sub">Montant: <span id="anneeMontant">0</span> TND</div></div>
            <div class="card"><h3>🗓️ Aujourd'hui</h3><div class="number" id="jourCount">0</div><div class="sub">échéances</div></div>
            <div class="card"><h3>🏦 Fournisseurs</h3><div class="number" id="fournisseurCount">0</div><div class="sub">actifs</div></div>
        </div>
        <div class="filters">
            <div><label>Fournisseur</label><select id="filterFournisseur"><option value="">Tous</option></select></div>
            <div><label>Banque</label><select id="filterBanque"><option value="">Toutes</option></select></div>
            <div><label>Échéance du</label><input type="text" id="filterDateDebut" placeholder="jj/mm/aaaa" class="date-input"></div>
            <div><label>au</label><input type="text" id="filterDateFin" placeholder="jj/mm/aaaa" class="date-input"></div>
            <div><label>Statut</label><select id="filterStatut"><option value="">Tous</option><option>En attente</option><option>Liquidée</option></select></div>
            <div><button id="resetFilters">Reset</button></div>
            <div><button id="openAddBtn">+ Nouvelle lettre</button></div>
        </div>
        <div class="export-buttons">
            <button id="exportExcelBtn" class="btn-export">📎 Exporter la liste en Excel</button>
            <button id="exportPdfBtn" class="btn-export">🖨️ Imprimer / PDF la liste</button>
        </div>
        <div style="overflow-x: auto;">
            <table id="tableLC"><thead><tr><th>Fournisseur</th><th>Banque</th><th>Montant (TND)</th><th>Émission</th><th>Échéance</th><th>Liquidée le</th><th>Statut</th><th>Actions</th></tr></thead><tbody id="tableBody"></tbody></table>
        </div>
    </div>

    <!-- TAB 2 -->
    <div id="tab2" class="tab-content">
        <h3>📇 Fournisseurs</h3>
        <div style="display: flex; gap: 10px; margin-bottom: 15px;"><input type="text" id="newFournisseur" placeholder="Nouveau fournisseur"><button id="addFournisseurBtn">+ Ajouter</button></div>
        <div id="fournisseurList"></div>
        <h3>🏦 Banques</h3>
        <div style="display: flex; gap: 10px; margin-bottom: 15px;"><input type="text" id="newBanque" placeholder="Nouvelle banque"><button id="addBanqueBtn">+ Ajouter</button></div>
        <div id="banqueList"></div>
    </div>

    <!-- TAB 3 -->
    <div id="tab3" class="tab-content">
        <h2>📊 Engagements par période</h2>
        <div class="report-filters">
            <div><label>Type de rapport :</label><select id="rapportPeriode"><option value="jour">Par jour</option><option value="semaine">Par semaine</option><option value="mois">Par mois</option><option value="trimestre">Par trimestre</option><option value="annee">Par année</option></select></div>
            <div><label>Année (optionnel) :</label><input type="number" id="rapportAnnee" placeholder="ex: 2025" style="width:100px"></div>
            <div><button id="genererRapportBtn">Générer le rapport</button></div>
        </div>
        <div class="export-buttons"><button id="exportRapportExcelBtn" class="btn-export">📎 Exporter rapport Excel</button><button id="exportRapportPdfBtn" class="btn-export">🖨️ Imprimer / PDF rapport</button></div>
        <div id="rapportContainer" class="report-table"><p>Sélectionnez une période et cliquez sur "Générer".</p></div>
        <div class="status-section">
            <h3>📌 Situation des lettres par rapport à l'échéance</h3>
            <div class="status-filters">
                <div><label>Filtrer par fournisseur :</label><select id="situationFournisseurFilter"><option value="">Tous les fournisseurs</option></select></div>
                <div><label>Filtrer par banque :</label><select id="situationBanqueFilter"><option value="">Toutes les banques</option></select></div>
            </div>
            <div class="status-buttons">
                <button id="showEchusRegles" class="btn-export">✅ Échus réglés</button>
                <button id="showEchusNonRegles" class="btn-warning">⚠️ Échus non réglés (retard)</button>
                <button id="showNonEchus" class="btn-export">⏳ Non échus (à venir)</button>
                <button id="showRecapSituation" class="btn-export">📊 Récapitulatif général</button>
            </div>
            <div id="situationContainer" class="report-table"><p>Cliquez sur un bouton pour afficher.</p></div>
            <div class="export-buttons" id="situationExportButtons" style="display: none;"><button id="exportSituationExcel" class="btn-export">📎 Exporter cette liste en Excel</button><button id="exportSituationPdf" class="btn-export">🖨️ Imprimer / PDF cette liste</button></div>
        </div>
    </div>

    <!-- TAB 4 (PERSONNALISATION + IMPORT/EXPORT JSON) -->
    <div id="tab4" class="tab-content">
        <h2>⚙️ Personnalisation pour le client</h2>
        <div style="background:#f8fafc; padding:20px; border-radius:20px;">
            <div class="form-group"><label>Nom de la société / client :</label><input type="text" id="clientNomInput" placeholder="Ex: SARL Bâtiment Tunis"></div>
            <div class="form-group"><label>Adresse :</label><input type="text" id="clientAdresseInput" placeholder="Adresse complète"></div>
            <div class="form-group"><label>Activité :</label><input type="text" id="clientActiviteInput" placeholder="Ex: Construction, BTP"></div>
            <div class="form-group"><label>Logo (URL ou base64) :</label><input type="text" id="clientLogoInput" placeholder="Collez une URL d'image ou une image base64"></div>
            <div class="form-group"><button id="saveClientInfos" class="btn-save">💾 Enregistrer les informations</button></div>
            <hr>
            <h3>🔄 Synchronisation manuelle entre appareils</h3>
            <p>Exportez vos données depuis votre PC, transférez le fichier sur votre téléphone (par email, WhatsApp, cloud), puis importez-le ici.</p>
            <div style="display: flex; gap: 15px; flex-wrap: wrap;">
                <button id="exportAllDataBtn" class="btn-export">💾 Exporter toutes les données (JSON)</button>
                <label for="importFileInput" class="btn-import" style="background: #8e44ad; padding: 8px 12px; border-radius: 12px; cursor: pointer;">📂 Importer des données (JSON)</label>
                <input type="file" id="importFileInput" accept=".json" style="display: none;">
            </div>
            <p><small>Le fichier JSON contient toutes vos lettres, fournisseurs, banques et informations client.</small></p>
        </div>
    </div>
</div>

<!-- Modals (inchangées) -->
<div id="modal" class="modal"><div class="modal-content"><h2 id="modalTitle">Lettre de change</h2>
    <div class="form-group"><label>Fournisseur *</label><select id="fournisseurSelect"></select><div class="inline-add"><input type="text" id="quickFournisseur" placeholder="Nouveau fournisseur"><button id="quickAddFournisseur">+ Ajouter</button></div></div>
    <div class="form-group"><label>Banque *</label><select id="banqueSelect"></select><div class="inline-add"><input type="text" id="quickBanque" placeholder="Nouvelle banque"><button id="quickAddBanque">+ Ajouter</button></div></div>
    <div class="form-group"><label>Montant (TND) *</label><input type="number" id="montant" step="0.001"></div>
    <div class="form-group"><label>Date d'émission</label><input type="text" id="dateEmission" placeholder="jj/mm/aaaa" class="date-input"></div>
    <div class="form-group"><label>Date d'échéance *</label><input type="text" id="dateEcheance" placeholder="jj/mm/aaaa" class="date-input"></div>
    <div class="form-group"><label>Date réelle de liquidation</label><input type="text" id="dateLiquidation" placeholder="jj/mm/aaaa" class="date-input"></div>
    <div style="display: flex; gap: 10px; justify-content: flex-end;"><button id="cancelModal">Annuler</button><button id="saveLettre" class="btn-save">Enregistrer</button></div>
</div></div>

<div id="modalEditRef" class="modal"><div class="modal-content"><h2 id="modalEditTitle">Modifier</h2><div class="form-group"><label>Nouveau nom</label><input type="text" id="editRefName"></div><div style="display: flex; gap: 10px; justify-content: flex-end;"><button id="cancelEditRef">Annuler</button><button id="saveEditRef" class="btn-save">Enregistrer</button></div></div></div>

<script>
    // ========== CONVERSION DATES (inchangée) ==========
    function dateToIso(str) {
        if (!str) return "";
        let parts = str.split('/');
        if (parts.length !== 3) return "";
        let day = parts[0].padStart(2,'0');
        let month = parts[1].padStart(2,'0');
        let year = parts[2];
        if (year.length === 2) year = "20" + year;
        return `${year}-${month}-${day}`;
    }
    function isoToDisplay(iso) {
        if (!iso) return "";
        let parts = iso.split('-');
        if (parts.length !== 3) return "";
        return `${parts[2]}/${parts[1]}/${parts[0]}`;
    }
    function isValidDateDMY(str) {
        if (!str) return false;
        let parts = str.split('/');
        if (parts.length !== 3) return false;
        let day = parseInt(parts[0],10);
        let month = parseInt(parts[1],10);
        let year = parseInt(parts[2],10);
        if (isNaN(day) || isNaN(month) || isNaN(year)) return false;
        if (month < 1 || month > 12) return false;
        let daysInMonth = new Date(year, month, 0).getDate();
        if (day < 1 || day > daysInMonth) return false;
        return true;
    }

    // ========== DONNÉES PRINCIPALES ==========
    let fournisseurs = ["Béton SARL", "Électricité Bât", "Menuiserie Moderne"];
    let banques = ["BNP Paribas", "Crédit Agricole", "Société Générale"];
    let lettres = [
        { id: 1, fournisseur: "Béton SARL", banque: "BNP Paribas", montant: 12500.500, dateEmission: "2025-01-10", dateEcheance: "2025-02-15", dateLiquidation: "2025-02-14", statut: "Liquidée" },
        { id: 2, fournisseur: "Électricité Bât", banque: "Crédit Agricole", montant: 8450.000, dateEmission: "2025-02-01", dateEcheance: "2025-03-10", dateLiquidation: "", statut: "En attente" }
    ];
    let nextId = 3;
    let editId = null;
    let editRefType = null, editRefOldName = null;

    function updateStatut(l) {
        l.statut = (l.dateLiquidation && l.dateLiquidation !== "") ? "Liquidée" : "En attente";
        return l;
    }
    function saveData() {
        localStorage.setItem("lc_fournisseurs_TN", JSON.stringify(fournisseurs));
        localStorage.setItem("lc_banques_TN", JSON.stringify(banques));
        localStorage.setItem("lc_lettres_TN", JSON.stringify(lettres));
        localStorage.setItem("lc_nextId_TN", nextId);
    }
    function loadData() {
        const f = localStorage.getItem("lc_fournisseurs_TN");
        if (f) fournisseurs = JSON.parse(f);
        const b = localStorage.getItem("lc_banques_TN");
        if (b) banques = JSON.parse(b);
        const l = localStorage.getItem("lc_lettres_TN");
        if (l) { lettres = JSON.parse(l); lettres.forEach(l => { if(!l.dateLiquidation) l.dateLiquidation=""; updateStatut(l); }); }
        const n = localStorage.getItem("lc_nextId_TN");
        if (n) nextId = parseInt(n);
    }
    function isUsedInLettres(type, name) {
        if (type === "fournisseur") return lettres.some(l => l.fournisseur === name);
        return lettres.some(l => l.banque === name);
    }
    function deleteRef(type, name) {
        if (isUsedInLettres(type, name)) {
            alert(`❌ Impossible de supprimer "${name}" car utilisé dans ${lettres.filter(l => type==="fournisseur"? l.fournisseur===name : l.banque===name).length} lettre(s). Supprimez d'abord ces lettres.`);
            return false;
        }
        if (type === "fournisseur") fournisseurs = fournisseurs.filter(f => f !== name);
        else banques = banques.filter(b => b !== name);
        saveData(); refreshAll(); return true;
    }
    function renameRef(type, oldName, newName) {
        if (!newName || newName.trim() === "") { alert("Le nom ne peut pas être vide"); return false; }
        newName = newName.trim();
        if (type === "fournisseur") {
            if (fournisseurs.includes(newName) && newName !== oldName) { alert("Ce fournisseur existe déjà"); return false; }
            lettres.forEach(l => { if (l.fournisseur === oldName) l.fournisseur = newName; });
            const index = fournisseurs.indexOf(oldName);
            if (index !== -1) fournisseurs[index] = newName;
        } else {
            if (banques.includes(newName) && newName !== oldName) { alert("Cette banque existe déjà"); return false; }
            lettres.forEach(l => { if (l.banque === oldName) l.banque = newName; });
            const index = banques.indexOf(oldName);
            if (index !== -1) banques[index] = newName;
        }
        saveData(); refreshAll(); return true;
    }
    function refreshAll() {
        refreshRefLists(); populateSelects(); computeDashboard(); renderTable();
        const fournFilter = document.getElementById("situationFournisseurFilter");
        if (fournFilter) {
            const current = fournFilter.value;
            fournFilter.innerHTML = '<option value="">Tous les fournisseurs</option>' + fournisseurs.map(f => `<option value="${f}">${f}</option>`).join('');
            if (current && fournisseurs.includes(current)) fournFilter.value = current;
        }
        const banqueFilter = document.getElementById("situationBanqueFilter");
        if (banqueFilter) {
            const current = banqueFilter.value;
            banqueFilter.innerHTML = '<option value="">Toutes les banques</option>' + banques.map(b => `<option value="${b}">${b}</option>`).join('');
            if (current && banques.includes(current)) banqueFilter.value = current;
        }
    }
    function refreshRefLists() {
        const fournDiv = document.getElementById("fournisseurList");
        fournDiv.innerHTML = "";
        fournisseurs.forEach(f => {
            const div = document.createElement("div"); div.className = "list-item";
            div.innerHTML = `<span>${f}</span><div class="list-actions"><button class="btn-small btn-edit" data-type="fournisseur" data-name="${f}">✏️ Modifier</button><button class="btn-small btn-delete" data-type="fournisseur" data-name="${f}">🗑️ Supprimer</button></div>`;
            fournDiv.appendChild(div);
        });
        const banqueDiv = document.getElementById("banqueList");
        banqueDiv.innerHTML = "";
        banques.forEach(b => {
            const div = document.createElement("div"); div.className = "list-item";
            div.innerHTML = `<span>${b}</span><div class="list-actions"><button class="btn-small btn-edit" data-type="banque" data-name="${b}">✏️ Modifier</button><button class="btn-small btn-delete" data-type="banque" data-name="${b}">🗑️ Supprimer</button></div>`;
            banqueDiv.appendChild(div);
        });
        document.querySelectorAll("#fournisseurList .btn-edit, #banqueList .btn-edit").forEach(btn => { btn.onclick = (e) => openEditRefModal(btn.dataset.type, btn.dataset.name); });
        document.querySelectorAll("#fournisseurList .btn-delete, #banqueList .btn-delete").forEach(btn => { btn.onclick = (e) => { if(confirm(`Supprimer ${btn.dataset.type==="fournisseur"?"le fournisseur":"la banque"} "${btn.dataset.name}" ?`)) deleteRef(btn.dataset.type, btn.dataset.name); }; });
    }
    function openEditRefModal(type, oldName) {
        editRefType = type; editRefOldName = oldName;
        document.getElementById("modalEditTitle").innerText = `Modifier ${type === "fournisseur" ? "le fournisseur" : "la banque"}`;
        document.getElementById("editRefName").value = oldName;
        document.getElementById("modalEditRef").style.display = "flex";
    }
    function populateSelects() {
        const fournSelect = document.getElementById("fournisseurSelect");
        fournSelect.innerHTML = fournisseurs.map(f => `<option>${f}</option>`).join('');
        const banqueSelect = document.getElementById("banqueSelect");
        banqueSelect.innerHTML = banques.map(b => `<option>${b}</option>`).join('');
        const filterFourn = document.getElementById("filterFournisseur");
        filterFourn.innerHTML = '<option value="">Tous</option>' + fournisseurs.map(f => `<option>${f}</option>`).join('');
        const filterBanque = document.getElementById("filterBanque");
        filterBanque.innerHTML = '<option value="">Toutes</option>' + banques.map(b => `<option>${b}</option>`).join('');
    }
    function computeDashboard() {
        const todayISO = new Date().toISOString().slice(0,10);
        const currentYear = new Date().getFullYear();
        const currentMonth = String(new Date().getMonth()+1).padStart(2,'0');
        let moisCount=0, moisMontant=0, anneeCount=0, anneeMontant=0, jourCount=0;
        const fournSet = new Set(lettres.map(l=>l.fournisseur));
        for(let l of lettres) {
            if(l.dateEcheance.startsWith(currentYear)) { anneeCount++; anneeMontant+=l.montant; }
            if(l.dateEcheance.startsWith(currentYear+"-"+currentMonth)) { moisCount++; moisMontant+=l.montant; }
            if(l.dateEcheance === todayISO) jourCount++;
        }
        document.getElementById("moisCount").innerText = moisCount;
        document.getElementById("moisMontant").innerText = moisMontant.toFixed(3);
        document.getElementById("anneeCount").innerText = anneeCount;
        document.getElementById("anneeMontant").innerText = anneeMontant.toFixed(3);
        document.getElementById("jourCount").innerText = jourCount;
        document.getElementById("fournisseurCount").innerText = fournSet.size;
    }
    function getFilteredLettres() {
        return lettres.filter(l => {
            const filFour = document.getElementById("filterFournisseur").value;
            if(filFour && l.fournisseur !== filFour) return false;
            const filBanque = document.getElementById("filterBanque").value;
            if(filBanque && l.banque !== filBanque) return false;
            const debutStr = document.getElementById("filterDateDebut").value;
            if (debutStr) { let di = dateToIso(debutStr); if (di && l.dateEcheance < di) return false; }
            const finStr = document.getElementById("filterDateFin").value;
            if (finStr) { let fi = dateToIso(finStr); if (fi && l.dateEcheance > fi) return false; }
            const statut = document.getElementById("filterStatut").value;
            if(statut && l.statut !== statut) return false;
            return true;
        });
    }
    function renderTable() {
        let filtered = getFilteredLettres();
        const tbody = document.getElementById("tableBody");
        tbody.innerHTML = "";
        for(let l of filtered) {
            const row = tbody.insertRow();
            row.insertCell(0).innerText = l.fournisseur;
            row.insertCell(1).innerText = l.banque;
            row.insertCell(2).innerText = l.montant.toFixed(3) + " TND";
            row.insertCell(3).innerText = isoToDisplay(l.dateEmission);
            row.insertCell(4).innerText = isoToDisplay(l.dateEcheance);
            row.insertCell(5).innerText = l.dateLiquidation ? isoToDisplay(l.dateLiquidation) : "Non liquidée";
            row.insertCell(6).innerText = l.statut;
            const actions = row.insertCell(7);
            const editBtn = document.createElement("button"); editBtn.innerText="✏️"; editBtn.style.marginRight="5px"; editBtn.onclick = () => openEditModal(l.id);
            const delBtn = document.createElement("button"); delBtn.innerText="🗑️"; delBtn.className="btn-delete"; delBtn.onclick = () => deleteLettre(l.id);
            actions.appendChild(editBtn); actions.appendChild(delBtn);
        }
    }
    function deleteLettre(id) {
        if(confirm("Supprimer cette lettre de change ?")) {
            lettres = lettres.filter(l=>l.id!==id);
            saveData(); refreshAll();
        }
    }
    function openEditModal(id) {
        const l = lettres.find(l=>l.id===id);
        if(l) {
            editId = id;
            document.getElementById("modalTitle").innerText = "Modifier la lettre";
            document.getElementById("fournisseurSelect").value = l.fournisseur;
            document.getElementById("banqueSelect").value = l.banque;
            document.getElementById("montant").value = l.montant;
            document.getElementById("dateEmission").value = isoToDisplay(l.dateEmission);
            document.getElementById("dateEcheance").value = isoToDisplay(l.dateEcheance);
            document.getElementById("dateLiquidation").value = isoToDisplay(l.dateLiquidation);
            document.getElementById("modal").style.display = "flex";
        }
    }
    function openAddModal() {
        editId = null;
        document.getElementById("modalTitle").innerText = "Ajouter une lettre";
        if(fournisseurs.length) document.getElementById("fournisseurSelect").value = fournisseurs[0];
        if(banques.length) document.getElementById("banqueSelect").value = banques[0];
        document.getElementById("montant").value = "";
        document.getElementById("dateEmission").value = "";
        document.getElementById("dateEcheance").value = "";
        document.getElementById("dateLiquidation").value = "";
        document.getElementById("modal").style.display = "flex";
    }
    function saveLettre() {
        const fournisseur = document.getElementById("fournisseurSelect").value;
        const banque = document.getElementById("banqueSelect").value;
        const montant = parseFloat(document.getElementById("montant").value);
        const dateEmissionStr = document.getElementById("dateEmission").value;
        const dateEcheanceStr = document.getElementById("dateEcheance").value;
        const dateLiquidationStr = document.getElementById("dateLiquidation").value;
        if(!fournisseur || !banque || isNaN(montant) || montant<=0) { alert("Veuillez remplir : fournisseur, banque, montant (>0)"); return; }
        if (!dateEcheanceStr || !isValidDateDMY(dateEcheanceStr)) { alert("Date d'échéance invalide (format jj/mm/aaaa)"); return; }
        if (dateEmissionStr && !isValidDateDMY(dateEmissionStr)) { alert("Date d'émission invalide (format jj/mm/aaaa)"); return; }
        if (dateLiquidationStr && !isValidDateDMY(dateLiquidationStr)) { alert("Date de liquidation invalide (format jj/mm/aaaa)"); return; }
        const dateEmissionIso = dateEmissionStr ? dateToIso(dateEmissionStr) : "";
        const dateEcheanceIso = dateToIso(dateEcheanceStr);
        const dateLiquidationIso = dateLiquidationStr ? dateToIso(dateLiquidationStr) : "";
        if(editId === null) {
            const newLettre = { id: nextId++, fournisseur, banque, montant, dateEmission: dateEmissionIso, dateEcheance: dateEcheanceIso, dateLiquidation: dateLiquidationIso, statut:"" };
            updateStatut(newLettre);
            lettres.push(newLettre);
        } else {
            const idx = lettres.findIndex(l=>l.id===editId);
            if(idx!==-1) {
                lettres[idx].fournisseur = fournisseur;
                lettres[idx].banque = banque;
                lettres[idx].montant = montant;
                lettres[idx].dateEmission = dateEmissionIso;
                lettres[idx].dateEcheance = dateEcheanceIso;
                lettres[idx].dateLiquidation = dateLiquidationIso;
                updateStatut(lettres[idx]);
            }
        }
        saveData(); document.getElementById("modal").style.display = "none"; refreshAll();
    }
    function quickAddFournisseur() {
        const nom = document.getElementById("quickFournisseur").value.trim();
        if(nom && !fournisseurs.includes(nom)) { fournisseurs.push(nom); saveData(); refreshAll(); document.getElementById("quickFournisseur").value = ""; document.getElementById("fournisseurSelect").value = nom; }
        else if(!nom) alert("Entrez un nom"); else alert("Ce fournisseur existe déjà");
    }
    function quickAddBanque() {
        const nom = document.getElementById("quickBanque").value.trim();
        if(nom && !banques.includes(nom)) { banques.push(nom); saveData(); refreshAll(); document.getElementById("quickBanque").value = ""; document.getElementById("banqueSelect").value = nom; }
        else if(!nom) alert("Entrez un nom"); else alert("Cette banque existe déjà");
    }
    function addFournisseurFromTab() {
        const nom = document.getElementById("newFournisseur").value.trim();
        if(nom && !fournisseurs.includes(nom)) { fournisseurs.push(nom); saveData(); refreshAll(); document.getElementById("newFournisseur").value=""; }
        else if(nom) alert("Déjà existant");
    }
    function addBanqueFromTab() {
        const nom = document.getElementById("newBanque").value.trim();
        if(nom && !banques.includes(nom)) { banques.push(nom); saveData(); refreshAll(); document.getElementById("newBanque").value=""; }
        else if(nom) alert("Déjà existante");
    }

    // ========== GESTION CLIENT ==========
    let clientInfos = { nom: "Ma Société", adresse: "Adresse : ...", activite: "Activité : ...", logo: "" };
    function loadClientInfos() {
        const saved = localStorage.getItem("lc_client_infos");
        if (saved) clientInfos = JSON.parse(saved);
        updateClientDisplay();
        document.getElementById("clientNomInput").value = clientInfos.nom;
        document.getElementById("clientAdresseInput").value = clientInfos.adresse;
        document.getElementById("clientActiviteInput").value = clientInfos.activite;
        document.getElementById("clientLogoInput").value = clientInfos.logo;
    }
    function saveClientInfos() {
        clientInfos.nom = document.getElementById("clientNomInput").value;
        clientInfos.adresse = document.getElementById("clientAdresseInput").value;
        clientInfos.activite = document.getElementById("clientActiviteInput").value;
        clientInfos.logo = document.getElementById("clientLogoInput").value;
        localStorage.setItem("lc_client_infos", JSON.stringify(clientInfos));
        updateClientDisplay();
        alert("Informations enregistrées !");
    }
    function updateClientDisplay() {
        document.getElementById("clientNom").innerText = clientInfos.nom;
        document.getElementById("clientAdresse").innerText = clientInfos.adresse;
        document.getElementById("clientActivite").innerText = clientInfos.activite;
        const logoContainer = document.getElementById("logoContainer");
        if (clientInfos.logo && clientInfos.logo.trim() !== "") {
            logoContainer.innerHTML = `<img src="${clientInfos.logo}" alt="Logo" style="max-width:150px; max-height:80px;" onerror="this.style.display='none'">`;
        } else logoContainer.innerHTML = "";
    }
    function getHeaderHTML() {
        let logoHtml = "";
        if (clientInfos.logo && clientInfos.logo.trim() !== "") logoHtml = `<div style="float:right;"><img src="${clientInfos.logo}" style="max-height:60px;"></div>`;
        return `<div style="border-bottom:2px solid #2c7da0; margin-bottom:15px; padding-bottom:10px;">${logoHtml}<h2 style="margin:0; color:#1e4a6e;">${clientInfos.nom}</h2><p style="margin:5px 0;">${clientInfos.adresse}<br>${clientInfos.activite}</p></div>`;
    }

    // ========== NOUVEAU : EXPORT/IMPORT JSON POUR SYNCHRONISATION ==========
    function exportAllData() {
        const exportObj = {
            version: "1.0",
            exportDate: new Date().toISOString(),
            fournisseurs: fournisseurs,
            banques: banques,
            lettres: lettres,
            nextId: nextId,
            clientInfos: clientInfos
        };
        const dataStr = JSON.stringify(exportObj, null, 2);
        const blob = new Blob([dataStr], {type: "application/json"});
        const url = URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url;
        a.download = `suivi_lettres_${new Date().toISOString().slice(0,19).replace(/:/g, '-')}.json`;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
        alert("✅ Données exportées avec succès !");
    }

    function importAllData(file) {
        const reader = new FileReader();
        reader.onload = function(e) {
            try {
                const imported = JSON.parse(e.target.result);
                if (!imported.fournisseurs || !imported.banques || !imported.lettres || imported.nextId === undefined) {
                    throw new Error("Fichier JSON invalide : structure incorrecte.");
                }
                // Remplacer les données actuelles
                fournisseurs = imported.fournisseurs;
                banques = imported.banques;
                lettres = imported.lettres.map(l => {
                    if (!l.dateLiquidation) l.dateLiquidation = "";
                    updateStatut(l);
                    return l;
                });
                nextId = imported.nextId;
                if (imported.clientInfos) clientInfos = imported.clientInfos;
                // Sauvegarder dans localStorage
                saveData();
                localStorage.setItem("lc_client_infos", JSON.stringify(clientInfos));
                // Rafraîchir l'interface
                refreshAll();
                updateClientDisplay();
                alert(`✅ Import réussi ! ${lettres.length} lettres, ${fournisseurs.length} fournisseurs, ${banques.length} banques.`);
            } catch (err) {
                alert("❌ Erreur lors de l'import : " + err.message);
            }
        };
        reader.onerror = () => alert("Erreur de lecture du fichier.");
        reader.readAsText(file);
    }

    // ========== EXPORTS EXCEL/PDF (inchangés) ==========
    function exportToExcelWithHeader(data, filename, sheetName) {
        const finalData = [[`Société: ${clientInfos.nom}`],[`Adresse: ${clientInfos.adresse}`],[`Activité: ${clientInfos.activite}`],[],...data.map(row => Object.values(row))];
        const ws = XLSX.utils.aoa_to_sheet(finalData);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, sheetName);
        XLSX.writeFile(wb, filename + ".xlsx");
    }
    function printWithHeader(htmlContent, title) {
        const fullHtml = `<html><head><title>${title}</title><style>body{font-family:Arial; padding:20px;} table{border-collapse:collapse; width:100%;} th,td{border:1px solid #ccc; padding:8px;}</style></head><body>${getHeaderHTML()}${htmlContent}</body></html>`;
        const win = window.open(); win.document.write(fullHtml); win.document.close(); win.print();
    }
    document.getElementById("exportExcelBtn").addEventListener("click", () => {
        const filtered = getFilteredLettres();
        const data = filtered.map(l => ({ Fournisseur: l.fournisseur, Banque: l.banque, Montant_TND: l.montant.toFixed(3), Emission: isoToDisplay(l.dateEmission), Echeance: isoToDisplay(l.dateEcheance), Liquidation: l.dateLiquidation ? isoToDisplay(l.dateLiquidation) : "", Statut: l.statut }));
        exportToExcelWithHeader(data, "lettres_change", "Lettres");
    });
    document.getElementById("exportPdfBtn").addEventListener("click", () => {
        const filtered = getFilteredLettres();
        let tableHtml = `<h3>Liste des lettres de change</h3><table border="1"><thead><tr><th>Fournisseur</th><th>Banque</th><th>Montant</th><th>Émission</th><th>Échéance</th><th>Liquidée le</th><th>Statut</th></tr></thead><tbody>`;
        filtered.forEach(l => { tableHtml += `<tr><td>${l.fournisseur}</td><td>${l.banque}</td><td>${l.montant.toFixed(3)} TND</td><td>${isoToDisplay(l.dateEmission)}</td><td>${isoToDisplay(l.dateEcheance)}</td><td>${l.dateLiquidation ? isoToDisplay(l.dateLiquidation) : ""}</td><td>${l.statut}</td></tr>`; });
        tableHtml += `</tbody></table>`;
        printWithHeader(tableHtml, "Liste des lettres de change");
    });
    document.getElementById("genererRapportBtn").addEventListener("click", genererRapport);
    function genererRapport() {
        const periode = document.getElementById("rapportPeriode").value;
        let anneeFilter = document.getElementById("rapportAnnee").value;
        if (anneeFilter) anneeFilter = parseInt(anneeFilter);
        let lettresFiltrees = lettres;
        if (anneeFilter && periode !== "annee") lettresFiltrees = lettres.filter(l => parseInt(l.dateEcheance.split('-')[0]) === anneeFilter);
        const engagements = new Map();
        for (let l of lettresFiltrees) {
            let key; const date = new Date(l.dateEcheance); const year = date.getFullYear();
            if (periode === "annee") key = `${year}`;
            else if (periode === "trimestre") key = `${year}-T${Math.floor(date.getMonth()/3)+1}`;
            else if (periode === "mois") key = `${year}-${String(date.getMonth()+1).padStart(2,'0')}`;
            else if (periode === "semaine") { const firstDay = new Date(year,0,1); const pastDays = (date-firstDay)/86400000; const week = Math.ceil((pastDays + firstDay.getDay()+1)/7); key = `${year}-S${week}`; }
            else key = l.dateEcheance;
            if (!engagements.has(key)) engagements.set(key, { count: 0, montant: 0 });
            const val = engagements.get(key); val.count++; val.montant += l.montant;
        }
        const sortedKeys = Array.from(engagements.keys()).sort();
        let html = `<h3>Engagements par ${periode}</h3><table border="1"><thead><tr><th>Période</th><th>Nb lettres</th><th>Montant total (TND)</th></tr></thead><tbody>`;
        for (let key of sortedKeys) {
            let disp = key;
            if (periode === "annee") disp = `Année ${key}`;
            else if (periode === "trimestre") { let [y,t]=key.split('-T'); disp=`${y} - Trimestre ${t}`; }
            else if (periode === "mois") disp = key;
            else if (periode === "semaine") { let [y,w]=key.split('-S'); disp=`Semaine ${w} de ${y}`; }
            else disp = isoToDisplay(key);
            const val = engagements.get(key);
            html += `<tr><td>${disp}</td><td>${val.count}</td><td>${val.montant.toFixed(3)}</td></tr>`;
        }
        html += `</tbody></table>`;
        document.getElementById("rapportContainer").innerHTML = html;
    }
    document.getElementById("exportRapportExcelBtn").addEventListener("click", () => {
        const table = document.querySelector("#rapportContainer table");
        if (!table) { alert("Générez d'abord un rapport."); return; }
        const rows = table.querySelectorAll("tr");
        const data = []; const headers = Array.from(rows[0].querySelectorAll("th")).map(h => h.innerText);
        for (let i=1; i<rows.length; i++) { const cells = rows[i].querySelectorAll("td"); const rowData = {}; cells.forEach((cell,idx) => { rowData[headers[idx]] = cell.innerText; }); data.push(rowData); }
        exportToExcelWithHeader(data, "rapport_engagements", "Rapport");
    });
    document.getElementById("exportRapportPdfBtn").addEventListener("click", () => {
        const rapportHtml = document.getElementById("rapportContainer").innerHTML;
        if (!rapportHtml || !rapportHtml.includes("<table")) { alert("Générez d'abord un rapport."); return; }
        printWithHeader(rapportHtml, "Rapport des engagements");
    });

    // SITUATION (inchangée)
    let currentSituationData = [];
    function getTodayISO() { return new Date().toISOString().slice(0,10); }
    function getSelectedFournisseur() { return document.getElementById("situationFournisseurFilter").value || null; }
    function getSelectedBanque() { return document.getElementById("situationBanqueFilter").value || null; }
    function filterBySituation(type) {
        const today = getTodayISO(); const f = getSelectedFournisseur(); const b = getSelectedBanque();
        return lettres.filter(l => {
            if (f && l.fournisseur !== f) return false; if (b && l.banque !== b) return false;
            if (type === "echus_regles") return l.dateEcheance < today && l.dateLiquidation && l.dateLiquidation !== "";
            if (type === "echus_non_regles") return l.dateEcheance < today && (!l.dateLiquidation || l.dateLiquidation === "");
            if (type === "non_echus") return l.dateEcheance >= today;
            return false;
        });
    }
    function displaySituationList(filtered) {
        currentSituationData = filtered.map(l => ({ Fournisseur: l.fournisseur, Banque: l.banque, Montant_TND: l.montant.toFixed(3), Emission: isoToDisplay(l.dateEmission), Echeance: isoToDisplay(l.dateEcheance), Liquidation: l.dateLiquidation ? isoToDisplay(l.dateLiquidation) : "", Statut: l.statut }));
        if (filtered.length === 0) { document.getElementById("situationContainer").innerHTML = `<p>Aucune lettre.</p>`; document.getElementById("situationExportButtons").style.display = "none"; return; }
        let html = `<table border="1"><thead><tr><th>Fournisseur</th><th>Banque</th><th>Montant</th><th>Émission</th><th>Échéance</th><th>Liquidée le</th><th>Statut</th></tr></thead><tbody>`;
        filtered.forEach(l => { html += `<tr><td>${l.fournisseur}</td><td>${l.banque}</td><td>${l.montant.toFixed(3)} TND</td><td>${isoToDisplay(l.dateEmission)}</td><td>${isoToDisplay(l.dateEcheance)}</td><td>${l.dateLiquidation ? isoToDisplay(l.dateLiquidation) : ""}</td><td>${l.statut}</td></tr>`; });
        html += `</tbody></table>`;
        document.getElementById("situationContainer").innerHTML = html;
        document.getElementById("situationExportButtons").style.display = "flex";
    }
    function displayRecapSituation() {
        const today = getTodayISO(); const f = getSelectedFournisseur(); const b = getSelectedBanque();
        let lettresFiltrees = lettres.filter(l => { if (f && l.fournisseur !== f) return false; if (b && l.banque !== b) return false; return true; });
        const echusRegles = lettresFiltrees.filter(l => l.dateEcheance < today && l.dateLiquidation && l.dateLiquidation !== "");
        const echusNonRegles = lettresFiltrees.filter(l => l.dateEcheance < today && (!l.dateLiquidation || l.dateLiquidation === ""));
        const nonEchus = lettresFiltrees.filter(l => l.dateEcheance >= today);
        const recap = [
            { Catégorie: "✅ Échus réglés", Nombre: echusRegles.length, Montant_TND: echusRegles.reduce((s,l)=>s+l.montant,0).toFixed(3) },
            { Catégorie: "⚠️ Échus non réglés (retard)", Nombre: echusNonRegles.length, Montant_TND: echusNonRegles.reduce((s,l)=>s+l.montant,0).toFixed(3) },
            { Catégorie: "⏳ Non échus (à venir)", Nombre: nonEchus.length, Montant_TND: nonEchus.reduce((s,l)=>s+l.montant,0).toFixed(3) }
        ];
        currentSituationData = recap;
        let html = `<table border="1"><thead><tr><th>Catégorie</th><th>Nombre de lettres</th><th>Montant total (TND)</th></tr></thead><tbody>`;
        recap.forEach(r => { html += `<tr><td>${r.Catégorie}</td><td>${r.Nombre}</td><td>${r.Montant_TND}</td></tr>`; });
        html += `</tbody></table>`;
        document.getElementById("situationContainer").innerHTML = html;
        document.getElementById("situationExportButtons").style.display = "flex";
    }
    document.getElementById("exportSituationExcel").addEventListener("click", () => { if (!currentSituationData.length) { alert("Aucune donnée à exporter."); return; } exportToExcelWithHeader(currentSituationData, "situation_lettres", "Situation"); });
    document.getElementById("exportSituationPdf").addEventListener("click", () => { if (!currentSituationData.length) { alert("Aucune donnée à exporter."); return; } printWithHeader(document.getElementById("situationContainer").innerHTML, "Situation des lettres"); });
    document.getElementById("showEchusRegles").addEventListener("click", () => displaySituationList(filterBySituation("echus_regles")));
    document.getElementById("showEchusNonRegles").addEventListener("click", () => displaySituationList(filterBySituation("echus_non_regles")));
    document.getElementById("showNonEchus").addEventListener("click", () => displaySituationList(filterBySituation("non_echus")));
    document.getElementById("showRecapSituation").addEventListener("click", displayRecapSituation);

    // Événements divers
    document.getElementById("saveEditRef").addEventListener("click", () => { const newName = document.getElementById("editRefName").value.trim(); if (newName && editRefType && editRefOldName) renameRef(editRefType, editRefOldName, newName); document.getElementById("modalEditRef").style.display = "none"; });
    document.getElementById("cancelEditRef").addEventListener("click", () => document.getElementById("modalEditRef").style.display = "none");
    document.getElementById("openAddBtn").addEventListener("click", openAddModal);
    document.getElementById("cancelModal").addEventListener("click", () => document.getElementById("modal").style.display = "none");
    document.getElementById("saveLettre").addEventListener("click", saveLettre);
    document.getElementById("resetFilters").addEventListener("click", () => { document.getElementById("filterFournisseur").value = ""; document.getElementById("filterBanque").value = ""; document.getElementById("filterDateDebut").value = ""; document.getElementById("filterDateFin").value = ""; document.getElementById("filterStatut").value = ""; renderTable(); });
    document.getElementById("filterFournisseur").addEventListener("change", renderTable);
    document.getElementById("filterBanque").addEventListener("change", renderTable);
    document.getElementById("filterDateDebut").addEventListener("input", renderTable);
    document.getElementById("filterDateFin").addEventListener("input", renderTable);
    document.getElementById("filterStatut").addEventListener("change", renderTable);
    document.getElementById("quickAddFournisseur").addEventListener("click", quickAddFournisseur);
    document.getElementById("quickAddBanque").addEventListener("click", quickAddBanque);
    document.getElementById("addFournisseurBtn").addEventListener("click", addFournisseurFromTab);
    document.getElementById("addBanqueBtn").addEventListener("click", addBanqueFromTab);
    document.getElementById("saveClientInfos").addEventListener("click", saveClientInfos);

    // Nouveaux événements pour l'export/import JSON
    document.getElementById("exportAllDataBtn").addEventListener("click", exportAllData);
    document.getElementById("importFileInput").addEventListener("change", (e) => {
        if (e.target.files && e.target.files[0]) importAllData(e.target.files[0]);
        e.target.value = ""; // permet de réimporter le même fichier plus tard
    });

    // Tabs
    document.querySelectorAll(".tab-btn").forEach(btn => {
        btn.addEventListener("click", () => {
            document.querySelectorAll(".tab-btn").forEach(b=>b.classList.remove("active"));
            btn.classList.add("active");
            document.querySelectorAll(".tab-content").forEach(tab=>tab.classList.remove("active"));
            document.getElementById(btn.dataset.tab).classList.add("active");
        });
    });

    loadData();
    loadClientInfos();
    refreshAll();

    // Service Worker (inchangé)
    if ('serviceWorker' in navigator) {
        navigator.serviceWorker.register('sw.js').then(reg => console.log('SW enregistré', reg)).catch(err => console.log('SW erreur', err));
    }

    window.onclick = (e) => { if (e.target === document.getElementById("modal")) document.getElementById("modal").style.display = "none"; if (e.target === document.getElementById("modalEditRef")) document.getElementById("modalEditRef").style.display = "none"; };
</script>
</body>
</html>