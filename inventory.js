// ===============================
// MSAL + GRAPH
// ===============================
const msalInstance = new msal.PublicClientApplication({
    auth: {
        clientId: "844062e7-0d5c-4852-88fc-5e9940f4ed66",
        authority: "https://login.microsoftonline.com/d05e1986-9d0f-4d67-8b0d-990eb3ae4ecd",
        redirectUri: window.location.href
    }
});



async function token(){
    let acc = msalInstance.getAllAccounts();

    if(acc.length === 0){
        await msalInstance.loginRedirect({
            scopes:["Sites.ReadWrite.All","User.Read"]
        });
        return;
    }

    let res = await msalInstance.acquireTokenSilent({
        scopes:["Sites.ReadWrite.All"],
        account: acc[0]
    });

    return res.accessToken;
}

async function graph(url, method="GET", body=null){
    let t = await token();

    return await fetch("https://graph.microsoft.com/v1.0" + url, {
        method,
        headers:{
            Authorization: "Bearer " + t,
            "Content-Type":"application/json"
        },
        body: body ? JSON.stringify(body) : null
    }).then(async r => {
        const text = await r.text();
        try { return JSON.parse(text); } catch { return text; }
    });
}

// ===============================
// GLOBALS
// ===============================
let siteId, inventoryId;
let products = [];
//Verleihsystem
let loansId;

// ===============================
// INIT
// ===============================
async function init(){

    await msalInstance.handleRedirectPromise();

    let t = await token();

    // ✅ Site holen
    let site = await (await fetch(
        "https://graph.microsoft.com/v1.0/sites/tsc1907.sharepoint.com:/sites/Kasse",
        {headers:{Authorization:"Bearer "+t}}
    )).json();

    siteId = site.id;

    // ✅ Listen laden
    let lists = await graph(`/sites/${siteId}/lists`);

    // ✅ Inventory
    let invList = lists.value.find(l =>
        l.displayName.toLowerCase() === "inventory"
    );

    if(!invList){
        alert("❌ Liste 'inventory' nicht gefunden!");
        return;
    }

    inventoryId = invList.id;

    // ✅ Loans (JETZT korrekt!)
    let loanList = lists.value.find(l =>
        l.displayName.toLowerCase() === "loans"
    );

    if(!loanList){
        alert("❌ Liste 'Loans' fehlt!");
        return;
    }

    loansId = loanList.id;

    // ✅ Daten laden
    loadProducts();
}


// ===============================
// LOAD PRODUCTS
// ===============================
async function loadProducts(){

    let res = await graph(`/sites/${siteId}/lists/${inventoryId}/items?expand=fields`);

    products = res.value || [];

    renderProducts();
}

// ===============================
// RENDER
// ===============================
function renderProducts(){

    let container = document.getElementById("products");
    container.innerHTML = "";

    let search = document.getElementById("search")?.value.toLowerCase() || "";
    let filter = document.getElementById("filter")?.value || "all";

    let categories = {
        "Getraenk": [],
        "Essen": [],
        "Verbrauch": [],
        "Inventar": []
    };

    products.forEach(p => {

        let f = p.fields;

        let name = (f.Title || "").toLowerCase();
        let cat = f.category || "Verbrauch";

        let low = isLowStock(p);

        // ✅ SUCHFILTER
        if(search && !name.includes(search)) return;

        // ✅ STATUS FILTER
        if(filter === "low" && !low) return;
        if(filter === "ok" && low) return;

        if(!categories[cat]) categories[cat] = [];
        categories[cat].push(p);
    });

    Object.keys(categories).forEach(cat => {

        if(categories[cat].length === 0) return;

        let id = "cat_" + cat;

        container.innerHTML += `
        <div class="category">
            <div class="categoryHeader" onclick="toggleCategory('${id}')">
                ${getCategoryLabel(cat)} (${categories[cat].length})
            </div>
            <div id="${id}" class="categoryContent">
                <div class="productGrid" id="${id}_grid"></div>
            </div>
        </div>
        `;

        let grid = document.getElementById(id + "_grid");

        categories[cat].forEach(p => {

            let f = p.fields;
            let low = isLowStock(p) ? "low" : "";

            // ✅ Unterschied Asset vs normal
            let actions = "";

            if(f.type === "asset"){
                actions = `
                <button onclick="lendItem('${p.id}')">📤</button>
                <button onclick="returnItem('${p.id}')">📥</button>
                `;
            } else {
                actions = `
                <button onclick="changeStock('${p.id}',1)">➕</button>
                <button onclick="changeStock('${p.id}',-1)">➖</button>
                `;
            }

            grid.innerHTML += `
            <div class="productItem ${low}">
                <b>${f.Title}</b><br>

                Bestand: ${f.stock || 0}<br>

                ${f.type !== "asset"
                    ? `Min: ${f.minstock || 0}<br>`
                    : ""}

                ${actions}
            </div>
            `;
        });
    });
}


function getCategoryLabel(cat){

    if(cat === "Getraenk") return "🍺 Getränke";
    if(cat === "Essen") return "🍔 Essen";
    if(cat === "Verbrauch") return "📦 Verbrauch";
 if(cat === "Inventar") return "🪑 Inventar";
    return cat;
}




// ===============================
// RECIPE DISPLAY
// ===============================
function renderRecipe(p){

    let f = p.fields;

    if(f.type !== "composite" || !f.recipe) return "";

    let recipe = JSON.parse(f.recipe);

    return `
    <ul>
        ${recipe.map(r => `<li>${r.quantity}x ${r.product}</li>`).join("")}
    </ul>
    `;
}

// ===============================
// LOW STOCK CHECK
// ===============================
function isLowStock(p){

    let f = p.fields;

    if(f.type === "simple"){
        return (f.stock || 0) <= (f.minstock || 0);
    }
    
if(f.type === "asset"){
    return false;
}

    if(f.type === "composite" && f.recipe){

        let recipe = JSON.parse(f.recipe);

        return recipe.some(r => {
            let sub = products.find(x => x.fields.Title === r.product);
            return sub && (sub.fields.stock <= sub.fields.minstock);
        });
    }
}

// ===============================
// CHANGE STOCK
// ===============================
async function changeStock(id, change){

    let p = products.find(x => x.id === id);
    let f = p.fields;

    let newStock = (f.stock || 0) + change;

    await graph(
        `/sites/${siteId}/lists/${inventoryId}/items/${id}/fields`,
        "PATCH",
        { stock: newStock }
    );

    loadProducts();
}

// ===============================
// CREATE PRODUCT
// ===============================
async function createProduct(){

    try {

        let title = document.getElementById("title").value.trim();
        let price = Number(document.getElementById("price").value) || 0;
        let stock = Number(document.getElementById("stock").value) || 0;
        let minstock = Number(document.getElementById("minstock").value) || 0;
        let type = document.getElementById("type").value;
let category = document.getElementById("category").value;

        let recipe = "";

        if(type === "asset"){
    category = "Inventar";
}

if(type === "asset"){
    minstock = 0;
}
        if(type === "composite"){
            let r = getRecipeData();
            if(r === "[]"){
                alert("Rezept fehlt!");
                return;
            }
            recipe = r;
        }

        let body = {
            fields: {
                Title: title,
                price: price,
                stock: stock,
                minstock: minstock,
                type: type,
                recipe: recipe,
                category: category
            }
        };

        console.log("SEND BODY:", JSON.stringify(body, null, 2));

        let res = await graph(
            `/sites/${siteId}/lists/${inventoryId}/items`,
            "POST",
            body
        );

        console.log("RESPONSE:", res);

        if(res.error){
            alert("❌ SharePoint Fehler: " + res.error.message);
            return;
        }

        alert("✅ OK gespeichert");

        loadProducts();

    } catch(e){
        console.error("HARTE ERROR:", e);
        alert("❌ Fehler – siehe Konsole");
    }
}


// ===============================
// RECIPE BUILDER
// ===============================
function addRecipeLine(){

    if(products.length === 0){
        alert("Produkte noch nicht geladen");
        return;
    }

    let div = document.getElementById("recipe");

    div.innerHTML += `
    <div class="recipe-line">
        <select class="recipe-product">
            ${products.map(p => `<option>${p.fields.Title}</option>`).join("")}
        </select>
        <input type="number" class="recipe-qty" placeholder="Menge">
    </div>
    `;
}

function getRecipeData(){

    let rows = document.querySelectorAll(".recipe-line");

    let recipe = [];

    rows.forEach(r => {
        let prod = r.querySelector(".recipe-product").value;
        let qty = parseFloat(r.querySelector(".recipe-qty").value);

        if(prod && qty){
            recipe.push({ product: prod, quantity: qty });
        }
    });

    return JSON.stringify(recipe);
}

async function lendItem(productId){

    let p = products.find(x => x.id === productId);
    let name = prompt("Wer leiht?");

    if(!name) return;

    let qty = Number(prompt("Menge:")) || 1;

    if(qty <= 0) return;

    let stock = p.fields.stock || 0;

    if(stock < qty){
        alert("Nicht genug Bestand!");
        return;
    }

    // Bestand reduzieren
    await changeStock(productId, -qty);

    // Eintrag speichern
    await graph(
        `/sites/${siteId}/lists/${loansId}/items`,
        "POST",
        {
            fields:{
                Title: "Loan",
                product: p.fields.Title,
                person: name,
                quantity: qty,
                action: "out",
                timestamp: new Date().toISOString()
            }
        }
    );

    alert("✅ Verliehen");
}

async function returnItem(productId){

    let p = products.find(x => x.id === productId);

    let name = prompt("Wer gibt zurück?");
    if(!name) return;

    let qty = Number(prompt("Menge:")) || 1;

    if(qty <= 0) return;

    // Bestand erhöhen
    await changeStock(productId, qty);

    // Log
    await graph(
        `/sites/${siteId}/lists/${loansId}/items`,
        "POST",
        {
            fields:{
                Title: "Return",
                product: p.fields.Title,
                person: name,
                quantity: qty,
                action: "back",
                timestamp: new Date().toISOString()
            }
        }
    );

    alert("✅ Zurückgegeben");
}
async function showLoans(){

    let res = await graph(`/sites/${siteId}/lists/${loansId}/items?expand=fields`);

    let list = res.value || [];

    // nach Datum sortieren (neueste zuerst)
    list.sort((a,b)=>
        new Date(b.fields.timestamp) - new Date(a.fields.timestamp)
    );

    let html = "<h2>📋 Verleih Übersicht</h2>";

    list.forEach(x => {

        let f = x.fields;

        let date = new Date(f.timestamp).toLocaleString("de-DE");

        html += `
        <div class="card">
            <b>${f.product}</b><br>
            👤 ${f.person}<br>
            📦 ${f.quantity}<br>
            ${f.action === "out" ? "📤 Verliehen" : "📥 Zurück"}<br>
            🕒 ${date}
        </div>
        `;
    });

    document.getElementById("products").innerHTML = html;
}

async function showOpenLoans(){

    let res = await graph(`/sites/${siteId}/lists/${loansId}/items?expand=fields`);

    let data = res.value || [];

    let summary = {};

    data.forEach(x => {

        let f = x.fields;

        let key = f.person + "|" + f.product;

        if(!summary[key]){
            summary[key] = {
                person: f.person,
                product: f.product,
                qty: 0
            };
        }

        if(f.action === "out"){
            summary[key].qty += Number(f.quantity || 0);
        }

        if(f.action === "back"){
            summary[key].qty -= Number(f.quantity || 0);
        }
    });

    let html = "<h2>📦 Offene Ausleihen</h2>";

    Object.values(summary)
        .filter(x => x.qty > 0)
        .forEach(x => {

            html += `
            <div class="card">
                <b>${x.product}</b><br>
                👤 ${x.person}<br>
                📦 Offen: ${x.qty}
            </div>
            `;
        });

    if(html === "<h2>📦 Offene Ausleihen</h2>"){
        html += "✅ Alles zurückgegeben!";
    }

    document.getElementById("products").innerHTML = html;
}

// ===============================
init();
