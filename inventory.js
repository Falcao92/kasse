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

// ===============================
// INIT
// ===============================
async function init(){

    await msalInstance.handleRedirectPromise();

    let t = await token();

    let site = await (await fetch(
        "https://graph.microsoft.com/v1.0/sites/tsc1907.sharepoint.com:/sites/Kasse",
        {headers:{Authorization:"Bearer "+t}}
    )).json();

    siteId = site.id;

    let lists = await graph(`/sites/${siteId}/lists`);

    let invList = lists.value.find(l =>
        l.displayName.toLowerCase() === "inventory"
    );

    if(!invList){
        alert("❌ Liste 'inventory' nicht gefunden!");
        return;
    }

    inventoryId = invList.id;

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

    products.forEach(p => {

        let f = p.fields;

        let low = isLowStock(p) ? "low" : "";

        container.innerHTML += `
        <div class="card ${low}">
            <h3>${f.Title}</h3>
            <p>Bestand: ${f.stock || 0}</p>
            <p>Min: ${f.minstock || 0}</p>

            ${renderRecipe(p)}

            <button onclick="changeStock('${p.id}',1)">➕</button>
            <button onclick="changeStock('${p.id}',-1)">➖</button>

            ${isLowStock(p) ? "<b>⚠️ Nachbestellen!</b>" : ""}
        </div>
        `;
    });
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

        let recipe = "";

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
                recipe: recipe
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

// ===============================
init();
