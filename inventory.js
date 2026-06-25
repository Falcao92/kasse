const siteUrl = _spPageContextInfo.webAbsoluteUrl;

let products = [];

// ===================== LOAD =====================
async function getList(name) {
  const res = await fetch(
    `${siteUrl}/_api/web/lists/getbytitle('${name}')/items`,
    { headers: { Accept: "application/json;odata=verbose" } }
  );
  const data = await res.json();
  return data.d.results;
}

async function loadData() {
  products = await getList("Products");
  renderProducts();
}

// ===================== UI =====================
function renderProducts() {
  const container = document.getElementById("products");
  container.innerHTML = "";

  products.forEach(p => {
    const low = isLowStock(p) ? "low" : "";

    container.innerHTML += `
      <div class="card ${low}">
        <h3>${p.Title}</h3>
        <p>Bestand: ${p.stock}</p>
        <p>Min: ${p.min_stock}</p>

        ${renderRecipe(p)}

        <button onclick="changeStock(${p.Id}, 1)">➕</button>
        <button onclick="changeStock(${p.Id}, -1)">➖</button>

        ${isLowStock(p) ? "<b>⚠️ Nachbestellen!</b>" : ""}
      </div>
    `;
  });
}

// ===================== RECEPT UI =====================
function addRecipeLine() {
  const div = document.getElementById("recipe");

  div.innerHTML += `
    <div class="recipe-line">
      <select class="recipe-product">
        ${products.map(p => `<option>${p.Title}</option>`).join("")}
      </select>
      <input type="number" class="recipe-qty" placeholder="Menge">
    </div>
  `;
}

function getRecipeData() {
  const items = document.querySelectorAll(".recipe-line");

  let recipe = [];

  items.forEach(row => {
    let product = row.querySelector(".recipe-product").value;
    let qty = parseFloat(row.querySelector(".recipe-qty").value);

    if (product && qty) {
      recipe.push({
        product: product,
        quantity: qty
      });
    }
  });

  return JSON.stringify(recipe);
}

// ===================== CREATE PRODUCT =====================
async function createProduct() {
  const title = document.getElementById("title").value;
  const price = parseFloat(document.getElementById("price").value);
  const stock = parseFloat(document.getElementById("stock").value);
  const min_stock = parseFloat(document.getElementById("min_stock").value);
  const type = document.getElementById("type").value;

  let recipe = null;
  if (type === "composite") {
    recipe = getRecipeData();
  }

  await fetch(
    `${siteUrl}/_api/web/lists/getbytitle('Products')/items`,
    {
      method: "POST",
      headers: {
        Accept: "application/json;odata=verbose",
        "Content-Type": "application/json;odata=verbose",
        "X-RequestDigest": document.getElementById("__REQUESTDIGEST").value
      },
      body: JSON.stringify({
        Title: title,
        price: price,
        stock: stock,
        min_stock: min_stock,
        type: type,
        recipe: recipe
      })
    }
  );

  alert("Produkt gespeichert!");
  loadData();
}

// ===================== STOCK =====================
async function changeStock(id, change) {
  let product = products.find(p => p.Id === id);
  let newStock = (product.stock || 0) + change;

  await fetch(
    `${siteUrl}/_api/web/lists/getbytitle('Products')/items(${id})`,
    {
      method: "POST",
      headers: {
        Accept: "application/json;odata=verbose",
        "Content-Type": "application/json;odata=verbose",
        "X-RequestDigest": document.getElementById("__REQUESTDIGEST").value,
        "IF-MATCH": "*",
        "X-HTTP-Method": "MERGE"
      },
      body: JSON.stringify({ stock: newStock })
    }
  );

  loadData();
}

// ===================== CHECK =====================
function isLowStock(p) {
  if (p.type === "simple") {
    return p.stock <= p.min_stock;
  }

  if (p.type === "composite" && p.recipe) {
    const recipe = JSON.parse(p.recipe);

    return recipe.some(r => {
      let sub = products.find(x => x.Title === r.product);
      return sub && sub.stock <= sub.min_stock;
    });
  }
}

// ===================== INIT =====================
loadData();
