import express from "express";
import fetch from "node-fetch";

const app = express();

const PORT = process.env.PORT || 3000;

// 🔒 ENV Variablen (in Render setzen!)
const TENANT = process.env.TENANT_ID;
const CLIENT_ID = process.env.CLIENT_ID;
const CLIENT_SECRET = process.env.CLIENT_SECRET;
const SITE_ID = process.env.SITE_ID;
const LIST_ID = process.env.CARDS_LIST_ID;

// ✅ LIMIT (0€ Schutz)
let counter = 0;
let lastReset = new Date().toDateString();

async function getToken(){

    let res = await fetch(`https://login.microsoftonline.com/${TENANT}/oauth2/v2.0/token`, {
        method: "POST",
        headers: { "Content-Type": "application/x-www-form-urlencoded" },
        body: new URLSearchParams({
            client_id: CLIENT_ID,
            client_secret: CLIENT_SECRET,
            grant_type: "client_credentials",
            scope: "https://graph.microsoft.com/.default"
        })
    });

    let data = await res.json();
    return data.access_token;
}

// ✅ API: Guthaben abrufen
app.get("/api/balance", async (req, res) => {

    // 🔒 LIMIT CHECK
    let today = new Date().toDateString();
    if(today !== lastReset){
        counter = 0;
        lastReset = today;
    }

    if(counter > 10000){
        return res.json({ error: "Limit erreicht" });
    }

    counter++;

    const cardId = req.query.id;

    if(!cardId){
        return res.json({ balance: 0 });
    }

    try {

        const token = await getToken();

        let r = await fetch(`https://graph.microsoft.com/v1.0/sites/${SITE_ID}/lists/${LIST_ID}/items?expand=fields`, {
            headers: {
                Authorization: "Bearer " + token
            }
        });

        let data = await r.json();

        let card = data.value.find(c => c.fields.Title === cardId);

        res.json({
            balance: card ? card.fields.balance : 0
        });

    } catch(err){
        res.json({ error: "Server Fehler" });
    }
});

app.listen(PORT, () => console.log("Server läuft"));
