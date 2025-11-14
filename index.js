import express from "express";
import cors from "cors";
import { Client } from "@microsoft/microsoft-graph-client";
import { ClientSecretCredential } from "@azure/identity";

const app = express();
app.use(cors());
app.use(express.json());

const TENANT_ID = process.env.TENANT_ID;
const CLIENT_ID = process.env.CLIENT_ID;
const CLIENT_SECRET = process.env.CLIENT_SECRET;
const EXCEL_ITEM_ID = process.env.EXCEL_ITEM_ID;
const EXCEL_TABLE_NAME = process.env.EXCEL_TABLE_NAME;

app.post("/addSteps", async (req, res) => {
    const { employeeId, date, steps } = req.body;

    try {
        const credential = new ClientSecretCredential(TENANT_ID, CLIENT_ID, CLIENT_SECRET);

        const graph = Client.initWithMiddleware({
            authProvider: {
                getAccessToken: async () => {
                    const token = await credential.getToken("https://graph.microsoft.com/.default");
                    return token.token;
                }
            }
        });

        await graph
            .api(`/me/drive/items/${EXCEL_ITEM_ID}/workbook/tables/${EXCEL_TABLE_NAME}/rows`)
            .post({ values: [[employeeId, date, steps]] });

        res.json({ success: true });
    } catch (err) {
        res.status(500).json({ error: err.toString() });
    }
});

app.get("/", (req, res) => res.send("Unity Step Backend Running"));

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server running on port ${PORT}`));

