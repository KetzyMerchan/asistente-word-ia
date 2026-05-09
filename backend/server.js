import express from "express";
import cors from "cors";
import dotenv from "dotenv";

dotenv.config();

const app = express();

app.use(cors());
app.use(express.json());

app.post("/ia", async (req, res) => {

    try {

        const prompt = req.body.prompt;

        const respuesta = await fetch(
            `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=${process.env.GEMINI_API_KEY}`,
            {
                method: "POST",
                headers: {
                    "Content-Type": "application/json"
                },
                body: JSON.stringify({
                    contents: [
                        {
                            parts: [
                                {
                                    text: prompt
                                }
                            ]
                        }
                    ]
                })
            }
        );

        const datos = await respuesta.json();

        res.json(datos);

    } catch (error) {

        console.error(error);

        res.status(500).json({
            error: "Error del servidor"
        });
    }
});

app.listen(3000, () => {
    console.log("Servidor activo en puerto 3000");
});