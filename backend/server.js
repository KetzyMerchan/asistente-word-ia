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
            "https://openrouter.ai/api/v1/chat/completions",
            {
                method: "POST",
                headers: {
                    "Authorization": `Bearer ${process.env.OPENROUTER_API_KEY}`,
                    "Content-Type": "application/json"
                },
                body: JSON.stringify({
                    model: "inclusionai/ring-2.6-1t:free", 
                    messages: [
                        {
                            role: "user",
                            content: prompt
                        }
                    ]
                })
            }
        );

        // 🔴 VALIDACIÓN CRÍTICA
        if (!respuesta.ok) {
            const errorText = await respuesta.text();
            console.error("Error API:", errorText);

            return res.status(respuesta.status).json({
                error: "Error desde OpenRouter",
                details: errorText
            });
        }

        const datos = await respuesta.json();

        console.log(JSON.stringify(datos, null, 2));

        const contenido = datos?.choices?.[0]?.message?.content;

        if (!contenido) {
            return res.status(500).json({
                error: "No se obtuvo respuesta del modelo."
            });
        }

        return res.json({
            respuesta: contenido
        });

    } catch (error) {

        console.error(error);

        return res.status(500).json({
            error: "Error del servidor."
        });
    }
});

app.listen(3000, () => {
    console.log("Servidor activo en puerto 3000");
});
