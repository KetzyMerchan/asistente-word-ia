// La API Key se carga desde config.js
//const GEMINI_API_URL = "https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent";

Office.onReady(function() {
    console.log("Office.js listo");
    updateSelectedText();
});

document.getElementById("explainBtn").onclick = function() {
    showLoading(true);
    clearResult();
    
    Word.run(async function(context) {
        try {
            // Obtener el texto seleccionado
            var selection = context.document.getSelection();
            selection.load("text");
            
            // Obtener contexto: buscar el parrafo completo
            var range = selection.getRange();
            range.load("text");
            
            await context.sync();
            
            var selectedText = selection.text;
            
            if (!selectedText || selectedText.trim() === "") {
                showResult("No hay texto seleccionado. Selecciona una palabra o frase en tu documento.", "error");
                showLoading(false);
                return;
            }
            
            // Actualizar display del texto seleccionado
            updateSelectedTextDisplay(selectedText);
            
            // Obtener contexto mas amplio (el parrafo completo)
            var contextoCompleto = range.text || selectedText;
            
            // Llamar a Gemini API para obtener la explicacion
            var explicacion = await llamarGemini(selectedText, contextoCompleto);
            
            showResult(explicacion, "success");
            showLoading(false);
            
        } catch (error) {
            console.error(error);
            showResult("Error al procesar la solicitud: " + error.message, "error");
            showLoading(false);
        }
    }).catch(function(error) {
        console.error(error);
        showResult("Error al leer el documento. Asegurate de estar en Word Online.", "error");
        showLoading(false);
    });
};
/*
async function llamarGemini(palabra, contexto) {
    if (!GEMINI_API_KEY) {
        return "Error: No se ha configurado la API Key de Gemini. Asegurate de que el archivo config.js existe y contiene la clave.";
    }
    
    // Construir el prompt para Gemini
    const prompt = `Eres un asistente academico especializado en explicar terminos en contexto.

El usuario ha seleccionado la siguiente palabra o frase: "${palabra}"

El texto completo donde aparece es: "${contexto}"

Por favor, explica que significa esta palabra o frase EN EL CONTEXTO ESPECIFICO de este texto.
Si la palabra tiene diferentes significados segun la disciplina (medicina, derecho, informatica, etc.), enfocate en el que corresponde al texto proporcionado.

Instrucciones:
- Responde en español
- Se claro, conciso y directo
- No uses markdown ni formato especial
- La explicacion debe ser de maximo 3 oraciones
- Enfocate en el significado contextual, no des definiciones genericas`;

    try {
        const respuesta = await fetch(`${GEMINI_API_URL}?key=${GEMINI_API_KEY}`, {
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
                ],
                generationConfig: {
                    temperature: 0.2,
                    maxOutputTokens: 300
                }
            })
        });
        
        if (!respuesta.ok) {
            const errorData = await respuesta.json();
            console.error("Error Gemini:", errorData);
            return "Error al consultar Gemini API. Verifica tu API Key o conexion a internet.";
        }
        
        const datos = await respuesta.json();
        
        // Extraer la explicacion de la respuesta
        if (datos.candidates && datos.candidates[0] && datos.candidates[0].content) {
            var explicacion = datos.candidates[0].content.parts[0].text;
            return explicacion;
        } else {
            return "No se pudo obtener una explicacion. Intenta de nuevo.";
        }
        
    } catch (error) {
        console.error("Error en llamada Gemini:", error);
        return "Error de conexion con Gemini API. Verifica tu conexion a internet.";
    }
}
*/

async function llamarGemini(palabra, contexto) {

    const prompt = `Eres un asistente academico especializado en explicar terminos en contexto.

El usuario ha seleccionado la siguiente palabra o frase: "${palabra}"

El texto completo donde aparece es: "${contexto}"

Por favor, explica que significa esta palabra o frase EN EL CONTEXTO ESPECIFICO de este texto.

Instrucciones:
- Responde en español
- Se claro y directo
- Maximo 3 oraciones`;

    try {

        const respuesta = await fetch("http://localhost:3000/ia", {
            method: "POST",
            headers: {
                "Content-Type": "application/json"
            },
            body: JSON.stringify({
                prompt: prompt
            })
        });

        if (!respuesta.ok) {
            return "Error al conectar con el servidor local.";
        }

        const datos = await respuesta.json();

        if (
            datos.candidates &&
            datos.candidates.length > 0 &&
            datos.candidates[0].content &&
            datos.candidates[0].content.parts &&
            datos.candidates[0].content.parts.length > 0
        ) {
            return datos.candidates[0].content.parts[0].text;
        }

        return "No se obtuvo respuesta de Gemini.";

    } catch (error) {

        console.error(error);

        return "Error de conexion con el backend.";
    }
}

function updateSelectedText() {
    Word.run(function(context) {
        var selection = context.document.getSelection();
        selection.load("text");
        return context.sync().then(function() {
            updateSelectedTextDisplay(selection.text);
        });
    }).catch(function() {});
}

function updateSelectedTextDisplay(text) {
    var el = document.getElementById("selectedDisplay");
    if (text && text.trim()) {
        var displayText = text.length > 100 ? text.substring(0, 100) + "..." : text;
        el.innerHTML = "\"" + displayText + "\"";
        el.classList.remove("empty");
    } else {
        el.innerHTML = "Ninguno";
        el.classList.add("empty");
    }
}

function showResult(msg, type) {
    var resultDiv = document.getElementById("result");
    var bg = "#e8f0fe";
    var border = "#667eea";
    if (type === "success") {
        bg = "#e6f4ea";
        border = "#34a853";
    } else if (type === "error") {
        bg = "#fce8e6";
        border = "#ea4335";
    }
    // Limpiar markdown basico (negritas, etc.)
    var mensajeLimpio = msg.replace(/\*\*/g, "").replace(/\*/g, "");
    resultDiv.innerHTML = "<div class=\"result-box\" style=\"background:" + bg + "; border-left-color:" + border + ";\">" + mensajeLimpio.replace(/\n/g, "<br>") + "</div>";
}

function clearResult() {
    document.getElementById("result").innerHTML = "";
}

function showLoading(show) {
    var loader = document.getElementById("loading");
    if (show) {
        loader.classList.remove("hidden");
    } else {
        loader.classList.add("hidden");
    }
}