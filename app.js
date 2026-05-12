Office.onReady(function() {
    console.log("Office.js listo");
    
    // Actualizar seleccion al abrir o volver al panel
    window.addEventListener("focus", function () {

        cargarSeleccionActual();
    });
});

document.getElementById("explainBtn").onclick = function() {
    showLoading(true);
    clearResult();
    
    Word.run(async function(context) {
        try {
            // Obtener el texto seleccionado
            var selection = context.document.getSelection();
            selection.load("text");
            
            // Obtener automaticamente el parrafo donde esta la seleccion
            var paragraph = selection.paragraphs.getFirst();
            paragraph.load("text");
            
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
            var contextoCompleto = paragraph.text || selectedText;
            
            // Llamar API para obtener la explicacion
            var explicacion = await consultarIA(selectedText, contextoCompleto);
            
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

async function consultarIA(palabra, contexto) {

    const prompt = `Eres un asistente academico especializado en explicar terminos en contexto.

El usuario ha seleccionado: "${palabra}"
Contexto: "${contexto}"
Explica de forma clara y breve (max 3 oraciones).`;
    try {

        const respuesta = await fetch("http://localhost:3000/ia", {
            method: "POST",
            headers: {
                "Content-Type": "application/json"
            },
            body: JSON.stringify({ prompt })
        });
        const data = await respuesta.json();

        if (!respuesta.ok) {
            console.error("Error backend:", data);
            return data.error || "Error al conectar con el servidor.";
        }

        return data.respuesta || "Sin respuesta del modelo.";

    } catch (error) {

        console.error("Error fetch:", error);
        return "Error de conexion con el backend.";
    }
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
    resultDiv.innerHTML =
    "<div class=\"result-box\" style=\"background:" + bg + "; border-left-color:" + border + ";\">" +
    mensajeLimpio.replace(/\n/g, "<br>") +
    "<br><br>" +
    "<button onclick='nuevaConsulta()' class='new-query-btn'>Nueva consulta</button>" +
    "</div>";
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

function nuevaConsulta() {
    clearResult();
    updateSelectedTextDisplay("");
}

function cargarSeleccionActual() {

    Word.run(function (context) {

        var selection = context.document.getSelection();
        selection.load("text");

        return context.sync().then(function () {

            updateSelectedTextDisplay(selection.text);

        });

    }).catch(function (error) {

        console.log("No se pudo actualizar seleccion:", error);

    });
}