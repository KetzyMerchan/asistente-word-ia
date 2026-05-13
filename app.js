Office.onReady(function() {
    console.log("Office.js listo");
    
    // Actualizar seleccion al abrir o volver al panel
    window.addEventListener("focus", function () {

        cargarSeleccionActual();
    });
});

let tipoExplicacion = "sencillo";

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

    const prompt = `
    Eres un asistente académico integrado en Microsoft Word.

    Texto seleccionado:
    "${palabra}"

    Contexto:
    "${contexto}"

    Tipo de explicación: ${tipoExplicacion}

    Instrucciones según tipo:
    - sencillo: explicación básica para estudiantes principiantes
    - tecnico: explicación formal, académica y precisa
    - ejemplo: incluye un ejemplo práctico para facilitar comprensión

    Reglas:
    - Español
    - Máximo 3 oraciones
    - No inventar información
    - No repetir el contexto
    `;

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
    "<div class='result-box' style='background:" + bg + "; border-left-color:" + border + ";'>" +
        mensajeLimpio.replace(/\n/g, "<br>") +
    "</div>" +

    "<button onclick='nuevaConsulta()' class='new-query-btn'>" +
        "Nueva consulta" +
    "</button>";

    // Cambiar texto del botón después de responder
    document.querySelector("#explainBtn span").textContent =
        "Generar otra explicación";
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
    // Restaurar texto original del botón
    document.querySelector("#explainBtn span").textContent =
        "Explicar";
}

function setTipo(tipo) {
    tipoExplicacion = tipo;

    const slider = document.getElementById("pill-slider");

    if (!slider) return;

    // mover slider (mejor controlado)
    const positions = {
        sencillo: "0%",
        tecnico: "100%",
        ejemplo: "200%"
    };

    slider.style.transform = `translateX(${positions[tipo]})`;

    // cambiar color dinámico del slider
    slider.className = "pill-slider " + tipo;

    // sincronizar botones visuales
    document.querySelectorAll(".pill-container button").forEach(btn => {
        btn.classList.remove("active");
    });

    const activeBtn = document.getElementById("btn-" + tipo);
    if (activeBtn) activeBtn.classList.add("active");

    // animación correcta (fluida real)
    slider.animate([
        { transform: `translateX(${positions[tipo]}) scale(0.95)` },
        { transform: `translateX(${positions[tipo]}) scale(1)` }
    ], {
        duration: 200,
        easing: "cubic-bezier(0.4, 0, 0.2, 1)"
    });
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

// vibración visual suave extra
window.addEventListener("load", () => {
    document.querySelector(".pill-container")?.animate([
        { transform: "scale(0.98)" },
        { transform: "scale(1)" }
    ], {
        duration: 120,
        easing: "ease-out"
    });
});