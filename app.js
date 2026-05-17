// ========== INICIALIZACIÓN ==========

Office.onReady(function () {

    console.log("Office.js listo");

    // Actualizar selección al abrir o volver al panel
    window.addEventListener("focus", function () {
        cargarSeleccionActual();
    });

    // Detectar consultas desde menú/contexto externo
    verificarDisparoExterno();

    window.addEventListener("storage", function(event) {

        if (
            event.key === "consultarDesdeMenu" &&
            event.newValue === "true"
        ) {
            verificarDisparoExterno();
        }
    });

    // Botón principal
    const explainBtn = document.getElementById("explainBtn");

    if (explainBtn) {
        explainBtn.onclick = ejecutarConsulta;
    }

    // Botón borrar todo
    const clearBtn = document.getElementById("clearAllBtn");

    if (clearBtn) {
        clearBtn.onclick = borrarTodo;
    }
});

// ========== VARIABLES GLOBALES ==========

let tipoExplicacion = "sencillo";

// ========== CONSULTA PRINCIPAL ==========

async function ejecutarConsulta() {

    showLoading(true);
    clearResult();

    Word.run(async function(context) {

        try {

            // Obtener texto seleccionado
            var selection = context.document.getSelection();
            selection.load("text");

            // Obtener párrafo completo como contexto
            var paragraph = selection.paragraphs.getFirst();
            paragraph.load("text");

            await context.sync();

            var selectedText = selection.text;

            // Validar selección
            if (!selectedText || selectedText.trim() === "") {

                showResult(
                    "No hay texto seleccionado. Selecciona una palabra o frase en tu documento.",
                    "error"
                );

                showLoading(false);
                return;
            }

            // Mostrar selección en pantalla
            updateSelectedTextDisplay(selectedText);

            // Contexto completo
            var contextoCompleto = paragraph.text || selectedText;

            // Consultar IA
            var explicacion = await consultarIA(
                selectedText,
                contextoCompleto
            );

            // Mostrar resultado
            showResult(explicacion, "success");

            showLoading(false);

        } catch (error) {

            console.error(error);

            showResult(
                "Error al procesar la solicitud: " + error.message,
                "error"
            );

            showLoading(false);
        }

    }).catch(function(error) {

        console.error(error);

        showResult(
            "Error al leer el documento. Asegúrate de estar en Word Online.",
            "error"
        );

        showLoading(false);
    });
}

// ========== CONSULTA IA ==========

async function consultarIA(palabra, contexto) {

    const tieneContextoSuficiente =
        contexto &&
        contexto.trim().length > 20;

    let prompt = "";

    if (tieneContextoSuficiente) {

        prompt = `
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
        - Explica el significado dentro del contexto proporcionado
        `;

    } else {

        prompt = `
        Define "${palabra}" de forma general.

        Tipo de explicación: ${tipoExplicacion}

        Instrucciones:
        - Español
        - Máximo 3 oraciones
        - Explicación clara y breve
        `;
    }

    try {

        // IMPORTANTE:
        // Mantener HTTP porque tu backend funciona así
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

        return "Error de conexión con el backend.";
    }
}

// ========== FUNCIONES UI ==========

function updateSelectedTextDisplay(text) {

    var el = document.getElementById("selectedDisplay");

    if (text && text.trim()) {

        var displayText =
            text.length > 100
            ? text.substring(0, 100) + "..."
            : text;

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

    // Limpiar markdown básico
    var mensajeLimpio =
        msg.replace(/\*\*/g, "")
           .replace(/\*/g, "");

    resultDiv.innerHTML =

        "<div class='result-box' style='background:" +
        bg +
        "; border-left-color:" +
        border +
        ";'>" +

        mensajeLimpio.replace(/\n/g, "<br>") +

        "</div>" +

        "<button onclick='nuevaConsulta()' class='new-query-btn'>" +
        "Nueva consulta" +
        "</button>";

    // Cambiar texto botón
    const btnText = document.querySelector("#explainBtn span");

    if (btnText) {
        btnText.textContent = "Generar otra explicación";
    }
}

function clearResult() {

    var resultDiv = document.getElementById("result");

    if (resultDiv) {
        resultDiv.innerHTML = "";
    }
}

function showLoading(show) {

    var loader = document.getElementById("loading");

    if (!loader) return;

    if (show) {

        loader.classList.remove("hidden");

    } else {

        loader.classList.add("hidden");
    }
}

function nuevaConsulta() {

    clearResult();

    updateSelectedTextDisplay("");

    const btnText = document.querySelector("#explainBtn span");

    if (btnText) {
        btnText.textContent = "Explicar";
    }
}

// ========== TIPO DE EXPLICACIÓN ==========

function setTipo(tipo) {

    tipoExplicacion = tipo;

    const slider = document.getElementById("pill-slider");

    if (!slider) return;

    const positions = {
        sencillo: "0%",
        tecnico: "100%",
        ejemplo: "200%"
    };

    slider.style.transform =
        `translateX(${positions[tipo]})`;

    slider.className = "pill-slider " + tipo;

    document
        .querySelectorAll(".pill-container button")
        .forEach(btn => {
            btn.classList.remove("active");
        });

    const activeBtn =
        document.getElementById("btn-" + tipo);

    if (activeBtn) {
        activeBtn.classList.add("active");
    }

    slider.animate([
        {
            transform:
            `translateX(${positions[tipo]}) scale(0.95)`
        },
        {
            transform:
            `translateX(${positions[tipo]}) scale(1)`
        }
    ], {
        duration: 200,
        easing: "cubic-bezier(0.4, 0, 0.2, 1)"
    });
}

// ========== SELECCIÓN ACTUAL ==========

function cargarSeleccionActual() {

    Word.run(function(context) {

        var selection =
            context.document.getSelection();

        selection.load("text");

        return context.sync().then(function() {

            updateSelectedTextDisplay(
                selection.text
            );
        });

    }).catch(function(error) {

        console.log(
            "No se pudo actualizar selección:",
            error
        );
    });
}

// ========== CONSULTA DESDE MENÚ EXTERNO ==========

function verificarDisparoExterno() {

    if (
        localStorage.getItem("consultarDesdeMenu")
        === "true"
    ) {

        localStorage.removeItem(
            "consultarDesdeMenu"
        );

        clearResult();

        setTimeout(function() {

            ejecutarConsulta();

        }, 100);
    }
}

// ========== BORRAR TODO ==========

function borrarTodo() {

    clearResult();

    updateSelectedTextDisplay("");

    showLoading(false);

    console.log("Todo borrado");
}

// ========== ANIMACIÓN INICIAL ==========

window.addEventListener("load", () => {

    document.querySelector(".pill-container")
    ?.animate([
        { transform: "scale(0.98)" },
        { transform: "scale(1)" }
    ], {
        duration: 120,
        easing: "ease-out"
    });
});