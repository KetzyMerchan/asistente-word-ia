// ========== INICIALIZACIÓN Y ESCUCHADORES ==========

Office.onReady(function(info) {
    console.log("Office.js listo - Asistente de Comprensión");
    
    if (info.host === Office.HostType.Word) {
        updateSelectedText();
        verificarDisparoExterno();

        window.addEventListener("storage", function(event) {
            if (event.key === "consultarDesdeMenu" && event.newValue === "true") {
                verificarDisparoExterno();
            }
        });

        //   Detectar foco en el panel
        window.addEventListener('focus', function() {
            console.log("Panel enfocado");
            if (localStorage.getItem("consultarDesdeMenu") === "true") {
                verificarDisparoExterno();
            }
        });

        //   Actualizar texto seleccionado al hacer clic en el panel
        window.addEventListener('focus', function() {
            console.log("Actualizar selección por foco");
            cargarSeleccionActual();
        });

        const btn = document.getElementById("explainBtn");
        if (btn) {
            btn.onclick = ejecutarConsulta;
        }

        //  Botón de papelera - Borrar todo
        const clearBtn = document.getElementById("clearAllBtn");
        if (clearBtn) {
            clearBtn.onclick = borrarTodo;
        }
    }
});

let tipoExplicacion = "sencillo";

function verificarDisparoExterno() {
    if (localStorage.getItem("consultarDesdeMenu") === "true") {
        localStorage.removeItem("consultarDesdeMenu");
        
        //  Borrado directo y forzado
        var resultDiv = document.getElementById("result");
        if (resultDiv) {
            resultDiv.innerHTML = "";
            console.log("Resultado anterior borrado");
        }
        
        setTimeout(function() {
            ejecutarConsulta();
        }, 100);
    }
}

// ========== LÓGICA DE CONSULTA E IA ==========

async function consultarIA(palabra, contexto) {
    const tieneContextoSuficiente = contexto && contexto.trim().length > 20;

    let prompt = "";
    if (tieneContextoSuficiente) {
        prompt = `Explica qué significa "${palabra}" en esta oración: "${contexto}". 
Tipo de explicación: ${tipoExplicacion}

    Instrucciones según tipo:
    - sencillo: explicación básica para estudiantes principiantes
    - tecnico: explicación formal, académica y precisa
    - ejemplo: incluye un ejemplo práctico para facilitar comprensión

Instrucciones:
- Respuesta de máximo 35 palabras.
- No copies la oración del contexto.
- Da una definición clara y, si es relevante, su propósito o una característica clave.
- Usa la estructura: "[Término] es un/una [definición]" o "[Término] se refiere a [definición]".`;
    } else {
        prompt = `Define "${palabra}" de forma general, indicando su propósito o los campos donde se usa. Usa la estructura: "[Término] es un/una [definición]" o "[Término] se refiere a [definición]". Máximo 35 palabras.`;
    }

    try {
        const respuesta = await fetch("http://localhost:3000/ia", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
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

function ejecutarConsulta() {
    showLoading(true);
    clearResult();

    Word.run(async function(context) {
        try {
            var selection = context.document.getSelection();
            selection.load("text");
            
            var paragraph = selection.paragraphs.getFirst();
            paragraph.load("text");
            
            await context.sync();
            
            var selectedText = selection.text;
            
            if (!selectedText || selectedText.trim() === "") {
                showResult("No hay texto seleccionado. Selecciona una palabra o frase en tu documento.", "error");
                showLoading(false);
                return;
            }

            updateSelectedTextDisplay(selectedText);
            
            var contextoCompleto = paragraph.text || selectedText;
            
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
        showResult("Error al leer el documento. Asegúrate de estar en Word Online.", "error");
        showLoading(false);
    });
}

// ========== FUNCIONES DE INTERFAZ (UI) ==========

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
    
    resultDiv.innerHTML =
    "<div class='result-box' style='background:" + bg + "; border-left-color:" + border + ";'>" +
        mensajeLimpio.replace(/\n/g, "<br>") +
    "</div>";

    // Cambiar texto del botón después de responder
    document.querySelector("#explainBtn span").textContent =
        "Generar otra explicación";
}

function clearResult() {
    var resultDiv = document.getElementById("result");
    if (resultDiv) {
        resultDiv.innerHTML = "";
        console.log("clearResult ejecutado");
    }
}

function showLoading(show) {
    var loader = document.getElementById("loading");
    if (loader) {
        if (show) {
            loader.classList.remove("hidden");
        } else {
            loader.classList.add("hidden");
        }
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

// ==========  NUEVA FUNCIÓN: BORRAR TODO ==========

function borrarTodo() {
    // Borrar resultado de la IA
    const resultDiv = document.getElementById("result");
    if (resultDiv) {
        resultDiv.innerHTML = "";
    }
    
    // Limpiar el texto seleccionado mostrado
    const selectedDisplay = document.getElementById("selectedDisplay");
    if (selectedDisplay) {
        selectedDisplay.innerHTML = "Ninguno";
        selectedDisplay.classList.add("empty");
    }
    
    // Ocultar loading si está visible
    const loader = document.getElementById("loading");
    if (loader) {
        loader.classList.add("hidden");
    }
    
    console.log(" Todo borrado por el usuario");
}

// ==========  NUEVA FUNCIÓN: CARGAR SELECCIÓN ACTUAL ==========

function cargarSeleccionActual() {
    Word.run(function(context) {
        var selection = context.document.getSelection();
        selection.load("text");
        return context.sync().then(function() {
            updateSelectedTextDisplay(selection.text);
        });
    }).catch(function(error) {
        console.log("No se pudo actualizar selección:", error);
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