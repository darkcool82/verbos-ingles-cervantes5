/* =====================================
   VARIABLES GLOBALES
===================================== */

let workbook = null;

let verbos = [];

let testActual = [];

let errores = [];

let resultados = [];

let historico = JSON.parse(
    localStorage.getItem("historico")
) || [];

let temporizador = null;

let segundosRestantes = 0;
let deferredPrompt = null;

/* =====================================
   INICIO
===================================== */

document.addEventListener(
    "DOMContentLoaded",
    async () => {

        registrarEventos();

        cargarHistorico();

        await cargarExcelPorDefecto();

    }
);

/* =====================================
   EVENTOS
===================================== */

function registrarEventos() {

    document
        .getElementById("btnImportar")
        .addEventListener(
            "click",
            importarExcel
        );

    document
        .getElementById("btnComenzar")
        .addEventListener(
            "click",
            iniciarTest
        );

    document
        .getElementById("btnEstudiar")
        .addEventListener(
            "click",
            estudiarTodos
        );

    document
        .getElementById(
            "btnRepetirErrores"
        )
        .addEventListener(
            "click",
            repetirErrores
        );

    document
        .getElementById(
            "btnExportarErrores"
        )
        .addEventListener(
            "click",
            exportarErrores
        );

    document
        .getElementById(
            "btnExportarResultados"
        )
        .addEventListener(
            "click",
            exportarResultados
        );

    document
        .getElementById(
            "btnHistorico"
        )
        .addEventListener(
            "click",
            mostrarHistorico
        );
}

/* =====================================
   CARGA AUTOMÁTICA EXCEL
===================================== */

async function cargarExcelPorDefecto() {

    try {

        const respuesta =
            await fetch(
                "data/Verbos_ingles.xlsx"
            );

        const buffer =
            await respuesta.arrayBuffer();

        workbook =
            XLSX.read(buffer);

        generarSelectorHojas();

    }
    catch(error) {

        console.error(error);

        console.log(
            "No se ha encontrado el Excel por defecto."
        );
    }
}

/* =====================================
   IMPORTAR OTRO EXCEL
===================================== */

async function importarExcel() {

    const archivo =
        document.getElementById(
            "excelFile"
        ).files[0];

    if(!archivo) {

        alert(
            "Seleccione un Excel"
        );

        return;
    }

    const buffer =
        await archivo.arrayBuffer();

    workbook =
        XLSX.read(buffer);

    generarSelectorHojas();

    alert(
        "Excel cargado correctamente"
    );
}

/* =====================================
   SELECTOR DE HOJAS
===================================== */

function generarSelectorHojas() {

    const contenedor =
        document.getElementById(
            "listaHojas"
        );

    contenedor.innerHTML = "";

    workbook.SheetNames
    .forEach(nombreHoja => {

        const item =
            document.createElement(
                "div"
            );

        item.className =
            "hoja-item";

        item.innerHTML = `
            <label>

                <input
                    type="checkbox"
                    class="hojaCheck"
                    value="${nombreHoja}">

                ${nombreHoja}

            </label>
        `;

        contenedor.appendChild(
            item
        );

    });
}

/* =====================================
   BOTÓN ESTUDIAR
===================================== */

function estudiarTodos() {

    if (!workbook) {

        alert(
            "No hay Excel cargado."
        );

        return;
    }

    const hojas =
        obtenerHojasSeleccionadas();

    if (
        hojas.length === 0
    ) {

        alert(
            "Seleccione al menos una hoja."
        );

        return;
    }

    const verbosEstudio =
        cargarVerbosSeleccionados();

    mostrarModoEstudio(
        verbosEstudio,
        hojas
    );
}

/* =====================================
   OBTENER VERBOS DE UNA HOJA
===================================== */

function obtenerVerbosHoja(
    nombreHoja
) {

    const hoja =
        workbook.Sheets[
            nombreHoja
        ];

    const filas =
        XLSX.utils.sheet_to_json(
            hoja
        );

    return filas.map(f => ({

        infinitivo:
            f["INFINITIVE"],

        pasado:
            f["PAST SIMPLE"],

        participio:
            f["PAST PARTICIPLE"],

        meaning:
            f["MEANING"]

    }));

}
/* =====================================
   HOJAS SELECCIONADAS
===================================== */

function obtenerHojasSeleccionadas() {

    const checks =
        document.querySelectorAll(
            ".hojaCheck:checked"
        );

    return [...checks]
        .map(c => c.value);
}

/* =====================================
   VERBOS SELECCIONADOS
===================================== */


function actualizarResumen() {

    const hojas =
        obtenerHojasSeleccionadas();

    let total = 0;

    hojas.forEach(h => {

        total +=
            obtenerVerbosHoja(
                h
            ).length;

    });

    document
        .getElementById(
            "resumenSeleccion"
        )
        .innerHTML =

        `
        Hojas seleccionadas:
        ${hojas.length}

        <br>

        Verbos disponibles:
        ${total}
        `;
}


/* =====================================
   UNIR VERBOS DE VARIAS HOJAS
===================================== */

function cargarVerbosSeleccionados() {

    const hojas =
        obtenerHojasSeleccionadas();

    let lista = [];

    hojas.forEach(nombre => {

        lista = lista.concat(
            obtenerVerbosHoja(
                nombre
            )
        );

    });

    return lista;
}

/* =====================================
   INICIAR TEST
===================================== */

function iniciarTest() {

    if (!workbook) {

        alert(
            "No hay Excel cargado"
        );

        return;
    }

    const hojas =
        obtenerHojasSeleccionadas();

    if (hojas.length === 0) {

        alert(
            "Seleccione al menos una hoja"
        );

        return;
    }

    verbos =
        cargarVerbosSeleccionados();

    let cantidad =
        parseInt(
            document.getElementById(
                "cantidadPreguntas"
            ).value
        );

    if (
        cantidad > verbos.length
    ) {

        cantidad =
            verbos.length;
    }

    testActual =
        [...verbos]
        .sort(
            () =>
            Math.random() - 0.5
        )
        .slice(
            0,
            cantidad
        );

    renderizarPreguntas();

    iniciarTemporizadorSiProcede();

    window.scrollTo({
        top: 0,
        behavior: "smooth"
    });

}

/* =====================================
   RENDERIZAR PREGUNTAS
===================================== */

function renderizarPreguntas() {

    const contenedor =
        document.getElementById(
            "contenido"
        );

    let html = "";

    testActual.forEach(
    (v, i) => {

        html += `

        <div
            class="pregunta">

            <h3>
                ${v.meaning}
            </h3>

            <div
                class="fila-verbos">

                <div
                    class="campo">

                    <input
                        type="text"

                        id="inf_${i}"

                        placeholder="INFINITIVE">

                    <div
                        id="error_inf_${i}"
                        class="error-correccion">
                    </div>

                </div>

                <div
                    class="campo">

                    <input
                        type="text"

                        id="pas_${i}"

                        placeholder="PAST SIMPLE">

                    <div
                        id="error_pas_${i}"
                        class="error-correccion">
                    </div>

                </div>

                <div
                    class="campo">

                    <input
                        type="text"

                        id="par_${i}"

                        placeholder="PAST PARTICIPLE">

                    <div
                        id="error_par_${i}"
                        class="error-correccion">
                    </div>

                </div>

            </div>

        </div>

        `;
    });

    html += `

        <button
            id="btnCorregir"
            onclick="corregir()">

            Corregir

        </button>

    `;

    contenedor.innerHTML =
        html;

    document
        .getElementById(
            "resultado"
        )
        .innerHTML = "";

}

/* =====================================
   TEMPORIZADOR
===================================== */

function iniciarTemporizadorSiProcede() {

    clearInterval(
        temporizador
    );

    const usar =

        document
        .getElementById(
            "usarTemporizador"
        )
        .checked;

    if (!usar) {

        document
        .getElementById(
            "temporizador"
        )
        .innerHTML = "";

        return;
    }

    segundosRestantes =
        testActual.length * 15;

    actualizarTemporizador();

    temporizador =
        setInterval(
            actualizarTemporizador,
            1000
        );
}

function actualizarTemporizador() {

    document
        .getElementById(
            "temporizador"
        )
        .innerHTML =

        `
        ⏱️
        ${segundosRestantes}
        segundos
        `;

    segundosRestantes--;

    if (
        segundosRestantes < 0
    ) {

        clearInterval(
            temporizador
        );

        corregir();
    }
}
/* =====================================
   CORREGIR TEST
===================================== */

function corregir() {

    clearInterval(
        temporizador
    );

    errores = [];

    resultados = [];

    let aciertos = 0;

    testActual.forEach(
    (verbo, i) => {

        const infInput =
            document.getElementById(
                `inf_${i}`
            );

        const pasInput =
            document.getElementById(
                `pas_${i}`
            );

        const parInput =
            document.getElementById(
                `par_${i}`
            );

        const infError =
            document.getElementById(
                `error_inf_${i}`
            );

        const pasError =
            document.getElementById(
                `error_pas_${i}`
            );

        const parError =
            document.getElementById(
                `error_par_${i}`
            );

        infError.innerHTML = "";
        pasError.innerHTML = "";
        parError.innerHTML = "";

        infInput.className = "";
        pasInput.className = "";
        parInput.className = "";

        let verboCorrecto = true;

        /* =====================
           INFINITIVE
        ===================== */

        const infUsuario =

            infInput.value
            .trim()
            .toLowerCase();

        const infCorrecto =

            String(
                verbo.infinitivo
            )
            .toLowerCase();

        if (
            infUsuario ===
            infCorrecto
        ) {

            infInput.classList.add(
                "correcto-input"
            );

        }
        else {

            verboCorrecto = false;

            infInput.classList.add(
                "incorrecto-input"
            );

            infError.innerHTML =

                verbo.infinitivo;
        }

        /* =====================
           PASADO
        ===================== */

        const pasUsuario =

            pasInput.value
            .trim()
            .toLowerCase();

        const pasCorrecto =

            String(
                verbo.pasado
            )
            .toLowerCase();

        if (
            pasUsuario ===
            pasCorrecto
        ) {

            pasInput.classList.add(
                "correcto-input"
            );

        }
        else {

            verboCorrecto = false;

            pasInput.classList.add(
                "incorrecto-input"
            );

            pasError.innerHTML =

                verbo.pasado;
        }

        /* =====================
           PARTICIPIO
        ===================== */

        const parUsuario =

            parInput.value
            .trim()
            .toLowerCase();

        const parCorrecto =

            String(
                verbo.participio
            )
            .toLowerCase();

        if (
            parUsuario ===
            parCorrecto
        ) {

            parInput.classList.add(
                "correcto-input"
            );

        }
        else {

            verboCorrecto = false;

            parInput.classList.add(
                "incorrecto-input"
            );

            parError.innerHTML =

                verbo.participio;
        }

        /* =====================
           RESULTADO DEL VERBO
        ===================== */

        if (
            verboCorrecto
        ) {

            aciertos++;

        }
        else {

            errores.push(
                verbo
            );
        }

        resultados.push({

            MEANING:
                verbo.meaning,

            ACIERTO:
                verboCorrecto
                ? "SI"
                : "NO"

        });

    });

    mostrarResultadoFinal(
        aciertos
    );
}

/* =====================================
   RESULTADO FINAL
===================================== */

function mostrarResultadoFinal(
    aciertos
) {

    const total =
        testActual.length;

    const erroresCantidad =
        total - aciertos;

    const porcentaje =

        Math.round(
            (
                aciertos
                * 100
            )
            /
            total
        );

    const nota =

        (
            aciertos
            * 10
        )
        /
        total;

    const notaFinal =
        nota.toFixed(1);

    document
        .getElementById(
            "resultado"
        )
        .innerHTML =

        `
        <div
            class="resultado-box">

            <h2>

                Nota:

                ${notaFinal}/10

            </h2>

            <p>

                ✅ Aciertos:
                ${aciertos}

            </p>

            <p>

                ❌ Errores:
                ${erroresCantidad}

            </p>

            <p>

                📊 Porcentaje:
                ${porcentaje}%

            </p>

        </div>
        `;

    guardarResultadoHistorico(
        notaFinal,
        porcentaje,
        aciertos,
        erroresCantidad,
        total
    );

    window.scrollTo({

        top: 0,

        behavior:
        "smooth"

    });

}

/* =====================================
   HISTÓRICO
===================================== */

function guardarResultadoHistorico(

    nota,

    porcentaje,

    aciertos,

    errores,

    total

) {

    const registro = {

        fecha:

            new Date()
            .toLocaleString(),

        nota,

        porcentaje,

        aciertos,

        errores,

        total

    };

    historico.push(
        registro
    );

    localStorage.setItem(

        "historico",

        JSON.stringify(
            historico
        )

    );
}

/* =====================================
   CARGAR HISTÓRICO
===================================== */

function cargarHistorico() {

    historico =

        JSON.parse(

            localStorage.getItem(
                "historico"
            )

        )

        ||

        [];

}
/* =====================================
   MOSTRAR HISTÓRICO
===================================== */

function mostrarHistorico() {

    if (
        historico.length === 0
    ) {

        alert(
            "Todavía no hay registros."
        );

        return;
    }

    let texto =

        "HISTÓRICO\n\n";

    historico
        .slice()
        .reverse()
        .forEach(r => {

            texto +=

`📅 ${r.fecha}

Nota: ${r.nota}/10

Aciertos: ${r.aciertos}
Errores: ${r.errores}

Porcentaje:
${r.porcentaje}%

------------------------

`;

        });

    alert(
        texto
    );
}

/* =====================================
   REPETIR ERRORES
===================================== */

function repetirErrores() {

    if (
        errores.length === 0
    ) {

        alert(
            "No hay errores para repetir."
        );

        return;
    }

    clearInterval(
        temporizador
    );

    testActual =
        [...errores];

    renderizarPreguntas();

    iniciarTemporizadorSiProcede();

    window.scrollTo({

        top: 0,

        behavior:
        "smooth"

    });
}

/* =====================================
   EXPORTAR ERRORES
===================================== */

function exportarErrores() {

    if (
        errores.length === 0
    ) {

        alert(
            "No existen errores."
        );

        return;
    }

    const datos =

        errores.map(v => ({

            "INFINITIVE":
                v.infinitivo,

            "PAST SIMPLE":
                v.pasado,

            "PAST PARTICIPLE":
                v.participio,

            "MEANING":
                v.meaning

        }));

    const ws =

        XLSX.utils
        .json_to_sheet(
            datos
        );

    const wb =

        XLSX.utils
        .book_new();

    XLSX.utils
    .book_append_sheet(

        wb,

        ws,

        "Errores"

    );

    XLSX.writeFile(

        wb,

        "Errores.xlsx"

    );
}

/* =====================================
   EXPORTAR RESULTADOS
===================================== */

function exportarResultados() {

    if (
        resultados.length === 0
    ) {

        alert(
            "No existen resultados."
        );

        return;
    }

    const ws =

        XLSX.utils
        .json_to_sheet(
            resultados
        );

    const wb =

        XLSX.utils
        .book_new();

    XLSX.utils
    .book_append_sheet(

        wb,

        ws,

        "Resultados"

    );

    XLSX.writeFile(

        wb,

        "Resultados.xlsx"

    );
}

/* =====================================
   LIMPIAR HISTÓRICO
===================================== */

function limpiarHistorico() {

    const confirmar =

        confirm(
            "¿Desea borrar el histórico?"
        );

    if (
        !confirmar
    ) {

        return;
    }

    historico = [];

    localStorage.removeItem(
        "historico"
    );

    alert(
        "Histórico eliminado."
    );
}

/* =====================================
   UTILIDAD MEZCLAR ARRAY
===================================== */

function mezclar(lista) {

    return [...lista]

    .sort(
        () =>
        Math.random() - 0.5
    );
}

/* =====================================
   UTILIDAD FECHA
===================================== */

function fechaActual() {

    return new Date()
        .toLocaleString(
            "es-ES"
        );
}


/* =====================================
   MODO ESTUDIO
===================================== */



function mostrarModoEstudio(
    verbos,
    hojas
) {

    const contenedor =
        document.getElementById(
            "contenido"
        );

    let html = `

        <div class="panel">

            <h2>

                📘 Modo estudio

            </h2>

            <p>

                Hojas seleccionadas:
                ${hojas.join(", ")}

            </p>

            <p>

                Total verbos:
                ${verbos.length}

            </p>

        </div>

    `;

    html += `

        <div
            class="tabla-estudio">

            <div
                class="cabecera-estudio">

                <div>
                    INFINITIVE
                </div>

                <div>
                    PAST SIMPLE
                </div>

                <div>
                    PAST PARTICIPLE
                </div>

                <div>
                    MEANING
                </div>

            </div>

    `;

    verbos.forEach(v => {

        html += `

            <div
                class="fila-estudio">

                <div>
                    ${v.infinitivo}
                </div>

                <div>
                    ${v.pasado}
                </div>

                <div>
                    ${v.participio}
                </div>

                <div>
                    ${v.meaning}
                </div>

            </div>

        `;
    });

    html += `
        </div>
    `;

    contenedor.innerHTML =
        html;

    document
        .getElementById(
            "resultado"
        )
        .innerHTML = "";

    window.scrollTo({

        top: 0,

        behavior:
        "smooth"

    });
}

/* =====================================
   INSTALACIÓN PWA
===================================== */

window.addEventListener(

    "beforeinstallprompt",

    event => {

        event.preventDefault();

        deferredPrompt = event;

        const boton =

            document.getElementById(
                "btnInstalar"
            );

        boton.hidden = false;

    }

);

document.addEventListener(

    "click",

    async event => {

        if (
            event.target.id !==
            "btnInstalar"
        ) {

            return;
        }

        if (
            !deferredPrompt
        ) {

            return;
        }

        deferredPrompt.prompt();

        const resultado =

            await deferredPrompt
            .userChoice;

        if (
            resultado.outcome ===
            "accepted"
        ) {

            console.log(
                "PWA instalada"
            );

        }

        deferredPrompt = null;

        document
        .getElementById(
            "btnInstalar"
        )
        .hidden = true;

    }

);

window.addEventListener(

    "appinstalled",

    () => {

        console.log(
            "Aplicación instalada"
        );

        const boton =

            document.getElementById(
                "btnInstalar"
            );

        boton.hidden = true;

    }

);
