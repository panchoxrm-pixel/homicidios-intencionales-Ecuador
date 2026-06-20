// ============================================================================
// BASE DE DATOS REAL EXTRACTA DE LOS EXCEL (CSV)
// ============================================================================

const dataMaster = {
    "2014": {
        meses: ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"],
        urbano: [99, 69, 83, 71, 79, 71, 55, 82, 64, 57, 62, 76],
        rural: [52, 32, 46, 49, 28, 44, 27, 32, 26, 25, 44, 37],
        hombres: [126, 88, 110, 105, 88, 100, 74, 99, 78, 67, 94, 95],
        mujeres: [25, 13, 19, 15, 19, 15, 8, 15, 12, 15, 12, 18],
        rangosEdad: { "Niño": 30, "Adolescente": 52, "Joven": 519, "Adulto": 646, "Adulto mayor": 63 },
        edadMedia: 34.01,
        edadMediana: 31,
        edadModa: ["20"],
        edadModaFrecuencia: 57,
        edadModaRelativa: "4.35%"
    },
    "2015": {
        meses: ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"],
        urbano: [54, 55, 81, 59, 68, 65, 47, 57, 49, 53, 52, 57],
        rural: [44, 27, 25, 26, 27, 35, 28, 21, 33, 36, 30, 21],
        hombres: [87, 73, 83, 74, 83, 82, 59, 63, 67, 75, 69, 62],
        mujeres: [11, 9, 23, 11, 12, 18, 16, 15, 15, 14, 13, 16],
        rangosEdad: { "Niño": 29, "Adolescente": 37, "Joven": 391, "Adulto": 529, "Adulto mayor": 64 },
        edadMedia: 35.32,
        edadMediana: 32,
        edadModa: ["28"],
        edadModaFrecuencia: 49,
        edadModaRelativa: "4.67%"
    },
    "2016": {
        meses: ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"],
        urbano: [68, 49, 48, 58, 60, 50, 50, 43, 49, 40, 47, 44],
        rural: [31, 24, 30, 19, 29, 34, 37, 29, 33, 16, 42, 29],
        hombres: [83, 54, 67, 66, 68, 65, 73, 60, 65, 46, 69, 62],
        mujeres: [16, 19, 11, 11, 21, 18, 14, 12, 17, 10, 17, 11],
        sinDatos: [4],
        rangosEdad: { "Niño": 24, "Adolescente": 35, "Joven": 331, "Adulto": 505, "Adulto mayor": 58 },
        edadMedia: 35.34,
        edadMediana: 32,
        edadModa: ["25"],
        edadModaFrecuencia: 38,
        edadModaRelativa: "3.96%"
    },
    "2017": {
        meses: ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"],
        urbano: [59, 52, 58, 50, 62, 55, 51, 40, 51, 40, 43, 50],
        rural: [36, 38, 35, 36, 28, 25, 24, 28, 20, 29, 22, 38],
        hombres: [69, 68, 77, 70, 75, 62, 58, 51, 59, 60, 56, 69],
        mujeres: [26, 22, 16, 16, 15, 18, 17, 17, 12, 9, 9, 19],
        rangosEdad: { "Niño": 17, "Adolescente": 30, "Joven": 339, "Adulto": 532, "Adulto mayor": 51 },
        edadMedia: 35.91,
        edadMediana: 33,
        edadModa: ["22", "24"],
        edadModaFrecuencia: 34,
        edadModaRelativa: "3.51%"
    },
    "2018": {
        meses: ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"],
        urbano: [48, 60, 52, 72, 67, 53, 49, 53, 69, 47, 51, 47],
        rural: [25, 26, 37, 27, 31, 23, 22, 23, 20, 29, 32, 33],
        hombres: [62, 77, 73, 85, 82, 66, 64, 65, 76, 65, 74, 67],
        mujeres: [11, 9, 16, 14, 16, 10, 7, 11, 13, 11, 9, 13],
        rangosEdad: { "Niño": 16, "Adolescente": 24, "Joven": 373, "Adulto": 520, "Adulto mayor": 63 },
        edadMedia: 35.92,
        edadMediana: 33,
        edadModa: ["24", "26"],
        edadModaFrecuencia: 41,
        edadModaRelativa: "4.12%"
    },
    "2019": {
        meses: ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"],
        urbano: [59, 56, 54, 77, 72, 62, 64, 56, 74, 73, 82, 81],
        rural: [24, 31, 29, 41, 31, 39, 26, 33, 21, 42, 29, 33],
        hombres: [73, 78, 69, 104, 92, 87, 78, 84, 81, 99, 93, 98],
        mujeres: [9, 9, 14, 14, 11, 14, 12, 5, 14, 15, 18, 16],
        sinDatos: [2],
        rangosEdad: { "Niño": 19, "Adolescente": 32, "Joven": 416, "Adulto": 653, "Adulto mayor": 64 },
        edadMedia: 35.39,
        edadMediana: 33,
        edadModa: ["25"],
        edadModaFrecuencia: 56,
        edadModaRelativa: "4.71%"
    },
    "2020": {
        meses: ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"],
        urbano: [86, 87, 55, 65, 62, 76, 70, 88, 75, 101, 113, 114],
        rural: [27, 29, 25, 23, 34, 39, 24, 34, 30, 37, 35, 43],
        hombres: [101, 104, 71, 74, 82, 96, 84, 109, 93, 128, 125, 139],
        mujeres: [12, 12, 9, 14, 14, 19, 10, 13, 12, 10, 23, 17],
        sinDatos: [1],
        rangosEdad: { "Niño": 29, "Adolescente": 35, "Joven": 431, "Adulto": 813, "Adulto mayor": 56 },
        edadMedia: 35.42,
        edadMediana: 34,
        edadModa: ["35"],
        edadModaFrecuencia: 66,
        edadModaRelativa: "4.81%"
    },
    "2021": {
        meses: ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"],
        urbano: [91, 179, 134, 136, 124, 132, 145, 159, 292, 176, 238, 180],
        rural: [31, 47, 39, 40, 48, 38, 53, 44, 33, 49, 44, 43],
        hombres: [113, 213, 147, 151, 150, 155, 180, 187, 300, 203, 262, 208],
        mujeres: [9, 13, 26, 25, 22, 15, 18, 16, 25, 22, 20, 15],
        rangosEdad: { "Niño": 23, "Adolescente": 57, "Joven": 935, "Adulto": 1384, "Adulto mayor": 56 },
        edadMedia: 33.53,
        edadMediana: 32,
        edadModa: ["30"],
        edadModaFrecuencia: 119,
        edadModaRelativa: "4.77%"
    },
    "2022": {
        meses: ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"],
        urbano: [257, 255, 280, 318, 364, 304, 323, 329, 390, 388, 354, 428],
        rural: [61, 62, 70, 55, 71, 59, 88, 72, 93, 89, 71, 105],
        hombres: [295, 286, 311, 337, 402, 327, 369, 363, 440, 435, 393, 481],
        mujeres: [22, 27, 33, 36, 33, 35, 40, 35, 40, 41, 30, 52],
        sinDatos: [23],
        rangosEdad: { "Niño": 46, "Adolescente": 158, "Joven": 2070, "Adulto": 2439, "Adulto mayor": 90 },
        edadMedia: 32.38,
        edadMediana: 30,
        edadModa: ["30"],
        edadModaFrecuencia: 242,
        edadModaRelativa: "4.95%"
    },
    "2023": {
        meses: ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"],
        urbano: [438, 430, 532, 534, 557, 582, 659, 480, 653, 643, 598, 601],
        rural: [96, 75, 127, 111, 105, 128, 136, 122, 128, 190, 168, 155],
        hombres: [500, 455, 596, 601, 618, 663, 723, 553, 715, 768, 712, 702],
        mujeres: [33, 47, 61, 42, 43, 40, 69, 49, 60, 61, 48, 50],
        sinDatos: [39],
        rangosEdad: { "Niño": 62, "Adolescente": 315, "Joven": 3483, "Adulto": 4006, "Adulto mayor": 113 },
        edadMedia: 31.89,
        edadMediana: 30,
        edadModa: ["30"],
        edadModaFrecuencia: 471,
        edadModaRelativa: "5.71%"
    },
    "2024": {
        meses: ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"],
        urbano: [387, 259, 401, 401, 449, 459, 445, 454, 487, 548, 523, 549],
        rural: [120, 129, 135, 125, 136, 152, 155, 167, 123, 152, 143, 164],
        hombres: [459, 351, 492, 491, 543, 561, 564, 574, 561, 640, 616, 655],
        mujeres: [44, 34, 43, 32, 40, 46, 35, 44, 47, 56, 47, 53],
        sinDatos: [35],
        rangosEdad: { "Niño": 42, "Adolescente": 373, "Joven": 2879, "Adulto": 3467, "Adulto mayor": 128 },
        edadMedia: 32.32,
        edadMediana: 30,
        edadModa: ["30"],
        edadModaFrecuencia: 315,
        edadModaRelativa: "4.46%"
    },
    "2025": {
        meses: ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"],
        urbano: [645, 612, 708, 574, 742, 446, 459, 572, 613, 550, 665, 624],
        rural: [155, 139, 149, 168, 193, 128, 178, 199, 158, 232, 193, 179],
        hombres: [727, 687, 791, 688, 831, 519, 591, 685, 701, 707, 781, 727],
        mujeres: [71, 63, 62, 49, 98, 51, 45, 74, 67, 72, 70, 66],
        sinDatos: [58],
        rangosEdad: { "Niño": 46, "Adolescente": 539, "Joven": 3824, "Adulto": 4433, "Adulto mayor": 144 },
        edadMedia: 30.00,
        edadMediana: 30,
        edadModa: ["25"],
        edadModaFrecuencia: 461,
        edadModaRelativa: "4.97%"
    },
    "2026": {
        meses: ["Ene", "Feb", "Mar", "Abr"],
        urbano: [574, 528, 498, 523],
        rural: [185, 154, 166, 150],
        hombres: [703, 630, 595, 603],
        mujeres: [32, 28, 30, 32],
        sinDatos: [10],
        rangosEdad: { "Niño": 11, "Adolescente": 194, "Joven": 1188, "Adulto": 1269, "Adulto mayor": 38 },
        edadMedia: 29.00,
        edadMediana: 29,
        edadModa: ["25"],
        edadModaFrecuencia: 129,
        edadModaRelativa: "4.64%"
    }
};

// Matriz oficial corregida según comparacionAnual.csv [cite: 1]
const dataHistorica = [
    { anio: 2014, total: 1310, tasa: 8.19, promedioMes: 109.17, promedioDia: 3.59 }, // [cite: 1]
    { anio: 2015, total: 1050, tasa: 6.45, promedioMes: 87.50, promedioDia: 2.88 },  // [cite: 1]
    { anio: 2016, total: 959,  tasa: 5.80, promedioMes: 79.92, promedioDia: 2.62 },  // [cite: 1]
    { anio: 2017, total: 970,  tasa: 5.78, promedioMes: 80.83, promedioDia: 2.66 },  // [cite: 1]
    { anio: 2018, total: 996,  tasa: 5.83, promedioMes: 83.00, promedioDia: 2.73 },  // [cite: 1]
    { anio: 2019, total: 1189, tasa: 6.85, promedioMes: 99.08, promedioDia: 3.26 },  // [cite: 1]
    { anio: 2020, total: 1372, tasa: 7.83, promedioMes: 114.33, promedioDia: 3.75 }, // [cite: 1]
    { anio: 2021, total: 2495, tasa: 14.16, promedioMes: 207.92, promedioDia: 6.84 }, // [cite: 1]
    { anio: 2022, total: 4886, tasa: 27.58, promedioMes: 407.17, promedioDia: 13.39 }, // [cite: 1]
    { anio: 2023, total: 8248, tasa: 46.25, promedioMes: 687.33, promedioDia: 22.60 }, // [cite: 1]
    { anio: 2024, total: 7063, tasa: 39.31, promedioMes: 588.58, promedioDia: 19.30 }, // [cite: 1]
    { anio: 2025, total: 9281, tasa: 51.27, promedioMes: 773.42, promedioDia: 25.43 }, // [cite: 1]
    { anio: 2026, total: 2778, tasa: 45.68, promedioMes: 694.50, promedioDia: 23.15 }  // [cite: 1]
];

// ============================================================================
// COMPONENTE DE SEGURO DE CANVAS (MANEJADOR DE MEMORIA PARA CHART.JS)
// ============================================================================
const chartInstances = {};
Chart.register(ChartDataLabels);

function destruirGraficoSiExiste(idCanvas) {
    if (chartInstances[idCanvas]) {
        chartInstances[idCanvas].destroy();
        delete chartInstances[idCanvas];
    }
}

// ============================================================================
// SISTEMA DE CONTROL DE PESTAÑAS (INTERFAZ DE USUARIO)
// ============================================================================
function ocultarSecciones(){
    const ids = ["inicio", "area", "sexo", "edad", "comparacionEntreAnios", "operacion"];
    ids.forEach(id => {
        const elemento = document.getElementById(id);
        if (elemento) elemento.classList.remove("activa");
    });
}

function mostrarSeccion(id){
    ocultarSecciones();
    
    const seccion = document.getElementById(id);
    if(seccion) {
        seccion.classList.add("activa");
    }
    
    const enlaces = document.querySelectorAll('.sidebar-menu label');
    enlaces.forEach(enlace => {
        enlace.classList.remove('active');
    });
    
    const enlaceActivo = Array.from(enlaces).find(enlace => {
        const accionClick = enlace.getAttribute('onclick');
        if (!accionClick) return false;
        return accionClick.includes(`'${id}'`) || accionClick.includes(`"${id}"`);
    });
    
    if(enlaceActivo) {
        enlaceActivo.classList.add('active');
    }

    const menuToggle = document.getElementById("menu-toggle");
    if(menuToggle) { 
        menuToggle.checked = false; 
    }

    // Inicializar de forma reactiva la tabla global comparativa
    if (id === 'comparacionEntreAnios') {
        renderizarPestañaComparativa();
    }
}

// ============================================================================
// FUNCIONES ESTADÍSTICAS AUTOMÁTICAS (SOPORTE 2014-2026)
// ============================================================================
function calcularMedia(arr) {
    const sum = arr.reduce((a, b) => a + b, 0);
    return (sum / arr.length).toFixed(2);
}

function calcularMediana(arr) {
    const sorted = [...arr].sort((a, b) => a - b);
    const mid = Math.floor(sorted.length / 2);
    if (sorted.length % 2 !== 0) {
        return sorted[mid];
    }
    return ((sorted[mid - 1] + sorted[mid]) / 2).toFixed(2);
}

// ============================================================================
// PROCESAMIENTO DINÁMICO POR SECCIÓN
// ============================================================================
function consultarPorAnio(seccion) {
    const inputId = "inputAnio" + seccion.charAt(0).toUpperCase() + seccion.slice(1);
    const anioInput = document.getElementById(inputId).value.trim();

    if (!dataMaster[anioInput]) {
        alert("Año fuera de rango. Por favor ingrese un año entre 2014 y 2026.");
        return;
    }

    const datos = dataMaster[anioInput];
    const panel = document.getElementById(`resultados${seccion.charAt(0).toUpperCase() + seccion.slice(1)}`);
    panel.style.display = "block";

    if (seccion === 'area') document.querySelector('#chartAreaLine').closest('.card').querySelector('h4').textContent = `Tendencia del año ${anioInput}`;
    if (seccion === 'sexo') document.querySelector('#chartSexoLine').closest('.card').querySelector('h4').textContent = `Tendencia del año ${anioInput}`;
    
const registroAnual = dataHistorica.find(item => item.anio == anioInput);
    const totalHomicidios = registroAnual ? registroAnual.total.toLocaleString() : "0";
    
    panel.querySelectorAll(".txt-anio").forEach(el => {
        el.textContent = `${anioInput} (total homicidios intencionales: ${totalHomicidios})`;
    });
    if (seccion === 'area') procesarArea(datos);
    if (seccion === 'sexo') procesarSexo(datos);
    if (seccion === 'edad') procesarEdad(datos);
}

// --- MÓDULO ÁREA GEOGRÁFICA ---
function procesarArea(d) {
    const totalU = d.urbano.reduce((a, b) => a + b, 0);
    const totalR = d.rural.reduce((a, b) => a + b, 0);
    const total = totalU + totalR;

    const pctU = ((totalU / total) * 100).toFixed(1);
    const pctR = ((totalR / total) * 100).toFixed(1);


    destruirGraficoSiExiste('chartAreaPie');
    chartInstances['chartAreaPie'] = new Chart(document.getElementById('chartAreaPie').getContext('2d'), {
        type: 'pie',
        data: {
            labels: [`Urbana (${totalU.toLocaleString()})`, `Rural (${totalR.toLocaleString()})`],
            datasets: [{ data: [pctU, pctR], backgroundColor: ['#2b5f9e', '#c8873a'], borderColor: '#1e293b', borderWidth: 2 }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { position: 'bottom', labels: { color: '#e2e8f0' } },
                tooltip: { callbacks: { label: (ctx) => ` ${parseFloat(ctx.parsed).toFixed(1)}%` } },
                datalabels: { display: false }
            }
        }
    });

    document.getElementById('tablaAreaContenedor').innerHTML = `
        <table class="tabla-estadistica">
            <thead>
                <tr><th>Métrica</th><th>Área urbana</th><th>Área rural</th></tr>
            </thead>
            <tbody>
                <tr><td><strong>Media (promedio)</strong></td><td>${calcularMedia(d.urbano)}</td><td>${calcularMedia(d.rural)}</td></tr>
                <tr><td><strong>Mediana</strong></td><td>${calcularMediana(d.urbano)}</td><td>${calcularMediana(d.rural)}</td></tr>
            </tbody>
        </table>
    `;

    destruirGraficoSiExiste('chartAreaLine');
    chartInstances['chartAreaLine'] = new Chart(document.getElementById('chartAreaLine').getContext('2d'), {
        type: 'line',
        data: {
            labels: d.meses,
            datasets: [
                { label: 'Urbano', data: d.urbano, borderColor: '#2b5f9e', backgroundColor: 'transparent', tension: 0.1, borderWidth: 3 },
                { label: 'Rural', data: d.rural, borderColor: '#c8873a', backgroundColor: 'transparent', tension: 0.1, borderWidth: 3 }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            scales: {
                x: { ticks: { color: '#94a3b8' }, grid: { color: 'rgba(255,255,255,0.03)' } },
                y: { ticks: { color: '#94a3b8' }, grid: { color: 'rgba(255,255,255,0.05)' } }
            },
            plugins: { legend: { labels: { color: '#e2e8f0' } } }
        }
    });
}

// --- MÓDULO SEXO DE LAS VÍCTIMAS ---
function procesarSexo(d) {
    const totalH = d.hombres.reduce((a, b) => a + b, 0);
    const totalM = d.mujeres.reduce((a, b) => a + b, 0);
    const totalSD = d.sinDatos ? d.sinDatos.reduce((a, b) => a + b, 0) : 0;
    const total = totalH + totalM + totalSD;


const pctH = ((totalH / total) * 100).toFixed(1);
const pctM = ((totalM / total) * 100).toFixed(1);
const pctSD = ((totalSD / total) * 100).toFixed(1);

    

    
    destruirGraficoSiExiste('chartSexoPie');
    chartInstances['chartSexoPie'] = new Chart(document.getElementById('chartSexoPie').getContext('2d'), {
        type: 'pie',
        data: {
            labels: [
    `Hombres (${totalH.toLocaleString()})`,
    `Mujeres (${totalM.toLocaleString()})`,
    `Sin datos (${totalSD.toLocaleString()})`
],

datasets: [{
    data: [pctH, pctM, pctSD],
    backgroundColor: [
        '#1a4070', // Hombres
        '#e9a95c', // Mujeres
        '#64748b'  // Sin datos
    ],
    borderColor: '#1e293b',
    borderWidth: 2
}]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { position: 'bottom', labels: { color: '#e2e8f0' } },
                tooltip: { callbacks: { label: (ctx) => ` ${parseFloat(ctx.parsed).toFixed(1)}%` } },
                datalabels: { display: false }
            }
        }
    });

    document.getElementById('tablaSexoContenedor').innerHTML = `
        <table class="tabla-estadistica">
            <thead>
                <tr><th>Métrica</th><th>Hombres</th><th>Mujeres</th></tr>
            </thead>
            <tbody>
                <tr><td><strong>Media (promedio)</strong></td><td>${calcularMedia(d.hombres)}</td><td>${calcularMedia(d.mujeres)}</td></tr>
                <tr><td><strong>Mediana</strong></td><td>${calcularMediana(d.hombres)}</td><td>${calcularMediana(d.mujeres)}</td></tr>
            </tbody>
        </table>
    `;

    destruirGraficoSiExiste('chartSexoLine');
    chartInstances['chartSexoLine'] = new Chart(document.getElementById('chartSexoLine').getContext('2d'), {
        type: 'line',
        data: {
            labels: d.meses,
            datasets: [
                { label: 'Hombres', data: d.hombres, borderColor: '#1a4070', backgroundColor: 'transparent', tension: 0.1, borderWidth: 3 },
                { label: 'Mujeres', data: d.mujeres, borderColor: '#e9a95c', backgroundColor: 'transparent', tension: 0.1, borderWidth: 3 }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            scales: {
                x: { ticks: { color: '#94a3b8' }, grid: { color: 'rgba(255,255,255,0.03)' } },
                y: { ticks: { color: '#94a3b8' }, grid: { color: 'rgba(255,255,255,0.05)' } }
            },
            plugins: { legend: { labels: { color: '#e2e8f0' } } }
        }
    });
}

// --- MÓDULO EDAD DE LAS VÍCTIMAS ---
function procesarEdad(d) {
    const etiquetas = Object.keys(d.rangosEdad);
    const valores = Object.values(d.rangosEdad);

    const totalEdad = valores.reduce((a, b) => a + b, 0);
    const etiquetasConCantidad = etiquetas.map((e, i) => `${e} (${valores[i].toLocaleString()})`);

    destruirGraficoSiExiste('chartEdadPie');
    chartInstances['chartEdadPie'] = new Chart(document.getElementById('chartEdadPie').getContext('2d'), {
        type: 'pie',
        data: {
            labels: etiquetasConCantidad,
            datasets: [{ data: valores, backgroundColor: ['#2b5f9e', '#c8873a', '#1a4070', '#e9a95c', '#64748b'], borderColor: '#1e293b', borderWidth: 2 }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { position: 'bottom', labels: { color: '#e2e8f0' } },
                tooltip: { callbacks: { label: (ctx) => ` ${((ctx.parsed / totalEdad) * 100).toFixed(1)}%` } },
                datalabels: { display: false }
            }
        }
    });

    const formatoModa = d.edadModa.map(m => `${m} años`).join(' y ');

    document.getElementById('tablaEdadContenedor').innerHTML = `
        <table class="tabla-estadistica">
            <thead>
                <tr><th>Métrica</th><th>Valor</th></tr>
            </thead>
            <tbody>
                <tr><td><strong>Edad promedio</strong></td><td>${d.edadMedia} años</td></tr>
                <tr><td><strong>Mediana de edad</strong></td><td>${d.edadMediana} años</td></tr>
                <tr><td><strong>Moda de edad</strong></td><td>${formatoModa}</td></tr>
                <tr><td><strong>Frecuencia absoluta de la moda</strong></td><td>${d.edadModaFrecuencia} casos</td></tr>
                <tr><td><strong>Frecuencia relativa de la moda</strong></td><td>${d.edadModaRelativa.replace('%', ' %')}</td></tr>
            </tbody>
        </table>
    `;
}

// --- RENDIMIENTO GLOBAL HISTÓRICO ---
function renderizarPestañaComparativa() {
    let rowsHtml = "";
    dataHistorica.forEach(f => {
        rowsHtml += `
            <tr>
                <td><strong>${f.anio}</strong></td>
                <td>${f.total.toLocaleString()} </td>
                <td>${f.promedioMes.toFixed(2)} </td>
                <td>${f.promedioDia.toFixed(2)} </td>
            </tr>
        `;
    });

    document.getElementById('tablaComparativaContenedor').innerHTML = `
        <table class="tabla-estadistica">
            <thead>
                <tr><th>Año de registro</th><th>Total homicidios intencionales</th><th>Promedio mensual</th><th>Promedio diario</th></tr>
            </thead>
            <tbody>
                ${rowsHtml}
            </tbody>
        </table>
    `;

    const labelsAnio = dataHistorica.map(h => h.anio);
    const tasasValores = dataHistorica.map(h => h.tasa);

    destruirGraficoSiExiste('chartTasaLine');
    chartInstances['chartTasaLine'] = new Chart(document.getElementById('chartTasaLine').getContext('2d'), {
        type: 'line',
        data: {
            labels: labelsAnio,
            datasets: [{
                label: 'Tasa nacional',
                data: tasasValores,
                borderColor: '#e9a95c',
                backgroundColor: 'rgba(233, 169, 92, 0.08)',
                fill: true,
                tension: 0.15,
                borderWidth: 3,
                pointRadius: 4,
                pointBackgroundColor: '#c8873a'
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            scales: {
                x: { ticks: { color: '#94a3b8' }, grid: { color: 'rgba(255,255,255,0.01)' } },
                y: { ticks: { color: '#94a3b8' }, grid: { color: 'rgba(255,255,255,0.05)' } }
            },
            plugins: { legend: { labels: { color: '#e2e8f0' } } }
        }
    });
}

// ============================================================================
// MÓDULO: OPERACIÓN ENTRE AÑOS (media, mediana, moda, porcentaje, tendencia)
// ============================================================================

// Recibe el texto del input (ej: "2020,2021,2022") y devuelve un arreglo de
// objetos { anio, total } válidos, tomando el total desde dataHistorica.
function obtenerDatosDeAnios(textoAnios) {
    const ANIO_MINIMO = 2014;
    const ANIO_MAXIMO = 2025;

    const aniosEscritos = textoAnios.split(",").map(a => a.trim()).filter(a => a !== "");

    const datos = [];
    for (const anioTexto of aniosEscritos) {
        const anio = Number(anioTexto);

        // Verifica que el año esté dentro del rango permitido (2014 - 2025)
        if (anio < ANIO_MINIMO || anio > ANIO_MAXIMO) {
            alert(`El año "${anioTexto}" no es válido. Por favor ingrese años entre ${ANIO_MINIMO} y ${ANIO_MAXIMO}, separados por coma.`);
            return null;
        }

        const registro = dataHistorica.find(item => item.anio == anio);
        if (!registro) {
            alert(`El año "${anioTexto}" no es válido. Por favor ingrese años entre ${ANIO_MINIMO} y ${ANIO_MAXIMO}, separados por coma.`);
            return null;
        }

        // Verifica que el año no haya sido escrito ya antes (no se permiten años repetidos)
        const yaExiste = datos.some(d => d.anio == anio);
        if (yaExiste) {
            console.log(`El año ${anio} está repetido.`);
            alert(`El año "${anioTexto}" está repetido. Por favor escriba cada año una sola vez.`);
            return null;
        }

        datos.push(registro);
    }

    if (datos.length === 0) {
        alert("Por favor escriba al menos un año.");
        return null;
    }

    // No se permite calcular con un solo año, se necesitan al menos dos para comparar
    if (datos.length === 1) {
        console.log("No se puede escribir un solo año. Debe ingresar al menos dos años separados por coma.");
        alert("No puede escribir un solo año. Por favor ingrese al menos dos años separados por coma.");
        return null;
    }

    return datos;
}

// Recibe el texto del input (ej: "2017-2025") y devuelve un arreglo de
// objetos { anio, total } válidos, tomando todos los años dentro del intervalo.
// Se usa solo para PORCENTAJE y TENDENCIA: aquí NO se permiten años sueltos
// ni separados por coma, únicamente un intervalo con guion.
function obtenerDatosDeIntervalo(textoIntervalo) {
    const ANIO_MINIMO = 2014;
    const ANIO_MAXIMO = 2025;

    // Debe tener exactamente la forma "numero-numero" (ej: 2017-2025)
    const patronIntervalo = /^(\d{4})-(\d{4})$/;
    const coincidencia = textoIntervalo.match(patronIntervalo);

    if (!coincidencia) {
        alert(`Formato no válido. Escriba únicamente un intervalo de años con guion, por ejemplo: ${ANIO_MINIMO}-${ANIO_MAXIMO}. No se permiten años sueltos ni separados por coma.`);
        return null;
    }

    const anioInicio = Number(coincidencia[1]);
    const anioFin = Number(coincidencia[2]);

    if (anioInicio < ANIO_MINIMO || anioInicio > ANIO_MAXIMO || anioFin < ANIO_MINIMO || anioFin > ANIO_MAXIMO) {
        alert(`Por favor ingrese años entre ${ANIO_MINIMO} y ${ANIO_MAXIMO}.`);
        return null;
    }

    if (anioInicio >= anioFin) {
        alert("El primer año del intervalo debe ser menor que el segundo. Ejemplo: 2017-2025.");
        return null;
    }

    // Genera todos los años entre anioInicio y anioFin (incluyendo ambos extremos)
    const datos = [];
    for (let anio = anioInicio; anio <= anioFin; anio++) {
        const registro = dataHistorica.find(item => item.anio == anio);
        if (registro) datos.push(registro);
    }

    return datos;
}

// Calcula la moda de un arreglo de números (puede haber más de un valor)
function calcularModaArreglo(arr) {
    const conteo = {};
    arr.forEach(num => { conteo[num] = (conteo[num] || 0) + 1; });

    const maxFrecuencia = Math.max(...Object.values(conteo));

    // Si todos los valores se repiten la misma cantidad de veces (1 vez cada uno),
    // no hay una moda real porque ningún valor se repite más que los demás.
    if (maxFrecuencia === 1) {
        return { valores: [], frecuencia: 1 };
    }

    const valoresModa = Object.keys(conteo)
        .filter(num => conteo[num] === maxFrecuencia)
        .map(Number);

    return { valores: valoresModa, frecuencia: maxFrecuencia };
}

// Función "router": según la operación elegida, llama a la función correspondiente
function calcularOperacion(operacion) {
    const inputId = "inputAnios" + operacion.charAt(0).toUpperCase() + operacion.slice(1);
    const textoAnios = document.getElementById(inputId).value.trim();

    // Porcentaje y Tendencia usan un intervalo con guion (ej: 2017-2025)
    // Media, Mediana y Moda usan años sueltos separados por coma (ej: 2020,2021,2022)
    const esOperacionDeIntervalo = (operacion === 'porcentaje' || operacion === 'tendencia');

    const datos = esOperacionDeIntervalo
        ? obtenerDatosDeIntervalo(textoAnios)
        : obtenerDatosDeAnios(textoAnios);

    if (!datos) return;

    const panel = document.getElementById("resultados" + operacion.charAt(0).toUpperCase() + operacion.slice(1));
    panel.style.display = "block";

    if (operacion === 'media') procesarOperacionMedia(datos);
    if (operacion === 'mediana') procesarOperacionMediana(datos);
    if (operacion === 'moda') procesarOperacionModa(datos);
    if (operacion === 'porcentaje') procesarOperacionPorcentaje(datos);
    if (operacion === 'tendencia') procesarOperacionTendencia(datos);
}

// Oculta el panel de resultados (tabla + gráfico) de una operación específica
function ocultarOperacion(operacion) {
    const panel = document.getElementById("resultados" + operacion.charAt(0).toUpperCase() + operacion.slice(1));
    panel.style.display = "none";
}

// Permite escribir solo números y comas en las cajas de Media, Mediana y Moda.
// Permite escribir solo números y un guion en las cajas de Porcentaje y Tendencia,
// ya que estas dos únicamente aceptan un intervalo (ej: 2017-2025).
function permitirSoloNumerosYComas() {
    const idsSolo = ['inputAnioEdad', 'inputAnioArea', 'inputAnioSexo' ];
    const idsConComa = ['inputAniosMedia', 'inputAniosMediana', 'inputAniosModa'];
    const idsConIntervalo = ['inputAniosPorcentaje', 'inputAniosTendencia'];


    idsSolo.forEach(id => {
        const input = document.getElementById(id);
        if (input) {
            input.addEventListener('input', () => {
                 //Elimina cualquier caracter que no sea dígito (0-9) o coma
              input.value = input.value.replace(/[^0-9]/g, '');
                
           });
        }
    }); 

    idsConComa.forEach(id => {
        const input = document.getElementById(id);
        if (input) {
            input.addEventListener('input', () => {
                // Elimina cualquier caracter que no sea dígito (0-9) o coma
                input.value = input.value.replace(/[^0-9,]/g, '');
            });
        }
    });

    idsConIntervalo.forEach(id => {
        const input = document.getElementById(id);
        if (input) {
            input.addEventListener('input', () => {
                // Elimina cualquier caracter que no sea dígito (0-9) o guion
                input.value = input.value.replace(/[^0-9-]/g, '');
            });
        }
    });
}

// --- MEDIA ---
function procesarOperacionMedia(datos) {
    const totales = datos.map(d => d.total);
    const media = calcularMedia(totales);

    document.getElementById('tablaMediaContenedor').innerHTML = `
        <table class="tabla-estadistica">
            <thead><tr><th>Años seleccionados</th><th>Media (promedio)</th></tr></thead>
            <tbody>
                <tr><td>${datos.map(d => d.anio).join(", ")}</td><td>${media}</td></tr>
            </tbody>
        </table>
    `;

    destruirGraficoSiExiste('chartMedia');
    chartInstances['chartMedia'] = new Chart(document.getElementById('chartMedia').getContext('2d'), {
        type: 'bar',
        data: {
            labels: datos.map(d => d.anio),
            datasets: [
                { label: 'Total homicidios por año', data: totales, backgroundColor: '#2b5f9e' },
                { label: 'Media', data: datos.map(() => media), type: 'line', borderColor: '#e9a95c', backgroundColor: 'transparent', borderWidth: 3, tension: 0 }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            scales: {
                x: { ticks: { color: '#94a3b8' }, grid: { color: 'rgba(255,255,255,0.03)' } },
                y: { ticks: { color: '#94a3b8' }, grid: { color: 'rgba(255,255,255,0.05)' } }
            },
            plugins: { legend: { labels: { color: '#e2e8f0' } } }
        }
    });
}

// --- MEDIANA ---
function procesarOperacionMediana(datos) {
    const totales = datos.map(d => d.total);
    const mediana = calcularMediana(totales);

    document.getElementById('tablaMedianaContenedor').innerHTML = `
        <table class="tabla-estadistica">
            <thead><tr><th>Años seleccionados</th><th>Mediana</th></tr></thead>
            <tbody>
                <tr><td>${datos.map(d => d.anio).join(", ")}</td><td>${mediana}</td></tr>
            </tbody>
        </table>
    `;

    destruirGraficoSiExiste('chartMediana');
    chartInstances['chartMediana'] = new Chart(document.getElementById('chartMediana').getContext('2d'), {
        type: 'bar',
        data: {
            labels: datos.map(d => d.anio),
            datasets: [
                { label: 'Total homicidios por año', data: totales, backgroundColor: '#c8873a' },
                { label: 'Mediana', data: datos.map(() => mediana), type: 'line', borderColor: '#e9a95c', backgroundColor: 'transparent', borderWidth: 3, tension: 0 }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            scales: {
                x: { ticks: { color: '#94a3b8' }, grid: { color: 'rgba(255,255,255,0.03)' } },
                y: { ticks: { color: '#94a3b8' }, grid: { color: 'rgba(255,255,255,0.05)' } }
            },
            plugins: { legend: { labels: { color: '#e2e8f0' } } }
        }
    });
}

// --- MODA ---
function procesarOperacionModa(datos) {
    const totales = datos.map(d => d.total);
    const resultadoModa = calcularModaArreglo(totales);

    const textoModa = resultadoModa.valores.length > 0
        ? resultadoModa.valores.map(v => v.toLocaleString()).join(" y ")
        : "No hay un valor que se repita (todos los totales son distintos)";

    document.getElementById('tablaModaContenedor').innerHTML = `
        <table class="tabla-estadistica">
            <thead><tr><th>Años seleccionados</th><th>Moda</th><th>Frecuencia</th></tr></thead>
            <tbody>
                <tr>
                    <td>${datos.map(d => d.anio).join(", ")}</td>
                    <td>${textoModa}</td>
                    <td>${resultadoModa.valores.length > 0 ? resultadoModa.frecuencia + " veces" : "-"}</td>
                </tr>
            </tbody>
        </table>
    `;

    // Para el gráfico, coloreamos en distinto color las barras que coinciden con la moda
    const coloresBarras = totales.map(t =>
        resultadoModa.valores.includes(t) ? '#e9a95c' : '#2b5f9e'
    );

    destruirGraficoSiExiste('chartModa');
    chartInstances['chartModa'] = new Chart(document.getElementById('chartModa').getContext('2d'), {
        type: 'bar',
        data: {
            labels: datos.map(d => d.anio),
            datasets: [
                { label: 'Total homicidios por año (resaltado = moda)', data: totales, backgroundColor: coloresBarras }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            scales: {
                x: { ticks: { color: '#94a3b8' }, grid: { color: 'rgba(255,255,255,0.03)' } },
                y: { ticks: { color: '#94a3b8' }, grid: { color: 'rgba(255,255,255,0.05)' } }
            },
            plugins: { legend: { labels: { color: '#e2e8f0' } } }
        }
    });
}

// --- PORCENTAJE ---
function procesarOperacionPorcentaje(datos) {
    const totales = datos.map(d => d.total);
    const sumaTotal = totales.reduce((a, b) => a + b, 0);
    const porcentajes = totales.map(t => ((t / sumaTotal) * 100).toFixed(1));

    let filas = "";
    datos.forEach((d, i) => {
        filas += `<tr><td>${d.anio}</td><td>${d.total.toLocaleString()}</td><td>${porcentajes[i]}%</td></tr>`;
    });

    document.getElementById('tablaPorcentajeContenedor').innerHTML = `
        <table class="tabla-estadistica">
            <thead><tr><th>Año</th><th>Total homicidios</th><th>Porcentaje del total seleccionado</th></tr></thead>
            <tbody>${filas}</tbody>
        </table>
    `;

    const coloresPie = ['#2b5f9e', '#c8873a', '#1a4070', '#e9a95c', '#64748b', '#3b82f6', '#0f2d52'];

    destruirGraficoSiExiste('chartPorcentaje');
    chartInstances['chartPorcentaje'] = new Chart(document.getElementById('chartPorcentaje').getContext('2d'), {
        type: 'pie',
        data: {
            labels: datos.map(d => `${d.anio} (${d.total.toLocaleString()})`),
            datasets: [{
                data: porcentajes,
                backgroundColor: datos.map((_, i) => coloresPie[i % coloresPie.length]),
                borderColor: '#1e293b',
                borderWidth: 2
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { position: 'bottom', labels: { color: '#e2e8f0' } },
                tooltip: { callbacks: { label: (ctx) => ` ${parseFloat(ctx.parsed).toFixed(1)}%` } },
                datalabels: { display: false }
            }
        }
    });
}

// --- TENDENCIA ---
function procesarOperacionTendencia(datos) {
    // Ordenamos los años de menor a mayor para que la tendencia tenga sentido
    const datosOrdenados = [...datos].sort((a, b) => a.anio - b.anio);
    const totales = datosOrdenados.map(d => d.total);

    let textoTendencia = "Estable";
    if (totales.length >= 2) {
        const primero = totales[0];
        const ultimo = totales[totales.length - 1];
        if (ultimo > primero) textoTendencia = "Creciente (aumentó)";
        else if (ultimo < primero) textoTendencia = "Decreciente (disminuyó)";
    }

    const cambioPorcentual = totales.length >= 2
        ? (((totales[totales.length - 1] - totales[0]) / totales[0]) * 100).toFixed(1)
        : "0.0";

    document.getElementById('tablaTendenciaContenedor').innerHTML = `
        <table class="tabla-estadistica">
            <thead><tr><th>Años seleccionados (ordenados)</th><th>Tendencia</th><th>Cambio porcentual del primero al último</th></tr></thead>
            <tbody>
                <tr>
                    <td>${datosOrdenados.map(d => d.anio).join(", ")}</td>
                    <td>${textoTendencia}</td>
                    <td>${cambioPorcentual}%</td>
                </tr>
            </tbody>
        </table>
    `;

    destruirGraficoSiExiste('chartTendencia');
    chartInstances['chartTendencia'] = new Chart(document.getElementById('chartTendencia').getContext('2d'), {
        type: 'line',
        data: {
            labels: datosOrdenados.map(d => d.anio),
            datasets: [{
                label: 'Total homicidios por año',
                data: totales,
                borderColor: '#e9a95c',
                backgroundColor: 'rgba(233, 169, 92, 0.08)',
                fill: true,
                tension: 0.15,
                borderWidth: 3,
                pointRadius: 5,
                pointBackgroundColor: '#c8873a'
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            scales: {
                x: { ticks: { color: '#94a3b8' }, grid: { color: 'rgba(255,255,255,0.01)' } },
                y: { ticks: { color: '#94a3b8' }, grid: { color: 'rgba(255,255,255,0.05)' } }
            },
            plugins: { legend: { labels: { color: '#e2e8f0' } } }
        }
    });
}

// Inicialización de la SPA al cargar el DOM
document.addEventListener("DOMContentLoaded", () => {
    mostrarSeccion("inicio");
    permitirSoloNumerosYComas();
});