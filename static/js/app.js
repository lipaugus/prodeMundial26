/**
 * LÓGICA DEL PRODE MUNDIAL 2026 - VERSIÓN MULTIPÁGINA COMPLETA
 */

document.addEventListener('DOMContentLoaded', () => {
    initNavigation();
    cargarDatosProde();
});

let prodeWorkbook = null;
let teamCodesMap = {}; 

/**
 * Normalizador de nombres de países para evitar discrepancies de escritura entre hojas
 */
function normalizarNombreEquipo(nombre) {
    if (!nombre) return '';
    let n = String(nombre).trim().toLowerCase();
    if (n === 'qatar') return 'catar';
    if (n === 'eeuu' || n === 'usa' || n === 'estados unidos') return 'eeuu';
    if (n === 'holanda' || n === 'paises bajos' || n === 'países bajos') return 'holanda';
    return n;
}

/**
 * 1. Control de Navegación Inteligente (Soporta index.html y ralfito.html)
 */
function initNavigation() {
    const menuToggle = document.getElementById('menuToggle');
    const navOverlay = document.getElementById('navOverlay');
    const navLinks = document.querySelectorAll('.nav-links a');
    const sections = document.querySelectorAll('.view-section');
    const roundSelector = document.getElementById('roundSelector');
    
    const prevBtn = document.getElementById('prevRoundBtn');
    const nextBtn = document.getElementById('nextRoundBtn');

    // Control seguro del menú desplegable móvil (solo si existe en la página actual)
    if (menuToggle && navOverlay) {
        menuToggle.addEventListener('click', () => {
            menuToggle.classList.toggle('open');
            navOverlay.classList.toggle('open');
        });
    }

    // Identificamos con precisión matemática en qué página estamos parados
    const path = window.location.pathname;
    const estaEnIndex = path.endsWith('index.html') || path.endsWith('/') || !path.includes('.html');

    navLinks.forEach(link => {
        link.addEventListener('click', (e) => {
            const target = link.getAttribute('data-target');
            
            // Regla de Oro: Solo interceptamos con SPA si estamos en index y el enlace pide una sección local
            if (estaEnIndex && target) {
                const targetSection = document.getElementById(`view-${target}`);
                if (targetSection) {
                    e.preventDefault(); // Detiene la recarga nativa de index.html
                    
                    navLinks.forEach(l => l.classList.remove('active'));
                    link.classList.add('active');

                    if (menuToggle) menuToggle.classList.remove('open');
                    if (navOverlay) navOverlay.classList.remove('open');

                    sections.forEach(sec => sec.classList.remove('active-view'));
                    targetSection.classList.add('active-view');
                }
            }
            // Si no se cumple la condición (ej: vas a ralfito.html o estás saliendo de él), 
            // el navegador ejecuta su comportamiento estándar y cambia de página de forma limpia.
        });
    });

    // Control manual seguro del selector de rondas
    if (roundSelector) {
        roundSelector.addEventListener('change', (e) => {
            renderizarRonda(e.target.value);
        });
    }

    // Control seguro de Flecha Izquierda
    if (prevBtn && roundSelector) {
        prevBtn.addEventListener('click', () => {
            if (roundSelector.selectedIndex > 0) {
                roundSelector.selectedIndex--;
                roundSelector.dispatchEvent(new Event('change'));
            }
        });
    }

    // Control seguro de Flecha Derecha
    if (nextBtn && roundSelector) {
        nextBtn.addEventListener('click', () => {
            if (roundSelector.selectedIndex < roundSelector.options.length - 1) {
                roundSelector.selectedIndex++;
                roundSelector.dispatchEvent(new Event('change'));
            }
        });
    }
}

/**
 * 2. Procesamiento de Fechas Híbridas (Soporta Strings y Números Seriales de Excel)
 */
function parsearFechaExcel(fechaRaw, horaRaw) {
    let d;
    
    if (typeof fechaRaw === 'number') {
        d = new Date(Math.round((fechaRaw - 25569) * 86400 * 1000));
    } else {
        d = new Date(fechaRaw);
        if (isNaN(d.getTime())) {
            const parts = String(fechaRaw).split('/');
            if (parts.length === 3) {
                d = new Date(parts[2], parts[1] - 1, parts[0]);
            }
        }
    }

    let horas = 0;
    let minutos = 0;
    
    if (typeof horaRaw === 'number') {
        const totalMinutes = Math.round(horaRaw * 24 * 60);
        horas = Math.floor(totalMinutes / 60);
        minutos = totalMinutes % 60;
    } else if (horaRaw) {
        const parts = String(horaRaw).split(':');
        horas = parseInt(parts[0] || 0, 10);
        minutos = parseInt(parts[1] || 0, 10);
    }

    d.setHours(horas, minutos, 0, 0);
    return d;
}

/**
 * 3. Carga e Inicialización de Datos del Excel
 */
async function cargarDatosProde() {
    // PROTECCIÓN: Si la página actual no tiene el contenedor de partidos (ej: ralfito.html), abortamos la carga
    if (!document.getElementById('matches-container')) return;

    try {
        const response = await fetch('prode_data.xlsx');
        if (!response.ok) throw new Error('No se pudo leer prode_data.xlsx');
        
        const arrayBuffer = await response.arrayBuffer();
        const workbook = XLSX.read(arrayBuffer, { type: 'array' });
        prodeWorkbook = workbook;

        procesarCodigosEquipos();

        const rondaActual = determinarRondaPorFecha();
        const selector = document.getElementById('roundSelector');
        if (selector) selector.value = rondaActual;
        
        renderizarRonda(rondaActual);

    } catch (error) {
        console.error(error);
        const container = document.getElementById('matches-container');
        if (container) {
            container.innerHTML = `<p class="loading-text" style="color: #ef4444;">Error al inicializar el Prode.</p>`;
        }
    }
}

function procesarCodigosEquipos() {
    if (!prodeWorkbook) return;
    const sheet = prodeWorkbook.Sheets['team-codes'];
    if (!sheet) return;
    const data = XLSX.utils.sheet_to_json(sheet);
    data.forEach(row => {
        if (row.team && row.fifa_code) {
            const key = normalizarNombreEquipo(row.team);
            teamCodesMap[key] = String(row.fifa_code).trim().toUpperCase();
        }
    });
}

function determinarRondaPorFecha() {
    if (!prodeWorkbook) return 'group-stage-3';
    const sheet = prodeWorkbook.Sheets['points-table'];
    if (!sheet) return 'group-stage-3'; 

    const data = XLSX.utils.sheet_to_json(sheet);
    const ahora = new Date(); 

    for (let row of data) {
        const roundId = String(row.round || '').trim();
        if (!roundId || roundId.toLowerCase() === 'total') continue;

        if (row.start_date) {
            const start = parsearFechaExcel(row.start_date, row.start_hour);
            
            if (roundId === 'final-3rd-place') {
                if (ahora >= start) return roundId; 
            } else if (row.end_date) {
                const end = parsearFechaExcel(row.end_date, row.end_hour);
                if (ahora >= start && ahora <= end) {
                    return roundId;
                }
            }
        }
    }
    return 'group-stage-3';
}

/**
 * 4. Renderizador de Boxes de Partidos
 */
function renderizarRonda(roundId) {
    if (!prodeWorkbook) return;

    const matchSheet = prodeWorkbook.Sheets[roundId];
    const matchTimesSheet = prodeWorkbook.Sheets['match-times'];
    const playerBlock = document.getElementById('player-picks-block');
    const container = document.getElementById('matches-container');

    if (!container) return; 

    if (!matchSheet) {
        container.innerHTML = '<p class="loading-text">No hay partidos en esta ronda.</p>';
        if (playerBlock) playerBlock.style.display = 'none';
        return;
    }

    const matches = XLSX.utils.sheet_to_json(matchSheet);
    const times = matchTimesSheet ? XLSX.utils.sheet_to_json(matchTimesSheet) : [];

    const horariosMap = {};
    times.forEach(t => {
        const idRaw = t['match-id'] !== undefined ? t['match-id'] : t['match_id'];
        if (idRaw !== undefined) horariosMap[String(idRaw).trim()] = t;
    });

    const lucasRow = matches.find(m => m.lucas_player_name);
    const tomasRow = matches.find(m => m.tomas_player_name);

    if (playerBlock) {
        if (roundId === 'group-stage-1' || (!lucasRow && !tomasRow)) {
            playerBlock.style.display = 'none';
        } else {
            playerBlock.style.display = 'block';
            const lp = document.getElementById('lucas-player');
            const lpp = document.getElementById('lucas-player-pts');
            const tp = document.getElementById('tomas-player');
            const tpp = document.getElementById('tomas-player-pts');

            if (lp) lp.innerText = lucasRow ? lucasRow.lucas_player_name : '-';
            if (lpp) lpp.innerText = lucasRow && lucasRow.lucas_player_pts !== undefined ? `${lucasRow.lucas_player_pts} pts` : '0.0 pts';
            if (tp) tp.innerText = tomasRow ? tomasRow.tomas_player_name : '-';
            if (tpp) tpp.innerText = tomasRow && tomasRow.tomas_player_pts !== undefined ? `${tomasRow.tomas_player_pts} pts` : '0.0 pts';
        }
    }

    // Vaciamos el contenedor para empezar a dibujar
    container.innerHTML = ''; 

    // =========================================================================
    // NUEVO: CÁLCULO DINÁMICO DE PUNTOS ACUMULADOS HASTA LA FASE SELECCIONADA
    // =========================================================================
    let acumLucas = 0;
    let acumTomas = 0;
    let faseLucas = 0;
    let faseTomas = 0;

    const pointsSheet = prodeWorkbook.Sheets['points-table'];
    if (pointsSheet) {
        const pointsData = XLSX.utils.sheet_to_json(pointsSheet);
        for (let row of pointsData) {
            const rId = String(row.round || '').trim();
            if (!rId || rId.toLowerCase() === 'total') continue;

            const pL = parseFloat(row.lucas || 0);
            const pT = parseFloat(row.tomas || 0);

            // Sumamos al acumulado histórico
            acumLucas += pL;
            acumTomas += pT;

            // Si llegamos a la ronda actual del filtro, guardamos los parciales y cortamos
            if (rId === roundId) {
                faseLucas = pL;
                faseTomas = pT;
                break;
            }
        }
    }

    // Inyectamos la tarjeta de puntuación global al inicio del contenedor
    const scorecardBox = document.createElement('div');
    scorecardBox.className = 'phase-scorecard';
    scorecardBox.innerHTML = `
        <div class="scorecard-player">
            <span class="scorecard-name">Lucas</span>
            <span class="scorecard-total">${Number(acumLucas.toFixed(1))} <small>pts</small></span>
            <span class="scorecard-partial">+${Number(faseLucas.toFixed(1))} fase</span>
        </div>
        <div class="scorecard-divider"></div>
        <div class="scorecard-player">
            <span class="scorecard-name">Tomás</span>
            <span class="scorecard-total">${Number(acumTomas.toFixed(1))} <small>pts</small></span>
            <span class="scorecard-partial">+${Number(faseTomas.toFixed(1))} fase</span>
        </div>
    `;
    container.appendChild(scorecardBox);
    // =========================================================================

    // Continúa el renderizado normal de los partidos...
    matches.forEach(m => {
        const idRaw = m.match_id !== undefined ? m.match_id : m['match-id'];
        if (idRaw === undefined) return;
        
        const idStr = String(idRaw).trim();
        const timeInfo = horariosMap[idStr];

        const res1 = m.result_team_1;
        const res2 = m.result_team_2;
        
        let centroHTML = '';
        let estado = 'antes';

        if (res1 !== undefined && res2 !== undefined && res1 !== '' && res2 !== '') {
            estado = 'pos';
            centroHTML = `
                <span class="score-main">${res1}  -  ${res2}</span>
                <span class="match-meta-text">FT</span>
            `;
        } else if (timeInfo) {
            const ahora = new Date();
            const fechaPartido = new Date(2026, (timeInfo.m - 1), timeInfo.d, timeInfo.hr, timeInfo.min || 0);

            if (ahora >= fechaPartido) {
                estado = 'en-vivo';
                centroHTML = `
                    <span class="score-main">-</span>
                    <span class="match-meta-text live">• En Vivo</span>
                `;
            } else {
                estado = 'antes';
                const minFormato = String(timeInfo.min || 0).padStart(2, '0');
                const esHoy = ahora.getDate() === parseInt(timeInfo.d) && (ahora.getMonth() + 1) === parseInt(timeInfo.m);
                
                if (esHoy) {
                    centroHTML = `
                        <span class="score-main">-</span>
                        <span class="match-meta-text">${timeInfo.hr}:${minFormato}</span>
                    `;
                } else {
                    const diaFormato = String(timeInfo.d).padStart(2, '0');
                    const mesFormato = String(timeInfo.m).padStart(2, '0');
                    centroHTML = `
                        <span class="score-main">-</span>
                        <span class="match-meta-text">${diaFormato}/${mesFormato}\n${timeInfo.hr}:${minFormato}</span>
                    `;
                }
            }
        } else {
            estado = 'antes';
            centroHTML = `
                <span class="score-main">-</span>
                <span class="match-meta-text">Horario<br>Pendiente</span>
            `;
        }

        const code1 = teamCodesMap[normalizarNombreEquipo(m.team_1)] || 'UNKNOWN';
        const code2 = teamCodesMap[normalizarNombreEquipo(m.team_2)] || 'UNKNOWN';

        const ptsLucasStr = estado === 'pos' ? `${m.pts_lucas ?? 0} ${m.pts_lucas === 1 ? 'pt.' : 'pts.'}` : '- pts.';
        const ptsTomasStr = estado === 'pos' ? `${m.pts_tomas ?? 0} ${m.pts_tomas === 1 ? 'pt.' : 'pts.'}` : '- pts.';

        const box = document.createElement('div');
        box.className = 'match-box';
        box.innerHTML = `
            <div class="match-top-row">
                <div class="team-block">
                    <img src="static/images/flags/${code1}.svg" alt="${m.team_1 || '?'}" class="flag-img" onerror="this.src='static/images/flags/placeholder.svg'">
                    <span class="team-name">${m.team_1 || 'Equipo 1'}</span>
                </div>
                <div class="center-score-block">
                    ${centroHTML}
                </div>
                <div class="team-block">
                    <img src="static/images/flags/${code2}.svg" alt="${m.team_2 || '?'}" class="flag-img" onerror="this.src='static/images/flags/placeholder.svg'">
                    <span class="team-name">${m.team_2 || 'Equipo 2'}</span>
                </div>
            </div>
            
            <hr class="match-divider">
            
            <div class="match-predictions-block">
                <div class="pred-row">
                    <span class="pred-user">Lucas</span>
                    <span class="pred-score">${m.lucas_result_team_1 ?? '-'}  -  ${m.lucas_result_team_2 ?? '-'}</span>
                    <span class="pred-pts">${ptsLucasStr}</span>
                </div>
                <div class="pred-row">
                    <span class="pred-user">Tomás</span>
                    <span class="pred-score">${m.tomas_result_team_1 ?? '-'}  -  ${m.tomas_result_team_2 ?? '-'}</span>
                    <span class="pred-pts">${ptsTomasStr}</span>
                </div>
            </div>
        `;
        container.appendChild(box);
    });
}