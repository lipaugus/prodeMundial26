
const CRONOGRAMA_FASES = [
    {
        id: 'group-stage-1',
        inicio: '2026-06-11T00:00:00-03:00',
        fin: '2026-06-18T09:59:59-03:00'
    },
    {
        id: 'group-stage-2',
        inicio: '2026-06-18T10:00:00-03:00',
        fin: '2026-06-24T09:59:59-03:00'
    },
    {
        id: 'group-stage-3',
        inicio: '2026-06-24T10:00:00-03:00',
        fin: '2026-06-28T12:59:59-03:00'
    },
    {
        id: 'round-of-32',
        inicio: '2026-06-28T13:00:00-03:00',
        fin: '2026-07-04T09:59:59-03:00'
    },
    {
        id: 'round-of-16',
        inicio: '2026-07-04T10:00:00-03:00',
        fin: '2026-07-08T14:59:59-03:00'
    },
    {
        id: 'quaterfinals',
        inicio: '2026-07-08T15:00:00-03:00',
        fin: '2026-07-13T14:59:59-03:00'
    },
    {
        id: 'semifinals',
        inicio: '2026-07-13T15:00:00-03:00',
        fin: '2026-07-17T14:59:59-03:00'
    },
    {
        id: 'final-3rd-place',
        inicio: '2026-07-17T15:00:00-03:00',
        fin: '2026-12-31T23:59:59-03:00'
    }
];




document.addEventListener('DOMContentLoaded', () => {
    initNavigation();
    cargarDatosProde();
});

let prodeWorkbook = null;
let teamCodesMap = {};


function normalizarNombreEquipo(nombre) {
    if (!nombre) return '';
    let n = String(nombre).trim().toLowerCase();
    if (n === 'qatar') return 'catar';
    if (n === 'eeuu' || n === 'usa' || n === 'estados unidos') return 'eeuu';
    if (n === 'holanda' || n === 'paises bajos' || n === 'países bajos') return 'holanda';
    return n;
}



function initNavigation() {
    const menuToggle = document.getElementById('menuToggle');
    const navOverlay = document.getElementById('navOverlay');
    const navLinks = document.querySelectorAll('.nav-links a');
    const sections = document.querySelectorAll('.view-section');
    const roundSelector = document.getElementById('roundSelector');

    const prevBtn = document.getElementById('prevRoundBtn');
    const nextBtn = document.getElementById('nextRoundBtn');


    if (menuToggle && navOverlay) {
        menuToggle.addEventListener('click', () => {
            menuToggle.classList.toggle('open');
            navOverlay.classList.toggle('open');
        });
    }


    const path = window.location.pathname;
    const estaEnIndex = path.endsWith('index.html') || path.endsWith('/') || !path.includes('.html');

    navLinks.forEach(link => {
        link.addEventListener('click', (e) => {
            const target = link.getAttribute('data-target');

            if (estaEnIndex && target) {
                const targetSection = document.getElementById(`view-${target}`);
                if (targetSection) {
                    e.preventDefault();

                    navLinks.forEach(l => l.classList.remove('active'));
                    link.classList.add('active');

                    if (menuToggle) menuToggle.classList.remove('open');
                    if (navOverlay) navOverlay.classList.remove('open');

                    sections.forEach(sec => sec.classList.remove('active-view'));
                    targetSection.classList.add('active-view');
                }
            }

        });
    });


    if (roundSelector) {
        roundSelector.addEventListener('change', (e) => {
            renderizarRonda(e.target.value);
        });
    }


    if (prevBtn && roundSelector) {
        prevBtn.addEventListener('click', () => {
            if (roundSelector.selectedIndex > 0) {
                roundSelector.selectedIndex--;
                roundSelector.dispatchEvent(new Event('change'));
            }
        });
    }


    if (nextBtn && roundSelector) {
        nextBtn.addEventListener('click', () => {
            if (roundSelector.selectedIndex < roundSelector.options.length - 1) {
                roundSelector.selectedIndex++;
                roundSelector.dispatchEvent(new Event('change'));
            }
        });
    }
}




async function cargarDatosProde() {

    if (!document.getElementById('matches-container')) return;

    try {
        const SHEET_ID = '1QZDecj0_Azvu6zKxToICv4vUZjqOMfnHGzRlTLI3J3o';
        const XLSX_URL = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/export?format=xlsx`;

        const response = await fetch(XLSX_URL);
        if (!response.ok) throw new Error('No se pudo leer Google Sheets');

        const arrayBuffer = await response.arrayBuffer();
        const workbook = XLSX.read(arrayBuffer, {
            type: 'array'
        });
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
    const ahora = new Date();

    console.log("Hora local:", ahora.toString());

    for (const fase of CRONOGRAMA_FASES) {
        const inicio = new Date(fase.inicio);
        const fin = new Date(fase.fin);

        console.log(
            `${fase.id}: ${inicio.toString()} -> ${fin.toString()}`
        );

        if (ahora >= inicio && ahora <= fin) {
            console.log("Fase seleccionada:", fase.id);
            return fase.id;
        }
    }

    console.warn("No se encontró ninguna fase activa.");
    return CRONOGRAMA_FASES[0].id;
}


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


    container.innerHTML = '';

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


            acumLucas += pL;
            acumTomas += pT;


            if (rId === roundId) {
                faseLucas = pL;
                faseTomas = pT;
                break;
            }
        }
    }

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

    matches.forEach(m => {
        const idRaw = m.match_id !== undefined ? m.match_id : m['match-id'];
        if (idRaw === undefined) return;

        const idStr = String(idRaw).trim();
        const timeInfo = horariosMap[idStr];

        const res1 = m.result_team_1;
        const res2 = m.result_team_2;

        let centroHTML = '';
        let estado = 'antes';

        const status = String(m.status || '').trim();

        console.log({
        match: idStr,
        status: m.status,
        res1,
        res2
        });

        if (status === 'FINISHED') {

            estado = 'pos';

            centroHTML = `
        <span class="score-main">${res1} - ${res2}</span>
        <span class="match-meta-text">FT</span>
    `;

        } else if (status === 'LIVE' || status === 'PAUSED' || status === "IN_PLAY") {

            estado = 'live';

            centroHTML = `
        <span class="score-main live-score">
            ${res1 ?? 0} - ${res2 ?? 0}
        </span>

        <span class="match-meta-text live">
            En vivo
        </span>
    `;

        } else if (status === 'EXTRA_TIME') {

            estado = 'extra-time';

            centroHTML = `
        <span class="score-main">
            ${res1 ?? 0} - ${res2 ?? 0}
        </span>

        <span class="match-meta-text extra-time">
            Tiempo Extra
        </span>
    `;

        } else if (status === 'PENALTY_SHOOTOUT') {

            estado = 'penalties';

            centroHTML = `
        <span class="score-main">
            ${res1 ?? 0} - ${res2 ?? 0}
        </span>

        <span class="match-meta-text penalties">
            Penales
        </span>
    `;

    console.log(idStr, status, timeInfo);

        } else if (timeInfo) {

    estado = 'antes';

    const fecha = new Date(
        2026,
        Number(timeInfo.m) - 1,
        Number(timeInfo.d),
        Number(timeInfo.hr),
        Number(timeInfo.min)
    );

    const fechaTexto = fecha.toLocaleDateString('es-AR', {
        weekday: 'short',
        day: 'numeric',
        month: 'short'
    });

    const horaTexto = fecha.toLocaleTimeString('es-AR', {
        hour: '2-digit',
        minute: '2-digit',
        hour12: false
    });

    centroHTML = `
        <span class="match-date">${fechaTexto}</span>
        <span class="match-time">${horaTexto}</span>
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


// Función para crear el botón flotante
function inicializarBotonAutoScroll() {

    if (document.getElementById('btn-next-match')) return;

    const btn = document.createElement('button');
    btn.id = 'btn-next-match';
    btn.title = 'Ir al próximo partido pendiente';


    btn.innerHTML = `
        <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" fill="currentColor" viewBox="0 0 16 16">
          <path d="M12.5 4a.5.5 0 0 0-1 0v3.248L5.233 3.612A.5.5 0 0 0 4.5 4.008v7.984a.5.5 0 0 0 .733.432L11.5 8.752V12a.5.5 0 0 0 1 0zM5.5 4.975 10.045 8 5.5 11.025z"/>
        </svg>
    `;


    btn.addEventListener('click', () => {

        const tarjetasPartidos = document.querySelectorAll('.match-box');
        let encontrado = false;

        for (let tarjeta of tarjetasPartidos) {

            if (!tarjeta.textContent.includes('FT')) {
                tarjeta.scrollIntoView({ behavior: 'smooth', block: 'center' });
                encontrado = true;
                break;
            }
        }

        if (!encontrado && tarjetasPartidos.length > 0) {
            tarjetasPartidos[tarjetasPartidos.length - 1].scrollIntoView({ behavior: 'smooth', block: 'center' });
        }
    });

    document.body.appendChild(btn);
}

document.addEventListener('DOMContentLoaded', inicializarBotonAutoScroll);