const SHEET_NAME_DATA = "Planilha Atual".trim().toLowerCase();
const SHEET_NAME_HISTORY = "historico".trim().toLowerCase();
const CUTOFF_DAY = 9;

const COL = { 
    CLIENTE: 0,     
    DATA_INICIAL: 4,
    DATA_FINAL: 5,  
    EMPRESTIMO: 6,  
    COL_H: 7,
    TOTAL_PAGAR: 9,  
    VALOR_PAGO: 10,  
    COL_L: 11,
    COL_M: 12,
    COL_N: 13,
    DIARIAS: 14,     
    OBSERVACAO: 15,  
    STATUS_CLI: 16,  
    DATA_ATUAL: 17   
};

let currentZoom = 85;
let currentTab = 'consolidated';
let allStatesData = [];

document.addEventListener('DOMContentLoaded', () => {
    updateDateTime();
    setupEventListeners();
    setInterval(updateDateTime, 1000);
    document.body.style.zoom = currentZoom + '%';
    const zoomEl = document.getElementById('zoomLevel');
    if (zoomEl) zoomEl.textContent = currentZoom + '%';
});

function setupEventListeners() {
    const area = document.getElementById('uploadArea');
    const input = document.getElementById('fileInput');
    
    if (area && input) {
        area.onclick = () => input.click();
        area.ondragover = (e) => { e.preventDefault(); area.classList.add('dragover'); };
        area.ondragleave = () => area.classList.remove('dragover');
        area.ondrop = (e) => { 
            e.preventDefault(); 
            area.classList.remove('dragover'); 
            if(e.dataTransfer.files.length) handleFiles(e.dataTransfer.files); 
        };
        input.onchange = (e) => { if(e.target.files.length) handleFiles(e.target.files); };
    }
}

async function handleFiles(files) {
    showLoading();
    allStatesData = [];
    
    const fileArray = Array.from(files);
    for (const file of fileArray) {
        try {
            const stateData = await processFile(file);
            if (stateData) {
                allStatesData.push(stateData);
            }
        } catch (err) {
            console.error(`Erro ao processar ${file.name}:`, err);
        }
    }
    
    if (allStatesData.length > 0) {
        renderAll();
        document.getElementById('uploadArea').style.display = 'none';
        document.getElementById('tabsContainer').style.display = 'flex';
        document.getElementById('bottomToolbar').style.display = 'flex';
        switchTab('consolidated');
    } else {
        showError('Nenhum dado válido encontrado nos arquivos.');
    }
    
    hideLoading();
}

function renderAll() {
    const consolidatedView = document.getElementById('consolidatedView');
    const monthlyView = document.getElementById('monthlyView');
    const historyView = document.getElementById('historyView');
    
    consolidatedView.innerHTML = '';
    monthlyView.innerHTML = '';
    historyView.innerHTML = '';
    
    renderConsolidatedReport(consolidatedView);
    
    allStatesData.forEach(state => {
        if (Object.keys(state.monthly).length > 0) {
            renderStateReport(state, 'monthly', monthlyView);
        }
        if (Object.keys(state.history).length > 0) {
            renderStateReport(state, 'history', historyView);
        }
    });
}

function processFile(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        const stateName = file.name.replace(/\.[^/.]+$/, ""); 

        reader.onload = (e) => {
            try {
                const workbook = XLSX.read(new Uint8Array(e.target.result), { type: 'array' });
                const stateResult = {
                    name: stateName,
                    monthly: {},
                    history: {},
                    filters: { monthly: '', history: '' }
                };

                const sheetNameFound = workbook.SheetNames.find(name => 
                    name.trim().toLowerCase() === SHEET_NAME_DATA
                );
                if (sheetNameFound) {
                    const worksheet = workbook.Sheets[sheetNameFound];
                    const data = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
                    stateResult.monthly = extractData(data);
                }
                
                const sheetNameHistoryFound = workbook.SheetNames.find(name => 
                    name.trim().toLowerCase() === SHEET_NAME_HISTORY
                );
                if (sheetNameHistoryFound) {
                    const worksheetHistory = workbook.Sheets[sheetNameHistoryFound];
                    const dataHistory = XLSX.utils.sheet_to_json(worksheetHistory, { header: 1, defval: '' });
                    stateResult.history = extractData(dataHistory);
                }
                
                resolve(stateResult);
            } catch (err) {
                reject(err);
            }
        };
        reader.onerror = reject;
        reader.readAsArrayBuffer(file);
    });
}

function extractData(rows) {
    const result = {};
    const hoje = new Date();
    
    for (let i = 1; i < rows.length; i++) {
        const row = rows[i];
        if (!row || row.length < 5) continue;

        const dataRefPeriodo = parseExcelDate(row[COL.DATA_INICIAL]);
        const periodo = extractPeriodo(dataRefPeriodo);
        if (periodo === 'Sem Data') continue;

        const valorEmprestado = getNum(row[COL.EMPRESTIMO]);
        const valorPrevisto = getNum(row[COL.TOTAL_PAGAR]);
        const valorPago = getNum(row[COL.VALOR_PAGO]);
        
        const colO = getNum(row[COL.DIARIAS]);
        const colH = getNum(row[COL.COL_H]);
        const colM = getNum(row[COL.COL_M]);
        const colL = getNum(row[COL.COL_L]);
        const colN = getNum(row[COL.COL_N]);
        
        // [MODIFICADO] Lucro Real = Valor Pago - Valor Previsto (Sinal positivo se Pago > Previsto, negativo se Pago < Previsto)
        const lucroReal = valorPago - valorPrevisto;
        // [MODIFICADO] Multas = (M * H) + (L * 20) + N (Ignorando coluna O)
        const multas = (colM * colH) + (colL * 20) + colN;
        
        const saldoAberto = valorPrevisto > valorPago ? valorPrevisto - valorPago : 0;
        const status = determinarStatusFinal(row, dataRefPeriodo, hoje, valorPrevisto, valorPago);

        if (!result[periodo]) {
            result[periodo] = createEmptyPeriod();
        }

        const p = result[periodo];
        p.totalEmprestado += valorEmprestado;
        p.totalPrevisto += valorPrevisto;
        p.totalPago += valorPago;
        p.totalLucro += lucroReal;
        // [MODIFICADO] Acumulando multas (mesmo valor do lucro real)
        p.totalMultas += multas;
        p.totalSaldoAberto += saldoAberto;

        if (dataRefPeriodo && dataRefPeriodo.getDate() <= CUTOFF_DAY) {
            p.emprestadoAte09 += valorEmprestado;
            p.previstoAte09 += valorPrevisto;
            p.pagoAte09 += valorPago;
            p.lucroRealAte09 += (valorPago - valorPrevisto); // [NOVO] Lucro Real até dia 9
        }

        if (p.status[status] !== undefined) {
            p.status[status]++;
        } else {
            p.status['ATIVO']++;
        }
    }
    return result;
}

function createEmptyPeriod() {
    return {
        totalEmprestado: 0, 
        totalPrevisto: 0, 
        totalPago: 0, 
        totalLucro: 0,
        totalMultas: 0, // [NOVO] Campo para multas
        totalSaldoAberto: 0,
        emprestadoAte09: 0, 
        previstoAte09: 0, 
        pagoAte09: 0,
        lucroRealAte09: 0, // [NOVO] Lucro Real até dia 9
        status: { 
            'ATIVO': 0, 
            'NÃO EMPRESTAR': 0, 
            'QUITADO': 0, 
            'AMARELADO': 0, 
            'COBRAR': 0, 
            'CLIENTE EM ACORDO': 0, 
            'VERDE': 0 
        }
    };
}

function renderConsolidatedReport(container) {
    let globalTotals = createEmptyPeriod();
    let allPeriods = {};
    let consolidatedFilter = '';
    
    allStatesData.forEach(state => {
        Object.entries(state.monthly).forEach(([period, d]) => {
            if (!allPeriods[period]) {
                allPeriods[period] = createEmptyPeriod();
            }
            allPeriods[period].totalEmprestado += d.totalEmprestado;
            allPeriods[period].totalPrevisto += d.totalPrevisto;
            allPeriods[period].totalPago += d.totalPago;
            allPeriods[period].totalLucro += d.totalLucro;
            allPeriods[period].totalMultas += d.totalMultas; // [NOVO]
            allPeriods[period].totalSaldoAberto += d.totalSaldoAberto;
            allPeriods[period].emprestadoAte09 += d.emprestadoAte09;
            allPeriods[period].previstoAte09 += d.previstoAte09;
            allPeriods[period].pagoAte09 += d.pagoAte09;
            allPeriods[period].lucroRealAte09 += d.lucroRealAte09;
            Object.keys(d.status).forEach(s => allPeriods[period].status[s] += d.status[s]);
        });
        
        Object.entries(state.history).forEach(([period, d]) => {
            if (!allPeriods[period]) {
                allPeriods[period] = createEmptyPeriod();
            }
            allPeriods[period].totalEmprestado += d.totalEmprestado;
            allPeriods[period].totalPrevisto += d.totalPrevisto;
            allPeriods[period].totalPago += d.totalPago;
            allPeriods[period].totalLucro += d.totalLucro;
            allPeriods[period].totalMultas += d.totalMultas; // [NOVO]
            allPeriods[period].totalSaldoAberto += d.totalSaldoAberto;
            allPeriods[period].emprestadoAte09 += d.emprestadoAte09;
            allPeriods[period].previstoAte09 += d.previstoAte09;
            allPeriods[period].pagoAte09 += d.pagoAte09;
            allPeriods[period].lucroRealAte09 += d.lucroRealAte09;
            Object.keys(d.status).forEach(s => allPeriods[period].status[s] += d.status[s]);
        });
    });
    
    Object.values(allPeriods).forEach(d => {
        globalTotals.totalEmprestado += d.totalEmprestado;
        globalTotals.totalPrevisto += d.totalPrevisto;
        globalTotals.totalPago += d.totalPago;
        globalTotals.totalLucro += d.totalLucro;
        globalTotals.totalMultas += d.totalMultas; // [NOVO]
        globalTotals.totalSaldoAberto += d.totalSaldoAberto;
        globalTotals.emprestadoAte09 += d.emprestadoAte09;
        globalTotals.previstoAte09 += d.previstoAte09;
        globalTotals.pagoAte09 += d.pagoAte09;
        globalTotals.lucroRealAte09 += d.lucroRealAte09;
        Object.keys(d.status).forEach(s => globalTotals.status[s] += d.status[s]);
    });
    
    const allPeriodsKeys = Object.keys(allPeriods).sort().reverse();
    const periodosPorAno = {};
    allPeriodsKeys.forEach(p => {
        const ano = p.split('-')[0];
        if (!periodosPorAno[ano]) periodosPorAno[ano] = [];
        periodosPorAno[ano].push(p);
    });
    
    let filteredPeriods = allPeriodsKeys;
    const storedFilter = localStorage.getItem('consolidatedFilter') || '';
    if (storedFilter) {
        consolidatedFilter = storedFilter;
        if (storedFilter.startsWith('ano:')) {
            const ano = storedFilter.split(':')[1];
            filteredPeriods = filteredPeriods.filter(p => p.startsWith(ano));
        } else {
            filteredPeriods = filteredPeriods.filter(p => p === storedFilter);
        }
    }
    
    let filteredTotals = createEmptyPeriod();
    filteredPeriods.forEach(p => {
        const d = allPeriods[p];
        filteredTotals.totalEmprestado += d.totalEmprestado;
        filteredTotals.totalPrevisto += d.totalPrevisto;
        filteredTotals.totalPago += d.totalPago;
        filteredTotals.totalLucro += d.totalLucro;
        filteredTotals.totalMultas += d.totalMultas; // [NOVO]
        filteredTotals.totalSaldoAberto += d.totalSaldoAberto;
        filteredTotals.emprestadoAte09 += d.emprestadoAte09;
        filteredTotals.previstoAte09 += d.previstoAte09;
        filteredTotals.pagoAte09 += d.pagoAte09;
        Object.keys(d.status).forEach(s => filteredTotals.status[s] += d.status[s]);
    });

    const stateNames = allStatesData.map(s => s.name).join(', ');
    const reportDiv = document.createElement('div');
    reportDiv.className = 'state-report';
    reportDiv.innerHTML = `
        <div class="state-title">
            <span>🌍</span> CONSOLIDADO GERAL - ${stateNames}
        </div>
        <div style="font-size: 12px; color: var(--text-muted); margin-bottom: 15px; padding: 10px; background: var(--bg-glass); border-radius: 6px; border-left: 3px solid var(--accent-primary);">
            <strong>Planilhas Selecionadas:</strong> ${stateNames}
        </div>
        
        <div class="filter-container">
            <span class="filter-label">Filtrar Período:</span>
            <select class="filter-select" onchange="applyConsolidatedFilter(this.value)">
                <option value="">📅 Todos os Períodos</option>
                ${Object.keys(periodosPorAno).sort().reverse().map(ano => `
                    <optgroup label="Ano ${ano}">
                        <option value="ano:${ano}" ${consolidatedFilter === 'ano:'+ano ? 'selected' : ''}>📅 Ano ${ano} (Completo)</option>
                        ${periodosPorAno[ano].map(p => `
                            <option value="${p}" ${consolidatedFilter === p ? 'selected' : ''}>${formatPeriodo(p)}</option>
                        `).join('')}
                    </optgroup>
                `).join('')}
            </select>
            <button class="filter-btn clear" onclick="applyConsolidatedFilter('')">Limpar</button>
        </div>
        
        <div class="card-grid">
            <div class="card" style="border-left-color: #7030A0;"><div class="card-title">Total Emprestado</div><div class="card-value">${formatCurrency(filteredTotals.totalEmprestado)}</div></div>
            <div class="card" style="border-left-color: #2E75B6;"><div class="card-title">Total Previsto</div><div class="card-value">${formatCurrency(filteredTotals.totalPrevisto)}</div></div>
            <div class="card" style="border-left-color: #00B0F0;"><div class="card-title">Total Recebido</div><div class="card-value">${formatCurrency(filteredTotals.totalPago)}</div></div>
            <div class="card" style="border-left-color: #ED7D31;"><div class="card-title">Saldo em Aberto (Girando)</div><div class="card-value">${formatCurrency(filteredTotals.totalSaldoAberto)}</div></div>
            <div class="card" style="border-left-color: #70AD47;"><div class="card-title">Lucro Real</div><div class="card-value">${formatCurrency(filteredTotals.totalLucro)}</div></div>
        </div>
        
        <div class="section-title">Detalhamento por Período (Consolidado)</div>
        <div class="table-container">
            <table>
                <thead>
                    <tr>
                        <th>Período</th>
                        <th class="text-right">Emprest. (Até 9)</th>
                        <th class="text-right">Previsto (Até 9)</th>
                        <th class="text-right">Pago (Até 9)</th>
                        <th class="text-right">Lucro Real (Até 9)</th>
                        <th class="text-right">Total Emprestado</th>
                        <th class="text-right">Total Previsto</th>
                        <th class="text-right">Total Pago</th>
                        <th class="text-right">Saldo Aberto</th>
                        <th class="text-right">Multas</th>
                        <th class="text-right">Lucro Real</th>
                    </tr>
                </thead>
                <tbody>
                    ${filteredPeriods.map(p => {
                        const d = allPeriods[p];
                        return `
                            <tr>
                                <td>${formatPeriodo(p)}</td>
                                <td class="text-right">${formatCurrency(d.emprestadoAte09)}</td>
                                <td class="text-right">${formatCurrency(d.previstoAte09)}</td>
                                <td class="text-right">${formatCurrency(d.pagoAte09)}</td>
                                <td class="text-right ${d.lucroRealAte09 > 0 ? 'text-success' : (d.lucroRealAte09 < 0 ? 'text-danger' : '')}">${formatCurrency(d.lucroRealAte09)}</td>
                                <td class="text-right">${formatCurrency(d.totalEmprestado)}</td>
                                <td class="text-right">${formatCurrency(d.totalPrevisto)}</td>
                                <td class="text-right">${formatCurrency(d.totalPago)}</td>
                                <td class="text-right">${formatCurrency(d.totalSaldoAberto)}</td>
                                <td class="text-right ${d.totalMultas > 0 ? 'text-warning' : ''}">${formatCurrency(d.totalMultas)}</td>
                                <td class="text-right ${d.totalLucro > 0 ? 'text-success' : (d.totalLucro < 0 ? 'text-danger' : '')}">${formatCurrency(d.totalLucro)}</td>
                            </tr>
                        `;
                    }).join('')}
                    <tr class="totals-row">
                        <td>TOTAL FILTRADO</td>
                        <td class="text-right">${formatCurrency(filteredTotals.emprestadoAte09)}</td>
                        <td class="text-right">${formatCurrency(filteredTotals.previstoAte09)}</td>
                        <td class="text-right">${formatCurrency(filteredTotals.pagoAte09)}</td>
                        <td class="text-right">${formatCurrency(filteredTotals.lucroRealAte09)}</td>
                        <td class="text-right">${formatCurrency(filteredTotals.totalEmprestado)}</td>
                        <td class="text-right">${formatCurrency(filteredTotals.totalPrevisto)}</td>
                        <td class="text-right">${formatCurrency(filteredTotals.totalPago)}</td>
                        <td class="text-right">${formatCurrency(filteredTotals.totalSaldoAberto)}</td>
                        <td class="text-right">${formatCurrency(filteredTotals.totalMultas)}</td>
                        <td class="text-right">${formatCurrency(filteredTotals.totalLucro)}</td>
                    </tr>
                </tbody>
            </table>
        </div>

        <div class="section-title">Status da Carteira (Consolidado)</div>
        <div class="card-grid">
            <div class="card" style="border-left-color: var(--status-active);"><div class="card-title">Ativos</div><div class="card-value">${filteredTotals.status['ATIVO']}</div></div>
            <div class="card" style="border-left-color: var(--status-verde);"><div class="card-title">Verde</div><div class="card-value">${filteredTotals.status['VERDE']}</div></div>
            <div class="card" style="border-left-color: var(--status-amarelo);"><div class="card-title">Amarelado</div><div class="card-value">${filteredTotals.status['AMARELADO']}</div></div>
            <div class="card" style="border-left-color: var(--status-vermelho);"><div class="card-title">Cobrar</div><div class="card-value">${filteredTotals.status['COBRAR']}</div></div>
            <div class="card" style="border-left-color: var(--status-acordo);"><div class="card-title">Em Acordo</div><div class="card-value">${filteredTotals.status['CLIENTE EM ACORDO']}</div></div>
            <div class="card" style="border-left-color: #8b949e;"><div class="card-title">Inconsistência</div><div class="card-value">${filteredTotals.status['QUITADO']}</div></div>
        </div>
    `;
    container.appendChild(reportDiv);
}

function renderStateReport(state, type, container) {
    const reportDiv = document.createElement('div');
    reportDiv.className = 'state-report';
    reportDiv.id = `report-${state.name}-${type}`;
    
    const data = type === 'monthly' ? state.monthly : state.history;
    const filterValue = state.filters[type];
    
    let filteredPeriods = Object.keys(data).sort().reverse();
    if (filterValue) {
        if (filterValue.startsWith('ano:')) {
            const ano = filterValue.split(':')[1];
            filteredPeriods = filteredPeriods.filter(p => p.startsWith(ano));
        } else {
            filteredPeriods = filteredPeriods.filter(p => p === filterValue);
        }
    }
    
    let totals = createEmptyPeriod();
    filteredPeriods.forEach(p => {
        const d = data[p];
        totals.totalEmprestado += d.totalEmprestado;
        totals.totalPrevisto += d.totalPrevisto;
        totals.totalPago += d.totalPago;
        totals.totalLucro += d.totalLucro;
        totals.totalMultas += d.totalMultas; // [NOVO]
        totals.totalSaldoAberto += d.totalSaldoAberto;
        totals.emprestadoAte09 += d.emprestadoAte09;
        totals.previstoAte09 += d.previstoAte09;
        totals.pagoAte09 += d.pagoAte09;
        totals.lucroRealAte09 += d.lucroRealAte09;
        Object.keys(d.status).forEach(s => totals.status[s] += d.status[s]);
    });

    const typeLabel = type === 'monthly' ? 'RELATÓRIO ATUAL' : 'HISTÓRICO GERAL';
    const allPeriods = Object.keys(data).sort().reverse();
    
    const periodosPorAno = {};
    allPeriods.forEach(p => {
        const ano = p.split('-')[0];
        if (!periodosPorAno[ano]) periodosPorAno[ano] = [];
        periodosPorAno[ano].push(p);
    });

    reportDiv.innerHTML = `
        <div class="state-title">
            <span>📍</span> ${state.name.toUpperCase()} - ${typeLabel}
        </div>

        <div class="filter-container">
            <span class="filter-label">Filtrar Período:</span>
            <select class="filter-select" onchange="applyStateFilter('${state.name}', '${type}', this.value)">
                <option value="">📅 Todos os Períodos</option>
                ${Object.keys(periodosPorAno).sort().reverse().map(ano => `
                    <optgroup label="Ano ${ano}">
                        <option value="ano:${ano}" ${filterValue === 'ano:'+ano ? 'selected' : ''}>📅 Ano ${ano} (Completo)</option>
                        ${periodosPorAno[ano].map(p => `
                            <option value="${p}" ${filterValue === p ? 'selected' : ''}>${formatPeriodo(p)}</option>
                        `).join('')}
                    </optgroup>
                `).join('')}
            </select>
            <button class="filter-btn clear" onclick="applyStateFilter('${state.name}', '${type}', '')">Limpar</button>
        </div>
        
        <div class="card-grid">
            <div class="card" style="border-left-color: #7030A0;"><div class="card-title">Total Emprestado</div><div class="card-value">${formatCurrency(totals.totalEmprestado)}</div></div>
            <div class="card" style="border-left-color: #2E75B6;"><div class="card-title">Total Previsto</div><div class="card-value">${formatCurrency(totals.totalPrevisto)}</div></div>
            <div class="card" style="border-left-color: #00B0F0;"><div class="card-title">Total Recebido</div><div class="card-value">${formatCurrency(totals.totalPago)}</div></div>
            <div class="card" style="border-left-color: #ED7D31;"><div class="card-title">Saldo em Aberto (Girando)</div><div class="card-value">${formatCurrency(totals.totalSaldoAberto)}</div></div>
            <div class="card" style="border-left-color: #70AD47;"><div class="card-title">Lucro Real</div><div class="card-value">${formatCurrency(totals.totalLucro)}</div></div>
        </div>

        <div class="section-title">Detalhamento por Período</div>
        <div class="table-container">
            <table>
                <thead>
                    <tr>
                        <th>Período</th>
                        <th class="text-right">Emprest. (Até 9)</th>
                        <th class="text-right">Previsto (Até 9)</th>
                        <th class="text-right">Pago (Até 9)</th>
                        <th class="text-right">Lucro Real (Até 9)</th>
                        <th class="text-right">Total Emprestado</th>
                        <th class="text-right">Total Previsto</th>
                        <th class="text-right">Total Pago</th>
                        <th class="text-right">Saldo Aberto</th>
                        <th class="text-right">Multas</th>
                        <th class="text-right">Lucro Real</th>
                    </tr>
                </thead>
                <tbody>
                    ${filteredPeriods.map(p => {
                        const d = data[p];
                        return `
                            <tr>
                                <td>${formatPeriodo(p)}</td>
                                <td class="text-right">${formatCurrency(d.emprestadoAte09)}</td>
                                <td class="text-right">${formatCurrency(d.previstoAte09)}</td>
                                <td class="text-right">${formatCurrency(d.pagoAte09)}</td>
                                <td class="text-right ${d.lucroRealAte09 > 0 ? 'text-success' : (d.lucroRealAte09 < 0 ? 'text-danger' : '')}">${formatCurrency(d.lucroRealAte09)}</td>
                                <td class="text-right">${formatCurrency(d.totalEmprestado)}</td>
                                <td class="text-right">${formatCurrency(d.totalPrevisto)}</td>
                                <td class="text-right">${formatCurrency(d.totalPago)}</td>
                                <td class="text-right">${formatCurrency(d.totalSaldoAberto)}</td>
                                <td class="text-right ${d.totalMultas > 0 ? 'text-warning' : ''}">${formatCurrency(d.totalMultas)}</td>
                                <td class="text-right ${d.totalLucro > 0 ? 'text-success' : (d.totalLucro < 0 ? 'text-danger' : '')}">${formatCurrency(d.totalLucro)}</td>
                            </tr>
                        `;
                    }).join('')}
                    <tr class="totals-row">
                        <td>TOTAL FILTRADO</td>
                        <td class="text-right">${formatCurrency(totals.emprestadoAte09)}</td>
                        <td class="text-right">${formatCurrency(totals.previstoAte09)}</td>
                        <td class="text-right">${formatCurrency(totals.pagoAte09)}</td>
                        <td class="text-right">${formatCurrency(totals.lucroRealAte09)}</td>
                        <td class="text-right">${formatCurrency(totals.totalEmprestado)}</td>
                        <td class="text-right">${formatCurrency(totals.totalPrevisto)}</td>
                        <td class="text-right">${formatCurrency(totals.totalPago)}</td>
                        <td class="text-right">${formatCurrency(totals.totalSaldoAberto)}</td>
                        <td class="text-right">${formatCurrency(totals.totalMultas)}</td>
                        <td class="text-right">${formatCurrency(totals.totalLucro)}</td>
                    </tr>
                </tbody>
            </table>
        </div>

        <div class="section-title">Status da Carteira</div>
        <div class="card-grid">
            <div class="card" style="border-left-color: var(--status-active);"><div class="card-title">Ativos</div><div class="card-value">${totals.status['ATIVO']}</div></div>
            <div class="card" style="border-left-color: var(--status-verde);"><div class="card-title">Verde</div><div class="card-value">${totals.status['VERDE']}</div></div>
            <div class="card" style="border-left-color: var(--status-amarelo);"><div class="card-title">Amarelado</div><div class="card-value">${totals.status['AMARELADO']}</div></div>
            <div class="card" style="border-left-color: var(--status-vermelho);"><div class="card-title">Cobrar</div><div class="card-value">${totals.status['COBRAR']}</div></div>
            <div class="card" style="border-left-color: var(--status-acordo);"><div class="card-title">Em Acordo</div><div class="card-value">${totals.status['CLIENTE EM ACORDO']}</div></div>
            <div class="card" style="border-left-color: #8b949e;"><div class="card-title">Inconsistência</div><div class="card-value">${totals.status['QUITADO']}</div></div>
        </div>
    `;
    
    container.appendChild(reportDiv);
}

function applyStateFilter(stateName, type, value) {
    const state = allStatesData.find(s => s.name === stateName);
    if (state) {
        state.filters[type] = value;
        renderAll();
    }
}

function switchTab(tab) {
    currentTab = tab;
    document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
    
    const tabConsolidated = document.getElementById('tabConsolidated');
    const tabMonthly = document.getElementById('tabMonthly');
    const tabHistory = document.getElementById('tabHistory');
    
    if (tab === 'consolidated' && tabConsolidated) tabConsolidated.classList.add('active');
    if (tab === 'monthly' && tabMonthly) tabMonthly.classList.add('active');
    if (tab === 'history' && tabHistory) tabHistory.classList.add('active');
    
    document.getElementById('consolidatedView').style.display = tab === 'consolidated' ? 'block' : 'none';
    document.getElementById('monthlyView').style.display = tab === 'monthly' ? 'block' : 'none';
    document.getElementById('historyView').style.display = tab === 'history' ? 'block' : 'none';
}

function determinarStatusFinal(linha, dataRef, hoje, valorPrevisto, valorPago) {
    const colO_Diarias = getNum(linha[COL.DIARIAS]);
    const colP_Obs = String(linha[COL.OBSERVACAO] || '').toLowerCase().trim();
    const colF_DataFinal = parseExcelDate(linha[COL.DATA_FINAL]);
    const colR_DataAtual = parseExcelDate(linha[COL.DATA_ATUAL]);
    
    if (colP_Obs.includes('acordo')) return 'CLIENTE EM ACORDO';
    if (colO_Diarias === 20) return 'VERDE';
    
    if (colF_DataFinal && colR_DataAtual && colO_Diarias !== 20) {
        const diffTempo = colR_DataAtual - colF_DataFinal;
        const diffDias = Math.ceil(diffTempo / (1000 * 60 * 60 * 24));
        
        if (diffDias >= 90) return 'COBRAR';
        if (diffDias >= 1) return 'AMARELADO';
    }
    
    const statusTxt = String(linha[COL.STATUS_CLI] || '').toUpperCase();
    if (statusTxt.includes('NÃO') || statusTxt.includes('NAO')) return 'NÃO EMPRESTAR';
    if (valorPago >= valorPrevisto && valorPrevisto > 0) return 'QUITADO';
    
    return 'ATIVO';
}

function parseExcelDate(val) {
    if (!val) return null;
    try {
        if (val instanceof Date) return val;
        if (typeof val === 'number') {
            const d = XLSX.SSF.parse_date_code(val);
            if (d) return new Date(d.y, d.m - 1, d.d);
        }
        if (typeof val === 'string') {
            if (val.includes('/')) {
                const p = val.split('/');
                if (p.length === 3) return new Date(p[2], p[1] - 1, p[0]);
            }
            const d = new Date(val);
            if (!isNaN(d.getTime())) return d;
        }
    } catch (e) {}
    return null;
}

function extractPeriodo(date) {
    if (!date) return 'Sem Data';
    return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}`;
}

function formatPeriodo(p) {
    if (p === 'Sem Data') return p;
    const parts = p.split('-');
    if (parts.length < 2) return p;
    const [year, month] = parts;
    const months = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez'];
    return `${months[parseInt(month)-1]}/${year}`;
}

function getNum(v) {
    if (v === null || v === undefined || v === '') return 0;
    if (typeof v === 'number') return v;
    try {
        let cleaned = String(v).replace(/[R$\s]/g, '').replace(/\./g, '').replace(',', '.');
        let n = parseFloat(cleaned);
        return isNaN(n) ? 0 : n;
    } catch (e) { return 0; }
}

function formatCurrency(v) {
    return new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL' }).format(v);
}

function updateDateTime() {
    const agora = new Date();
    const opcoes = { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric', hour: '2-digit', minute: '2-digit' };
    const el = document.getElementById('currentDateTime');
    if (el) el.textContent = 'Gerado em ' + agora.toLocaleDateString('pt-BR', opcoes);
}

function changeZoom(delta) {
    currentZoom += delta;
    if (currentZoom < 50) currentZoom = 50;
    if (currentZoom > 150) currentZoom = 150;
    document.body.style.zoom = currentZoom + '%';
    const el = document.getElementById('zoomLevel');
    if (el) el.textContent = currentZoom + '%';
}

function showLoading() { 
    const el = document.getElementById('loadingOverlay');
    if (el) el.style.display = 'flex'; 
}
function hideLoading() { 
    const el = document.getElementById('loadingOverlay');
    if (el) el.style.display = 'none'; 
}
function showError(msg) {
    const err = document.getElementById('errorMessage');
    if (err) {
        err.textContent = msg;
        err.style.display = 'block';
    }
}

function applyConsolidatedFilter(value) {
    localStorage.setItem('consolidatedFilter', value);
    renderAll();
}

function exportToPDF() {
    const printWindow = window.open('', '', 'width=1400,height=900');
    const htmlContent = document.documentElement.innerHTML;
    printWindow.document.write(htmlContent);
    printWindow.document.close();
    setTimeout(() => {
        printWindow.print();
    }, 250);
}
