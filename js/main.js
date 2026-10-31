// js/main.js - Lógica do Dashboard Analítico Claro PEPR & Integração de Modais/Toasts

let globalRawData = [];
let opChartInstance = null;
let doughnutChartInstance = null;

document.addEventListener("DOMContentLoaded", () => {
  // Inicialização do listener de upload de arquivos Excel
  const loadBtn = document.getElementById("loadBtn");
  if (loadBtn) {
    loadBtn.addEventListener("click", function () {
      const fileInput = document.getElementById("excelFile");
      const alertBox = document.getElementById("alertBox");
      if (alertBox) {
        alertBox.classList.add("d-none");
        alertBox.innerText = "";
      }

      if (!fileInput || fileInput.files.length === 0) {
        showAppModal(
          "Seleção de Arquivo Requerida",
          "Por favor, selecione uma base de dados no formato Excel (.xlsx ou .xls) para iniciar a análise.",
          "warning"
        );
        return;
      }

      const file = fileInput.files[0];
      const reader = new FileReader();

      reader.onload = function (e) {
        try {
          const data = new Uint8Array(e.target.result);
          const workbook = XLSX.read(data, { type: "array" });

          if (!workbook.SheetNames || workbook.SheetNames.length === 0) {
            showAppModal(
              "Arquivo Inválido",
              "Não foi possível encontrar planilhas válidas dentro do arquivo fornecido.",
              "danger"
            );
            return;
          }

          const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
          globalRawData = XLSX.utils.sheet_to_json(firstSheet);

          if (!globalRawData || globalRawData.length === 0) {
            showAlert(
              "A planilha está vazia ou os dados não puderam ser lidos."
            );
            showAppModal(
              "Planilha Vazia",
              "A planilha selecionada está sem registros ou possui formato inconsistente.",
              "warning"
            );
            return;
          }

          populateMonthFilter(globalRawData);
          processData(globalRawData);

          const dashContent = document.getElementById("dashboardContent");
          if (dashContent) dashContent.style.display = "block";

          // Toast dinâmico de sucesso
          showToast(
            `Base carregada com sucesso! ${globalRawData.length} registros processados.`,
            "success"
          );
        } catch (error) {
          showAlert("Erro ao processar a planilha: " + error.message);
          showAppModal(
            "Erro de Processamento",
            `Falha técnica ao tentar ler a estrutura da planilha: <strong>${error.message}</strong>`,
            "danger"
          );
        }
      };

      reader.readAsArrayBuffer(file);
    });
  }

  // Listener para mudança de filtro por mês
  const monthFilter = document.getElementById("monthFilter");
  if (monthFilter) {
    monthFilter.addEventListener("change", function () {
      processData(globalRawData);
      showToast("Filtro de período atualizado com sucesso.", "info");
    });
  }

  // Listener para exibição do nome do arquivo selecionado
  const excelFileInput = document.getElementById("excelFile");
  if (excelFileInput) {
    excelFileInput.addEventListener("change", function (e) {
      const fileName = e.target.files[0]
        ? e.target.files[0].name
        : "Nenhum arquivo selecionado...";
      const fileNameDisplay = document.getElementById("fileNameDisplay");
      if (fileNameDisplay) {
        fileNameDisplay.innerHTML = `<i class="fa-solid fa-file-excel me-2 text-success"></i>${fileName}`;
      }
    });
  }

  // Controle de alternância de Tema (Light / Dark)
  const themeToggleBtn = document.getElementById("themeToggleBtn");
  const themeLabel = document.getElementById("themeLabel");
  const themeIcon = document.getElementById("themeIcon");

  if (themeToggleBtn) {
    themeToggleBtn.addEventListener("click", () => {
      const currentTheme = document.documentElement.getAttribute("data-theme");
      const newTheme = currentTheme === "dark" ? "light" : "dark";

      document.documentElement.setAttribute("data-theme", newTheme);

      if (newTheme === "dark") {
        if (themeLabel) themeLabel.innerText = "Modo Claro";
        if (themeIcon) themeIcon.className = "fa-solid fa-sun";
      } else {
        if (themeLabel) themeLabel.innerText = "Modo Escuro";
        if (themeIcon) themeIcon.className = "fa-solid fa-moon";
      }
    });
  }
});

// Exibição de alertas fallback
function showAlert(msg) {
  const box = document.getElementById("alertBox");
  if (box) {
    box.innerText = msg;
    box.classList.remove("d-none");
  }
}

// Busca tolerante de nomes de colunas
function getColValue(row, possibleNames) {
  const keys = Object.keys(row);
  for (let name of possibleNames) {
    let match = keys.find(
      (k) => k.trim().toLowerCase() === name.trim().toLowerCase()
    );
    if (match !== undefined) {
      return row[match];
    }
  }
  return "";
}

// Preenchimento do filtro de meses
function populateMonthFilter(data) {
  const monthSelect = document.getElementById("monthFilter");
  if (!monthSelect) return;

  monthSelect.innerHTML =
    '<option value="ALL">📅 Todos os Meses (Consolidado Geral)</option>';

  const monthsSet = new Set();
  data.forEach((row) => {
    let m = getColValue(row, [
      "Início do Mês",
      "Mes",
      "Mês",
      "Data de Abertura",
      "Periodo",
      "Período",
    ]);
    if (m !== "" && m !== null && m !== undefined) {
      monthsSet.add(String(m).trim());
    }
  });

  const sortedMonths = Array.from(monthsSet).sort();
  sortedMonths.forEach((m) => {
    let opt = document.createElement("option");
    opt.value = m;
    opt.innerText = m;
    monthSelect.appendChild(opt);
  });

  monthSelect.disabled = sortedMonths.length === 0;
}

// Processamento central da base e cruzamento de dados
function processData(data) {
  const monthSelect = document.getElementById("monthFilter");
  const selectedMonth = monthSelect ? monthSelect.value : "ALL";

  const filteredData = data.filter((row) => {
    if (selectedMonth === "ALL") return true;
    let m = getColValue(row, [
      "Início do Mês",
      "Mes",
      "Mês",
      "Data de Abertura",
      "Periodo",
      "Período",
    ]);
    return String(m).trim() === selectedMonth;
  });

  const primeIdFrequency = {};
  const primeIdList = {};
  const causesFrequency = {};
  const solutionsFrequency = {};

  filteredData.forEach((row, index) => {
    // Prime Id
    let primeId = getColValue(row, [
      "Prime Id",
      "PrimeId",
      "PRIME ID",
      "Prime_Id",
    ]);
    if (primeId !== "" && primeId !== null && primeId !== undefined) {
      let cleanId = String(primeId).trim();
      primeIdFrequency[cleanId] = (primeIdFrequency[cleanId] || 0) + 1;

      if (!primeIdList[cleanId]) primeIdList[cleanId] = [];
      primeIdList[cleanId].push({ rowIndex: index, rowData: row });
    }

    // Descrição Solução
    let causa = getColValue(row, [
      "Descrição Solução",
      "Descricao Solucao",
      "Descricao da Solucao",
      "Descricao_Solucao",
    ]);
    if (causa !== "" && causa !== null && causa !== undefined) {
      let cleanCausa = String(causa).trim();
      causesFrequency[cleanCausa] = (causesFrequency[cleanCausa] || 0) + 1;
    }

    // Solução Detalhada
    let sol = getColValue(row, [
      "Solução Detalhada",
      "Solucao Detalhada",
      "Solucao_Detalhada",
    ]);
    if (sol !== "" && sol !== null && sol !== undefined) {
      let cleanSol = String(sol).trim();
      solutionsFrequency[cleanSol] = (solutionsFrequency[cleanSol] || 0) + 1;
    }
  });

  // Top 10 Prime IDs
  let topPrimeArray = [];
  for (let pid in primeIdFrequency) {
    if (primeIdFrequency[pid] > 1) {
      topPrimeArray.push({ primeId: pid, count: primeIdFrequency[pid] });
    }
  }
  topPrimeArray.sort((a, b) => b.count - a.count);
  let top10Primes = topPrimeArray.slice(0, 10);

  // Top 10 Descrição Solução
  let topCausesArray = [];
  for (let c in causesFrequency) {
    topCausesArray.push({ cause: c, count: causesFrequency[c] });
  }
  topCausesArray.sort((a, b) => b.count - a.count);
  let top10Causes = topCausesArray.slice(0, 10);

  // Top 10 Solução Detalhada
  let topSolArray = [];
  for (let s in solutionsFrequency) {
    topSolArray.push({ solution: s, count: solutionsFrequency[s] });
  }
  topSolArray.sort((a, b) => b.count - a.count);
  let top10Solutions = topSolArray.slice(0, 10);

  const operators = {};
  let totalSlaOkCount = 0;
  let totalSlaNoCount = 0;
  let totalDefinitiveCount = 0;
  let totalLoopCount = 0;

  let inopTotal = 0,
    inopOk = 0,
    inopNo = 0;
  let parcTotal = 0,
    parcOk = 0,
    parcNo = 0;

  filteredData.forEach((row, index) => {
    let operador = getColValue(row, [
      "Login Solucionador",
      "Solucionador",
      "Operador",
      "Cargo Solucionador",
    ]);
    operador = operador ? String(operador).trim() : "Não Identificado";

    if (!operators[operador]) {
      operators[operador] = {
        total: 0,
        slaOk: 0,
        slaNo: 0,
        definitive: 0,
        loop: 0,
      };
    }

    operators[operador].total++;

    let tempoVal = getColValue(row, [
      "Tempo Atendido",
      "Tempo Atendido Média Visual",
      "Tempo S/Agendamento",
    ]);
    let numTempo = parseFloat(String(tempoVal).replace(",", ".")) || 0;
    let estouroExplicit = String(
      getColValue(row, ["CAUSA TELMEX > 4H", "Estouros", "Estouro Piramide"])
    ).toLowerCase();

    let isSlaNo =
      numTempo > 4 ||
      estouroExplicit.includes("sim") ||
      estouroExplicit.includes("> 4");

    let impVal = String(
      getColValue(row, ["Impacto", "Gravidade", "Tipo Falha", "Status Serviço"])
    ).toLowerCase();
    let isInop =
      impVal.includes("inoperante") ||
      impVal.includes("total") ||
      impVal.includes("critico");

    if (isInop) {
      inopTotal++;
      if (isSlaNo) {
        inopNo++;
      } else {
        inopOk++;
      }
    } else {
      parcTotal++;
      if (isSlaNo) {
        parcNo++;
      } else {
        parcOk++;
      }
    }

    if (isSlaNo) {
      operators[operador].slaNo++;
      totalSlaNoCount++;
    } else {
      operators[operador].slaOk++;
      totalSlaOkCount++;
    }

    let primeId = getColValue(row, [
      "Prime Id",
      "PrimeId",
      "PRIME ID",
      "Prime_Id",
    ]);
    if (primeId !== "" && primeId !== null && primeId !== undefined) {
      let cleanId = String(primeId).trim();
      let ocorrencias = primeIdList[cleanId];

      if (ocorrencias && ocorrencias.length > 1) {
        let ultimaOcorrenciaIndex =
          ocorrencias[ocorrencias.length - 1].rowIndex;

        if (index === ultimaOcorrenciaIndex) {
          operators[operador].definitive++;
          totalDefinitiveCount++;
        } else {
          operators[operador].loop++;
          totalLoopCount++;
        }
      }
    }
  });

  const totalRegs = filteredData.length;
  const slaOkPercent =
    totalRegs > 0 ? ((totalSlaOkCount / totalRegs) * 100).toFixed(1) : 0;
  const defPercent =
    totalRegs > 0 ? ((totalDefinitiveCount / totalRegs) * 100).toFixed(1) : 0;
  const loopPercent =
    totalRegs > 0 ? ((totalLoopCount / totalRegs) * 100).toFixed(1) : 0;
  const noPercent =
    totalRegs > 0 ? ((totalSlaNoCount / totalRegs) * 100).toFixed(1) : 0;

  let inopNoPercent =
    inopTotal > 0 ? ((inopNo / inopTotal) * 100).toFixed(1) : 0;
  let parcNoPercent =
    parcTotal > 0 ? ((parcNo / parcTotal) * 100).toFixed(1) : 0;

  const totalTicketsEl = document.getElementById("totalTickets");
  if (totalTicketsEl) totalTicketsEl.innerText = totalRegs;

  const totalSlaOkEl = document.getElementById("totalSlaOk");
  if (totalSlaOkEl) {
    totalSlaOkEl.innerHTML = `${totalSlaOkCount} <small class="fs-6 text-success fw-normal">(${slaOkPercent}%)</small>`;
  }
  const barSlaEl = document.getElementById("barSla");
  if (barSlaEl) barSlaEl.style.width = slaOkPercent + "%";

  const totalDefinitiveEl = document.getElementById("totalDefinitive");
  if (totalDefinitiveEl) totalDefinitiveEl.innerText = totalDefinitiveCount;
  const barDefEl = document.getElementById("barDef");
  if (barDefEl) barDefEl.style.width = defPercent + "%";

  const totalLoopEl = document.getElementById("totalLoop");
  if (totalLoopEl) totalLoopEl.innerText = totalLoopCount;
  const barLoopEl = document.getElementById("barLoop");
  if (barLoopEl) barLoopEl.style.width = loopPercent + "%";

  const totalSlaNoEl = document.getElementById("totalSlaNo");
  if (totalSlaNoEl) totalSlaNoEl.innerText = totalSlaNoCount;
  const barNoEl = document.getElementById("barNo");
  if (barNoEl) barNoEl.style.width = noPercent + "%";

  const slaInopTextEl = document.getElementById("slaInoperanteText");
  if (slaInopTextEl) slaInopTextEl.innerText = `${inopTotal} chamados`;

  const slaInopSubEl = document.getElementById("slaInoperanteSub");
  if (slaInopSubEl)
    slaInopSubEl.innerText = `No prazo: ${inopOk} | Estouros: ${inopNo}`;

  const badgeInopEl = document.getElementById("badgeInop");
  if (badgeInopEl) badgeInopEl.innerText = `${inopNoPercent}% estouro`;

  const slaParcTextEl = document.getElementById("slaParcialText");
  if (slaParcTextEl) slaParcTextEl.innerText = `${parcTotal} chamados`;

  const slaParcSubEl = document.getElementById("slaParcialSub");
  if (slaParcSubEl)
    slaParcSubEl.innerText = `No prazo: ${parcOk} | Estouros: ${parcNo}`;

  const badgeParcialEl = document.getElementById("badgeParcial");
  if (badgeParcialEl) badgeParcialEl.innerText = `${parcNoPercent}% estouro`;

  let operatorRanking = [];
  for (let op in operators) {
    let d = operators[op];
    let taxaSla = d.total > 0 ? (d.slaOk / d.total) * 100 : 0;
    let score = d.definitive * 10 + taxaSla * 2 - d.loop * 15 - d.slaNo * 5;
    operatorRanking.push({
      name: op,
      score: score,
      ...d,
      taxaSla: taxaSla.toFixed(1),
    });
  }

  operatorRanking.sort((a, b) => b.score - a.score);

  // Ranking de Produtividade (Top 10 por maior volume total)
  let productivityRanking = [...operatorRanking].sort(
    (a, b) => b.total - a.total
  );
  let top10Productivity = productivityRanking.slice(0, 10);

  let topHtml = '<div class="row g-3">';
  let top3 = operatorRanking.slice(0, 3);
  if (top3.length === 0 || top3[0].total === 0) {
    topHtml =
      '<p class="text-muted m-0">Nenhum dado suficiente para ranking no período.</p>';
  } else {
    top3.forEach((op, idx) => {
      let medals = ["🥇 1º Lugar", "🥈 2º Lugar", "🥉 3º Lugar"];
      let badgeColor =
        idx === 0
          ? "bg-warning text-dark"
          : idx === 1
          ? "bg-secondary text-white"
          : "bg-dark text-white";
      topHtml += `
            <div class="col-md-4">
                <div class="podium-item h-100">
                    <span class="badge ${badgeColor} mb-2">${medals[idx]}</span>
                    <h6 class="text-dark fw-bold mb-1">${op.name}</h6>
                    <p class="mb-1 text-success small"><strong>SLA Cumprido:</strong> ${op.taxaSla}% (${op.slaOk}/${op.total})</p>
                    <p class="mb-0 text-muted small">Definitivas: <strong>${op.definitive}</strong> | Loops: <strong>${op.loop}</strong></p>
                </div>
            </div>
        `;
    });
    topHtml += "</div>";
  }
  const topPerfEl = document.getElementById("topPerformersText");
  if (topPerfEl) topPerfEl.innerHTML = topHtml;

  let topProducer =
    productivityRanking.length > 0
      ? productivityRanking[0]
      : { name: "N/A", total: 0 };
  let worstPrime =
    top10Primes.length > 0 ? top10Primes[0] : { primeId: "N/A", count: 0 };
  let topCause =
    top10Causes.length > 0 ? top10Causes[0] : { cause: "N/A", count: 0 };

  const insightsEl = document.getElementById("insightsText");
  if (insightsEl) {
    insightsEl.innerHTML = `
      <div class="row g-3">
          <div class="col-md-6">
              <div class="p-3 border rounded bg-light h-100">
                  <h6 class="fw-bold text-success mb-2"><i class="fa-solid fa-circle-check me-1"></i> Pontos Fortes da Operação</h6>
                  <p class="mb-1 small"><strong>Maior Produtividade:</strong> O operador <strong>${topProducer.name}</strong> lidera o volume de encerramentos com <strong>${topProducer.total}</strong> chamados finalizados.</p>
                  <p class="mb-1 small"><strong>Conformidade Geral de Prazo:</strong> ${slaOkPercent}% dos chamados resolvidos dentro do SLA (&le; 4h).</p>
                  <p class="mb-0 small"><strong>Resolução Raiz:</strong> ${totalDefinitiveCount} tickets reincidentes estancados permanentemente.</p>
              </div>
          </div>
          <div class="col-md-6">
              <div class="p-3 border rounded bg-light h-100">
                  <h6 class="fw-bold text-danger mb-2"><i class="fa-solid fa-triangle-exclamation me-1"></i> Alertas e Gargalos Críticos</h6>
                  <p class="mb-1 small"><strong>Impacto Crítico:</strong> O volume de falhas <em>Inoperantes</em> apresenta <strong>${inopNoPercent}%</strong> de estouro de SLA.</p>
                  <p class="mb-1 small"><strong>Caso Crônico de Rede:</strong> O Prime ID <strong>${worstPrime.primeId}</strong> acumula ${worstPrime.count} retornos.</p>
                  <p class="mb-0 small"><strong>Principal Falha:</strong> "${topCause.cause}" lidera as ocorrências na base.</p>
              </div>
          </div>
      </div>
    `;
  }

  const actionPlanEl = document.getElementById("actionPlanText");
  if (actionPlanEl) {
    actionPlanEl.innerHTML = `
      <div class="row g-3">
          <div class="col-md-4">
              <div class="p-3 border rounded bg-white h-100 border-start border-4 border-danger">
                  <h6 class="fw-bold text-danger mb-2"><i class="fa-solid fa-bullseye me-1"></i> Foco em Inoperantes</h6>
                  <p class="small text-muted mb-0">Como chamados inoperantes impactam totalmente o cliente, priorize a fila de atendimento imediato para este impacto e reduza os ${inopNo} estouros atuais.</p>
              </div>
          </div>
          <div class="col-md-4">
              <div class="p-3 border rounded bg-white h-100 border-start border-4 border-warning">
                  <h6 class="fw-bold text-warning mb-2" style="color: var(--claro-orange) !important;"><i class="fa-solid fa-arrows-rotate me-1"></i> Tratativa de Prime IDs</h6>
                  <p class="small text-muted mb-0">Encaminhar os Top 10 Prime IDs listados abaixo para a equipe de manutenção corretiva externa (campo/infraestrutura) para eliminar a raiz do problema.</p>
              </div>
          </div>
          <div class="col-md-4">
              <div class="p-3 border rounded bg-white h-100 border-start border-4 border-success">
                  <h6 class="fw-bold text-success mb-2"><i class="fa-solid fa-diagram-project me-1"></i> Padronização de Soluções</h6>
                  <p class="small text-muted mb-0">Disseminar o conhecimento das <strong>Soluções Detalhadas</strong> mais eficazes utilizadas pelos operadores do Grupo 1 (Eficazes) para reduzir os loops de reincidência.</p>
              </div>
          </div>
      </div>
    `;
  }

  // Tabela Top 10 Produtividade
  const tableTopProductivityBody = document.querySelector(
    "#tableTopProductivity tbody"
  );
  if (tableTopProductivityBody) {
    tableTopProductivityBody.innerHTML = "";
    if (top10Productivity.length === 0) {
      tableTopProductivityBody.innerHTML = `<tr><td colspan="5" class="text-center text-muted py-3">Nenhum operador encontrado.</td></tr>`;
    } else {
      top10Productivity.forEach((op, idx) => {
        let medals = ["🥇", "🥈", "🥉"];
        let rankDisplay =
          idx < 3 ? `${medals[idx]} ${idx + 1}º` : `${idx + 1}º`;
        tableTopProductivityBody.innerHTML += `
            <tr>
                <td class="fw-bold text-secondary">${rankDisplay}</td>
                <td class="fw-bold">${op.name}</td>
                <td><span class="badge bg-danger">${op.total} chamados</span></td>
                <td class="text-success fw-bold">${op.slaOk}</td>
                <td><span class="badge bg-success">${op.taxaSla}%</span></td>
            </tr>
        `;
      });
    }
  }

  // Tabela Top 10 Prime IDs
  const tableTopPrimeBody = document.querySelector("#tableTopPrime tbody");
  if (tableTopPrimeBody) {
    tableTopPrimeBody.innerHTML = "";
    if (top10Primes.length === 0) {
      tableTopPrimeBody.innerHTML = `<tr><td colspan="4" class="text-center text-muted py-3">Nenhum Prime ID reincidente encontrado no período.</td></tr>`;
    } else {
      top10Primes.forEach((item, idx) => {
        let badgeSeverity =
          item.count >= 5
            ? '<span class="badge bg-danger">Crítico (>=5)</span>'
            : '<span class="badge bg-warning text-dark">Alerta (2-4)</span>';
        tableTopPrimeBody.innerHTML += `
            <tr>
                <td class="fw-bold text-secondary">#${idx + 1}</td>
                <td class="fw-bold font-monospace">${item.primeId}</td>
                <td><span class="badge bg-dark">${
                  item.count
                } ocorrências</span></td>
                <td>${badgeSeverity}</td>
            </tr>
        `;
      });
    }
  }

  // Tabela Top 10 Causas
  const tableTopCausesBody = document.querySelector("#tableTopCauses tbody");
  if (tableTopCausesBody) {
    tableTopCausesBody.innerHTML = "";
    if (top10Causes.length === 0) {
      tableTopCausesBody.innerHTML = `<tr><td colspan="3" class="text-center text-muted py-3">Coluna "Descrição Solução" não encontrada na planilha.</td></tr>`;
    } else {
      top10Causes.forEach((item, idx) => {
        tableTopCausesBody.innerHTML += `
            <tr>
                <td class="fw-bold text-secondary">#${idx + 1}</td>
                <td>${item.cause}</td>
                <td><span class="badge bg-danger">${item.count}</span></td>
            </tr>
        `;
      });
    }
  }

  // Tabela Top 10 Soluções
  const tableTopSolutionsBody = document.querySelector(
    "#tableTopSolutions tbody"
  );
  if (tableTopSolutionsBody) {
    tableTopSolutionsBody.innerHTML = "";
    if (top10Solutions.length === 0) {
      tableTopSolutionsBody.innerHTML = `<tr><td colspan="3" class="text-center text-muted py-3">Coluna "Solução Detalhada" não encontrada na planilha.</td></tr>`;
    } else {
      top10Solutions.forEach((item, idx) => {
        tableTopSolutionsBody.innerHTML += `
            <tr>
                <td class="fw-bold text-secondary">#${idx + 1}</td>
                <td>${item.solution}</td>
                <td><span class="badge bg-success">${item.count}</span></td>
            </tr>
        `;
      });
    }
  }

  const tableDefinitiveBody = document.querySelector("#tableDefinitive tbody");
  const tableLoopBody = document.querySelector("#tableLoop tbody");

  if (tableDefinitiveBody) tableDefinitiveBody.innerHTML = "";
  if (tableLoopBody) tableLoopBody.innerHTML = "";

  let chartSortedOps = [...operatorRanking].sort((a, b) => b.total - a.total);
  const chartLabels = chartSortedOps.map((d) => d.name);
  const chartTotals = chartSortedOps.map((d) => d.total);

  operatorRanking.forEach((d) => {
    if (d.definitive > 0 && tableDefinitiveBody) {
      tableDefinitiveBody.innerHTML += `
        <tr>
            <td class="fw-bold">${d.name}</td>
            <td>${d.total}</td>
            <td class="text-success fw-bold">${d.slaOk}</td>
            <td class="text-danger fw-bold">${d.slaNo}</td>
            <td><span class="badge bg-success px-3 py-2">${d.definitive} resolução(ões) definitiva(s)</span></td>
        </tr>`;
    }

    if (d.loop > 0 && tableLoopBody) {
      tableLoopBody.innerHTML += `
        <tr>
            <td class="fw-bold">${d.name}</td>
            <td>${d.total}</td>
            <td class="text-success fw-bold">${d.slaOk}</td>
            <td class="text-danger fw-bold">${d.slaNo}</td>
            <td><span class="badge bg-warning text-dark px-3 py-2" style="background-color: var(--claro-orange) !important; color: white !important;">${d.loop} ocorrência(s) em loop</span></td>
        </tr>`;
    }
  });

  if (tableDefinitiveBody && tableDefinitiveBody.innerHTML === "") {
    tableDefinitiveBody.innerHTML = `<tr><td colspan="5" class="text-center text-muted py-3">Nenhum registro encontrado para o filtro.</td></tr>`;
  }
  if (tableLoopBody && tableLoopBody.innerHTML === "") {
    tableLoopBody.innerHTML = `<tr><td colspan="5" class="text-center text-muted py-3">Nenhum registro encontrado para o filtro.</td></tr>`;
  }

  renderCharts(
    chartLabels,
    chartTotals,
    totalSlaOkCount,
    totalLoopCount,
    totalSlaNoCount
  );
}

// Renderização dos gráficos interativos via Chart.js
function renderCharts(labels, totals, slaOkCount, loopCount, slaNoCount) {
  if (opChartInstance) opChartInstance.destroy();
  if (doughnutChartInstance) doughnutChartInstance.destroy();

  const canvasEl = document.getElementById("operatorChart");
  if (!canvasEl) return;

  const dynamicHeight = Math.max(300, labels.length * 30);
  canvasEl.parentNode.style.height = dynamicHeight + "px";

  const ctx1 = canvasEl.getContext("2d");
  opChartInstance = new Chart(ctx1, {
    type: "bar",
    data: {
      labels: labels,
      datasets: [
        {
          label: "Total de Atendimentos",
          data: totals,
          backgroundColor: "#DA291C",
          borderRadius: 6,
          barThickness: 18,
        },
      ],
    },
    options: {
      indexAxis: "y",
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: {
          callbacks: {
            label: function (context) {
              return ` Atendimentos: ${context.raw} chamados`;
            },
          },
        },
      },
      scales: {
        x: { grid: { color: "#f1f3f5" }, ticks: { font: { size: 11 } } },
        y: {
          grid: { display: false },
          ticks: { font: { size: 12, family: "Segoe UI" } },
        },
      },
    },
  });

  const doughnutCanvas = document.getElementById("slaDoughnutChart");
  if (!doughnutCanvas) return;

  const ctx2 = doughnutCanvas.getContext("2d");
  doughnutChartInstance = new Chart(ctx2, {
    type: "doughnut",
    data: {
      labels: [
        "SLA Cumprido (<= 4h)",
        "Em Loop de Reincidência",
        "Estouro de SLA (> 4h)",
      ],
      datasets: [
        {
          data: [slaOkCount, loopCount, slaNoCount],
          backgroundColor: ["#0097A9", "#F5842B", "#DA291C"],
          borderWidth: 3,
          borderColor: "#ffffff",
        },
      ],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: {
          position: "bottom",
          labels: {
            boxWidth: 14,
            font: { size: 11, family: "Segoe UI" },
            padding: 15,
          },
        },
      },
    },
  });
}

// ==========================================================================
// Sistema Unificado de Modais e Toasts (Substitui os alerts nativos)
// ==========================================================================

// Exibe um Modal customizado no tema ativo
function showAppModal(title, message, type = "info", onConfirm = null) {
  const modalEl = document.getElementById("appModal");
  if (!modalEl) return alert(message); // Fallback de emergência

  const modalTitle = document.getElementById("appModalLabel");
  const modalBody = document.getElementById("appModalBody");
  const modalIcon = document.getElementById("modalIcon");
  const modalIconBox = document.getElementById("modalIconBox");
  const confirmBtn = document.getElementById("modalConfirmBtn");

  if (modalTitle) modalTitle.innerText = title;
  if (modalBody) modalBody.innerHTML = message;

  // Ajusta ícones e cores de acordo com o tipo de mensagem
  if (modalIcon && modalIconBox) {
    if (type === "danger" || type === "error") {
      modalIcon.className = "fa-solid fa-triangle-exclamation text-danger fs-5";
      modalIconBox.className = "icon-shape bg-danger-subtle rounded-3 p-2";
    } else if (type === "warning") {
      modalIcon.className = "fa-solid fa-circle-exclamation text-warning fs-5";
      modalIconBox.className = "icon-shape bg-warning-subtle rounded-3 p-2";
    } else if (type === "success") {
      modalIcon.className = "fa-solid fa-circle-check text-success fs-5";
      modalIconBox.className = "icon-shape bg-success-subtle rounded-3 p-2";
    } else {
      modalIcon.className = "fa-solid fa-circle-info text-primary fs-5";
      modalIconBox.className = "icon-shape bg-primary-subtle rounded-3 p-2";
    }
  }

  // Gerencia o botão de confirmação (para ações destrutivas ou de prosseguir)
  if (confirmBtn) {
    if (typeof onConfirm === "function") {
      confirmBtn.classList.remove("d-none");
      confirmBtn.onclick = () => {
        onConfirm();
        const bsModal = bootstrap.Modal.getInstance(modalEl);
        if (bsModal) bsModal.hide();
      };
    } else {
      confirmBtn.classList.add("d-none");
    }
  }

  const bsModal = new bootstrap.Modal(modalEl);
  bsModal.show();
}

// Exibe um Toast rápido e discreto no canto inferior direito
function showToast(message, type = "success") {
  const toastEl = document.getElementById("appToast");
  if (!toastEl) return;

  const toastMessage = document.getElementById("toastMessage");
  const toastIcon = document.getElementById("toastIcon");

  if (toastMessage) toastMessage.innerText = message;

  if (toastIcon) {
    if (type === "error" || type === "danger") {
      toastIcon.className = "fa-solid fa-circle-xmark text-danger fs-5";
    } else if (type === "warning") {
      toastIcon.className =
        "fa-solid fa-triangle-exclamation text-warning fs-5";
    } else if (type === "info") {
      toastIcon.className = "fa-solid fa-circle-info text-primary fs-5";
    } else {
      toastIcon.className = "fa-solid fa-circle-check text-success fs-5";
    }
  }

  const toast = new bootstrap.Toast(toastEl, { delay: 4000 });
  toast.show();
}
