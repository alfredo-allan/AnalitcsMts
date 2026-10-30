let globalRawData = [];
let opChartInstance = null;
let doughnutChartInstance = null;

document.getElementById("loadBtn").addEventListener("click", function () {
  const fileInput = document.getElementById("excelFile");
  const alertBox = document.getElementById("alertBox");
  alertBox.classList.add("d-none");
  alertBox.innerText = "";

  if (fileInput.files.length === 0) {
    alert("Por favor, selecione um arquivo Excel (.xlsx).");
    return;
  }

  const file = fileInput.files[0];
  const reader = new FileReader();

  reader.onload = function (e) {
    try {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: "array" });
      const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
      globalRawData = XLSX.utils.sheet_to_json(firstSheet);

      if (!globalRawData || globalRawData.length === 0) {
        showAlert("A planilha está vazia ou os dados não puderam ser lidos.");
        return;
      }

      populateMonthFilter(globalRawData);
      processData(globalRawData);

      document.getElementById("dashboardContent").style.display = "block";
    } catch (error) {
      showAlert("Erro ao processar a planilha: " + error.message);
    }
  };
  reader.readAsArrayBuffer(file);
});

document.getElementById("monthFilter").addEventListener("change", function () {
  processData(globalRawData);
});

function showAlert(msg) {
  const box = document.getElementById("alertBox");
  box.innerText = msg;
  box.classList.remove("d-none");
}

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

function populateMonthFilter(data) {
  const monthSelect = document.getElementById("monthFilter");
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

function processData(data) {
  const selectedMonth = document.getElementById("monthFilter").value;

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

  document.getElementById("totalTickets").innerText = totalRegs;
  document.getElementById(
    "totalSlaOk"
  ).innerHTML = `${totalSlaOkCount} <small class="fs-6 text-success fw-normal">(${slaOkPercent}%)</small>`;
  document.getElementById("barSla").style.width = slaOkPercent + "%";

  document.getElementById("totalDefinitive").innerText = totalDefinitiveCount;
  document.getElementById("barDef").style.width = defPercent + "%";

  document.getElementById("totalLoop").innerText = totalLoopCount;
  document.getElementById("barLoop").style.width = loopPercent + "%";

  document.getElementById("totalSlaNo").innerText = totalSlaNoCount;
  document.getElementById("barNo").style.width = noPercent + "%";

  document.getElementById(
    "slaInoperanteText"
  ).innerText = `${inopTotal} chamados`;
  document.getElementById(
    "slaInoperanteSub"
  ).innerText = `No prazo: ${inopOk} | Estouros: ${inopNo}`;
  document.getElementById("badgeInop").innerText = `${inopNoPercent}% estouro`;

  document.getElementById("slaParcialText").innerText = `${parcTotal} chamados`;
  document.getElementById(
    "slaParcialSub"
  ).innerText = `No prazo: ${parcOk} | Estouros: ${parcNo}`;
  document.getElementById(
    "badgeParcial"
  ).innerText = `${parcNoPercent}% estouro`;

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
  document.getElementById("topPerformersText").innerHTML = topHtml;

  let bestSlaOp = operatorRanking.reduce(
    (prev, curr) =>
      parseFloat(curr.taxaSla) > parseFloat(prev.taxaSla)
        ? curr
        : operatorRanking[0],
    operatorRanking[0] || { name: "N/A", taxaSla: 0 }
  );
  let topProducer =
    productivityRanking.length > 0
      ? productivityRanking[0]
      : { name: "N/A", total: 0 };
  let worstPrime =
    top10Primes.length > 0 ? top10Primes[0] : { primeId: "N/A", count: 0 };
  let topCause =
    top10Causes.length > 0 ? top10Causes[0] : { cause: "N/A", count: 0 };

  document.getElementById("insightsText").innerHTML = `
                <div class="row g-3">
                    <div class="col-md-6">
                        <div class="p-3 border rounded bg-light h-100">
                            <h6 class="fw-bold text-success mb-2"><i class="bi bi-check-circle-fill me-1"></i> Pontos Fortes da Operação</h6>
                            <p class="mb-1 small"><strong>Maior Produtividade:</strong> O operador <strong>${topProducer.name}</strong> lidera o volume de encerramentos com <strong>${topProducer.total}</strong> chamados finalizados.</p>
                            <p class="mb-1 small"><strong>Conformidade Geral de Prazo:</strong> ${slaOkPercent}% dos chamados resolvidos dentro do SLA (&le; 4h).</p>
                            <p class="mb-0 small"><strong>Resolução Raiz:</strong> ${totalDefinitiveCount} tickets reincidentes estancados permanentemente.</p>
                        </div>
                    </div>
                    <div class="col-md-6">
                        <div class="p-3 border rounded bg-light h-100">
                            <h6 class="fw-bold text-danger mb-2"><i class="bi bi-exclamation-octagon-fill me-1"></i> Alertas e Gargalos Críticos</h6>
                            <p class="mb-1 small"><strong>Impacto Crítico:</strong> O volume de falhas <em>Inoperantes</em> apresenta <strong>${inopNoPercent}%</strong> de estouro de SLA.</p>
                            <p class="mb-1 small"><strong>Caso Crônico de Rede:</strong> O Prime ID <strong>${worstPrime.primeId}</strong> acumula ${worstPrime.count} retornos.</p>
                            <p class="mb-0 small"><strong>Principal Falha:</strong> "${topCause.cause}" lidera as ocorrências na base.</p>
                        </div>
                    </div>
                </div>
            `;

  document.getElementById("actionPlanText").innerHTML = `
                <div class="row g-3">
                    <div class="col-md-4">
                        <div class="p-3 border rounded bg-white h-100 border-start border-4 border-danger">
                            <h6 class="fw-bold text-danger mb-2"><i class="bi bi-1-circle-fill me-1"></i> Foco em Inoperantes</h6>
                            <p class="small text-muted mb-0">Como chamados inoperantes impactam totalmente o cliente, priorize a fila de atendimento imediato para este impacto e reduza os ${inopNo} estouros atuais.</p>
                        </div>
                    </div>
                    <div class="col-md-4">
                        <div class="p-3 border rounded bg-white h-100 border-start border-4 border-warning">
                            <h6 class="fw-bold text-warning mb-2" style="color: var(--claro-orange) !important;"><i class="bi bi-2-circle-fill me-1"></i> Tratativa de Prime IDs</h6>
                            <p class="small text-muted mb-0">Encaminhar os Top 10 Prime IDs listados abaixo para a equipe de manutenção corretiva externa (campo/infraestrutura) para eliminar a raiz do problema.</p>
                        </div>
                    </div>
                    <div class="col-md-4">
                        <div class="p-3 border rounded bg-white h-100 border-start border-4 border-success">
                            <h6 class="fw-bold text-success mb-2"><i class="bi bi-3-circle-fill me-1"></i> Padronização de Soluções</h6>
                            <p class="small text-muted mb-0">Disseminar o conhecimento das <strong>Soluções Detalhadas</strong> mais eficazes utilizadas pelos operadores do Grupo 1 (Eficazes) para reduzir os loops de reincidência.</p>
                        </div>
                    </div>
                </div>
            `;

  // Tabela Top 10 Produtividade
  const tableTopProductivityBody = document.querySelector(
    "#tableTopProductivity tbody"
  );
  tableTopProductivityBody.innerHTML = "";
  if (top10Productivity.length === 0) {
    tableTopProductivityBody.innerHTML = `<tr><td colspan="5" class="text-center text-muted py-3">Nenhum operador encontrado.</td></tr>`;
  } else {
    top10Productivity.forEach((op, idx) => {
      let medals = ["🥇", "🥈", "🥉"];
      let rankDisplay = idx < 3 ? `${medals[idx]} ${idx + 1}º` : `${idx + 1}º`;
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

  // Tabela Top 10 Prime IDs
  const tableTopPrimeBody = document.querySelector("#tableTopPrime tbody");
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
                            <td class="fw-bold font-monospace">${
                              item.primeId
                            }</td>
                            <td><span class="badge bg-dark">${
                              item.count
                            } ocorrências</span></td>
                            <td>${badgeSeverity}</td>
                        </tr>
                    `;
    });
  }

  // Tabela Top 10 Causas
  const tableTopCausesBody = document.querySelector("#tableTopCauses tbody");
  tableTopCausesBody.innerHTML = "";
  if (top10Causes.length === 0) {
    tableTopCausesBody.innerHTML = `<tr><td colspan="3" class="text-center text-muted py-3">Coluna "Descrição Solução" não encontrada na planilha.</td></tr>`;
  } else {
    top10Causes.forEach((item, idx) => {
      tableTopCausesBody.innerHTML += `
                        <tr>
                            <td class="fw-bold text-secondary">#${idx + 1}</td>
                            <td>${item.cause}</td>
                            <td><span class="badge bg-danger">${
                              item.count
                            }</span></td>
                        </tr>
                    `;
    });
  }

  // Tabela Top 10 Soluções
  const tableTopSolutionsBody = document.querySelector(
    "#tableTopSolutions tbody"
  );
  tableTopSolutionsBody.innerHTML = "";
  if (top10Solutions.length === 0) {
    tableTopSolutionsBody.innerHTML = `<tr><td colspan="3" class="text-center text-muted py-3">Coluna "Solução Detalhada" não encontrada na planilha.</td></tr>`;
  } else {
    top10Solutions.forEach((item, idx) => {
      tableTopSolutionsBody.innerHTML += `
                        <tr>
                            <td class="fw-bold text-secondary">#${idx + 1}</td>
                            <td>${item.solution}</td>
                            <td><span class="badge bg-success">${
                              item.count
                            }</span></td>
                        </tr>
                    `;
    });
  }

  const tableDefinitiveBody = document.querySelector("#tableDefinitive tbody");
  const tableLoopBody = document.querySelector("#tableLoop tbody");
  tableDefinitiveBody.innerHTML = "";
  tableLoopBody.innerHTML = "";

  let chartSortedOps = [...operatorRanking].sort((a, b) => b.total - a.total);
  const chartLabels = chartSortedOps.map((d) => d.name);
  const chartTotals = chartSortedOps.map((d) => d.total);

  operatorRanking.forEach((d) => {
    if (d.definitive > 0) {
      tableDefinitiveBody.innerHTML += `
                    <tr>
                        <td class="fw-bold">${d.name}</td>
                        <td>${d.total}</td>
                        <td class="text-success fw-bold">${d.slaOk}</td>
                        <td class="text-danger fw-bold">${d.slaNo}</td>
                        <td><span class="badge bg-success px-3 py-2">${d.definitive} resolução(ões) definitiva(s)</span></td>
                    </tr>`;
    }

    if (d.loop > 0) {
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

  if (tableDefinitiveBody.innerHTML === "") {
    tableDefinitiveBody.innerHTML = `<tr><td colspan="5" class="text-center text-muted py-3">Nenhum registro encontrado para o filtro.</td></tr>`;
  }
  if (tableLoopBody.innerHTML === "") {
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

function renderCharts(labels, totals, slaOkCount, loopCount, slaNoCount) {
  if (opChartInstance) opChartInstance.destroy();
  if (doughnutChartInstance) doughnutChartInstance.destroy();

  const canvasEl = document.getElementById("operatorChart");
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

  const ctx2 = document.getElementById("slaDoughnutChart").getContext("2d");
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

themeToggleBtn.addEventListener("click", () => {
  const currentTheme = document.documentElement.getAttribute("data-theme");
  const newTheme = currentTheme === "dark" ? "light" : "dark";

  document.documentElement.setAttribute("data-theme", newTheme);

  if (newTheme === "dark") {
    themeLabel.innerText = "Modo Claro";
    themeIcon.className = "fa-solid fa-moon";
  } else {
    themeLabel.innerText = "Modo Escuro";
    themeIcon.className = "fa-solid fa-sun";
  }
});

document.getElementById("excelFile").addEventListener("change", function (e) {
  const fileName = e.target.files[0]
    ? e.target.files[0].name
    : "Nenhum arquivo selecionado...";
  document.getElementById(
    "fileNameDisplay"
  ).innerHTML = `<i class="fa-solid fa-file-excel me-2 text-success"></i>${fileName}`;
});
