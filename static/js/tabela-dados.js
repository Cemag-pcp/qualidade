// Variáveis para controle de edição
let currentId = null;
let isEditing = false;
let tableData = [];      // fonte única

function buscarItens() {
    return fetch('/api/items', {
      method: 'GET',
      headers: { 'Content-Type': 'application/json' }
    })
    .then(response => {
      if (!response.ok) {
        throw new Error('Erro na requisição: ' + response.status);
      }
      return response.json();
    });
}

function loadTableData(data = tableData) {   // ← default
    const tbody = document.getElementById('dataTableBody');
    tbody.innerHTML = '';
  
    if (!Array.isArray(data)) return;          // segurança extra
    
    data.forEach(item => {
      const row = document.createElement('tr');
      row.innerHTML = `
        <td>${item.carreta}</td>
        <td>${item.codigo}</td>
        <td>${item.desc}</td>
        <td>${item.item_code}</td>
        <td>${item.qt}</td>
        <td>${item.item_description}</td>
        <td>${item.tipo}</td>
        <td class="text-center">
          <button class="btn btn-sm btn-primary me-1" onclick="editRecord(${item.id})">
            <i class="bi bi-pencil"></i>
          </button>
          <button class="btn btn-sm btn-danger"  onclick="deleteRecord(${item.id})">
            <i class="bi bi-trash"></i>
          </button>
        </td>`;
      tbody.appendChild(row);
    });
}

// Função para preparar o modal para adicionar um novo registro
function prepareModalForAdd() {
    document.getElementById('modalLabel').textContent = 'Adicionar Novo Registro';
    document.getElementById('recordForm').reset();
    currentId = null;
    isEditing = false;
}

// Função para editar um registro
function editRecord(id) {
    const record = tableData.find(item => item.id === id);
    if (record) {
        document.getElementById('modalLabel').textContent = 'Editar Registro';
        document.getElementById('recordId').value = record.id;
        document.getElementById('carreta').value = record.carreta;
        document.getElementById('codigo').value = record.codigo;
        document.getElementById('desc').value = record.desc;
        document.getElementById('itemCode').value = record.item_code;
        document.getElementById('qt').value = record.qt;
        document.getElementById('itemDescription').value = record.item_description;
        document.getElementById('tipo').value = record.tipo;
        
        currentId = id;
        isEditing = true;
        
        // Abrir o modal
        const modal = new bootstrap.Modal(document.getElementById('addEditModal'));
        modal.show();
    }
}

// Função para excluir um registro
function deleteRecord(id) {
    currentId = id;
    const modal = new bootstrap.Modal(document.getElementById('deleteModal'));
    modal.show();
}

function setDeleteLoading(isLoading) {
    const confirmButton = document.getElementById("confirmDeleteButton")
    const cancelButton = document.getElementById("cancelDeleteButton")
    const loadingSpinner = document.getElementById("deleteButtonLoading")
    const buttonText = document.getElementById("deleteButtonText")
  
    if (isLoading) {
      // Ativar estado de loading
      confirmButton.disabled = true
      cancelButton.disabled = true
      loadingSpinner.style.display = "inline-block"
      buttonText.textContent = "Excluindo"
    } else {
      // Desativar estado de loading
      confirmButton.disabled = false
      cancelButton.disabled = false
      loadingSpinner.style.display = "none"
      buttonText.textContent = "Excluir"
    }
}

function setAddLoading(isLoading) {
    const confirmButton = document.getElementById("saveButton")
    const cancelButton = document.getElementById("cancelSaveButton")
    const loadingSpinner = document.getElementById("saveButtonLoading")
    const buttonText = document.getElementById("saveButtonText")
  
    if (isLoading) {
      // Ativar estado de loading
      confirmButton.disabled = true
      cancelButton.disabled = true
      loadingSpinner.style.display = "inline-block"
      buttonText.textContent = "Salvando"
    } else {
      // Desativar estado de loading
      confirmButton.disabled = false
      cancelButton.disabled = false
      loadingSpinner.style.display = "none"
      buttonText.textContent = "Salvar"
    }
}

function setGenerateReport(isLoading) {
    const confirmButton = document.getElementById("downloadReportButton")
    const cancelButton = document.getElementById("cancelDownloadReportButton")
    const loadingSpinner = document.getElementById("saveDownloadReportButton")
    const buttonText = document.getElementById("saveDownloadText")
  
    if (isLoading) {
      // Ativar estado de loading
      confirmButton.disabled = true
      cancelButton.disabled = true
      loadingSpinner.style.display = "inline-block"
      buttonText.textContent = "Gerando"
    } else {
      // Desativar estado de loading
      confirmButton.disabled = false
      cancelButton.disabled = false
      loadingSpinner.style.display = "none"
      buttonText.textContent = "Gerar"
    }
}

async function confirmDelete() {
    try {
      // Ativar estado de loading
      setDeleteLoading(true)
  
      const response = await fetch(`/api/item/${currentId}`, {
        method: "DELETE",
    })
  
    if (!response.ok) {
        const erro = await response.json()
        alert(erro.error || "Erro ao excluir o item.")
        // Desativar estado de loading em caso de erro
        setDeleteLoading(false)
        return
    }
  
        // Remove da lista e atualiza a tabela
        tableData = tableData.filter((item) => item.id !== currentId)
        loadTableData()

        // Desativar estado de loading antes de fechar o modal
        setDeleteLoading(false)

        // Fecha o modal
        const modal = bootstrap.Modal.getInstance(document.getElementById("deleteModal"))
        modal.hide()
    } catch (err) {
        console.error("Erro na exclusão:", err)
        alert("Erro ao excluir item.")
        // Desativar estado de loading em caso de erro
        setDeleteLoading(false)
    }
}

// Função para salvar um registro (novo ou editado)
async function saveRecord() {
    const carreta = document.getElementById('carreta').value;
    const codigo = document.getElementById('codigo').value;
    const desc = document.getElementById('desc').value;
    const itemCode = document.getElementById('itemCode').value;
    const qt = parseInt(document.getElementById('qt').value);
    const itemDescription = document.getElementById('itemDescription').value;
    const tipo = document.getElementById('tipo').value;

    const payload = {
        carreta,
        codigo,
        desc,
        itemCode,
        qt,
        itemDescription,
        tipo
    };

    let url = '/api/item';
    let method = 'POST';

    if (isEditing) {
        url = `/api/item/${currentId}`;
        method = 'PUT';
    }

    try {
        setAddLoading(true);

        const response = await fetch(url, {
            method: method,
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(payload)
        });

        if (!response.ok) {
            const erro = await response.json();
            alert(erro.error || 'Erro ao salvar');
            setAddLoading(false);
            return;
        }

        const result = await response.json();

        if (isEditing) {
            const index = tableData.findIndex(item => item.id === currentId);
            if (index !== -1) tableData[index] = result;
        } else {
            tableData.push(result);
        }

        setAddLoading(false);
        loadTableData();
        bootstrap.Modal.getInstance(document.getElementById('addEditModal')).hide();
    } catch (err) {
        setAddLoading(false);
        console.error('Erro:', err);
        alert('Erro ao salvar item');
    }
}

// Função para gerar relatório
function generateReport() {
    const reportDate = document.getElementById('reportDate').value;
    if (reportDate) {
        document.getElementById('reportDateDisplay').textContent = reportDate;
        const modal = new bootstrap.Modal(document.getElementById('reportModal'));
        modal.show();
    } else {
        alert('Por favor, selecione uma data para gerar o relatório.');
    }
}

// Função para pesquisar na tabela
function searchTable() {
    const searchTerm = document.getElementById('searchInput').value.toLowerCase();
    if (searchTerm.trim() === '') {
        loadTableData();
        return;
    }
    
    const filteredData = tableData.filter(item => 
        item.carreta.toLowerCase().includes(searchTerm) ||
        item.codigo.toLowerCase().includes(searchTerm) ||
        item.desc.toLowerCase().includes(searchTerm) ||
        item.itemCode.toLowerCase().includes(searchTerm) ||
        item.itemDescription.toLowerCase().includes(searchTerm) ||
        item.tipo.toLowerCase().includes(searchTerm)
    );
    
    const tableBody = document.getElementById('dataTableBody');
    tableBody.innerHTML = '';
    
    filteredData.forEach(item => {
        const row = document.createElement('tr');
        row.innerHTML = `
            <td>${item.carreta}</td>
            <td>${item.codigo}</td>
            <td>${item.desc}</td>
            <td>${item.itemCode}</td>
            <td>${item.qt}</td>
            <td>${item.itemDescription}</td>
            <td>${item.tipo}</td>
            <td class="text-center">
                <button class="btn btn-sm btn-primary me-1" onclick="editRecord(${item.id})">
                    <i class="bi bi-pencil"></i>
                </button>
                <button class="btn btn-sm btn-danger" onclick="deleteRecord(${item.id})">
                    <i class="bi bi-trash"></i>
                </button>
            </td>
        `;
        tableBody.appendChild(row);
    });
}

function buscarItens(filtros = {}) {
    // monta ?carreta=...&codigo=...
    const params = new URLSearchParams();
    Object.entries(filtros).forEach(([k, v]) => {
      if (v && v.trim() !== '') params.append(k, v.trim());
    });
  
    const url = params.toString() ? `/api/items?${params}` : '/api/items';
  
    return fetch(url)
      .then(r => r.ok ? r.json() : Promise.reject(r.status));
}
  
function carregarComFiltros() {
    const filtros = {
        carreta: document.getElementById('filtroCarreta').value,
        item_code : document.getElementById('filtroCodigo').value,
        tipo   : document.getElementById('filtroTipo').value
    };

    buscarItens(filtros)
        .then(data => { tableData = data; loadTableData(tableData); })
        .catch(console.error);
}

function gerarModelo(dataCarga) {
    setGenerateReport(true);

    fetch(`/api/carretas?dataCarga=${encodeURIComponent(dataCarga)}`)
    .then(r => r.json())
    .then(json => {
        const carretas = json.carretas || [];
        carretas.forEach(carreta => {
            const url = new URL("/modelo", window.location.origin);
            url.searchParams.set("dataCarga", dataCarga);
            url.searchParams.set("recurso", carreta[0]);
            url.searchParams.set("serie", carreta[1]);
            
            window.open(url.toString(), "_blank");
        });

        // Gerar e baixar PDFs
        gerarPDFs(dataCarga, carretas);

        const modal = bootstrap.Modal.getInstance(document.getElementById('reportModal'));
        modal.hide();
    })
    .catch(err => {
        console.error("Erro ao gerar modelo:", err);
    })
    .finally(() => {
        setGenerateReport(false);
    });
}

async function gerarPDFs(dataCarga, carretas) {
    try {
        // Opção 1: Gerar PDFs individuais
        // for (const carreta of carretas) {
        //     await baixarPDFIndividual(dataCarga, carreta[0], carreta[1]);
        // }

        // Opção 2: Gerar ZIP com todos os PDFs (descomente se preferir)
        await baixarZipPDFs(dataCarga, carretas);
        
    } catch (error) {
        console.error("Erro ao gerar PDFs:", error);
        alert("Erro ao gerar os PDFs. Tente novamente.");
    }
}

async function baixarZipPDFs(dataCarga, carretas) {
    try {
        const response = await fetch('/api/gerar-pdfs-zip', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
            },
            body: JSON.stringify({
                dataCarga: dataCarga,
                carretas: carretas
            })
        });

        if (!response.ok) {
            throw new Error(`Erro HTTP: ${response.status}`);
        }

        const blob = await response.blob();
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `checklists_${dataCarga}.zip`;
        document.body.appendChild(a);
        a.click();
        window.URL.revokeObjectURL(url);
        document.body.removeChild(a);
        
    } catch (error) {
        console.error("Erro ao baixar ZIP dos PDFs:", error);
    }
}

document.getElementById("resetCacheButton").addEventListener("click", function () {
    const dateInput = document.getElementById("reportDate");
    const data = dateInput.value;

    if (!data) {
        alert("Por favor, selecione uma data antes de resetar o cache.");
        return;
    }

    const url = `/planilha-cargas?data=${encodeURIComponent(data)}&reset=true`;

    fetch(url)
        .then(response => {
            if (!response.ok) {
                throw new Error("Erro ao resetar o cache.");
            }
            return response.json();
        })
        .then(data => {
            alert("Cache resetado com sucesso!");
            console.log("Resposta da API:", data);
        })
        .catch(error => {
            console.error("Erro:", error);
            alert("Falha ao resetar o cache.");
        });
});

document.addEventListener('DOMContentLoaded', () => {
    
    buscarItens()
    .then(dados => {
        tableData = dados; 
        loadTableData(tableData);    
    })
    .catch(err => console.error(err));

    // Botão de salvar no modal
    document.getElementById('saveButton').addEventListener('click', saveRecord);
        
    // Botão de confirmar exclusão
    document.getElementById('confirmDeleteButton').addEventListener('click', confirmDelete);
    
    // Botão de gerar relatório
    document.getElementById('generateButton').addEventListener('click', generateReport);
    
    document.getElementById('downloadReportButton').addEventListener('click', () => {
        
        const dataCarga = document.getElementById('reportDate').value;
        
        gerarModelo(dataCarga);
        
    });

    carregarComFiltros();                                // tabela inicial

    document.getElementById('searchButton')
            .addEventListener('click', carregarComFiltros);
  
    document.getElementById('filtroCarreta')
            .addEventListener('keypress', e => e.key === 'Enter' && carregarComFiltros());
    document.getElementById('filtroCodigo')
            .addEventListener('keypress', e => e.key === 'Enter' && carregarComFiltros());
  
    document.getElementById("confirmBulkInsertButton").addEventListener("click", async function () {
        const fileInput = document.getElementById("bulkFile");
        const feedback = document.getElementById("bulkImportFeedback");
        const button = this;
        const spinner = button.querySelector(".spinner-border");
    
        feedback.innerHTML = "";
    
        if (!fileInput.files.length) {
            feedback.innerHTML = `<div class="alert alert-warning">Por favor, selecione um arquivo .xlsx antes de continuar.</div>`;
            return;
        }
    
        const formData = new FormData();
        formData.append("arquivo", fileInput.files[0]);
    
        // Ativa loading
        button.disabled = true;
        spinner.style.display = "inline-block";
    
        try {
            const response = await fetch("/api/importar-recurso", {
                method: "POST",
                body: formData
            });
    
            const result = await response.json();
    
            if (response.ok) {
                feedback.innerHTML =
                    `<div class="alert alert-success">
                        ${result.inseridos} recurso(s) importado(s) com sucesso.
                        ${result.duplicados.length ? `<br><strong>Ignorados (duplicados):</strong> ${result.duplicados.map(d => d.carreta + ' + ' + d.recurso).join(', ')}` : ''}
                    </div>`;
            } else {
                feedback.innerHTML =
                    `<div class="alert alert-danger">${result.error}</div>`;
            }
        } catch (err) {
            console.error("Erro ao importar:", err);
            feedback.innerHTML =
                `<div class="alert alert-danger">Erro inesperado ao importar.</div>`;
        } finally {
            button.disabled = false;
            spinner.style.display = "none";
        }
    });

});
  