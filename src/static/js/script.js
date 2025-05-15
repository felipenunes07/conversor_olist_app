document.addEventListener("DOMContentLoaded", () => {
    const clienteSelect = document.getElementById("cliente-select");
    const usarArquivoExcelBtn = document.getElementById("usar-arquivo-excel-btn");
    const arquivoExcelInput = document.getElementById("arquivo-excel-input");
    const processarBtn = document.getElementById("processar-btn");
    const previewArea = document.getElementById("preview-area");
    const uploadArea = document.getElementById("upload-area");
    const fileNameDisplay = document.getElementById("file-name-display");

    // Elementos da nova seção de upload de mapeamento
    const clientesFileInput = document.getElementById("clientes-file-input");
    const produtosFileInput = document.getElementById("produtos-file-input");
    const clientesFileNameDisplay = document.getElementById("clientes-file-name-display");
    const produtosFileNameDisplay = document.getElementById("produtos-file-name-display");
    const salvarMapeamentoBtn = document.getElementById("salvar-mapeamento-btn");
    const mappingUploadStatus = document.getElementById("mapping-upload-status");

    let arquivoSelecionado = null;
    let arquivoClientesSelecionado = null;
    let arquivoProdutosSelecionado = null;

    // Função para carregar clientes no select
    function carregarClientes() {
        // Limpa opções existentes, exceto a primeira ("Selecione um cliente...")
        while (clienteSelect.options.length > 1) {
            clienteSelect.remove(1);
        }

        fetch("/clientes")
            .then(response => {
                if (!response.ok) {
                    return response.json().then(err => { 
                        throw new Error(err.error || "Erro ao carregar clientes do servidor."); 
                    });
                }
                return response.json();
            })
            .then(data => {
                if (data.error) {
                    console.error("Erro ao carregar clientes:", data.error);
                    // Não sobrescrever a previewArea principal se ela já estiver mostrando algo do processamento
                    if (!previewArea.querySelector("#download-link")){
                        previewArea.innerHTML = `<p style='color:red;'>Erro ao carregar lista de clientes: ${data.error}</p>`;
                    }
                    return;
                }
                if (data.clientes && data.clientes.length > 0) {
                    const clientesOrdenados = data.clientes.sort((a, b) => {
                        const nomeA = String(a.Nome);
                        const nomeB = String(b.Nome);
                        const matchA = nomeA.match(/^CL(\d+)/);
                        const matchB = nomeB.match(/^CL(\d+)/);
                        const idNumA = matchA ? parseInt(matchA[1], 10) : Infinity;
                        const idNumB = matchB ? parseInt(matchB[1], 10) : Infinity;
                        if (idNumA !== Infinity && idNumB !== Infinity) return idNumA - idNumB;
                        if (idNumA !== Infinity) return -1;
                        if (idNumB !== Infinity) return 1;
                        return nomeA.localeCompare(nomeB);
                    });
                    clientesOrdenados.forEach(cliente => {
                        const option = document.createElement("option");
                        option.value = cliente.ID;
                        option.textContent = cliente.Nome;
                        clienteSelect.appendChild(option);
                    });
                } else {
                     if (!previewArea.querySelector("#download-link")){
                        previewArea.innerHTML = "<p>Nenhum cliente encontrado. Verifique os arquivos de mapeamento.</p>";
                    }
                }
            })
            .catch(error => {
                console.error("Erro na requisição para buscar clientes:", error);
                 if (!previewArea.querySelector("#download-link")){
                    previewArea.innerHTML = `<p style='color:red;'>Falha ao buscar clientes: ${error.message}</p>`;
                }
            });
    }

    carregarClientes(); // Carrega clientes ao iniciar

    usarArquivoExcelBtn.addEventListener("click", () => {
        arquivoExcelInput.click();
    });

    arquivoExcelInput.addEventListener("change", (event) => {
        if (event.target.files.length > 0) {
            arquivoSelecionado = event.target.files[0];
            fileNameDisplay.textContent = `Arquivo selecionado: ${arquivoSelecionado.name}`;
            previewArea.innerHTML = `<p>Arquivo "${arquivoSelecionado.name}" pronto para processar.</p><p>Selecione o cliente e clique em processar.</p>`;
        }
    });

    uploadArea.addEventListener("dragover", (event) => {
        event.preventDefault();
        uploadArea.classList.add("dragover");
    });

    uploadArea.addEventListener("dragleave", () => {
        uploadArea.classList.remove("dragover");
    });

    uploadArea.addEventListener("drop", (event) => {
        event.preventDefault();
        uploadArea.classList.remove("dragover");
        const files = event.dataTransfer.files;
        if (files.length > 0) {
            const file = files[0];
            if (file.name.endsWith(".xlsx") || file.name.endsWith(".xls")) {
                arquivoSelecionado = file;
                fileNameDisplay.textContent = `Arquivo selecionado: ${arquivoSelecionado.name}`;
                previewArea.innerHTML = `<p>Arquivo "${arquivoSelecionado.name}" pronto para processar.</p><p>Selecione o cliente e clique em processar.</p>`;
            } else {
                fileNameDisplay.textContent = "";
                previewArea.innerHTML = "<p style='color:red;'>Por favor, solte um arquivo Excel (.xlsx ou .xls).</p>";
                arquivoSelecionado = null;
            }
        }
    });

    processarBtn.addEventListener("click", () => {
        const clienteId = clienteSelect.value;
        if (!clienteId) {
            previewArea.innerHTML = "<p style='color:red;'>Por favor, selecione um cliente.</p>";
            return;
        }
        if (!arquivoSelecionado) {
            previewArea.innerHTML = "<p style='color:red;'>Por favor, selecione um arquivo Excel de orçamento.</p>";
            return;
        }
        previewArea.innerHTML = "<p>Processando...</p>";
        const formData = new FormData();
        formData.append("cliente_id", clienteId);
        formData.append("arquivo_excel", arquivoSelecionado);
        fetch("/processar", { method: "POST", body: formData })
            .then(response => {
                if (!response.ok) {
                    return response.json().then(err => { throw new Error(err.error || "Erro no servidor ao processar o orçamento") });
                }
                const contentType = response.headers.get("content-type");
                if (contentType && contentType.indexOf("application/json") !== -1) {
                    return response.json().then(err => { throw new Error(err.error || "Erro retornado pelo servidor ao processar o orçamento") });
                }
                return response.blob();
            })
            .then(blob => {
                if (blob) {
                    const url = window.URL.createObjectURL(blob);
                    const a = document.createElement("a");
                    a.href = url;
                    a.download = "orcamento_convertido_olist.xlsx";
                    document.body.appendChild(a);
                    a.click();
                    a.remove();
                    window.URL.revokeObjectURL(url);
                    previewArea.innerHTML = `<div id='download-link'><p>Arquivo de orçamento processado com sucesso!</p><a href='${url}' download='orcamento_convertido_olist.xlsx'>Clique aqui para baixar novamente</a></div>`;
                    fileNameDisplay.textContent = "";
                    arquivoSelecionado = null;
                    arquivoExcelInput.value = null;
                    clienteSelect.selectedIndex = 0;
                }
            })
            .catch(error => {
                console.error("Erro no processamento do orçamento:", error);
                previewArea.innerHTML = `<p style='color:red;'>Erro no processamento do orçamento: ${error.message}</p>`;
            });
    });

    // Lógica para upload de arquivos de mapeamento
    clientesFileInput.addEventListener("change", (event) => {
        if (event.target.files.length > 0) {
            arquivoClientesSelecionado = event.target.files[0];
            clientesFileNameDisplay.textContent = `Arquivo selecionado: ${arquivoClientesSelecionado.name}`;
        } else {
            arquivoClientesSelecionado = null;
            clientesFileNameDisplay.textContent = "";
        }
    });

    produtosFileInput.addEventListener("change", (event) => {
        if (event.target.files.length > 0) {
            arquivoProdutosSelecionado = event.target.files[0];
            produtosFileNameDisplay.textContent = `Arquivo selecionado: ${arquivoProdutosSelecionado.name}`;
        } else {
            arquivoProdutosSelecionado = null;
            produtosFileNameDisplay.textContent = "";
        }
    });

    salvarMapeamentoBtn.addEventListener("click", () => {
        mappingUploadStatus.innerHTML = ""; // Limpa status anterior
        let uploadsPromises = [];

        if (arquivoClientesSelecionado) {
            const formDataClientes = new FormData();
            formDataClientes.append("file", arquivoClientesSelecionado);
            formDataClientes.append("file_type", "clientes");
            uploadsPromises.push(
                fetch("/upload_mapeamento", { method: "POST", body: formDataClientes })
                    .then(response => response.json())
                    .then(data => {
                        if (data.error) throw new Error(`Clientes: ${data.error}`);
                        return data.message;
                    })
            );
        }

        if (arquivoProdutosSelecionado) {
            const formDataProdutos = new FormData();
            formDataProdutos.append("file", arquivoProdutosSelecionado);
            formDataProdutos.append("file_type", "produtos");
            uploadsPromises.push(
                fetch("/upload_mapeamento", { method: "POST", body: formDataProdutos })
                    .then(response => response.json())
                    .then(data => {
                        if (data.error) throw new Error(`Produtos: ${data.error}`);
                        return data.message;
                    })
            );
        }

        if (uploadsPromises.length === 0) {
            mappingUploadStatus.innerHTML = "<p style='color:orange;'>Nenhum arquivo selecionado para upload.</p>";
            return;
        }

        mappingUploadStatus.innerHTML = "<p>Enviando arquivos...</p>";

        Promise.all(uploadsPromises)
            .then(messages => {
                mappingUploadStatus.innerHTML = messages.map(msg => `<p style='color:green;'>${msg}</p>`).join("");
                // Limpa os inputs e nomes de arquivos após sucesso
                if(arquivoClientesSelecionado) {
                    clientesFileInput.value = null;
                    clientesFileNameDisplay.textContent = "";
                    arquivoClientesSelecionado = null;
                    carregarClientes(); // Recarrega a lista de clientes
                }
                if(arquivoProdutosSelecionado) {
                    produtosFileInput.value = null;
                    produtosFileNameDisplay.textContent = "";
                    arquivoProdutosSelecionado = null;
                }
            })
            .catch(error => {
                mappingUploadStatus.innerHTML = `<p style='color:red;'>Erro no upload: ${error.message}</p>`;
            });
    });
});

