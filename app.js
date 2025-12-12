// Mapeamento de exames com tipos de dados
const MAPEAMENTO_EXAMES = {
    "36486": { colunas: ["C_Hb", "hb", "HB"], nome: "Hemoglobina (Hb)", tipo: "NUMBER" },
    "36485": { colunas: ["C_Ht", "ht", "HT"], nome: "Hematócrito (Ht)", tipo: "NUMBER" },
    "36452": { colunas: ["UREI"], nome: "Uréia Pré", tipo: "NUMBER" },
    "36581": { colunas: ["UPD"], nome: "Uréia Pós-Diálise", tipo: "NUMBER" },
    "36434": { colunas: ["CREA", "crea"], nome: "Creatinina", tipo: "NUMBER" },
    "36433": { colunas: ["CALCIO", "cálcio"], nome: "Cálcio", tipo: "NUMBER" },
    "36435": { colunas: ["FOSFS"], nome: "Fósforo", tipo: "NUMBER" },
    "36461": { colunas: ["Na"], nome: "Sódio", tipo: "NUMBER" },
    "36436": { colunas: ["POTAS"], nome: "Potássio (K)", tipo: "NUMBER" },
    "36437": { colunas: ["TGP"], nome: "TGP (ALT)", tipo: "NUMBER" },
    "36438": { colunas: ["GLIC"], nome: "Glicose", tipo: "NUMBER" },
    "36439": { colunas: ["ANTI_HIV", "HIV", "HIV_AB"], nome: "Anti - HIV - Anticorpo", tipo: "VARCHAR2" },
    "36447": { colunas: ["CTOT"], nome: "Colesterol Total", tipo: "NUMBER" },
    "36448": { colunas: ["ALU_SER"], nome: "Alumínio Sérico", tipo: "NUMBER" },
    "36449": { colunas: ["TRIG"], nome: "Triglicerídeos", tipo: "NUMBER" },
    "36450": { colunas: ["T4"], nome: "T4 Total", tipo: "NUMBER" },
    "36453": { colunas: ["Hb_A1c"], nome: "Hemoglobina Glicada (HbA1c)", tipo: "VARCHAR2" },
    "36455": { colunas: ["TSH"], nome: "TSH", tipo: "NUMBER" },
    "36456": { colunas: ["ANTI_HBS"], nome: "Anti-HBs (Hepatite B)", tipo: "NUMBER" },
    "36457": { colunas: ["VITD25OH"], nome: "Vitamina D (25-OH)", tipo: "NUMBER" },
    "36483": { colunas: ["PLAQ"], nome: "Plaquetas", tipo: "NUMBER" },
    "36501": { colunas: ["LEUC", "WBC", "LEUCOCITOS"], nome: "Leucócitos (07)", tipo: "NUMBER" },
    "36502": { colunas: ["PTH_DB"], nome: "PTH (Paratormônio)", tipo: "NUMBER" },
    "36518": { colunas: ["Ferritina"], nome: "Ferritina", tipo: "NUMBER" },
    "36520": { colunas: ["IST"], nome: "Índice Saturação Transferrina", tipo: "NUMBER" },
    "36522": { colunas: ["FA"], nome: "Fosfatase Alcalina", tipo: "NUMBER" },
    "36523": { colunas: ["PT"], nome: "Proteínas Totais", tipo: "NUMBER" },
    "36567": { colunas: ["FER"], nome: "Ferro Sérico", tipo: "NUMBER" },
    "36574": { colunas: ["HCV"], nome: "Anti-HCV (Hepatite C)", tipo: "VARCHAR2" },
    "36578": { colunas: ["AAU"], nome: "Anticorpo Anti-HBs", tipo: "VARCHAR2" },
    "36579": { colunas: ["HDL"], nome: "HDL Colesterol", tipo: "NUMBER" },
    "36580": { colunas: ["Col_LDL"], nome: "LDL Colesterol", tipo: "NUMBER" },
    "36582": { colunas: ["CTT"], nome: "Capacidade Total de Fixação", tipo: "NUMBER" },
    "36584": { colunas: ["ALB"], nome: "Albumina", tipo: "NUMBER" },
    "36585": { colunas: ["GLB"], nome: "Globulinas", tipo: "NUMBER" },
    "36587": { colunas: ["Rel_Alb_Gl"], nome: "Relação Albumina/Globulina", tipo: "NUMBER" },
    "36588": { colunas: ["CTF", "TIBC", "CLF"], nome: "Capacidade de Ligação de Ferro (CTF)", tipo: "NUMBER" }
};

const ORDEM_COLUNAS = [
    "NM_PACIENTE",
    "NR_ATENDIMENTO",
    "DT_RESULTADO",
    "DS_PROTOCOLO",
    "CD_ESTABELECIMENTO",
    "NR_EXAME_36433", "NR_EXAME_36434", "NR_EXAME_36435", "NR_EXAME_36436",
    "NR_EXAME_36437", "NR_EXAME_36438", "NR_EXAME_36439", "NR_EXAME_36447",
    "NR_EXAME_36448", "NR_EXAME_36449", "NR_EXAME_36450", "NR_EXAME_36452",
    "NR_EXAME_36453", "NR_EXAME_36455", "NR_EXAME_36456", "NR_EXAME_36457",
    "NR_EXAME_36461", "NR_EXAME_36483", "NR_EXAME_36485", "NR_EXAME_36486",
    "NR_EXAME_36501", "NR_EXAME_36502", "NR_EXAME_36518", "NR_EXAME_36520",
    "NR_EXAME_36522", "NR_EXAME_36523", "NR_EXAME_36567", "NR_EXAME_36574",
    "NR_EXAME_36578", "NR_EXAME_36579", "NR_EXAME_36580", "NR_EXAME_36581",
    "NR_EXAME_36582", "NR_EXAME_36584", "NR_EXAME_36585", "NR_EXAME_36587",
    "NR_EXAME_36588"
];

const ESTABELECIMENTOS = {
    "1": "MATRIZ",
    "3": "MONTE_SERRAT",
    "4": "CONVENIOS",
    "5": "RIO_VERMELHO",
    "7": "SANTO_ESTEVAO"
};

let dadosTasy = null;
let dadosLab1 = null;
let dadosLab2 = null;

// Normalizar nome para comparação
function normalizarNome(nome) {
    if (!nome) return '';
    return nome.toString().trim().toUpperCase()
        .normalize('NFD').replace(/[\u0300-\u036f]/g, '');
}

// Converter data do laboratório para formato do sistema
function converterData(dataStr) {
    if (!dataStr) return '';
    
    try {
        // Formato do lab: "01/12/2025 09:10:08"
        // Manter o mesmo formato
        return dataStr.toString().trim();
    } catch (e) {
        console.error('Erro ao converter data:', dataStr, e);
    }
    
    return '';
}

// Limpar e extrair valor numérico
function limparValorNumerico(valor) {
    if (valor === null || valor === undefined || valor === '') {
        return '';
    }
    
    // Se já é número, retornar
    if (typeof valor === 'number') {
        return valor;
    }
    
    // Converter para string
    const valorStr = valor.toString().trim();
    
    // Se está vazio após trim
    if (valorStr === '') return '';
    
    // Tentar extrair número
    // Remove caracteres como <, >, =, letras, etc
    // Mantém números, vírgula, ponto e sinal negativo
    let numLimpo = valorStr.replace(/[^0-9.,\-]/g, '');
    
    // Se não sobrou nada numérico, retornar vazio
    if (!numLimpo || numLimpo === '-') return '';
    
    // Substituir vírgula por ponto para conversão
    numLimpo = numLimpo.replace(',', '.');
    
    // Tentar converter
    const numero = parseFloat(numLimpo);
    
    // Se conversão falhou, retornar vazio
    if (isNaN(numero)) return '';
    
    return numero;
}

// Encontrar valor de exame e data nas planilhas do laboratório
function encontrarDadosExame(codigoExame, nomePaciente, lab1Data, lab2Data) {
    const config = MAPEAMENTO_EXAMES[codigoExame];
    if (!config) return { valor: null, data: null };

    const nomeNormalizado = normalizarNome(nomePaciente);
    
    // Buscar em ambas as planilhas
    const datasources = [lab1Data, lab2Data].filter(d => d !== null);
    
    for (const data of datasources) {
        for (const row of data) {
            const nomeRowNormalizado = normalizarNome(row.nome);
            
            if (nomeRowNormalizado === nomeNormalizado) {
                // Procurar em todas as possíveis colunas
                for (const coluna of config.colunas) {
                    if (row[coluna] !== undefined && row[coluna] !== null && row[coluna] !== '') {
                        let valorFinal = row[coluna];
                        
                        // Se o tipo é NUMBER, aplicar limpeza
                        if (config.tipo === 'NUMBER') {
                            valorFinal = limparValorNumerico(valorFinal);
                        }
                        
                        return { 
                            valor: valorFinal,
                            data: row.dthr_os || null
                        };
                    }
                }
            }
        }
    }
    
    return { valor: null, data: null };
}

// Ler arquivo Excel/CSV
function lerArquivo(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        
        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                const jsonData = XLSX.utils.sheet_to_json(firstSheet, { raw: false });
                resolve(jsonData);
            } catch (error) {
                reject(error);
            }
        };
        
        reader.onerror = reject;
        reader.readAsArrayBuffer(file);
    });
}

// Mostrar alerta
function mostrarAlerta(mensagem, tipo = 'info') {
    const alert = document.getElementById('alert');
    alert.className = `alert alert-${tipo} show`;
    alert.textContent = mensagem;
    
    setTimeout(() => {
        alert.classList.remove('show');
    }, 5000);
}

// Atualizar progresso
function atualizarProgresso(percentual) {
    const progressContainer = document.getElementById('progressContainer');
    const progressFill = document.getElementById('progressFill');
    
    progressContainer.style.display = 'block';
    progressFill.style.width = percentual + '%';
    progressFill.textContent = percentual + '%';
}

// Validar uploads
document.getElementById('fileTasy').addEventListener('change', async function(e) {
    const file = e.target.files[0];
    if (file) {
        try {
            dadosTasy = await lerArquivo(file);
            document.getElementById('statusTasy').textContent = `✓ ${dadosTasy.length} pacientes carregados`;
            document.getElementById('statusTasy').classList.add('show');
            mostrarAlerta('Planilha Tasy carregada com sucesso!', 'success');
        } catch (error) {
            mostrarAlerta('Erro ao ler planilha Tasy: ' + error.message, 'error');
        }
    }
});

document.getElementById('fileLab1').addEventListener('change', async function(e) {
    const file = e.target.files[0];
    if (file) {
        try {
            dadosLab1 = await lerArquivo(file);
            document.getElementById('statusLab1').textContent = `✓ ${dadosLab1.length} registros carregados`;
            document.getElementById('statusLab1').classList.add('show');
            mostrarAlerta('Planilha Laboratório 1 carregada com sucesso!', 'success');
        } catch (error) {
            mostrarAlerta('Erro ao ler planilha Lab 1: ' + error.message, 'error');
        }
    }
});

document.getElementById('fileLab2').addEventListener('change', async function(e) {
    const file = e.target.files[0];
    if (file) {
        try {
            dadosLab2 = await lerArquivo(file);
            document.getElementById('statusLab2').textContent = `✓ ${dadosLab2.length} registros carregados`;
            document.getElementById('statusLab2').classList.add('show');
            mostrarAlerta('Planilha Laboratório 2 carregada com sucesso!', 'success');
        } catch (error) {
            mostrarAlerta('Erro ao ler planilha Lab 2: ' + error.message, 'error');
        }
    }
});

// Processar arquivos
async function processarArquivos() {
    try {
        // Validações
        const estabelecimento = document.getElementById('estabelecimento').value;
        const protocolo = document.getElementById('protocolo').value;

        if (!estabelecimento || !protocolo) {
            mostrarAlerta('Preencha todos os campos obrigatórios!', 'error');
            return;
        }

        if (!dadosTasy) {
            mostrarAlerta('Carregue a planilha Tasy!', 'error');
            return;
        }

        if (!dadosLab1 && !dadosLab2) {
            mostrarAlerta('Carregue pelo menos uma planilha do laboratório!', 'error');
            return;
        }

        // Desabilitar botão
        const btn = document.getElementById('btnProcessar');
        btn.disabled = true;
        btn.textContent = '⏳ Processando...';

        atualizarProgresso(10);

        // Processar dados
        const resultados = [];
        const total = dadosTasy.length;
        
        for (let i = 0; i < dadosTasy.length; i++) {
            const paciente = dadosTasy[i];
            const percentual = Math.round(((i + 1) / total) * 80) + 10;
            atualizarProgresso(percentual);

            // Buscar dados do primeiro exame encontrado para pegar a data
            let dataResultado = '';
            const nomePaciente = paciente.nm_paciente || paciente.nome || '';
            
            // Tentar pegar data de qualquer exame deste paciente
            for (const codigoExame of Object.keys(MAPEAMENTO_EXAMES)) {
                const dados = encontrarDadosExame(codigoExame, nomePaciente, dadosLab1, dadosLab2);
                if (dados.data) {
                    dataResultado = converterData(dados.data);
                    break;
                }
            }

            // Criar linha base
            const linha = {
                NM_PACIENTE: nomePaciente,
                NR_ATENDIMENTO: paciente.nr_atendimento || '',
                DT_RESULTADO: dataResultado,
                DS_PROTOCOLO: protocolo,
                CD_ESTABELECIMENTO: estabelecimento
            };

            // Buscar valores de todos os exames
            Object.keys(MAPEAMENTO_EXAMES).forEach(codigoExame => {
                const dados = encontrarDadosExame(
                    codigoExame,
                    linha.NM_PACIENTE,
                    dadosLab1,
                    dadosLab2
                );
                linha[`NR_EXAME_${codigoExame}`] = dados.valor !== null ? dados.valor : '';
            });

            resultados.push(linha);
        }

        atualizarProgresso(95);

        // Criar arquivo Excel
        const ws = XLSX.utils.json_to_sheet(resultados, { header: ORDEM_COLUNAS });
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'Exames');

        // Gerar nome do arquivo
        const hoje = new Date();
        const dataFormatada = hoje.toISOString().split('T')[0].replace(/-/g, '');
        const nomeEstabelecimento = ESTABELECIMENTOS[estabelecimento];
        const nomeArquivo = `Exames_${nomeEstabelecimento}_${dataFormatada}.xlsx`;

        // Download
        XLSX.writeFile(wb, nomeArquivo);

        atualizarProgresso(100);
        mostrarAlerta(`✓ Arquivo gerado com sucesso! ${resultados.length} pacientes processados.`, 'success');

        // Reabilitar botão
        setTimeout(() => {
            btn.disabled = false;
            btn.textContent = '🚀 Processar e Gerar Arquivo';
            document.getElementById('progressContainer').style.display = 'none';
        }, 2000);

    } catch (error) {
        mostrarAlerta('Erro ao processar: ' + error.message, 'error');
        document.getElementById('btnProcessar').disabled = false;
        document.getElementById('btnProcessar').textContent = '🚀 Processar e Gerar Arquivo';
    }
}

// Remover função de setar data padrão
document.addEventListener('DOMContentLoaded', function() {
    // Preencher tabela de referência
    const tbody = document.getElementById('corpoTabelaExames');
    
    Object.keys(MAPEAMENTO_EXAMES).forEach(codigo => {
        const exame = MAPEAMENTO_EXAMES[codigo];
        const tr = document.createElement('tr');
        tr.style.borderBottom = '1px solid #ddd';
        
        // Cor de fundo alternada
        if (tbody.children.length % 2 === 0) {
            tr.style.background = '#f8f9fa';
        }
        
        // Cor especial para VARCHAR2
        const tipoBadge = exame.tipo === 'VARCHAR2' 
            ? '<span style="background: #17a2b8; color: white; padding: 3px 8px; border-radius: 4px; font-size: 11px;">VARCHAR2</span>'
            : '<span style="background: #28a745; color: white; padding: 3px 8px; border-radius: 4px; font-size: 11px;">NUMBER</span>';
        
        tr.innerHTML = `
            <td style="padding: 10px; border: 1px solid #ddd; font-weight: 600;">NR_EXAME_${codigo}</td>
            <td style="padding: 10px; border: 1px solid #ddd;">${exame.nome}</td>
            <td style="padding: 10px; border: 1px solid #ddd; text-align: center;">${tipoBadge}</td>
            <td style="padding: 10px; border: 1px solid #ddd; font-family: monospace; font-size: 12px; color: #666;">${exame.colunas.join(', ')}</td>
        `;
        
        tbody.appendChild(tr);
    });
});
