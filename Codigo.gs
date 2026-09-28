/**
 * Cria o menu na planilha.
 */
function onOpen() {
  try {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('🤖 Automação Fiscal')
      .addItem('📥 1. Configurar/Importar CSV', 'configurarCSV')
      .addItem('▶️ 2. Conferir guia na pasta', 'processarFechamentoMensal')
      .addToUi();
  } catch (e) {
    // Executado manualmente pelo editor (sem contexto de planilha aberta).
    // Isso é esperado: abra a planilha normalmente para ver o menu.
    console.log('onOpen executado fora do contexto da planilha: ' + e.message);
  }
}

/**
 * Função para receber o link do CSV.
 */
function configurarCSV() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    'Importar CSV do Acessórias',
    'Por favor, cole abaixo o LINK de compartilhamento ou o ID do arquivo CSV gerado (certifique-se de que o acesso esteja livre como Leitor):',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() === ui.Button.OK) {
    const input = response.getResponseText().trim();
    const fileId = extrairIdDoLink(input);
    
    if (fileId && fileId.length > 15) {
      PropertiesService.getDocumentProperties().setProperty('ID_CSV_ACESSORIAS', fileId);
      ui.alert('✅ Sucesso!', 'O CSV foi configurado corretamente.', ui.ButtonSet.OK);
    } else {
      ui.alert('❌ Erro', 'Link inválido.', ui.ButtonSet.OK);
    }
  }
}

/**
 * Função robusta para extração de ID do link do Google Drive (limpa parâmetros adicionais).
 */
function extrairIdDoLink(url) {
  if (!url) return null;
  if (url.indexOf('/') === -1 && url.length > 15) {
    return url;
  }
  const match = url.match(/[-\w]{25,}(?!.*[-\w]{25,})/);
  return match ? match[0] : null;
}

/**
 * Script principal OTIMIZADO (Escrita em Lote e Filtro de CSV com ponto e vírgula)
 */
function processarFechamentoMensal() {
  const ui = SpreadsheetApp.getUi();
  
  // ================= CONFIGURAÇÕES =================
  const ID_PLANILHA = "ID_EXEMPLO";
  const NOME_ABA = "07-2026";
  const ID_PASTA_RAIZ = "ID_EXEMPLO";
  // =================================================
  
  const ID_CSV_ACESSORIAS = PropertiesService.getDocumentProperties().getProperty('ID_CSV_ACESSORIAS');
  
  if (!ID_CSV_ACESSORIAS) {
    ui.alert('⚠️ Atenção', 'Você precisa configurar o CSV primeiro! Vá no menu e clique em "1. Configurar/Importar CSV".', ui.ButtonSet.OK);
    return;
  }
  
  const ss = SpreadsheetApp.openById(ID_PLANILHA);
  const sheet = ss.getSheetByName(NOME_ABA);
  
  const ultimaLinha = sheet.getLastRow();
  if (ultimaLinha < 2) return; 
  
  const data = sheet.getRange(1, 1, ultimaLinha, 12).getValues(); 
  const rootFolder = DriveApp.getFolderById(ID_PASTA_RAIZ);
  
  // 1. Lendo e filtrando o CSV do Acessórias na memória usando ';' como separador
  let enviosValidos = []; 
  try {
    const csvFile = DriveApp.getFileById(ID_CSV_ACESSORIAS);
    const csvContent = csvFile.getBlob().getDataAsString("UTF-8");
    
    // CORREÇÃO: Forçando o delimitador a ser Ponto e Vírgula (';')
    const linhasCsv = Utilities.parseCsv(csvContent, ';'); 
    
    for (let i = 1; i < linhasCsv.length; i++) {
      if (linhasCsv[i].length > 11) {
        let status = linhasCsv[i][10] ? String(linhasCsv[i][10]).toUpperCase().trim() : "";
        let departamento = linhasCsv[i][11] ? String(linhasCsv[i][11]).toUpperCase().trim() : "";
        
        if (status.includes("ENT.") && departamento.includes("SERVIÇO")) {
           let empresa = linhasCsv[i][2] ? String(linhasCsv[i][2]).toUpperCase().trim() : "";
           enviosValidos.push(empresa);
        }
      }
    }
  } catch(e) {
    ui.alert('❌ Erro', 'Não foi possível ler o CSV. Certifique-se de que ele é de fato um CSV válido e que você liberou a permissão de compartilhamento dele no Google Drive.', ui.ButtonSet.OK);
    return; 
  }

  const variacoesMes = gerarVariacoesMes(NOME_ABA);
  
  let colunaJ_saida = [];
  let colunaK_saida = [];
  let colunaL_saida = [];

  // 2. Percorrendo a planilha na memória
  for (let i = 1; i < data.length; i++) { 
    let row = data[i];
    let empresaCompleta = row[0]; 
    let valorColunaJ = row[9];
    let valorColunaK = row[10];
    let marcadoColunaL = row[11];
    
    if (!empresaCompleta || (marcadoColunaL !== true && String(marcadoColunaL).toLowerCase() !== "true")) {
       colunaJ_saida.push([valorColunaJ]);
       colunaK_saida.push([valorColunaK]);
       colunaL_saida.push([marcadoColunaL]);
       continue; 
    }
    
    let tresPrimeirosNomes = obterTresPrimeirosNomes(empresaCompleta);
    let primeiraPalavra = extrairPrimeiraPalavra(empresaCompleta);
    
    let empresaFolder = buscarPastaPorPrefixo(rootFolder, primeiraPalavra);
    let temGuia = false;
    
    if (empresaFolder) {
       let subfolders = empresaFolder.getFolders();
       let monthFolder = null;
       
       while (subfolders.hasNext()) {
         let subF = subfolders.next();
         let n = subF.getName().toUpperCase();
         
         if (variacoesMes.includes(n)) {
            monthFolder = subF;
            break;
         }
       }
       
       if (monthFolder) {
         let files = monthFolder.getFilesByType(MimeType.PDF); 
         if (files.hasNext()) {
             temGuia = true;
         }
       }
    }
    
    // 3. Cruza informações com o CSV
    let csvMatched = enviosValidos.some(empresaCsv => empresaCsv.includes(tresPrimeirosNomes));
    
    // 4. Salva as respostas
    colunaJ_saida.push([temGuia ? "Com guia na pasta" : "Sem DARF na pasta"]);
    colunaK_saida.push([csvMatched ? "Enviado - Acessórias" : "Pendente"]);
    colunaL_saida.push([false]); 
  }
  
  // 5. Escrita em lote rápida
  if (colunaJ_saida.length > 0) {
    sheet.getRange(2, 10, colunaJ_saida.length, 1).setValues(colunaJ_saida);
    sheet.getRange(2, 11, colunaK_saida.length, 1).setValues(colunaK_saida);
    sheet.getRange(2, 12, colunaL_saida.length, 1).setValues(colunaL_saida);
  }
  
  ui.alert('✅ Concluído', 'Conferência finalizada com sucesso!', ui.ButtonSet.OK);
}

/**
 * Funções Auxiliares
 */

function obterTresPrimeirosNomes(nomeCompleto) {
  if (!nomeCompleto) return "";
  let partes = String(nomeCompleto).trim().split(/\s+/);
  return partes.slice(0, 3).join(" ").toUpperCase();
}

function extrairPrimeiraPalavra(nomeCompleto) {
  if (!nomeCompleto) return "";
  let partes = String(nomeCompleto).replace(/['"]/g, '').trim().split(/\s+/);
  return partes[0].toUpperCase();
}

function gerarVariacoesMes(abaNome) {
  let partes = abaNome.split('-');
  if (partes.length === 2) {
    let mes = partes[0];
    let ano = partes[1];
    return [
      abaNome,              
      mes + ano,            
      ano + "-" + mes,      
      ano + mes,            
      mes + "." + ano,      
      mes + "_" + ano       
    ];
  }
  return [abaNome];
}

function buscarPastaPorPrefixo(pastaRaiz, primeiraPalavra) {
  if (!primeiraPalavra || primeiraPalavra === "") return null; 
  
  try {
    let pastas = pastaRaiz.searchFolders(`title contains '${primeiraPalavra}'`);
    
    while (pastas.hasNext()) {
      let f = pastas.next();
      let nomePasta = f.getName().toUpperCase();
      
      if (nomePasta.includes(primeiraPalavra)) {
         return f;
      }
    }
  } catch(e) { }
  return null;
}