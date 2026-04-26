// ======================================================
// Sistema de Acesso de Alunos via QR Code e Web App
// Projeto educacional com Google Apps Script
//
// Funcionalidades:
// - Lê turmas em abas diferentes da planilha
// - Cria documentos individuais para cada aluno
// - Gera links individuais
// - Gera QR Codes
// - Organiza documentos no Google Drive por turma
// - Atualiza documento apenas se a senha mudar
// - Disponibiliza Web App para consulta por matrícula + nome completo
// ======================================================


// ======================================================
// FUNÇÃO PRINCIPAL
// Cria documentos, links e QR Codes para todas as turmas
// ======================================================
function criarLinksTodasTurmas() {

  // Nome da pasta principal no Google Drive
  const NOME_PASTA_RAIZ = 'Sistema de Acesso de Alunos via QR Code';

  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  const abas = planilha.getSheets();

  // Cria ou reutiliza a pasta raiz do projeto
  const pastaRaiz = getOuCriarPasta(NOME_PASTA_RAIZ);

  // Percorre todas as abas da planilha
  for (let a = 0; a < abas.length; a++) {

    const aba = abas[a];
    const nomeTurma = aba.getName();
    const dados = aba.getDataRange().getValues();

    // Se a aba não tiver dados suficientes, pula
    if (dados.length < 2) continue;

    // Cria ou reutiliza pasta da turma
    const pastaTurma = getOuCriarPasta('Turma ' + nomeTurma, pastaRaiz);

    // Define os cabeçalhos principais
    aba.getRange(1, 4).setValue('Link Individual');
    aba.getRange(1, 5).setValue('QR Code');
    aba.getRange(1, 6).setValue('Link do QR Code');

    // Percorre os alunos, começando na linha 2
    for (let i = 1; i < dados.length; i++) {

      const nome = dados[i][0];          // Coluna A: ALUNOS
      const email = dados[i][1];         // Coluna B: EMAIL
      const senha = dados[i][2];         // Coluna C: SENHA
      const linkExistente = dados[i][3]; // Coluna D: Link Individual

      // Se faltar nome ou e-mail, pula a linha
      if (!nome || !email) continue;

      // ======================================================
      // CASO 1: ALUNO JÁ TEM LINK
      // Abre o documento existente e atualiza apenas se necessário
      // ======================================================
      if (linkExistente && linkExistente.toString().includes('docs.google.com')) {

        const match = linkExistente.toString().match(/[-\w]{25,}/);

        // Se não conseguir extrair o ID do documento, pula
        if (!match) continue;

        const idDoc = match[0];

        try {
          const doc = DocumentApp.openById(idDoc);
          const body = doc.getBody();
          const textoAtual = body.getText();

          // Atualiza o documento somente se a senha mudou
          if (!textoAtual.includes('Senha: ' + senha)) {
            preencherDocumentoAluno(body, nomeTurma, nome, email, senha);
            doc.saveAndClose();
          }

          // Gera novamente o QR Code com o mesmo link
          const qr = gerarUrlQRCode(linkExistente);

          aba.getRange(i + 1, 5).setFormula('=IMAGE("' + qr + '"; 1)');
          aba.getRange(i + 1, 6).setValue(qr);

        } catch (e) {
          // Se der erro ao abrir documento, não trava o script inteiro
          continue;
        }

        continue;
      }

      // ======================================================
      // CASO 2: ALUNO AINDA NÃO TEM LINK
      // Cria novo documento, gera link e QR Code
      // ======================================================
      try {

        // Pequena pausa para evitar limite de serviço do Google
        Utilities.sleep(300);

        const doc = DocumentApp.create('Dados - ' + nome);
        const body = doc.getBody();

        preencherDocumentoAluno(body, nomeTurma, nome, email, senha);

        doc.saveAndClose();

        const arquivo = DriveApp.getFileById(doc.getId());

        // Move o documento para a pasta da turma
        pastaTurma.addFile(arquivo);
        DriveApp.getRootFolder().removeFile(arquivo);

        // Permite acesso para qualquer pessoa com o link
        arquivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

        const link = doc.getUrl();
        const qr = gerarUrlQRCode(link);

        // Salva link e QR Code na planilha
        aba.getRange(i + 1, 4).setValue(link);
        aba.getRange(i + 1, 5).setFormula('=IMAGE("' + qr + '"; 1)');
        aba.getRange(i + 1, 6).setValue(qr);

      } catch (e) {
        // Se houver erro em uma linha, o script segue para a próxima
        continue;
      }
    }
  }
}


// ======================================================
// WEB APP
// Abre a interface HTML quando o aluno acessa o link
// ======================================================
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Acesso do Estudante');
}


// ======================================================
// BUSCA DO ALUNO NO WEB APP
// Valida matrícula + nome completo
// ======================================================
function buscarAluno(matricula, nome) {

  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  const abas = planilha.getSheets();

  const matriculaDigitada = matricula.toString().trim();
  const nomeDigitado = normalizarTexto(nome);

  // Percorre todas as abas
  for (let a = 0; a < abas.length; a++) {

    const aba = abas[a];
    const dados = aba.getDataRange().getValues();

    if (dados.length < 2) continue;

    const cabecalho = dados[0];

    // Identifica as colunas pelo nome do cabeçalho
    const idxNome = cabecalho.findIndex(col =>
      col.toString().toLowerCase().includes('aluno')
    );

    const idxMatricula = cabecalho.findIndex(col =>
      col.toString().toLowerCase().includes('matr')
    );

    const idxLink = cabecalho.findIndex(col =>
      col.toString().toLowerCase().includes('link individual')
    );

    // Se alguma coluna essencial não existir, ignora a aba
    if (idxNome === -1 || idxMatricula === -1 || idxLink === -1) continue;

    // Percorre os alunos da aba
    for (let i = 1; i < dados.length; i++) {

      const nomePlanilha = normalizarTexto(dados[i][idxNome]);
      const matriculaPlanilha = dados[i][idxMatricula]?.toString().trim();
      const link = dados[i][idxLink];

      // Validação segura:
      // matrícula exata + nome completo normalizado
      if (
        matriculaDigitada === matriculaPlanilha &&
        nomeDigitado === nomePlanilha
      ) {
        return link;
      }
    }
  }

  return 'NOT_FOUND';
}


// ======================================================
// FUNÇÃO AUXILIAR
// Cria ou reutiliza uma pasta no Google Drive
// ======================================================
function getOuCriarPasta(nome, pastaPai = null) {

  const pastas = pastaPai
    ? pastaPai.getFoldersByName(nome)
    : DriveApp.getFoldersByName(nome);

  if (pastas.hasNext()) {
    return pastas.next();
  }

  return pastaPai
    ? pastaPai.createFolder(nome)
    : DriveApp.createFolder(nome);
}


// ======================================================
// FUNÇÃO AUXILIAR
// Preenche o documento individual do aluno
// ======================================================
function preencherDocumentoAluno(body, nomeTurma, nome, email, senha) {

  body.clear();

  const titulo = body.appendParagraph('ACESSO DO ESTUDANTE');
  titulo.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  titulo.setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  body.appendParagraph('');

  const turma = body.appendParagraph('Turma: ' + nomeTurma);
  turma.setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  body.appendParagraph('');

  const nomeParagrafo = body.appendParagraph(nome);
  nomeParagrafo.setBold(true);
  nomeParagrafo.setFontSize(14);
  nomeParagrafo.setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  body.appendParagraph('');

  const emailParagrafo = body.appendParagraph('E-mail: ' + email);
  emailParagrafo.setBold(true);

  const senhaParagrafo = body.appendParagraph('Senha: ' + senha);
  senhaParagrafo.setBold(true);

  body.appendParagraph('');

  const aviso = body.appendParagraph('Guarde essas informações com segurança.');
  aviso.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  aviso.setBold(true);
}


// ======================================================
// FUNÇÃO AUXILIAR
// Gera URL do QR Code usando QuickChart
// ======================================================
function gerarUrlQRCode(link) {
  return 'https://quickchart.io/qr?text=' + encodeURIComponent(link);
}


// ======================================================
// FUNÇÃO AUXILIAR
// Normaliza texto para comparação
// Remove acentos, espaços extras e ignora maiúsculas/minúsculas
// ======================================================
function normalizarTexto(texto) {
  return texto
    .toString()
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/\s+/g, ' ')
    .trim();
}
