// Sistema de acesso de alunos via QR Code e Web App
// Projeto educacional com Google Apps Script

function criarLinksTodasTurmas() {

  const NOME_PASTA_RAIZ = 'Sistema de Acesso de Alunos via QR Code';

  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  const abas = planilha.getSheets();

  // ================================
  // FUNÇÃO: pega ou cria pasta
  // ================================
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

  // ================================
  // PASTA RAIZ DO PROJETO
  // ================================
  const pastaRaiz = getOuCriarPasta(NOME_PASTA_RAIZ);

  // ================================
  // LOOP DAS TURMAS
  // ================================
  for (let a = 0; a < abas.length; a++) {

    const aba = abas[a];
    const nomeTurma = aba.getName();
    const dados = aba.getDataRange().getValues();

    if (dados.length < 2) continue;

    // Pasta da turma dentro da raiz
    const pastaTurma = getOuCriarPasta('Turma ' + nomeTurma, pastaRaiz);

    // Cabeçalhos
    aba.getRange(1, 4).setValue('Link Individual');
    aba.getRange(1, 5).setValue('QR Code');
    aba.getRange(1, 6).setValue('Link do QR Code');

    // ================================
    // LOOP DOS ALUNOS
    // ================================
    for (let i = 1; i < dados.length; i++) {

      const nome = dados[i][0];
      const email = dados[i][1];
      const senha = dados[i][2];
      const linkExistente = dados[i][3];

      if (!nome || !email) continue;

      // =============================
      // SE JÁ EXISTE DOCUMENTO
      // =============================
      if (linkExistente && linkExistente.includes('docs.google.com')) {

        const match = linkExistente.match(/[-\\w]{25,}/);
        if (!match) continue;

        const idDoc = match[0];

        try {
          const doc = DocumentApp.openById(idDoc);
          const body = doc.getBody();
          const textoAtual = body.getText();

          // Atualiza só se a senha mudou
          if (!textoAtual.includes('Senha: ' + senha)) {

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

            body.appendParagraph('E-mail: ' + email);
            body.appendParagraph('Senha: ' + senha);

            body.appendParagraph('');

            const aviso = body.appendParagraph('Guarde essas informações com segurança.');
            aviso.setAlignment(DocumentApp.HorizontalAlignment.CENTER);

            doc.saveAndClose();
          }

          const qr = 'https://quickchart.io/qr?text=' + encodeURIComponent(linkExistente);

          aba.getRange(i + 1, 5).setFormula('=IMAGE("' + qr + '"; 1)');
          aba.getRange(i + 1, 6).setValue(qr);

        } catch (e) {
          // Se der erro ao abrir doc, ignora
          continue;
        }

        continue;
      }

      // =============================
      // CRIA NOVO DOCUMENTO
      // =============================
      try {

        Utilities.sleep(300);

        const doc = DocumentApp.create(`Dados - ${nome}`);
        const body = doc.getBody();

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

        body.appendParagraph('E-mail: ' + email);
        body.appendParagraph('Senha: ' + senha);

        body.appendParagraph('');

        const aviso = body.appendParagraph('Guarde essas informações com segurança.');
        aviso.setAlignment(DocumentApp.HorizontalAlignment.CENTER);

        doc.saveAndClose();

        const arquivo = DriveApp.getFileById(doc.getId());

        // Move para pasta da turma
        pastaTurma.addFile(arquivo);
        DriveApp.getRootFolder().removeFile(arquivo);

        // Permissão pública
        arquivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

        const link = doc.getUrl();
        const qr = 'https://quickchart.io/qr?text=' + encodeURIComponent(link);

        aba.getRange(i + 1, 4).setValue(link);
        aba.getRange(i + 1, 5).setFormula('=IMAGE("' + qr + '"; 1)');
        aba.getRange(i + 1, 6).setValue(qr);

      } catch (e) {
        continue;
      }
    }
  }
}

// ================================
// WEB APP - PONTO DE ENTRADA
// ================================
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index');
}


// ================================
// FUNÇÃO PRINCIPAL DE BUSCA
// ================================
function buscarAluno(matricula, nome) {

  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  const abas = planilha.getSheets();

  // normaliza entrada do usuário
  const matriculaDigitada = matricula.toString().trim();
  const nomeDigitado = normalizarTexto(nome);

  // percorre todas as abas (turmas)
  for (let a = 0; a < abas.length; a++) {

    const aba = abas[a];
    const dados = aba.getDataRange().getValues();

    if (dados.length < 2) continue; // pula abas vazias

    const cabecalho = dados[0];

    // ================================
    // IDENTIFICA COLUNAS DINAMICAMENTE
    // ================================

    // Nome (aceita ALUNOS, Nome, etc.)
    const idxNome = cabecalho.findIndex(col =>
      col.toString().toLowerCase().includes("aluno")
    );

    // Matrícula
    const idxMatricula = cabecalho.findIndex(col =>
      col.toString().toLowerCase().includes("matr")
    );

    // Link do documento
    const idxLink = cabecalho.findIndex(col =>
      col.toString().toLowerCase().includes("link individual")
    );

    // se não encontrar as colunas, ignora a aba
    if (idxNome === -1 || idxMatricula === -1 || idxLink === -1) continue;

    // ================================
    // PERCORRE OS ALUNOS
    // ================================
    for (let i = 1; i < dados.length; i++) {

      const nomePlanilha = normalizarTexto(dados[i][idxNome]);
      const matriculaPlanilha = dados[i][idxMatricula]?.toString().trim();
      const link = dados[i][idxLink];

      // ================================
      // VALIDAÇÃO SEGURA
      // ================================
      if (
        matriculaDigitada === matriculaPlanilha &&
        nomeDigitado === nomePlanilha
      ) {
        return link;
      }
    }
  }

  // se não encontrou
  return "NOT_FOUND";
}


// ================================
// FUNÇÃO PARA NORMALIZAR TEXTO
// ================================
function normalizarTexto(texto) {
  return texto
    .toString()
    .toLowerCase()
    .normalize("NFD") // separa acentos
    .replace(/[\u0300-\u036f]/g, "") // remove acentos
    .replace(/\s+/g, " ") // remove espaços duplicados
    .trim();
}
