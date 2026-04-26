# 🎓 Sistema de Acesso de Alunos via QR Code

Automação com Google Apps Script para gerar acessos individuais de alunos por meio de QR Codes e Web App, garantindo privacidade, organização e escalabilidade.

---

## 📌 Problema

Em ambiente escolar, é comum a necessidade de compartilhar e-mails e senhas dos alunos.  
Fazer isso de forma manual pode:

- expor dados de toda a turma  
- gerar desorganização  
- demandar muito tempo  

---

## ✅ Solução

Este projeto automatiza o processo:

- Cria um documento individual para cada aluno  
- Gera um link exclusivo de acesso  
- Cria QR Codes individuais  
- Disponibiliza um Web App com QR Code único  
- Mantém os dados atualizados sem alterar os links  

---

## 🏫 Aplicação prática

Este sistema foi desenvolvido e aplicado em contexto escolar real para distribuição de acessos individuais de alunos.

A solução permitiu:

- maior organização no processo  
- redução de tempo operacional  
- proteção dos dados dos estudantes  

---

## ⚙️ Tecnologias utilizadas

- Google Sheets  
- Google Apps Script (JavaScript)  
- Google Drive  
- HTML + JavaScript (interface do Web App)  
- QuickChart (geração de QR Code)  

---

## 🧠 Como funciona

### 🔹 Modelo 1 — QR Code individual

1. Cada aluno possui um documento individual  
2. Um QR Code leva diretamente ao documento  
3. O aluno acessa apenas suas informações  

---

### 🔹 Modelo 2 — QR Code único (Web App)

1. O aluno escaneia um QR Code geral  
2. Acessa uma página (Web App)  
3. Informa:
   - matrícula  
   - nome completo  
4. O sistema valida os dados  
5. O aluno acessa seu documento individual  

---

## 📊 Estrutura da planilha

| ALUNOS | EMAIL | SENHA | Link Individual | QR Code | Link do QR Code | Matrícula |
|--------|-------|-------|----------------|--------|----------------|-----------|

---

## 🔁 Atualizações inteligentes

- O QR Code não muda  
- O link permanece o mesmo  
- O documento é atualizado apenas quando necessário  
- Novos alunos são reconhecidos automaticamente  

---

## 🔐 Considerações de segurança

- Os links são individuais  
- O acesso é somente leitura  
- O Web App exige validação por matrícula e nome  
- O sistema ignora acentos e diferenças de escrita  
- Recomenda-se não compartilhar dados pessoais  

---

## 🧩 Estrutura do código

O sistema é dividido em:

- geração de documentos e QR Codes  
- leitura dinâmica da planilha  
- validação de usuários  
- Web App para consulta  
- organização automática no Google Drive  

---

## 💡 Como esse projeto surgiu

Este projeto nasceu de uma necessidade real em ambiente escolar:  
disponibilizar e-mails e senhas dos alunos de forma individual, sem expor os dados da turma.

Inicialmente, o processo era manual e pouco seguro.

A solução evoluiu para uma automação com Google Apps Script que gera documentos individuais, QR Codes e um sistema de consulta via Web App.

---

## 🚀 Como usar

1. Abra o Google Sheets  
2. Vá em **Extensões → Apps Script**  
3. Cole o código do projeto  
4. Execute a função principal  
5. Publique o Web App  
6. Gere os QR Codes  

---

## 📸 Exemplos

### Planilha com QR Codes individuais

![Planilha](imagens/planilha.png)

---

### Documento individual do aluno

![Documento](imagens/documento.png)

---

### Web App de acesso (QR Code único)

![Web App](imagens/webapp.png)

---

## 📱 Uso prático

O sistema pode ser utilizado por meio de:

- carteirinhas com QR Code  
- QR Code em murais ou salas  
- acesso via link compartilhado  

---

## 👩‍🏫 Contexto

Projeto desenvolvido para uso real em ambiente escolar, com foco em:

- organização  
- eficiência  
- privacidade dos dados  

---

## 📄 Licença

Este projeto está sob a licença MIT.
