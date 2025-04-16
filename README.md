<h1 align="center">
  📝 Conversor de Markdown para Docx
</h1>

<p align="center">
  Transforme textos em Markdown diretamente em documentos Word (.docx) formatados segundo as <strong>normas da ABNT</strong>. Ideal para gerar relatórios, planos de aula, artigos e trabalhos acadêmicos direto do que o ChatGPT (ou você) produz em Markdown.
</p>

---

## 🎯 Proposta do Projeto

O objetivo principal do `ConvertMarkdownToDocx` é automatizar a geração de documentos `.docx` formatados com base no padrão ABNT, a partir de textos em Markdown — formato amplamente utilizado por desenvolvedores, escritores técnicos e plataformas como o ChatGPT.

> Imagine gerar um relatório acadêmico com negritos, títulos, tabelas e sumário **direto do conteúdo que o ChatGPT já retorna**. É isso que este projeto faz.

---

## 🧪 Quando usar?

- Para transformar resumos, relatórios e planejamentos em Markdown em documentos profissionais.
- Em projetos acadêmicos, educacionais ou corporativos que precisam de **conformidade com a ABNT**.
- Para automatizar a criação de `.docx` sem depender de ferramentas online ou formatação manual.

---

## 🚀 Como utilizar o projeto

### 1️⃣ Clone o repositório:

```bash
git clone https://github.com/BrayanPletsch/ConvertMarkdownToDocx.git
cd ConvertMarkdownToDocx
```

### 2️⃣ Crie o ambiente virtual:

```bash
python3 -m venv .venv      # no Windows: python -m venv .venv
source .venv/bin/activate  # no Windows: .venv\Scripts\activate
```

### 3️⃣ Instale as dependências:

```bash
pip install -r requirements.txt
```

---

### 4️⃣ Execute o script:

Edite a variável `markdown_text` no `app.py` com seu conteúdo Markdown.

Depois, execute:

```bash
python3 app.py  # no Windows: python .\app.py
```

✅ Um documento `.docx` formatado em ABNT será salvo na pasta `output/`, com nome único baseado no timestamp.

---

## 📂 Estrutura do Projeto

```
ConvertMarkdownToDocx/
├── .venv/                 # Ambiente virtual (ignorado no Git)
├── app.py                 # Script principal para conversão e formatação
├── output/                # Onde os arquivos .docx são salvos
├── requirements.txt       # Bibliotecas utilizadas
└── .gitignore             # Arquivos e pastas ignoradas pelo Git
```

---

## 🔍 O que o script faz?

(❗ Em breve será detalhado em `docs/`)

- Lê texto Markdown com títulos, listas, tabelas e separadores.
- Cria um `.docx` com:
  - Títulos formatados conforme ABNT (tamanhos 18, 16, 14).
  - Corpo com fonte Times New Roman 12, justificado, espaçamento 1.5.
  - Sumário automático gerado no início (basta atualizar no Word).
  - Tabelas com bordas e suporte a **negrito dentro de células**.
  - Linhas horizontais a partir de `---`.
  - Numeração de página no rodapé.

---

## 🤝 Como contribuir

Contribuições são bem-vindas! Para colaborar:

1. Faça um **fork** do projeto
2. Crie uma branch com sua feature:
   ```bash
   git checkout -b minha-melhoria
   ```
3. Faça os commits:
   ```bash
   git commit -m "feat: adicionei nova funcionalidade"
   ```
4. Envie sua branch:
   ```bash
   git push origin minha-melhoria
   ```
5. Abra um Pull Request 🚀

---

## 🧾 Licença

Distribuído sob a licença MIT. Veja `LICENSE` para mais detalhes.

---

<p align="center"><i>Desenvolvido com por Brayan Pletsch</i></p>
