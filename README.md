<h1 align="center">📝Conversor Markdown para Docx📝</h1>

O **ConvertMarkdownToDocx** é um projeto desenvolvido para automatizar a conversão de arquivos Markdown para documentos Word (.docx) formatados segundo as normas ABNT. É especialmente útil para acadêmicos, profissionais da educação, redatores técnicos e empresas que desejam uma forma prática e rápida de gerar documentos estruturados e profissionais a partir de conteúdo escrito em Markdown.

## Proposta do Projeto

O objetivo principal deste projeto é simplificar e agilizar o processo de criação de documentos formatados corretamente, eliminando o tempo gasto com a formatação manual no Word. É uma solução ideal para quem utiliza regularmente Markdown e precisa gerar documentos em conformidade com as normas ABNT.

O projeto lê o conteúdo Markdown, interpreta títulos, listas, tabelas e outros elementos comuns, e gera automaticamente um documento Word formatado, pronto para revisão ou impressão.

## Quando utilizar este projeto?

- Criação rápida e eficiente de relatórios técnicos e acadêmicos.
- Automatização da geração de documentos formatados conforme ABNT.
- Simplificação do processo de conversão de conteúdos produzidos por plataformas como o ChatGPT.

<h2 align="center">Como utilizar</h2>

### Opção 1: Instalação a partir do PyPI

A forma mais rápida de começar a usar o projeto é instalando diretamente do PyPI:

```bash
pip install convert-markdown-docx
```

Após instalado, você pode executar o conversor diretamente:

1. Para executar com o nome do arquivo gerado automaticamente:

```bash
convert_markdown input/seu_arquivo.md
```

2. Caso deseje definir manualmente o nome do arquivo que será gerado:

```bash
convert_markdown input/seu_arquivo.md -o nome_final.docx
```

O arquivo convertido será salvo na pasta `output`.

### Opção 2: Clonando o repositório

Se preferir trabalhar diretamente com o código-fonte, siga os passos abaixo:

#### Passo 1: Clone o repositório

```bash
git clone https://github.com/BrayanPletsch/ConvertMarkdownToDocx.git
cd ConvertMarkdownToDocx
```

#### Passo 2: Crie um ambiente virtual

```bash
python3 -m venv .venv      # Windows: python -m venv .venv 
source .venv/bin/activate  # Windows: .venv\Scripts\activate
```

#### Passo 3: Instale as dependências

```bash
pip install --upgrade pip
pip install -e .
```

Após esses passos, o comando `convert_markdown` estará disponível no seu terminal.

#### Passo 4: Execute o comando

```bash
convert_markdown input/seu_arquivo.md
# ou
convert_markdown input/seu_arquivo.md -o nome_final.docx
```

<h2 align="center">Estrutura do Projeto</h2>

```
ConvertMarkdownToDocx/
├── src/
│   └── convert_markdown/
│       ├── cli.py                # Interface de linha de comando
│       ├── converter.py          # Limpa e trata o conteúdo Markdown
│       ├── docx_generator.py     # Gera e formata o documento Word
│       └── utils/
│           └── markdown_utils.py # Funções auxiliares
├── tests/                        # Testes unitários
├── docs/                         # Documentação detalhada
├── input/                        # Arquivos Markdown de entrada
├── output/                       # Documentos Word gerados
├── LICENSE
├── pyproject.toml
├── requirements.txt
└── .gitignore
```

## Funcionalidades principais

- Lê texto Markdown com títulos, listas, tabelas e separadores.
- Cria um `.docx` com:
  - Títulos formatados conforme ABNT (tamanhos 18, 16, 14).
  - Corpo com fonte Times New Roman 12, justificado, espaçamento 1.5.
  - Geração automática de sumário (necessita atualização manual no Word).
  - Tabelas com bordas e suporte a **negrito dentro de células**.
  - Rodapé com numeração automática das páginas.

## Limitações atuais

- Não processa imagens no formato Markdown (`![]()`).
- Itálico (`*texto*`) ainda não é suportado.
- Blocos de código (```) não são tratados adequadamente.
- O sumário deve ser atualizado manualmente no Word após a geração.

## Futuras melhorias

- Implementação de suporte completo para imagens e links.
- Melhoria na conversão de listas aninhadas.
- Opção para exportação também em formato PDF.
- Interface web interativa via Swagger/OpenAPI.

<h2 align="center">Como contribuir</h2>

Contribuições são muito bem-vindas! Caso queira contribuir, siga estes passos:

1. Faça um **fork** do projeto.
2. Crie uma branch com sua feature:

```bash
git checkout -b minha-melhoria
```

3. Commit suas alterações:

```bash
git commit -m "feat: descrição da nova funcionalidade"
```

4. Envie sua branch para o repositório:

```bash
git push origin minha-melhoria
```

5. Abra um Pull Request com uma descrição clara das mudanças feitas.

## Licença

Este projeto está sob a licença MIT. Veja o arquivo [LICENSE](LICENSE) para mais detalhes.

<p align="center"><i>Desenvolvido por Brayan Pletsch.</i></p>