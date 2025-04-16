<h1 align="center"> 
    Documentação do Código
</h1>
<p align="center">
  Este documento explica, passo a passo, o funcionamento interno do script `app.py` utilizado no projeto **ConvertMarkdownToDocx**, que transforma um texto em Markdown em um documento `.docx` formatado conforme as normas da ABNT.
</p>

---

## ✅ Objetivo do Script
Transformar um texto escrito em Markdown em um documento `.docx`, com:
- Estilo ABNT (margens, fonte, espaçamento)
- Títulos hierárquicos personalizados
- Tabelas com suporte a negrito
- Linhas horizontais visuais (a partir de `---`)
- Sumário automático
- Numeração de páginas

---

## **📎 Índice**
1. [Criação do Documento e Configuração das Margens (ABNT)](#1-criação-do-documento-e-configuração-das-margens-abnt)
2. [Configuração dos Estilos para Títulos (Heading 1, 2, 3)](#2-configuração-dos-estilos-para-títulos-heading-1-2-3)
3. [Inserção de Numeração de Página no Rodapé (Centralizado)](#3-inserção-de-numeração-de-página-no-rodapé-centralizado)
4. [Criação de Página de Sumário (TOC) Isolada](#4-criação-de-página-de-sumário-toc-isolada)
5. [Processar Conteúdo Markdown e Adicionar ao Documento](#5-processar-conteúdo-markdown-e-adicionar-ao-documento)
6. [Salvar o Documento na Pasta `output/` com Timestamp no Nome](#6-salvar-o-documento-na-pasta-output-com-timestamp-no-nome)
7. [❌ Limitações atuais](#-limitações-atuais)
8. [📆 Futuras melhorias](#-futuras-melhorias)

---

## ⌚ Bibliotecas Utilizadas

### ✨ Bibliotecas internas (Python):
- `os`, `re`, `datetime`: manipulação de arquivos, expressões regulares e timestamps

### 📊 Bibliotecas externas:
- `python-docx`: gera e manipula arquivos `.docx`
- `docx.oxml`, `docx.shared`, `docx.enum`: customiza estilos, espaçamento, alinhamento, campos (ex: numeração e TOC)

---

## ✏️ Estrutura do Script

### 1. Criação do Documento e Configuração das Margens (ABNT)

O documento é iniciado com a função `Document()` da biblioteca `python-docx`, que cria um novo arquivo Word em branco. Em seguida, são configuradas as margens do documento conforme os padrões da ABNT:

- **Margem superior**: 3 cm
- **Margem inferior**: 2 cm
- **Margem esquerda**: 3 cm
- **Margem direita**: 2 cm

Esses ajustes são aplicados a todas as seções do documento com o loop `for section in doc.sections`:

```python
doc = Document()

for section in doc.sections:
    section.top_margin = Cm(3)
    section.bottom_margin = Cm(2)
    section.left_margin = Cm(3)
    section.right_margin = Cm(2)
```

Após configurar as margens, define-se o estilo padrão de parágrafo com o estilo `Normal`, ajustando:

- **Fonte**: Times New Roman
- **Tamanho da fonte**: 12 pt
- **Cor**: preta (`RGBColor(0, 0, 0)`)
- **Alinhamento**: justificado (`WD_ALIGN_PARAGRAPH.JUSTIFY`)
- **Espaçamento entre linhas**: 1,5 linhas

```python
# Definir estilo Normal: Times New Roman 12, preto, justificado, espaçamento 1.5
normal_style = doc.styles['Normal']
normal_style.font.name = 'Times New Roman'
normal_style.font.size = Pt(12)
normal_style.font.color.rgb = RGBColor(0, 0, 0)
normal_style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
normal_style.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
```

---

### 2. Configuração dos Estilos para Títulos (Heading 1, 2, 3)

Para manter a hierarquia de seções do documento e a conformidade com as normas da ABNT, os títulos também são personalizados.

- `#` Heading 1: fonte Times New Roman, 18 pt, negrito
- `##` Heading 2: fonte Times New Roman, 16 pt, negrito
- `###` Heading 3: fonte Times New Roman, 14 pt, negrito

Todos os títulos são alinhados à esquerda para seguir o estilo acadêmico tradicional.

O código percorre os níveis de título (1 a 3) e aplica as configurações de estilo adequadas:

```python
heading_levels = {1: 18, 2: 16, 3: 14}
for level, size in heading_levels.items():
    style = doc.styles[f'Heading {level}']
    style.font.name = 'Times New Roman'
    style.font.size = Pt(size)
    style.font.color.rgb = RGBColor(0, 0, 0)
    style.font.bold = True
    style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
```

Essa configuração garante que os títulos adicionados posteriormente ao documento, seja via `add_heading` ou por conversão de Markdown, estejam com estilo visual adequado e respeitem a hierarquia informacional do texto.

---

### 3. Inserção de Numeração de Página no Rodapé (Centralizado)

Para seguir o formato acadêmico exigido pelas normas ABNT, o documento inclui a numeração de página no rodapé, centralizada. A numeração é feita de forma automática utilizando campos especiais do Word (campo `PAGE`). Esse campo é interpretado pelo Microsoft Word como um marcador dinâmico, exibindo o número da página atual.

#### Lógica de implementação:
- O código percorre todas as seções do documento.
- Em cada uma, localiza o primeiro parágrafo do rodapé.
- Alinha esse parágrafo ao centro.
- Insere um campo Word dinâmico que representa a numeração de página.

#### Código correspondente:

```python
for section in doc.sections:
    footer_paragraph = section.footer.paragraphs[0]
    footer_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    # Inserir campo de número de página { PAGE }
    run = footer_paragraph.add_run()

    fld_begin = OxmlElement('w:fldChar')
    fld_begin.set(qn('w:fldCharType'), 'begin')

    instr = OxmlElement('w:instrText')
    instr.set(qn('xml:space'), 'preserve')
    instr.text = "PAGE"

    fld_end = OxmlElement('w:fldChar')
    fld_end.set(qn('w:fldCharType'), 'end')

    run._r.append(fld_begin)
    run._r.append(instr)
    run._r.append(fld_end)
```

Esse bloco de código garante que todas as páginas do documento gerado exibam corretamente sua numeração no rodapé. A renderização do número será automática ao abrir o `.docx` no Microsoft Word, sem necessidade de intervenção do usuário.

---

### 4. Criação de Página de Sumário (TOC) Isolada

Um sumário (TOC - Table of Contents) permite que o leitor visualize rapidamente a estrutura do documento, com os títulos e seções devidamente hierarquizados. Este bloco do código insere o sumário em uma página exclusiva e gera o campo de índice dinâmico do Word.

#### Lógica de implementação:
- Adiciona uma **quebra de página** antes do sumário, caso o documento já tenha conteúdo.
- Cria o título "SUMÁRIO" centralizado e em negrito.
- Insere o campo TOC (`TOC \o "1-3" \h \z \u`) que inclui títulos de níveis 1 a 3.
- Insere um texto de aviso para que o usuário atualize o sumário.
- Adiciona uma nova quebra de página após o sumário para isolar essa seção.

#### Código correspondente:

```python
# Quebra de página antes do Sumário (se já houver conteúdo no documento)
if len(doc.paragraphs) > 0:
    doc.add_page_break()

# Título do sumário
toc_title_paragraph = doc.add_paragraph("SUMÁRIO")
toc_title_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
toc_title_paragraph.runs[0].bold = True

# Inserir campo de Sumário (TOC) para níveis 1-3
toc_paragraph = doc.add_paragraph()
run = toc_paragraph.add_run()

fld_begin = OxmlElement('w:fldChar')
fld_begin.set(qn('w:fldCharType'), 'begin')

instr = OxmlElement('w:instrText')
instr.set(qn('xml:space'), 'preserve')
instr.text = 'TOC \\o "1-3" \\h \\z \\u'

fld_separate = OxmlElement('w:fldChar')
fld_separate.set(qn('w:fldCharType'), 'separate')

# Texto opcional mostrado antes de atualizar o campo
fld_separator_text = OxmlElement('w:t')
fld_separator_text.text = "Atualize o campo para gerar o sumário.(F9 ou clique direito → Atualizar campo)"
fld_separate.append(fld_separator_text)

fld_end = OxmlElement('w:fldChar')
fld_end.set(qn('w:fldCharType'), 'end')

# Montar os elementos do campo TOC no parágrafo
run._r.append(fld_begin)
run._r.append(instr)
run._r.append(fld_separate)
run._r.append(fld_end)

# Quebra de página após o sumário
doc.add_page_break()
```

> O campo de sumário só será exibido corretamente no Word após atualizar os campos (F9). O script já adiciona a instrução de atualização visível para o usuário.

---

### 5. Processar Conteúdo Markdown e Adicionar ao Documento

Essa etapa é responsável por interpretar cada linha do conteúdo Markdown e convertê-la para o formato Word equivalente, respeitando estilos e formatações. O Markdown é processado linha a linha e diferentes estruturas são detectadas e tratadas individualmente.

#### Lógica de Implementação:
- Cabeçalhos (`#`, `##`, `###`) → convertidos em títulos com hierarquia
- Tabelas → renderizadas com estilo e suporte a negrito nas células
- Listas → listas com marcadores e numeradas, estilizadas automaticamente
- Linhas horizontais (`---`) → inseridas como borda inferior
- Negrito (`**texto**`) → aplicado com run em negrito
- Parágrafos comuns → adicionados com estilo normal

#### Código Correspondente:

```python
lines = markdown_text.splitlines()
i = 0
while i < len(lines):
    line = lines[i]

    # Cabeçalhos (#, ##, ###)
    if line.strip().startswith('#'):
        level = 0
        while level < len(line) and line[level] == '#':
            level += 1
        heading_text = line[level:].strip()
        if heading_text:
            doc.add_heading(heading_text, level=level)
        i += 1
        continue

    # Tabelas Markdown
    if '|' in line and i + 1 < len(lines) and re.match(r'^\s*[\|\:\-\s]+\s*$', lines[i + 1]):
        header_line = line
        j = i + 2
        body_lines = []
        while j < len(lines) and '|' in lines[j]:
            body_lines.append(lines[j])
            j += 1

        def split_cells(row):
            parts = row.strip().strip('|').split('|')
            return [cell_text.strip() for cell_text in parts]

        header_cells = split_cells(header_line)
        col_count = len(header_cells)
        row_count = len(body_lines)
        table = doc.add_table(rows=row_count + 1, cols=col_count)
        table.style = 'Table Grid'

        for ci, cell_text in enumerate(header_cells):
            cell = table.cell(0, ci)
            cell_para = cell.paragraphs[0]
            parts = re.split(r'(\*\*[^*]+\*\*)', cell_text)
            for part in parts:
                if part.startswith('**') and part.endswith('**'):
                    cell_para.add_run(part[2:-2]).bold = True
                else:
                    cell_para.add_run(part)

        for ri, body_line in enumerate(body_lines, start=1):
            cells = split_cells(body_line)
            for ci, cell_text in enumerate(cells):
                cell = table.cell(ri, ci)
                cell_para = cell.paragraphs[0]
                parts = re.split(r'(\*\*[^*]+\*\*)', cell_text)
                for part in parts:
                    if part.startswith('**') and part.endswith('**'):
                        cell_para.add_run(part[2:-2]).bold = True
                    else:
                        cell_para.add_run(part)
        i = j
        continue

    # Listas com marcadores
    if line.strip().startswith('- ') or line.strip().startswith('* '):
        item_text = line.strip()[2:].strip()
        doc.add_paragraph(item_text, style='List Bullet')
        i += 1
        continue

    # Listas numeradas
    if re.match(r'^\d+\.\s', line.strip()):
        item_text = re.sub(r'^\d+\.\s*', '', line.strip())
        doc.add_paragraph(item_text, style='List Number')
        i += 1
        continue

    # Linha horizontal (---)
    if re.match(r'^\s*[-_*]{3,}\s*$', line):
        hr_para = doc.add_paragraph()
        p = hr_para._p
        pPr = p.get_or_add_pPr()
        pBdr = OxmlElement('w:pBdr')
        bottom = OxmlElement('w:bottom')
        bottom.set(qn('w:val'), 'single')
        bottom.set(qn('w:sz'), '6')
        bottom.set(qn('w:space'), '1')
        bottom.set(qn('w:color'), 'auto')
        pBdr.append(bottom)
        pPr.append(pBdr)
        i += 1
        continue

    # Linha vazia
    if line.strip() == "":
        doc.add_paragraph("")
        i += 1
        continue

    # Parágrafo comum com negrito
    paragraph = doc.add_paragraph()
    parts = re.split(r'(\*\*[^*]+\*\*)', line)
    for part in parts:
        if part.startswith('**') and part.endswith('**'):
            paragraph.add_run(part[2:-2]).bold = True
        else:
            paragraph.add_run(part)
    i += 1
```

> Esta seção é o coração do script. Ela trata o conteúdo Markdown dinamicamente, garantindo a conversão adequada para `.docx` com formatação ABNT.

---

### 6. Salvar o Documento na Pasta `output/` com Timestamp no Nome

Para evitar sobrescrever documentos e organizar os arquivos gerados, o script salva o `.docx` final na pasta `output/`, utilizando um nome baseado na data e hora da execução.

#### Lógica de Implementação:
- Garante que a pasta `output/` exista (cria se necessário)
- Gera o nome do arquivo com base no timestamp atual (formato `YYYYMMDD_HHMM`)
- Verifica se já existe um arquivo com o mesmo nome
- Se sim, adiciona um sufixo numérico incremental (`_1`, `_2`, etc.)
- Salva o arquivo com tratamento de exceções (`try/except`) para garantir robustez

#### Código Correspondente:

```python
os.makedirs('output', exist_ok=True)
timestamp = datetime.now().strftime("%Y%m%d_%H%M")
base_name = f"relatorio_{timestamp}"
filename = base_name + ".docx"
filepath = os.path.join("output", filename)
counter = 1

# Se o arquivo já existir, acrescentar um sufixo numérico
while os.path.exists(filepath):
    filename = f"{base_name}_{counter}.docx"
    filepath = os.path.join("output", filename)
    counter += 1

try:
    doc.save(filepath)
    print(f"Documento salvo com sucesso: {filepath}")
except Exception as e:
    print(f"Falha ao salvar o documento: {e}")
```

> Esse padrão garante que nenhum documento seja sobrescrito acidentalmente, mantendo um histórico de versões com nomes únicos e organizados.

---

## ❌ Limitações atuais
- Imagens em Markdown (`![]()`) ainda não são processadas
- Não suporta itálico (`*texto*`)
- Código em bloco (```) não é tratado
- O sumário precisa ser atualizado manualmente no Word

---

## 📆 Futuras melhorias
- Suporte a **imagens e links**
- Conversão completa de listas aninhadas
- Exportação também para `.pdf`
- Interface web via Swagger para colar o markdown

---

Para mais informações sobre o projeto, consulte o `README.md` na raiz ou abra um [issue](https://github.com/BrayanPletsch/ConvertMarkdownToDocx/issues).

> Documentação gerada para a versão inicial do script `app.py`