import re


def remove_citation_tokens(text: str) -> str:
    text = re.sub(r'.*?', '', text)
    text = re.sub(r'[ ]+\.', '.', text)
    text = re.sub(r'\. {2,}', '. ', text)
    text = re.sub(r' {2,}', ' ', text)
    return text


def add_markdown_runs(paragraph, text: str):
    parts = re.split(r'(\*\*[^*]+\*\*)', text)
    for part in parts:
        if not part:
            continue
        if part.startswith('**') and part.endswith('**'):
            content = part[2:-2]
            run = paragraph.add_run(content)
            run.bold = True
        else:
            paragraph.add_run(part)