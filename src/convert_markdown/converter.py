from .utils.markdown_utils import remove_citation_tokens


def clean_markdown(path_md: str) -> str:
    with open(path_md, encoding="utf-8") as f:
        raw = f.read()
    cleaned = remove_citation_tokens(raw)
    return cleaned