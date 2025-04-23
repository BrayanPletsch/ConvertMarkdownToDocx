from convert_markdown.converter import clean_markdown


def test_clean_markdown_basic(tmp_path):
    md_file = tmp_path / "sample.md"
    raw_content = (
        "Hello citeturn1 world .  Next.\n"
        "Another line foo123.   End."
    )
    md_file.write_text(raw_content, encoding="utf-8")

    cleaned = clean_markdown(str(md_file))

    expected = (
        "Hello world. Next.\n"
        "Another line. End."
    )
    assert cleaned == expected


def test_clean_markdown_preserves_content_without_tokens(tmp_path):
    md_file = tmp_path / "raw.md"
    raw_content = (
        "Just a normal line.\n"
        "Second line without tokens."
    )
    md_file.write_text(raw_content, encoding="utf-8")

    cleaned = clean_markdown(str(md_file))
    assert cleaned == raw_content