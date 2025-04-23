import sys
from pathlib import Path
import argparse
from datetime import datetime

from .converter import clean_markdown
from .docx_generator import generate_docx


def unique_filename(base: str, ext: str = ".docx") -> str:
    ts = datetime.now().strftime("%Y%m%d_%H%M")
    return f"{base}_{ts}{ext}"


def main():
    parser = argparse.ArgumentParser(
        description="Converte Markdown em DOCX (formatado em ABNT)."
    )
    parser.add_argument("input", help="arquivo .md em input/")
    parser.add_argument(
        "-o", "--output",
        help="nome opcional para o arquivo .docx (em output/).",
        default=None
    )
    args = parser.parse_args()

    input_path = Path("./") / args.input

    if input_path.suffix.lower() != ".md":
        parser.error(f"Apenas arquivos Markdown (*.md) são válidos: {input_path.name}")

    if not input_path.is_file():
        parser.error(f"Arquivo de entrada não encontrado: {input_path}")

    out_dir = Path("output")
    out_dir.mkdir(exist_ok=True)

    if args.output:
        out_name = args.output if args.output.lower().endswith(".docx") else args.output + ".docx"
    else:
        base = input_path.stem
        out_name = unique_filename(base)
        candidate = out_dir / out_name
        counter = 1
        while candidate.exists():
            out_name = f"{base}_{datetime.now().strftime('%Y%m%d_%H%M')}_{counter}.docx"
            candidate = out_dir / out_name
            counter += 1

    output_path = out_dir / out_name

    markdown = clean_markdown(str(input_path))

    try:
        generate_docx(markdown, str(output_path))
    except Exception as e:
        print(f"Falha ao gerar o documento: {e}", file=sys.stderr)
        sys.exit(1)
    else:
        print(f"Documento salvo em: {output_path}")
        sys.exit(0)


if __name__ == "__main__":
    main()