# docx_to_jsonl.py
"""
DOCX -> cleaned JSONL
Reads .docx (or folder of .docx), extracts text, splits into paragraphs,
cleans PII, outputs line-per-json: {"text": "...", "source": "file.docx", "para_idx": 3}
"""
import re, json, argparse
from pathlib import Path
import docx

PII_PATTERNS = [
    (re.compile(r'\b\d{12,19}\b'), '[CARD]'),
    (re.compile(r'\+?\d{7,15}'), '[PHONE]'),
    (re.compile(r'\S+@\S+\.\S+'), '[EMAIL]'),
]

def mask_pii(s: str):
    if not s: return s
    t = s
    for pat, repl in PII_PATTERNS:
        t = pat.sub(repl, t)
    return t

def read_docx(path: Path):
    doc = docx.Document(path)
    paras = [p.text.strip() for p in doc.paragraphs if p.text and p.text.strip()]
    return paras

def to_jsonl(docx_path, out_path):
    p = Path(docx_path)
    out = Path(out_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    count = 0
    with open(out, "w", encoding="utf-8") as fout:
        if p.is_file():
            files = [p]
        else:
            files = list(p.glob("**/*.docx"))
        for fp in files:
            paras = read_docx(fp)
            for i, para in enumerate(paras):
                text = mask_pii(para)
                entry = {"text": text, "source": str(fp.name), "para_idx": i}
                fout.write(json.dumps(entry, ensure_ascii=False) + "\n")
                count += 1
    print(f"Saved {count} paras -> {out}")

if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--in", dest="inp", required=True, help="docx file or folder")
    parser.add_argument("--out", dest="out", default="raw_docx.jsonl", help="output jsonl")
    args = parser.parse_args()
    to_jsonl(args.inp, args.out)
