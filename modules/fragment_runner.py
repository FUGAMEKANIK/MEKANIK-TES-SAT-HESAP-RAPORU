from pathlib import Path

def run_fragment(fragment_name, namespace):
    path = Path(__file__).resolve().parent / fragment_name
    source = path.read_text(encoding="utf-8")
    code = compile(source, str(path), "exec")
    exec(code, namespace, namespace)
