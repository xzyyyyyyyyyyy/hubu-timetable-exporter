import os
import sys


def get_base_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def resolve_input_path(filename):
    base_dir = get_base_dir()
    candidates = [
        os.path.join(os.getcwd(), filename),
        os.path.join(base_dir, filename),
        os.path.join(base_dir, "..", filename),
    ]
    for path in candidates:
        if os.path.exists(path):
            return os.path.abspath(path)
    raise FileNotFoundError(
        f"未找到文件: {filename}。请将其放在 {base_dir} 或其上一级目录。"
    )


def resolve_output_path(filename):
    base_dir = get_base_dir()
    return os.path.abspath(os.path.join(base_dir, filename))
