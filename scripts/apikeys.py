"""Read API keys from the project's .env file. Keys are never printed or written anywhere else."""
import os
import re

ENV = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), ".env")


def key(name):
    if os.environ.get(name):
        return os.environ[name]
    if os.path.exists(ENV):
        for line in open(ENV, encoding="utf-8"):
            m = re.match(r"\s*([A-Z_]+)\s*=\s*(.*?)\s*$", line)
            if m and m.group(1) == name and m.group(2):
                return m.group(2).strip().strip('"').strip("'")
    raise SystemExit(f"{name} is not set in .env")
