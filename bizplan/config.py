from pathlib import Path
import yaml

DEFAULT_MAPPINGS_PATH = Path("config/mappings.yaml")
DEFAULT_MODEL_PATH = Path("config/model.yaml")

def load_mappings(path: Path = DEFAULT_MAPPINGS_PATH) -> dict:
    if not path.exists():
        return {}
    return yaml.safe_load(path.read_text()) or {}

def save_mappings(mappings: dict, path: Path = DEFAULT_MAPPINGS_PATH) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(yaml.dump(mappings, default_flow_style=False, sort_keys=True))

def load_model(path: Path = DEFAULT_MODEL_PATH) -> dict:
    if not path.exists():
        return {}
    return yaml.safe_load(path.read_text()) or {}

def save_model(model: dict, path: Path = DEFAULT_MODEL_PATH) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(yaml.dump(model, default_flow_style=False, sort_keys=True))
