"""
Configuration Loader for audit note filler engine.
Loads all YAML configuration files from the specified directory.

Supports PyYAML (preferred) and ruamel.yaml as fallback.
"""

import sys
from pathlib import Path

try:
    import yaml
except ImportError:
    try:
        from ruamel import yaml
    except ImportError:
        print("ERROR: No YAML library found. Install pyyaml: pip install pyyaml")
        sys.exit(1)


def load_config(config_dir="config"):
    """
    Load all YAML configuration files from the specified directory.

    Args:
        config_dir: Path to config directory (relative to CWD or absolute)

    Returns:
        dict with keys: 'mappings', 'name_synonyms', 'semantic_tables',
                        'cross_validations', 'concept_map', 'auto_sum_rules'
    """
    config_path = Path(config_dir)
    if not config_path.exists():
        print(f"ERROR: Config directory not found: {config_path.resolve()}")
        sys.exit(1)
    if not config_path.is_dir():
        print(f"ERROR: Config path is not a directory: {config_path.resolve()}")
        sys.exit(1)

    config_files = {
        "mappings": "mappings.yaml",
        "name_synonyms": "name_synonyms.yaml",
        "semantic_tables": "semantic_tables.yaml",
        "cross_validations": "cross_validations.yaml",
        "concept_map": "concept_map.yaml",
        "auto_sum_rules": "auto_sum_rules.yaml",
    }

    config = {}

    for key, filename in config_files.items():
        filepath = config_path / filename
        if not filepath.exists():
            print(f"WARNING: Config file not found: {filepath}")
            # Use appropriate default type based on config category
            if key in (
                "name_synonyms",
                "semantic_tables",
                "concept_map",
                "auto_sum_rules",
            ):
                config[key] = {}
            else:
                config[key] = []
            continue

        try:
            with open(filepath, "r", encoding="utf-8") as f:
                data = yaml.safe_load(f)
            if data is None:
                data = (
                    {}
                    if key
                    in (
                        "name_synonyms",
                        "semantic_tables",
                        "concept_map",
                        "auto_sum_rules",
                    )
                    else []
                )
            config[key] = data
        except Exception as e:
            print(f"ERROR: Failed to load {filepath}: {e}")
            sys.exit(1)

    return config
