import tomli

def load_toml_config(
    config_filename: str = "config.toml",
) -> dict:
    with open(config_filename, "rb") as f:
        return tomli.load(f)