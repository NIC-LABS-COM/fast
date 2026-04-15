"""Entrypoint para build do .exe DEV via PyInstaller."""

import os
os.environ["ENV_FILE"] = ".env.dev"

from consumer.main import main

if __name__ == "__main__":
    main()
