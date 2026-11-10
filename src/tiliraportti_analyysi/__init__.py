def main() -> None:
    import sys
    from pathlib import Path

    from streamlit.web import cli

    sys.argv = ["streamlit", "run", str(Path(__file__).with_name("app.py")),
                "--server.address=localhost", "--browser.gatherUsageStats=false", *sys.argv[1:]]
    cli.main()
