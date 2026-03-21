"""Console I/O and formatting utilities."""


def ask_yes_no(prompt: str, lang: dict) -> bool:
    """Ask a yes/no question and return the boolean result."""
    while True:
        answer = input(prompt + " ").strip().lower()
        if answer in [lang["yes"], lang["no"]]:
            return answer == lang["yes"]


def print_green(text: str):
    """Print text in green color."""
    print(f"\033[92m{text}\033[0m")


def print_yellow(text: str):
    """Print text in yellow color."""
    print(f"\033[93m{text}\033[0m")


def print_red(text: str):
    """Print text in red color."""
    print(f"\033[91m{text}\033[0m")
