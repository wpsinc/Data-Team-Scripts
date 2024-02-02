from rich.console import Console
from rich.table import Table
from rich.progress import track
from prompt_toolkit import prompt
from prompt_toolkit.completion import WordCompleter
import subprocess
import time
import os 

def main():
    console = Console()

    current_dir = os.path.dirname(os.path.abspath(__file__))

    scripts = {
        'EOM_MegaReport': (os.path.join(current_dir, "..", "project", "End Of Month", "MegaReport.py"), "End of month mega report transformation and creation/edit"),
        'EOM_Directory_Cleanup': (os.path.join(current_dir, "..", "project", "london_eom_reporting", "reportDirCleanup.py"), "Transfers and renames files in a directory to a new directory structure for their respective date created"),
        'Catalog_Photo_File_Renamer': (os.path.join(current_dir, "..", "project", "catalog", "photo_file_renamer.py"), "Renames files in a directory to a new naming convention"),
    }

    script_completer = WordCompleter(scripts.keys(), ignore_case=True)

    console.print("Welcome to the script runner dashboard!", style="bold red")
    console.print("Here are the available scripts:", style="bold underline")

    table = Table(show_header=True, header_style="bold magenta")
    table.add_column("Script")
    table.add_column("Description")
    for script, (path, description) in scripts.items():
        table.add_row(script, description)
    console.print(table)

    selected_script = prompt("Which script do you want to run?\n", completer=script_completer)

    console.print(f"Running {selected_script}...", style="bold green")

    # Run the selected script
    subprocess.run(f'python "{scripts[selected_script][0]}"', shell=True)

    console.print(f"{selected_script} completed.", style="bold blue")

if __name__ == '__main__':
    main()