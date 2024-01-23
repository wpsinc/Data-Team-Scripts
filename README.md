# Data Team Misc Scripts

This repository contains miscellaneous Python scripts used for unique cases on an as needed basis.

## Table of Contents

- [Getting Started](#getting-started)
  - [Prerequisites](#prerequisites)
  - [Installation](#installation)
- [Project List](#project-list)
- [Author](#author)

## Getting Started

These instructions will get you a copy of the project up and running on your local machine.

### Prerequisites

- Git
- Python 3
- pip
- venv (optional)

### Installation

1. **Clone the repository**:

    ```bash
    git clone https://github.com/wpsinc/Data-Team-Scripts.git
    ```

2. **Create and activate a virtual environment** (optional):

    ```bash
    python3 -m venv /path/to/new/virtual/environment
    ```

    On Windows, activate the virtual environment with:

    ```bash
    /path/to/new/virtual/environment/Scripts/activate
    ```

    On Unix or MacOS, use:

    ```bash
    source /path/to/new/virtual/environment/bin/activate
    ```

3. **Install the requirements**:

    ```bash
    pip install -r requirements.txt
    ```

    Alternatively, install all dependencies and run any of the dedicated scripts in the terminal without having to navigate to the file itself:

    ```bash
    pip install -e .
    ```

    Available Commands:

    ```bash
    EOM_MegaReport
    EOM_Directory_Cleanup
    ```

    or

    ```bash
    python terminal_dashboard\run_scripts.py
    ```

4. **Create your own branch**:

    ```bash
    git checkout -b <branch-name>
    ```

5. **Make your changes**:

    ```bash
    git add .
    git commit -m "Your commit message"
    ```

6. **Push your changes**:

    ```bash
    git push origin <branch-name>
    ```

7. **Create a pull request**: Create a pull request to merge your changes into the master branch. Once your pull request is approved, you can merge your changes into the master branch.

## Project List

- `london_eom_reporting`: Monthly reporting done for specified brands completed by London
- `End of Month`: Waiting for Description
    - **MegaReport.py**: Generates monthly MegaReport that then updates existing MegaReport file in the shared drive
- `transform_scripts`: Used to transform data based on variable file schema structures

## Author

- London Perry
- Date: 11/13/2023