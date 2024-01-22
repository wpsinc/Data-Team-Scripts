## Data Team Misc Scripts

> This repository contains miscellaneous Python scripts used for unique cases on an as needed basis

### Getting Started

Follow these steps to get started with this project:

1. **Clone the repository**: Use the following command to clone this repository to your local machine:
    
```bash
git clone https://github.com/wpsinc/Data-Team-Scripts.git
```

2. Create a virtual environment: It's recommended to create a virtual environment to isolate the dependencies of this project. If you're using venv, you can create a virtual environment using the following command:

```bash
python3 -m venv /path/to/new/virtual/environment
```

Then, activate the virtual environment. On Windows, use:
`/path/to/new/virtual/environment/Scripts/activate`

On Unix or MacOS, use:  
`source /path/to/new/virtual/environment/bin/activate`

3. Install the requirements: This project has some dependencies which are listed in the requirements.txt file. After activating the virtual environment, install these dependencies using the following command:
    
```bash
pip install -r requirements.txt
```

3.1 (Alternative) Install all dependencies and run any of the dedicated scripts in the terminal without having to navigate to the file itself
    
```bash
pip install -e . # This installs all dependencies required for any of the scripts
```
Available Commands:
```bash
EOM_MegaReport
EOM_Directory_Cleanup
```
or
```bash
python terminal_dashboard\run_scripts.py # Starts an interactive dashboard to run any of the scripts
```

4. Create your own branch: It's recommended to create your own branch to work on for changes you are applying to the repository. It is also recommended to create a branch dedicated to a major feature implementation to isolate large scale changes. You can create a new branch using the following command:
    
```bash
git checkout -b <branch-name>
```

5. Make your changes: Make any changes you want to make to the project. Once you're done, you can commit your changes using the following commands:
    
```bash
git add .
git commit -m "Your commit message"
```

6. Push your changes: Push your changes to the remote repository using the following command:
    
```bash
git push origin <branch-name>
```

7. Create a pull request: Create a pull request to merge your changes into the master branch. Once your pull request is approved, you can merge your changes into the master branch.


### Project List:

- london_eom_reporting: Monthly reporting done for specified brands completed by London
- monthly_sales_reporting: Waiting for Description
- transform_scripts: Used to transform data based on variable file schema structures


Author: London Perry
Date: 11/13/2023
