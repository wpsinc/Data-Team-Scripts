from setuptools import setup, find_packages

def main():
    with open('requirements.txt', 'r', encoding='utf-16') as f:
        required_packages = f.read().splitlines()

    setup_args = {
        'name': 'DataTeamScripts',
        'version': '0.1',
        'packages': find_packages(),
        'install_requires': required_packages,
        'entry_points': {
            'console_scripts': [
                'EOM_MegaReport=project.End_Of_Month.MegaReport:main',
                'EOM_Directory_Cleanup=project.london_eom_reporting.reportDirCleanup:main',
                'Catalog_Photo_File_Renamer=project.catalog.photo_file_renamer:main',
            ],
        }
    }

    setup(**setup_args)

if __name__ == '__main__':
    main()