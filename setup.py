from setuptools import setup, find_packages

setup(
    name='DataTeamScripts',
    version='0.1',
    packages=find_packages(),
    install_requires=[
        # add your project dependencies here
    ],
    entry_points={
        'console_scripts': [
            'EOM_MegaReport=project.End_Of_Month.MegaReport:main',
            'EOM_Directory_Cleanup=project.london_eom_reporting.reportDirCleanup:main',
        ],
    }
)

