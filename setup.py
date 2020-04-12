from setuptools import find_packages, setup

setup(
    name="stock_analyze",
    version="1.0",
    packages=find_packages(),
    license="Private",
    description="Analyze indian stocks using screen.in excel sheet",
    author="sukhbinder",
    author_email="sukh2010@yahoo.com",
    entry_points={
        'console_scripts': [' analyze = stock_analyze.AnalyzeQuarters:main']
    }
)
