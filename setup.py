from setuptools import setup, find_packages

setup(
    name="ukcs",
    version="0.1.0",
    packages=find_packages(),
    install_requires=[
        "streamlit>=1.31.0",
        "pandas>=2.2.0",
        "openpyxl>=3.1.2",
        "python-docx>=1.1.0",
    ],
) 