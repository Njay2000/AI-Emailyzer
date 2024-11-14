from setuptools import setup, find_packages

setup(
    name="AI_EXCEL_GENERATOR",
    version="0.1",
    packages=find_packages(),
    install_requires=[
        "pandas",
        "requests",
        "loguru",
        "msal",
        "python-dotenv",
        "openpyxl",
        "pytest",
        "pytest-cov",
        "pytest-html",
        "black",
        "pytest-mock" "openai",
    ],
)
