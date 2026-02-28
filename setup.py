from setuptools import setup, find_packages

setup(
    name="excel_agent",
    version="1.0.0",
    packages=find_packages(),
    install_requires=[
        "langgraph>=0.2.0",
        "langchain-openai>=0.1.0",
        "langchain-core>=0.2.0",
        "openpyxl>=3.1.0",
        "python-dotenv>=1.0.0",
    ],
)
