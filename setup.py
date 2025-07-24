from setuptools import setup, find_packages

with open("README.md", "r", encoding="utf-8") as fh:
    long_description = fh.read()

setup(
    name="excelllm",
    version="0.1.0",
    author="ExcelLLM",
    description="Convert Excel files to LLM-friendly formats (JSON/Markdown)",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/yourusername/excelllm",
    packages=find_packages(),
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ],
    python_requires=">=3.6",
    install_requires=[
        "openpyxl>=3.0.0",
    ],
    entry_points={
        "console_scripts": [
            "excelllm=excelllm:main",
        ],
    },
)
