"""
Setup script for Excel MCP Plugin
"""

from setuptools import setup, find_packages

with open("README.md", "r", encoding="utf-8") as fh:
    long_description = fh.read()

setup(
    name="excel-mcp-plugin",
    version="1.0.0",
    author="Excel MCP Plugin Team",
    description="Excel plugin for connecting to Model Context Protocol (MCP) servers",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/yourusername/excel-mcp-plugin",
    packages=find_packages(),
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
        "Development Status :: 4 - Beta",
        "Intended Audience :: End Users/Desktop",
        "Topic :: Office/Business :: Financial :: Spreadsheet",
    ],
    python_requires=">=3.7",
    install_requires=[
        "xlwings>=0.30.0",
        "openpyxl>=3.1.0",
    ],
    entry_points={
        "console_scripts": [
            "excel-mcp-plugin=mcp_plugin:main",
        ],
    },
)
