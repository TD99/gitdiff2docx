"""Setup script for gitdiff2docx package."""

from setuptools import setup, find_packages
import os

# Read the contents of README file
with open("README.md", "r", encoding="utf-8") as f:
    long_description = f.read()

# Read requirements
with open("requirements.txt", "r", encoding="utf-8") as f:
    requirements = [line.strip() for line in f if line.strip() and not line.startswith("#")]

setup(
    name="gitdiff2docx",
    version="1.0.0",
    author="TD99",
    description="A Python utility that converts git diffs into formatted Word documents",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/TD99/gitdiff2docx",
    packages=find_packages(),
    package_data={
        "gitdiff2docx": [
            "config.json",
            "lang/*.json",
            "schemas/*.json",
        ],
    },
    install_requires=requirements,
    entry_points={
        "console_scripts": [
            "gdd=gitdiff2docx.main:main",
        ],
    },
    classifiers=[
        "Development Status :: 4 - Beta",
        "Intended Audience :: Developers",
        "Topic :: Software Development :: Version Control :: Git",
        "License :: OSI Approved :: MIT License",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.7",
        "Programming Language :: Python :: 3.8",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
        "Programming Language :: Python :: 3.11",
        "Programming Language :: Python :: 3.12",
        "Operating System :: OS Independent",
    ],
    python_requires=">=3.7",
    keywords="git diff docx word document converter",
)
