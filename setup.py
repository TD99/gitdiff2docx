"""Setup configuration for GitDiff2Docx."""

from setuptools import setup, find_packages

with open("README.md", "r", encoding="utf-8") as fh:
    long_description = fh.read()

with open("requirements.txt", "r", encoding="utf-8") as fh:
    requirements = [line.strip() for line in fh if line.strip() and not line.startswith("#")]

setup(
    name="gitdiff2docx",
    version="1.0.0",
    author="TD99",
    description="Convert git diffs to formatted Word documents",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/TD99/gitdiff2docx",
    packages=find_packages(),
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ],
    python_requires=">=3.7",
    install_requires=requirements,
    entry_points={
        "console_scripts": [
            "gitdiff2docx=gitdiff2docx.main:main",
        ],
    },
    include_package_data=True,
)
