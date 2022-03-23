#!/usr/bin/env python
# coding: utf-8

#python setup.py sdist build
#twine upload dist/*

from setuptools import setup
from pathlib import Path

this_directory = Path(__file__).parent
long_description = (this_directory / "README.md").read_text()

setup(
    name="pptxtpl",
    version="0.0.4",
    author="kalavt",
    author_email="unknow@unknow.com",
    url="https://github.com/kalavt/Pptxtpl",
    description="generate pptx file with jinja2 template",
    long_description=long_description,
    packages=["pptxtpl"],
    install_requires=["python-pptx", "jinja2"],
    entry_points={"console_scripts": []},
)
