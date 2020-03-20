#!/usr/bin/python

from setuptools import setup
import sys

if sys.version_info < (3,):
    raise RuntimeError("xlsx-reader requires Python 3 or higher")

setup(
    name='xlsx-reader',
    version="0.1",
    description="Read very large Excel XLSX files efficiently",
    author='David Megginson',
    author_email='megginson@un.org',
    install_requires=[],
    packages=['xlsxr'],
    test_suite='tests'
)
