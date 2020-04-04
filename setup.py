import os
import sys
from setuptools import setup, find_packages

# python setup.py sdist bdist_wheel
# twine upload --repository-url https://test.pypi.org/legacy/ dist/*
# twine upload dist/*

path = os.path.abspath(os.path.dirname(__file__))
sys.path.append(path)

setup(
    name = "ppt2gif",
    version = "1.0.2",
    keywords = ["pip", "ppt", "gif"],
    description = "https://github.com/AdjWang/ppt2gif",
    long_description = "github 链接: https://github.com/AdjWang/ppt2gif 说明详见`README.md`",
    long_description_content_type='text/markdown',
    python_requires=">=3.5.0",
    license = "MIT Licence",

    url = "https://github.com/AdjWang/ppt2gif",
    author = "AdjWang",
    author_email = "491918260@qq.com",

    packages = ['ppt2gif'],
    include_package_data = True,
    install_requires = [
        "tqdm >= 4.28.1", 
        "imageio >= 2.4.1",
        "pypiwin32 >= 223"
    ],
    platforms = "Windows",
)