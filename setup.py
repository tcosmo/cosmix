import pathlib
import re

import pkg_resources
import setuptools


VERSIONFILE = "cosmix/_version.py"

with open(VERSIONFILE, "rt") as f:
    verstrline = f.read()
    VSRE = r"^__version__ = ['\"]([^'\"]*)['\"]"
    mo = re.search(VSRE, verstrline, re.M)
    if mo:
        verstr = mo.group(1)
    else:
        raise RuntimeError(f"Unable to find version string in {VERSIONFILE}.")

with pathlib.Path("requirements.txt").open() as requirements_txt:
    install_requires = [
        str(requirement)
        for requirement in pkg_resources.parse_requirements(requirements_txt)
    ]


with open("README.md", "r") as f:
    long_description = f.read()

setuptools.setup(
    name="cosmix-wetlab",
    version=verstr,
    author="Tristan Stérin",
    author_email="tristan.sterin@mu.ie",
    description="Utility to create wet-lab mixes and integrations with third-parties such as Google Sheets",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/tcosmo/cosmix",
    packages=setuptools.find_packages(),
    include_package_data=True,
    install_requires=install_requires,
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ],
)
