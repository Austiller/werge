from setuptools import setup, find_packages
from pip._internal.req import parse_requirements

setup(
    name="Werge",
    version="0.0.61",
    description="A library used to parse word files to a json structure that can be converted to PDF files",
    author="Austin Miller",
    author_email="Werge@Austins.site",
    license="GNU General Public License, version 2",
    packages=find_packages(),
    package_data={"werge":["config/.*."]},
    include_package_data=True,
    install_requires = ["pandas>=0.25.3","reportlab>=3.5.65","PyPdf2>=1.26.0","python-docx>=0.8.10"]



)