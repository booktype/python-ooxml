from setuptools import setup

setup(
    name = 'Python-OOXML',
    version = '0.12',

    author = 'Aleksandar Erkalovic',
    author_email = 'aerkalov@gmail.com',

    packages = ['ooxml'],

    url = 'https://github.com/booktype/python-ooxml',

    license = 'GNU Affero General Public License',
    description = 'Python-OOXML is a Python library for parsing Office Open XML files.',
    long_description = open('README.rst').read(),
    keywords = ['ooxml', 'docx'],

    classifiers = [
        "Development Status :: 4 - Beta",
        "Intended Audience :: Developers",
        "Operating System :: OS Independent",
        "Programming Language :: Python :: 2.7",
        "Topic :: Software Development :: Libraries :: Python Modules"
    ],

    install_requires = [
       "lxml", "six"
    ],

    test_suite = "tests"
)
