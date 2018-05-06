"""
setup.py
"""
from setuptools import setup

setup(**{
    'name': 'md2docx',
    'version': '0.9.0',
    'author': 'Ilya Voronin',
    'author_email': 'ivoronin@gmail.com',
    'url': 'https://github.com/ivoronin/md2docx',
    'packages': ['md2docx', 'md2docx.styles'],
    'entry_points': {
        'console_scripts': [
            'md2docx = md2docx:main'
        ],
    },
    'install_requires': ['python-docx>=0.8.6', 'mistune>=0.8.3']
})
