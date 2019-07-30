# -*- coding: utf-8 -*-

from setuptools import setup, find_packages

setup(
    name='pyWinVirtualDesktop',
    author='Kevin Schlosser',
    author_email='kevin.g.schlosser@gmail.com',
    version='0.1.0b',
    zip_safe=False,
    dependency_links=[
        'https://github.com/kdschlosser/comtypes/tarball'
        '/zip_safe#egg=comtypes'
    ],
    url='https://github.com/kdschlosser/pyWinVirtualDesktop',
    install_requires=['six', 'comtypes'],
    packages=['pyWinVirtualDesktop'] + find_packages('pyWinVirtualDesktop'),
    package_dir=dict(
        pyWinVirtualDesktop='pyWinVirtualDesktop',
    ),
    description='Windows 10 Virtual Desktop Management.',
    long_description='Windows 10 Virtual Desktop Management.',
    keywords=['virtual desktop', 'windows 10'],
    classifiers=[
        "Operating System :: Microsoft :: Windows",
        "Programming Language :: Python :: 2",
        "Programming Language :: Python :: 3",
        (
            "License :: OSI Approved :: "
            "GNU General Public License v3 or later (GPLv3+)"
        ),
    ],
)
