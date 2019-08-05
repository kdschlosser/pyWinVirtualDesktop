# -*- coding: utf-8 -*-
from __future__ import print_function
import msvc
from setuptools import setup, find_packages, Extension
import os
import sys


print(msvc.environment)

def_template = '''\
LIBRARY DLLMAIN
EXPORTS
    {exports}
'''
def_file = os.path.join('src', 'dllmain.def')

if sys.version_info[0] >= 3:
    exports = 'PyInit_libWinVirtualDesktop'
else:
    exports = 'initlibWinVirtualDesktop'

with open(def_file, 'w') as f:
    f.write(def_template.format(exports=exports))


libWinVirtualDesktop = Extension(
    'libWinVirtualDesktop',
    sources=[
        os.path.join('src', 'dllmain.cpp'),
        os.path.join('src', 'stdafx.cpp'),
    ],
    include_dirs=['src'] + msvc.environment.python.includes,
    library_dirs=(
        msvc.environment.windows_sdk.lib +
        msvc.environment.python.libraries
    ),
    extra_link_args=[
        '/def:"' + def_file + '"'
    ],
    extra_compile_args=[
        '/Gy',
        '/O2',
        '/Gd',
        '/Oy',
        '/Oi',
        '/fp:precise',
        '/Zc:wchar_t',
        '/Zc:forScope',
        '/EHsc',
    ],
    define_macros=[
        ('WIN64', 1),
        ('NDEBUG', 1),
        ('_WINDOWS', 1),
        ('_USRDLL', 1),
        ('LIBWINVIRTUALDESKTOP_EXPORTS', 1),
    ],
    libraries=[
        'Rpcrt4',
        'Ole32',
        'User32',
        msvc.environment.python.dependency[:-4]
    ]
)


setup(
    name='pyWinVirtualDesktop',
    author='Kevin Schlosser',
    author_email='kevin.g.schlosser@gmail.com',
    version='0.1.0',
    zip_safe=False,
    ext_modules=[libWinVirtualDesktop],
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
