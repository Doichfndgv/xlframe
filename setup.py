from setuptools import setup

__major__ = 0
__minor__ = 0
__release__ = 5

__version__ = '{}.{}.{}'.format(__major__, __minor__, __release__)
__desc__ = 'A Python library for styling dataframes when exporting to excel.'

setup(
    name='xlframe',
    version=__version__,
    author='Doichfndgv',
    author_email='Doichfndgv@gmail.com',
    description=__desc__,
    url='https://github.com/Doichfndgv/xlframe',
    packages=[
        'xlframe',
    ],
    install_requires=[
        'pandas>=0.21.0',
        'openpyxl>=2.4.1',
    ],
    classifiers=[
        'Development Status :: 4 - Beta',
        'Intended Audience :: Developers',
        'License :: OSI Approved :: MIT License',
        'Programming Language :: Python :: 3',
        'Programming Language :: Python :: 3.5',
        'Programming Language :: Python :: 3.6',
        'Programming Language :: Python :: 3.7',
    ],
)


if __name__ == '__main__':
    pass
