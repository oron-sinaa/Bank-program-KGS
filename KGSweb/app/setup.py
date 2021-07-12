from setuptools import setup

setup(
    name='KGSweb',
    version='1.0',
    author='Aanis Noor',
    packages=['KGSweb'],
    include_package_data=True,
    install_requires=[
        'flask',
        'flask_wtf',
        'werkzeug',
        'pandas',
    ],
)