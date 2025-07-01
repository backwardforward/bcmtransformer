from setuptools import setup, find_packages

setup(
    name='bcm_transformer',
    version='0.1.0',
    description='Business Capability Map Generator Web App',
    author='Markus Paszek',
    packages=find_packages(),
    include_package_data=True,
    install_requires=[
        'Flask>=2.0',
        'pandas>=1.3',
        'python-pptx>=0.6.21',
        'flask-cors>=3.0',
    ],
    entry_points={
        'console_scripts': [
            'bcm_transformer_app=bcm_transformer.app:main',
            'generate_presentation=bcm_transformer.generate_presentation:main',
        ],
    },
)
