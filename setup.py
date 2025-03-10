from setuptools import setup

setup(
    name="sapylot",
    version="0.0.1", 
    packages=['sapylot'],  
    install_requires=[  
        'psutil',
        'pywin32'  
    ],

    long_description_content_type="text/markdown",
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ],
    python_requires='>=3.6',
)
