import setuptools

setuptools.setup(
    name="its_easy",
    version="0.1.0",
    url="https://github.com/soonraah/its_easy",

    author="soonraah",
    author_email="soonraah.dev@gmail.com",

    description="A Python package to make ITS tour procedure easy.",
    long_description=open('README.rst').read(),

    packages=setuptools.find_packages(),

    install_requires=[],

    classifiers=[
        'Development Status :: 2 - Pre-Alpha',
        'Programming Language :: Python',
        'Programming Language :: Python :: 3.6',
    ],
)
