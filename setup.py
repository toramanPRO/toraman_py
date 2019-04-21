import setuptools

with open("README.md", "r") as input_file:
    long_description = input_file.read()

setuptools.setup(
    name="toraman",
    version="0.0.1",
    author="Çağatay Onur Şengör",
    author_email="contact@csengor.com",
    description="A computer assisted translation tool package",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/csengor/toraman-py",
    packages=setuptools.find_packages(),
    install_requires=[
        "lxml",
        "regex"
    ],
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ],
)