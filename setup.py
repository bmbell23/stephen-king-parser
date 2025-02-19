from setuptools import setup, find_packages

setup(
    name="stephen-king-parser",
    version="0.1.0",
    packages=find_packages(where="src"),
    package_dir={"": "src"},
    install_requires=[
        line.strip()
        for line in open("requirements.txt")
        if line.strip() and not line.startswith("#")
    ],
    entry_points={
        "console_scripts": [
            "stephen-king-parser=stephen_king_parser.cli:main",
        ],
    },
    python_requires=">=3.8",
)
