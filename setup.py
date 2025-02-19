from setuptools import find_packages, setup

setup(
    name="stephen-king-parser",
    version="3.0.2",
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
