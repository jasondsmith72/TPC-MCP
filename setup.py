from setuptools import setup, find_packages

setup(
    name="tpc-mcp",
    version="0.1.0",
    description="Total PC Commander - Remote PC control using Model Context Protocol",
    author="Jason Smith",
    author_email="jason.smith@mtusa.com",
    packages=find_packages(),
    install_requires=[
        "mcp>=1.2.0",
        "pywin32>=303",
        "pillow>=9.0.0",
        "psutil>=5.9.0",
    ],
    entry_points={
        "console_scripts": [
            "tpc-mcp=tpc_server:main",
        ],
    },
    classifiers=[
        "Development Status :: 3 - Alpha",
        "Intended Audience :: System Administrators",
        "License :: OSI Approved :: MIT License",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.7",
        "Programming Language :: Python :: 3.8",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
        "Topic :: System :: Systems Administration",
    ],
    python_requires=">=3.7",
)
