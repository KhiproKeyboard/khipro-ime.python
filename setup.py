from setuptools import setup, find_packages

setup(
    name="khipro-ime",
    version="1.0.0",
    packages=find_packages(),
    install_requires=[
        "pynput>=1.7.3",
        "pywin32>=300",
        "pystray>=0.19.0",
        "Pillow>=8.3.1",
        "winshell>=0.6"
    ],
    entry_points={
        "gui_scripts": [
            "khipro-ime = khipro_ime:main",
        ]
    },
)
