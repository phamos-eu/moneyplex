from setuptools import setup, find_packages

with open("requirements.txt") as f:
	install_requires = f.read().strip().split("\n")

# get version from __version__ variable in moneyplex/__init__.py
from moneyplex import __version__ as version

setup(
	name="moneyplex",
	version=version,
	description="Moneyplex integration",
	author="Phamos GmbH",
	author_email="wolfram.schmidt@phamos.eu",
	packages=find_packages(),
	zip_safe=False,
	include_package_data=True,
	install_requires=install_requires
)
