#!/usr/bin/env python

from distutils.core import setup

setup(name="odswriter",
      version="0.2.0",
      description="A pure-Python module for writing OpenDocument spreadsheets (similar to csv.writer).",
      author="Michael Mulqueen",
      author_email="michael@mulqueen.me.uk",
      url="https://github.com/mmulqueen/odswriter",
      packages=["odswriter"],
      license="MIT License",
      classifiers=[
          "License :: OSI Approved :: MIT License",
          "Topic :: Office/Business :: Financial :: Spreadsheet",
          "Programming Language :: Python :: 2.7",
          "Programming Language :: Python :: 3.2",
          "Programming Language :: Python :: 3.3",
          "Programming Language :: Python :: 3.4",
          "Programming Language :: Python :: Implementation :: PyPy"
      ]
     )



