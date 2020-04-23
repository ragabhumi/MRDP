# MRDP: Monthly Report Data Processing

Overview
========

MRDP is a program to create routine geomagnetic monthly report at geomagnetic observatory.
It reads IAGA-2002 geomagnetic data format and write the report in Microsoft
Excel format. It is also can be used to remove step in magnetogram.


Requirements
===============================

Officially, MRDP requires Python version 3.7 and KASm installed on your computer.
Works better using Anaconda.
You can get KASm from <http://www.intermagnet.org/software/Kasm_1.09.zip>

Python module requirements :
- PyQt5
- Numpy
- Openpyxl
- Matplotlib
- Pillow
- Scipy
- Pandas
- Lxml


Instant running
===============================

You can run MRDP from this directory by clicking MRDP.bat directly or by creating desktop shortcut.
Adjust station.ini to change station parameters. 

