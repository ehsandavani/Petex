# Petex
Code for running Petex OpenServer commands in Python

### Prerequisites

To run OpenServer commands in Python you need to install the pywin32 extension (Provides access to much of the Win32 API, the ability to create and use COM objects, and the Pythonwin environment). Check out https://pypi.org/project/pywin32/

### Getting started

Download the PetexOpenServer.py file and import it to your Python file with following code:
```
from PetexOpenServer import *
```
then use standard OpenServer functionality.

## Example 

The following code will import the OpenServer module, start Prosper, open a file named C-2 on root drive and adding a comment into the comment section in Prosper.

```
from PetexOpenServer import *

DoCmd('PROSPER.START()')
DoCmd('PROSPER.OPENFILE("C:\C-2.OUT")')
DoSet('PROSPER.SIN.SUM.Comments', 'Testing OpenServer from Python')
```
