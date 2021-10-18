#!/usr/bin/env python
import analizatorkotow.guilib
# from multiprocessing import freeze_support
#
# if __name__ == "__main__":
#     freeze_support() #Keeps windows binary working for threading code
#     analizatorkotow.guilib.mainGui()
import sys
from os.path import expanduser
from multiprocessing import freeze_support
# home = expanduser("~")
# sys.stdout = open(home+'/.pies.log', 'w')
# sys.stderr = open(home+'/.kot.log', 'w') #to prevent program from logging.
if __name__ == "__main__":
    freeze_support()

    analizatorkotow.guilib.mainGui()