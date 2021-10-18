#!/usr/bin/env python
import analizatorkotow.guilib
from multiprocessing import freeze_support

if __name__ == "__main__":
     freeze_support() #Keeps windows binary working for threading code
     analizatorkotow.guilib.mainGui()

