# -*- coding: UTF-8 -*-
#
# Copyright 2011-2019 by Dirk Gorissen, Stephen Rauch and Contributors
# All rights reserved.
# This file is part of the Pycel Library, Licensed under GPLv3 (the 'License')
# You may not use this work except in compliance with the License.
# You may obtain a copy of the Licence at:
#   https://www.gnu.org/licenses/gpl-3.0.en.html

"""
Simple example file showing how a spreadsheet can be translated to python
and executed
"""
import logging
import os
import sys
import time

from pycel import ExcelCompiler


def pycel_logging_to_console(enable=True):
    if enable:
        logger = logging.getLogger('pycel')
        logger.setLevel('INFO')

        console = logging.StreamHandler(sys.stdout)
        console.setLevel(logging.INFO)
        logger.addHandler(console)


if __name__ == '__main__':
    # pycel_logging_to_console()

    path = os.path.dirname(__file__)
    fname = os.path.join(path, "big-sheet.xlsx")

    print(f"Loading sheet with pycel {fname}...")

    pycel_start = time.time()

    # load & compile the file to a graph
    excel = ExcelCompiler(filename=fname)

    # test evaluation
    print(f"A2 is {excel.evaluate('Test!A2')}")
    print(f"B2 is {excel.evaluate('Test!B2')}")

    print("Setting A2 to new value")
    excel.set_value('Test!A2', "10.372")

    print(f"A2 is {excel.evaluate('Test!A2')}")
    print(f"B2 is {excel.evaluate('Test!B2')}")

    pycel_end = time.time()

    print("Done. Took: ", pycel_end - pycel_start, " seconds")
