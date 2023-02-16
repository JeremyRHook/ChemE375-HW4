# ChemE375-HW4
# ChemE 375: HW4

The assignment is due midnight **before** the beginning of the next class period.

Note: Use Excel for all problems.  Remember that on all Excel assignments, **some points will be given for following the recommended practice of naming variables and for presentation (formatting)**.  For each problem please use a separate worksheet and highlight your answers!

## Problem 1

Label a new worksheet "Problem 1". Record an Excel macro, "Time12" that changes the selected cells' font to "Times New Roman", size 12 and colors the cell light blue. Program it to run with the shortcut key combination Ctrl + Shift + T. Add a button labelled "Times12" which will run the macro.

## Problem 2

Label a worksheet "Problem 2" and complete the following. In HW2, you performed a bubble point calculation for the benzene-toluene system at a single set of y1/y2 values. Repeat this problem over the complete range of y1 values for Benzene (0 to 1) in 0.05 intervals at P = 120 kPa to construct a Txy diagram. In order to automate this process, use a for loop within a macro to avoid running Solver 21 times by hand. After calculating all the temperatures, plot the Txy diagram for the benzene-toluene system.

*Hint*: Solver will not work by default in a VBA macro. You will need to activate it by going to Tools --> References --> Solver within the VBA interface.
