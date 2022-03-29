# About U2C
A pre-processor for KLayout DRC engine and Mentor Calibre.  
For Ux Corporation Internal Boolean Opeartion Code for Auto-generation.

# Manual
## WorkSheet
+ LayerTable
  + Layer map and data type setting for databases.
+ CtrlTable
  + Boolean Operation Code in original format thru Ux MES.
+ MES
  + Boolean Operation Code separate by rows without any translation.
+ Result
  + Final deck format for Calibre or KLayout. --> Save as *.txt to be run in Calibre or KLayout.
+ NewTable
  + Boolean Operation Code in Original Format which merged from MES worksheet.

## Command Bar
+ Initial
  + Separate the original format(CtrlTable) into separate rows(MES).
+ Execute(KLayout)
  + Translate Codes into KLayout DRC engine codes.
+ Execute(Calibre)
  + Translate Codes into Calibre DRC deck.
+ MergeRows
  + Merge separate rows into original ofrmat thru Ux MES.
+ Scaling
  + Operate on Result sheet with a shrinking ratio. 
+ Version. 1.13
  + Show version informaition

# Copyright 
@author: [Rain Hu](https://intervalrain.github.io/posts/aboutme/)  
@email: [intervalrain@gmail.com](intervalrain@gmail.com)  
@github: [https://github.com/intervalrain](https://github.com/intervalrain)  
@website: [https://intervalrain.github.io](https://intervalrain.github.io)

# Version
+ Version 1.00: Macro released.  
+ Version 1.01: Revise "GROW" operator.  
+ Version 1.02: Add scaling function in macro list.  
+ Version 1.03: Add "SHRINK" operator.  
+ Version 1.04: Fix "SHRINK" and "GROW".  
+ Version 1.05: Fix compling problem.  
+ Version 1.06: Print out errors at one time.  
+ Version 1.07: Add "AREA", "HOLES", "RECTANGLE" opearator.  
+ Version 1.08: Adjust "Scaling" function for "AREA" operator.  
+ Version 1.09: Fix "SHRINK" and "GROW" um/side to um.  
+ Version 1.10: Fix BEoL layer fail issue.  
+ Version 1.11: Fix "INTERACT" and "NOT INTERACT".  
+ Version 1.12: Fix Calibre deck problem.  
+ Version 1.13: Fix DRCS number capital problem.  
