CompMDB

This program will allow the user to select a MDB file
to compact.  The size of the file is captured and a
calculation of twice that size is made to determine
the amount of free space required to compact the
database.  Half that amount is used for a backup copy
of the original database and the other half is for
the compacted database.  if there is not enough space,
the user is prompted to select another path in which
to perform this operation or leave the application.
After the database is compacted, the original is deleted
and the new version is moved back into the place of the
original.

This program now recognizes drives greater than 2gb and
you can use command line parameters to point to your
favorite database.


-----------------------------------------------------------------
Written by Kenneth Ives                    kenaso@home.com

All of my routines have been compiled with VB6 Service Pack 3.
There are several locations on the web to obtain these
modules.

Whenever I use someone else's code, I will give them credit.  
This is my way of saying thank you for your efforts.  I would
appreciate the same consideration.

Read all of the documentation within this program.  It is very
informative.  Also, if you learn to document properly now, you
will not be scratching your head next year trying to figure out
exactly what you were programming today.  Been there, done that.

This software is FREEWARE. You may use it as you see fit for 
your own projects but you may not re-sell the original or the 
source code. If you redistribute it you must include this 
disclaimer and all original copyright notices. 

No warranty express or implied, is given as to the use of this
program. Use at your own risk.

If you have any suggestions or questions, I'd be happy to
hear from you.
-----------------------------------------------------------------

 
 