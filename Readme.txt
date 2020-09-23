Original Source by: Scott Pierce (webmaster@calclinks.net)
Date: April 02, 2000

This is free source code for you to openly use.  It is provided without warranty and 
the author will not be held responsible for any possible problems which may arise.
I only ask you give credit where credit is due and feel free to e-mail me if you found this source helpful.


Know problems:
   If you try to open too many connections simutaniously (more than 80 winsock controls)
   the computer will run out of buffer space.  As far as I can tell, the buffer space
   depends on the number of open programs, so the less programs you have running, the 
   more connections you can open.

   Also, there is no error checking in this code.  It was simply to show how to make a 
   simple port scanner using the MS Winsock control.