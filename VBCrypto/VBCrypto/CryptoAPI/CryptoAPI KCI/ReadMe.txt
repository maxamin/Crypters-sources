CryptoAPI Demo v1.7
by Kenneth Ives  kenaso@home.com

-----------------------------------------------------------------
Modification History:

09-SEP-2001  1.7  Fixed several bugs relating to using the enhanced
                  provider and block ciphers.  Modified and tested
                  the demo program.

24-JUL-2001  1.6  Cleaned up some code and did a lot of documenting.         
             
05-JUN-2001  1.5  Removed clsCryptoAPI.cls from demo program and created
                  a DLL.

20-JAN-2001  1.4  According to theory, whenever you leave a text box, the
                  lost focus event is supposed to fire.  I came upon
                  multiple instances where it did not.  This would happen 
                  when I pressed the ENTER key while still inside the text 
                  box and executing the command button.  I decided to move 
                  the lost focus logic to the validate event and added a 
                  piece of code in the keypress event to force the 
                  validate event to fire.  I could had done the same with 
                  the lost focus but I had already moved my code.
                  
18-JAN-2001  1.3  The decoded file was be one byte larger than the 
                  source.  To fix this, subtract 1 from the file size 
                  to accomodate the zero based array.  Fix suggested 
                  by Harbinder Gill  hgill@altavista.net
                  
                  Also found that when you use PUT to write a byte array
                  to a file, the last character is converted to a NULL.
                  To get around this quirk, I converted the decrypted data
                  to a text string and then PUT it in the output file.
                  See frmEncFiles(cmdChoice_Click)

10-JAN-2001  1.2  Converted data to byte array and then encrypt/decrypt
                  the data.  For display purposes, I use a hex display
                  because if an encrypted character returned is a Null,
                  then I would end up with a null terminated string.
                  The text box control will not display anything after
                  the NULL character.  Therefore, when I would read from
                  the text box to get the data to decrypt, I would not
                  have all the data.  Thanks to Haakan Gustavsson
                  for pointing me in the right direction.
                  See frmEncStrings(cmdChoice_Click)

08 JAN 2001  1.1  For file and string testing, I have converted the
                  data and password to a byte array first.

30 DEC 2000  1.0  Wrote CryptoAPI demo program.

=============================================================================
                     CryptKci.dll
                     
Copy the \DLL\CryptKci.dll to the system directory where all your
other DLL's are stored.

	Windows 9x, ME         \Windows\System
	Windows NT4, 2000      \Winnt\System32

Now register the DLL so it will be recognized by the system.

Select the START button, RUN 

    for Windows 9x, ME  type:  
          regsvr32 c:\windows\system\CryptKci.dll
          
    for NT4, 2000  Type:      
          regsvr32 c:\winnt\system32\CryptKci.dll

In the VB IDE, to use this DLL, you must first reference it via
Projects, References on the toolbar menu.  Scroll down the list
and place a checkmark next to CryptKci                 

=============================================================================

You will need the VB6 runtime files.

This is freeware.  Since security is of the upmost these days,
a tool such as this should assist you in protecting your data.
This is well documented and should help you understand what is
happening.  I have tried to give everyone credit on their code
snippet contributions.  If you recognize something I missed, let
me know and I will update that portion with your name and email
address (I must have both).

To begin with, I used a lot of screens to demonstrate each function.
This is to better illustrate what is going on without getting lost in
performing multiple functions within a single form.

Next, I use a database for network security because the user would
never have access to the directory where this database is located.
Also, I doubt if they would recognize any of the data in it.

For test purposes, I have entered one user into the database.

	Name:       JohnDoe
	password:   moneytree
	
	Options:	Case senitive for both entries
	            MD5 (Message Digest) hash algorithm
	            Default provider

**  Brief Overview  ***************
Whenever a user logs onto a network, a server application is executed
from within the login script.  This server application has the only
access to the database as far as the user is concerned.  The user's
logon data is extracted from the workstation screen, manipulated and
applied to the database for verification.  If the logon data is
authenticated, the user is allowed onto the network.
***********************************

This database is very limited as it is for demonstration purposes only.

One of the things you could add to your code is the number of tries a
user can make trying to remember their password. Add a couple of fields
to the database to deny the user access for 15 minutes before being
allowed to try again.  In other words, set a flag field in the database
for a "1" or "0" and another field for the current timestamp.  If the
user is locked out after three tries, a "1" is entered into the flag
field and the system timestamp in the other.  Whenever the user attempts
a logon, first see if there is a "1" in the flag field.  If so, then test
to see if 15 minutes have elaspsed since the "1" was entered.  If 15
minutes or more have elapsed then enter a "0" in the flag field and NULL
in the timestamp field and continue processing.

Using this scenario is a definite thorn in the side of individuals trying
to gain unauthorized access to your system by way of brute force entry.

If you are using local security (the user's workstation), you can
apply these same principles for the Windows registry.  See the "Hash Test"
and you will see where, in my opinion, a Message Digest (MDn) algorithm
was probably used for registry entries.

=============================================================================

Written by Kenneth Ives                    kenaso@home.com

All of my routines have been compiled with VB6 Service Pack 5.
There are several locations on the web to obtain these
runtime modules.

This software is FREEWARE.  You may use it as you see fit for
your own projects but you may not re-sell the original or the
source code.

If there is anything in here that you want to use and I wrote
it, please give me credit.  This is a way of saying "Thank you"
to another programmer.

No warranty expressed or implied is given as to the use of this
program.  Use at your own risk.

If you have any suggestions or questions, I would be happy to
hear from you.
=============================================================================
