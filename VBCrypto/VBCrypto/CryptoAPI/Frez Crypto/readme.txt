FrezCryto - available from http://www.frez.co.uk

Should you wish to commission any consultancy work, please do not hesitate to contact us; sales@frez.co.uk

FrezCrypto is an ActiveX DLL that wraps the Microsoft CryptoAPI and allows the user to encrypt or decrypt strings based on the RC2 Block Algorithm or RC4 Stream Algorithm. The encrypted text can either be returned as is, or converted to hex codes of the ascii characters for easy storage (albeit twice the size). Block encryption is the more secure method in the majority of cases, but the encrypted string returned is usually larger than the string that is supplied. Stream encryption returns a string of the same size but is usually less secure.

If used on very large strings it is recommended that you perform some form of compression on the string first.

A sample use of this technology would be to encypt a password for storing in a database that could otherwise be read in clear.

For added security you may want to put the code in a standard module or internal class rather than have it invoked in an external DLL.

This is 'free' software with the following restrictions:

You may not redistribute this code as a 'sample' or 'demo'. However, you are free to use the source code or DLL in your own code, but you may not claim that you created the sample code or compiled component. It is expressly forbidden to sell or profit from this source code or DLL other than by the knowledge gained or the enhanced value added by your own code.

Use of this software is also done so at your own risk. The code and component are supplied as is without warranty or guarantee of any kind.

Amendment History:
1.0  18-Feb-2000  Initial version.
1.1  21-Feb-2000  Fix for Windows NT 4.0. AquireContext required a verify.
1.2  22-Feb-2000  Minor fix to the source of an Err.Raise statement
