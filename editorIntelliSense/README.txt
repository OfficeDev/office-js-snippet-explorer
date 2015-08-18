Files in this directory are copied here for the Monaco Editor's IntelliSense functionality.

They are marked as .txt rather than d.ts to avoid conflicts with Visual Studio's handling of d.ts files,
and to avoid issues with IIS not serving them.

The jquery.txt file is taken from the DefinitelyTyped repository 
	(https://github.com/borisyankov/DefinitelyTyped/blob/master/jquery/jquery.d.ts)

The Helpers.txt was written manually, for some helper functions used by the scripts (for logging and reporting success)

The rest of the files are copied automatically, either via running:
	%SRCROOT%\osfclient\RichApi\Test\RichApiAgaveWeb\CopyJsFromTarget.bat
Or (in the case of FakeXlapiTest.bat), by a post-build event on compiling Client.Test
