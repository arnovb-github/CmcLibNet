# CmcLibNet

## What is it
CmcLibNet is a .NET wrapper for the Commence RM API. I created it to make working with Commence RM from .Net easier, particularly PowerShell. Powershell can talk to Commence natively, but it cannot use the familiar dot notation like you can in other languages like VBscript, C# and Python to name a few. In Powershell you have to use the cumbersome `InvokeMember()` method and I really wanted an alternative.

Also, the Commence API is not very user-friendly. It is quirky, verbose and in some cases inconsistent. 

CmcLibNet intends to provide a more consistent interface wrapping the Commence API for use in .NET (and COM, but see below). It completely hides the Commence DDE interface and it has some functionality to make simple things actually simple(r)

## Benefits
One of the simple but helpful concepts that CmcLibNet offers is getting rid of all the string parsing you have to do when talking to Commence. For instance, when you want to retrieve a list of categories in a database, you can just call a method and you get returned an `IEnumerable`. The tedious stuff is handled by CmcLibNet.

Another concept that CmcLibNet uses is that filters on a cursor are actual objects, not unintelligible strings. Slightly more verbose but easier to use. Setting columns on a cursor is also simplified.

You can still make almost all API calls that are in the Commence RM API using CmcLibNet. The good bits are implemented either as an overload and/or extension method. The main exception to that is that CmcLibNet does not support `GetConversation()`. All methods exposed by `ICommenceConversation` in Commence RM are implemented as native methods in CmcLibNet.

CmcLibNet can export to more formats than Commence natively does, like JSON and XML, and extends some built-in ones. It does so pretty fast as well.

## Audience
This assembly still assumes (intimate) familiarity with the Commenc RM API. It just make some things easier and does some things better.

## Summary 
* use this assembly if you want to talk to Commence RM from Powershell.
* use this assembly if you want something easier than the vanilla Commence RM COM interface, in any .Net language.

## Binaries and documentation
See the [Wiki](https://github.com/arnovb-github/CmcLibNet/wiki) for download links and full documentation including basic examples.

## COM Interop
The assembly is _COM-visible_, meaning you can call CmcLibNet from any COM-capable language. That is why some of the code may appear counter-intuitive or overly complex in some places. In retrospect, this was a questionable design decision. It is great that you can call CmcLibNet from Commence Detail Form Scripts or Microsoft Office, but since this project was my first serious attempt at a .Net application, diving into the world of black magic that is COM-Interop was probably not the best idea. Don't worry though, I did not just tick 'Make assembly COM-visible' in the assembly properties, I did try to follow recommended practices. 

So before you start calling CmcLibNet from your VBA or Python or even Commence Detail Forms: you are then doing COM Interop->.NET->COM Interop. Be my guest but be aware of that!