# CmcLibNet

## What is it
CmcLibNet is a .NET wrapper for the Commence RM API. I created it to make working with Commence RM from .Net easier, particularly PowerShell. PowerShell can already talk to Commence natively, but you have to use the cumbersome IDispatch interface. In other .Net languages like C# you can talk to Commence just as you would from non-.Net languages such as VBScript or Python.

However, the Commence API isn't very user-friendly. It is quirky, verbose and at times inconsistent. CmcLibNet intends to provide a more consistent interface. It completely hides Commence RM's DDE interface and it has some functionality to make simple things actually simple(r). For instance, getting and therefore exporting data from Commence can be a bit of a hassle. CmcLibNet makes this a little bit easier. It provides an interface for exporting to more formats than Commence natively does, like JSON and XML, and extends some built-in ones.

## Benefits
One of the simple but helpful concepts that CmcLibNet offers is getting rid of all the string parsing you have to do when talking to Commence. For instance, when you want to retrieve a list of categories in a database, you can just call a method and you get returned an IEnumerable. The tedious stuff is handled by CmcLibNet.

Another concept that CmcLibNet uses is that filters on a cursor are actual objects, not unintelligible strings. Slightly more verbose but easier to use. Setting columns on a cursor is also greatly simplified.

It should be noted that in CmcLibNet, you can still make all API calls using CmcLibNet. The good stuff is implemented as an overload and/or extension method. The big exception being that CmcLibNet does not support `GetConversation()`. All methods exposed by `ICommenceConversation` are implemented as native methods in CmcLibNet.

## Audience
This assembly still assumes (intimate) familiarity with the Commenc RM API. It just make some things easier and does some things better.

## Summary 
* use this assembly if you want to talk to Commence RM from Powershell.
* use this assembly if you want something easier than the vanilla Commence RM COM interface, in any .Net language.

## Binaries and documentation
See the [Wiki](https://github.com/arnovb-github/CmcLibNet/wiki) for download links and full documentation including basic examples.

## COM Interop
The assembly is _COM-visible_, meaning you can call CmcLibNet from any COM-capable language. That is why some of the code may appear counter-intuitive or overly complex in some places. In retrospect, this was a questionable design decision. It is great that you can call CmcLibNet from Commence Detail Form Scripts or Microsoft Office, but since this project was my first serious attempt at a .Net application, diving into the world of black magic that is COM-Interop was probably not the best idea. Don't worry though, I did not just tick 'Make assembly COM-visible' in the assembly properties, I did try to follow recommended practices. It is the stuff of headaches, though.

On the other hand, the only way to talk to Commence is through COM-Interop, so headaches were pretty much unavoidable anyway.