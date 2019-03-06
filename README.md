# CmcLibNet
CmcLibNet is a .Net wrapper around the Commence API designed to make working with Commence from a .Net language easier. My primary motive for developing it was to make talking to Commence from PowerShell easier. PowerShell will allow you to talk to Commence as a COM object, like any other COM-capable language, but it seems you have to use the nasty IDispatch interface.

The Commence API is a funny beast anyway, with it's quirks and flaws. CmcLibNet intends to alleviate some (not all!) of that. It completely hides the DDE interface (=anything to do with ICommenceConversation), and it has some functionality to make simple things actually simple(r). For instance, exporting data from Commence can be a bit of a hassle. CmcLibNet makes it all a little easier.

The assembly is COM-enabled, meaning you can call CmcLibNet from any COM-capable language. That is why some of the code may appear overly complex in places.

This was the first C# project I've ever worked on. I've learned a lot since, but not all of that made it back to this code. Be gentle :)