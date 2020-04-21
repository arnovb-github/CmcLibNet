# CmcLibNet
CmcLibNet is a .NET wrapper around the Commence RM API. It was designed to make working with Commence RM from a .NET language easier, especially PowerShell. PowerShell will already allow you to talk to Commence as a COM object, but you have to do it using the rather unfriendly IDispatch interface.

Furthermore, the Commence API is a funny animal, with its quirks and flaws. CmcLibNet intends to alleviate some (certainly not all) of those. It completely hides the DDE interface (=anything to do with ICommenceConversation), and it has some functionality to make simple things actually simple(r). For instance, exporting data from Commence can be a bit of a hassle. CmcLibNet makes it all a little bit easier and, more importantly, more consistent.

One of the simple but helpful things that CmcLibNet does is get rid of a lot of the string parsing you usually have to do when talking to Commence. For instance, when you want to retrieve a list of categories in a database, you can just call a method and you get returned an IEnumerable. That is much easier than constructing a DDE Request, opening a DDE conversation, execute the request and the parse the delimited return-string. In fact, CmcLibNet completely hides the ICommenceConversation (=DDE) interface.

Another concept that CmcLibNet uses is that filters on a cursor are actual objects, not unintelligible strings. Slightly more verbose but so much easier to use.

See the Wiki for links to more details and some basic examples of actually using the assembly.

The assembly is COM-visible, meaning you can call CmcLibNet from any COM-capable language. That is why some of the code may appear overly complex in places. In retrospect, this was a questionable design decision. It is great that you can call CmcLibNet from Commence Detail Form Scripts or Excel, but since this project was my first serious attempt at a .Net application, diving into the world of black magic that is COM-Interop was probably not the best idea. Don't worry though, I did not just tick 'Make assembly COM-visible' in the assembly info, I did follow recommended practices. It is the stuff of headaches, though.

On the other hand, the only way to talk to Commence is through COM-Interop, so headaches were pretty much unavoidable.