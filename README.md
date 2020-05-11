# CmcLibNet
CmcLibNet is a .NET wrapper around the Commence RM API. It was designed to make working with Commence RM from a .NET language easier, PowerShell in particular. PowerShell can talk to Commence natively, but you have to use the rather cumbersome IDispatch interface.

Furthermore, the Commence API is a little quirky and at times inconsistent. CmcLibNet intends to provide a more modern-day, more consistent interface. It completely hides Commence's DDE interface (i.e. ICommenceConversation), and it has some functionality to make simple things actually simple(r). For instance, exporting data from Commence can be a bit of a hassle. CmcLibNet makes it a little bit easier.

One of the simple but helpful concepts that CmcLibNet employs is getting rid of all the string parsing you usually have to do when talking to Commence. For instance, when you want to retrieve a list of categories in a database, you can just call a method and you get returned an IEnumerable. The tedious stuff is handled internally by CmcLibNet.

Another concept that CmcLibNet uses is that filters on a cursor are actual objects, not unintelligible strings. Slightly more verbose but much easier to use. 

See the Wiki for links to more details on this and some basic examples of how to use the assembly.

The assembly is COM-visible, meaning you can call CmcLibNet from any COM-capable language. That is why some of the code may appear counter-intuitive or overly complex in some places. In retrospect, this was a questionable design decision. It is great that you can call CmcLibNet from Commence Detail Form Scripts or Microsoft Office, but since this project was my first serious attempt at a .Net application, diving into the world of black magic that is COM-Interop was probably not the best idea. Don't worry though, I did not just tick 'Make assembly COM-visible' in the assembly info, I tried to follow recommended practices. It is the stuff of headaches, though.

On the other hand, the only way to talk to Commence is through COM-Interop anyway, so headaches were pretty much unavoidable.