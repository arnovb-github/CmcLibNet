# CmcLibNet
CmcLibNet is a .Net wrapper around the Commence API designed to make working with Commence from a .Net language easier. My primary motive for developing it was to make talking to Commence from PowerShell easier. PowerShell will allow you to talk to Commence as a COM object, but you have to use the nasty IDispatch interface.

You can talk to Commence in a 'normal' COM way from any COM-capable language (except Powershell). It isn't always easy. The Commence API is hard to master, and it has it's flaws. CmcLibNet intends to alleviate some of that. It completely hides the DDE interface which Commence has sneakily implemented as ICommenceConversation (which is really just a proxy for DDE calls), and it adds some functionality to speed up things. Exporting data from Commence can be a bit of a bitch, as is querying the database. CmcLibNet makes all that a little easier.

The assembly is COM-enabled, which why some of the code may seem overly complex. You can call CmcLibNet from any COM-capable language.
