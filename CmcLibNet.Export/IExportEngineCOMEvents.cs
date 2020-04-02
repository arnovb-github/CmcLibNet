using System.Runtime.InteropServices;

namespace Vovin.CmcLibNet.Export
{
    /// <summary>
    ///  Interface for COM clients to implement events raised by this asssembly.
    /// </summary>
    /// <remarks>When used in VBA (Microsoft Office macros) you have to be very specific as to how to implement the events:
    /// <code language="VB">
    /// 'Create a class module called 'Class1'
    /// 'Declare the assembly using the WithEvents keyword.
    /// Private WithEvents cmclibnet as Vovin_CmcLibNet.ExportEngine
    /// Sub DoExport ' have a method that performs export
    ///     Set cmclibnet = New Vovin_CmcLibNet.ExportEngine
    ///     'do something
    ///     'cmclibnet.ExportView(..) ..
    ///     'cmclibnet.ExportCategory(..) ..
    ///     cmclibnet.Close
    /// End Sub
    /// 
    /// 'Subscribe to the event
    /// 'The sender argument has to be of type Variant, not Object.
    /// 'Also, do not leave ByVal out, and do not use ByRef.
    /// Private Sub cmclibnet_ExportProgressChanged(ByVal sender As Variant, ByVal args As ExportProgressChangedArgs)
    ///     Debug.Print args.CurrentRow
    /// End Sub
    /// </code>
    /// <para>In a regular module, you would use this as:</para>
    /// <code language="VB">
    /// Sub Test
    ///     Dim x As New Class1
    ///     Call x.DoExport
    /// End Sub
    /// </code> 
    /// </remarks>
    [ComVisible(true)]
    [Guid("521C2F4E-4596-4D04-BA57-0B0A6377F7F6")]
    // The ComInterfaceType.InterfaceIsIDispatch argument for the InterfaceTypeAttribute is important especially for VB6 clients.
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IExportEngineCOMEvents // no interfaces should inherit this interface!
    {
        /// <summary>
        /// Represents a batch of rows read from Commence.
        /// </summary>
        /// <param name="sender">sender.</param>
        /// <param name="e">ExportProgressChangedArgs.</param>
        [DispId(1)]
        void ExportProgressChanged(object sender, ExportProgressChangedArgs e);

        /// <summary>
        /// Export completed
        /// </summary>
        /// <param name="sender">sender</param>
        /// <param name="e">ExportCompleteArgs</param>
        [DispId(2)]
        void ExportCompleted(object sender, ExportCompleteArgs e);
    }
}