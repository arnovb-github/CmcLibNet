﻿using System.Runtime.InteropServices;

namespace Vovin.CmcLibNet.Export
{
    /// <summary>
    ///  Interface for COM clients to implement events raised by this asssembly.
    ///  <remarks>When used in VBA (Microsoft Office macro's) you have to be quite specific as to how to implement the events:
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
    /// 'Note that you have to be very specific with the signature, you explicitly need the ByVal keyword
    /// 'Don't leave ByVal out, and don't use ByRef.
    /// 'Also note that the sender argument has to be of type Variant, not Object.
    /// Private Sub cmclibnet_ExportProgressChanged(ByVal sender As Variant, ByVal args As ExportProgressAsJsonChangedArgs)
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
    /// </summary>
    [ComVisible(true)]
    [GuidAttribute("521C2F4E-4596-4D04-BA57-0B0A6377F7F6")]
    // The ComInterfaceType.InterfaceIsIDispatch argument for the InterfaceTypeAttribute is important especially for VB6 clients.
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IExportEngineCOMEvents // no interfaces should inherit this interface!
    {
        /// <summary>
        /// Represents a batch of rows read from Commence.
        /// </summary>
        /// <param name="sender">sender</param>
        /// <param name="e">ExportProgressAsJsonChangedArgs</param>
        [DispId(1)]
        void ExportProgressChanged(object sender, ExportProgressAsJsonChangedArgs e);

        /// <summary>
        /// Export completed
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        [DispId(2)]
        void ExportCompleted(object sender, ExportCompleteArgs e);
    }
}