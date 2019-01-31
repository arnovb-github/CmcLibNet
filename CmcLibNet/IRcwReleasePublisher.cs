using System;

namespace Vovin.CmcLibNet
{
    /// <summary>
    /// Interface for RCWReleasePublisher.
    /// </summary>
    internal interface IRcwReleasePublisher
    {
        event EventHandler RCWRelease;
        void ReleaseRCWReferences();
    }
}
