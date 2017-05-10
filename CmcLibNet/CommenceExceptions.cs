using System;
using System.Runtime.Serialization;

namespace Vovin.CmcLibNet
{
    /// <summary>
    /// Custom Commence DDE Exception thrown when DDE call failed.
    /// </summary>
    [Serializable]
    public class CommenceDDEException : System.Exception
    {
        /// <summary>
        /// constructor.
        /// </summary>
        public CommenceDDEException()
        {
        }
        /// <summary>
        /// overloaded constructor.
        /// </summary>
        /// <param name="message">Exception message.</param>
        public CommenceDDEException(string message)
            : base(message)
        {
        }
        /// <summary>
        /// overloaded constructor.
        /// </summary>
        /// <param name="message">Exception message.</param>
        /// <param name="innerException">Inner exception.</param>
        public CommenceDDEException(string message,
            Exception innerException)
            : base(message, innerException)
        {
        }
        /// <summary>
        /// overloaded constructor.
        /// </summary>
        /// <param name="info">SerializationInfo</param>
        /// <param name="context">CommenceDDEException</param>
        protected CommenceDDEException(SerializationInfo info,
            StreamingContext context)
            : base(info, context)
        {
        }
    }

    // Commence threw a COM exception
    /// <summary>
    /// Custom Commence COM exception thrown when a COM error occurs while talking to Commence.
    /// </summary>
    [Serializable]
    public class CommenceCOMException : System.Exception
    {
        /// <summary>
        /// default constructor.
        /// </summary>
        public CommenceCOMException()
        {
        }
        /// <summary>
        /// overloaded constructor.
        /// </summary>
        /// <param name="message">Exception message.</param>
        public CommenceCOMException(string message)
            : base(message)
        {
        }
        /// <summary>
        /// overloaded constructor.
        /// </summary>
        /// <param name="message">Exception message.</param>
        /// <param name="innerException">Inner exception.</param>
        public CommenceCOMException(string message,
            Exception innerException)
            : base(message, innerException)
        {
        }
        /// <summary>
        /// overloaded constructor.
        /// </summary>
        /// <param name="info">SerializationInfo</param>
        /// <param name="context">CommenceDDEException</param>
        protected CommenceCOMException(SerializationInfo info,
            StreamingContext context)
            : base(info, context)
        {
        }
    }

    // Commence is not running exception
    /// <summary>
    /// Custom exception thrown when assembly is invoked while Commence is not running.
    /// </summary>
    [Serializable]
    public class CommenceNotRunningException : System.Exception
    {
        /// <summary>
        /// default constructor.
        /// </summary>
        public CommenceNotRunningException()
        {
        }
        /// <summary>
        /// overloaded constructor.
        /// </summary>
        /// <param name="message">Exception message.</param>
        public CommenceNotRunningException(string message)
            : base(message)
        {
        }
        /// <summary>
        /// overloaded constructor.
        /// </summary>
        /// <param name="message">Exception message.</param>
        /// <param name="innerException">Inner exception.</param>
        public CommenceNotRunningException(string message,
            Exception innerException)
            : base(message, innerException)
        {
        }
        /// <summary>
        /// overloaded constructor.
        /// </summary>
        /// <param name="info">SerializationInfo</param>
        /// <param name="context">CommenceDDEException</param>
        protected CommenceNotRunningException(SerializationInfo info,
            StreamingContext context)
            : base(info, context)
        {
        }
    }

    // Commence is running multiple instances
    /// <summary>
    /// Custom exception thrown when assembly is invoked while Commence is running more than 1 instance.
    /// </summary>
    [Serializable]
    public class CommenceMultipleInstancesException : System.Exception
    {
        /// <summary>
        /// default constructor.
        /// </summary>
        public CommenceMultipleInstancesException()
        {
        }
        /// <summary>
        /// overloaded constructor.
        /// </summary>
        /// <param name="message">Exception message.</param>
        public CommenceMultipleInstancesException(string message)
            : base(message)
        {
        }
        /// <summary>
        /// overloaded constructor.
        /// </summary>
        /// <param name="message">Exception message.</param>
        /// <param name="innerException">Inner exception.</param>
        public CommenceMultipleInstancesException(string message,
            Exception innerException)
            : base(message, innerException)
        {
        }
        /// <summary>
        /// overloaded constructor.
        /// </summary>
        /// <param name="info">SerializationInfo</param>
        /// <param name="context">CommenceDDEException</param>
        protected CommenceMultipleInstancesException(SerializationInfo info,
            StreamingContext context)
            : base(info, context)
        {
        }
    }
}