using System;
using System.Diagnostics;
using System.Runtime.Serialization;
using System.Web;

namespace Microsoft.Operations
{
    /// <author>Riza Marhaban (Adecco)</author>
    /// <summary>
    /// Error Detail class to be use for <see cref="System.ServiceModel.FaultException{TDetail}"/>
    /// </summary>
    /// <remarks>
    /// (notes by Warren) We don't often use all this functionality (which is heavily orientated
    /// towards debugging) Should be:
    /// 1. Trimmed to minimal
    /// 2. Remove the 'Web App' stuff.
    /// 3. Made Compatible with 'Azure Logger' object.
    /// 4. overload 'base' functionality for common loads
    /// 5. Remove the 'gets or sets' default xml 'help' i.e. give more context around usage.
    /// </remarks>
    [DataContract(Name = "ErrorDetail", Namespace = "")]
    public class ErrorDetail
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ErrorDetail"/> class.
        /// </summary>
        public ErrorDetail() { }

        /// <summary>
        /// Initializes a new instance of the <see cref="ErrorDetail"/> class.
        /// </summary>
        public ErrorDetail(Exception ex) : base()
        {
            if (HttpContext.Current != null)
            {
                UserAlias = HttpContext.Current.User.Identity.Name;
            }
            else
            {
                UserAlias = "(not a web app)";
            }

            StackTrace st = new StackTrace(ex, true);
            StackFrame frame = st.GetFrame(st.GetFrames().Length - 1);
            LineNumber = frame.GetFileLineNumber();
            Filename = frame.GetFileName();
            StackTrace = ex.StackTrace;
            Message = ex.Message;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ErrorDetail"/> class.
        /// </summary>
        public ErrorDetail(Exception ex, string reason) : base()
        {
            if (HttpContext.Current != null)
            {
                UserAlias = HttpContext.Current.User.Identity.Name;
            }
            else
            {
                UserAlias = "(not a web app)";
            }

            Reason = reason;
            if (ex == null)
            {
                Message = reason;
                return;
            }

            StackTrace st = new StackTrace(ex, true);
            StackFrame frame = st.GetFrame(st.GetFrames().Length - 1);
            LineNumber = frame.GetFileLineNumber();
            Filename = frame.GetFileName();
            StackTrace = ex.StackTrace;
            Message = ex.Message;
        }

        public ErrorDetail(Exception ex, string reason, string friendlyMessage) : base()
        {
            if (HttpContext.Current != null)
            {
                UserAlias = HttpContext.Current.User.Identity.Name;
            }
            else
            {
                UserAlias = "(not a web app)";
            }

            Reason = reason;
            FriendlyMessage = friendlyMessage;
            if (ex == null)
            {
                Message = friendlyMessage;
                return;
            }

            StackTrace st = new StackTrace(ex, true);
            StackFrame frame = st.GetFrame(st.GetFrames().Length - 1);
            LineNumber = frame.GetFileLineNumber();
            Filename = frame.GetFileName();
            StackTrace = ex.StackTrace;
            Message = ex.Message;
        }

        /// <summary>
        /// Gets or sets the error id.
        /// </summary>
        [DataMember]
        public int ErrorId { get; set; }

        /// <summary>
        /// Gets or sets the filename.
        /// </summary>
        [DataMember]
        public string Filename { get; set; }

        /// <summary>
        /// Gets or sets the simple message for client user.
        /// </summary>
        [DataMember]
        public string FriendlyMessage { get; set; }

        /// <summary>
        /// Gets or sets the line number.
        /// </summary>
        [DataMember]
        public int LineNumber { get; set; }

        /// <summary>
        /// Gets or sets the message.
        /// </summary>
        [DataMember]
        public string Message { get; set; }

        /// <summary>
        /// Gets or sets the reason.
        /// </summary>
        [DataMember]
        public string Reason { get; set; }

        /// <summary>
        /// Gets or sets the stack trace.
        /// </summary>
        [DataMember]
        public string StackTrace { get; set; }

        /// <summary>
        /// Gets or sets the user alias.
        /// </summary>
        [DataMember]
        public string UserAlias { get; set; }
    }
}