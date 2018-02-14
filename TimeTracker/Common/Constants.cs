using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TimeTracker.Common
{
    public class Constants
    {
        #region control constants

        public const string project = "project_";
        public const string title = "title_";
        public const string description = "description_";
        public const string time = "time_";
        public const string isBillable = "isBillable_";
        public const string isSubmitted = "isSubmitted_";

        #endregion control constants

        #region regular expressions

        public const string HLSCaseNumberRegex = @"^HLS-[0-9]{5}-[a-zA-Z0-9]{6}";

        #endregion regulare expressions

        #region status messages

        public const string CRMNotConnected = @"Time Track is not connected to CRM
Please enter your credentials in the connect.xml file";

        #endregion status messages
    }
}
