using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.UpdateServices.Administration;

namespace UpdatesReport
{
    class Update
    {
        public string Name;
        public UpdateInstallationState State;
        public UpdateRevisionId Id;
        public string ErrorMessage;

        public Update(string updateName, UpdateInstallationState updateState, UpdateRevisionId updateID)
        {
            Name = updateName;
            State = updateState;
            Id = updateID;
            ErrorMessage = "";
        }

        public string toString()
        {
            string uis = "";
            switch (State)
            {
                case UpdateInstallationState.Downloaded: uis = "Downloaded"; break;
                case UpdateInstallationState.Failed: uis = "Failed"; break;
                case UpdateInstallationState.Installed: uis = "Installed"; break;
                case UpdateInstallationState.InstalledPendingReboot: uis = "Pending Reboot"; break;
                case UpdateInstallationState.NotApplicable: uis = "Not Applicable"; break;
                case UpdateInstallationState.NotInstalled: uis = "Not Installed"; break;
                case UpdateInstallationState.Unknown: uis = "Unknown"; break;
            }

            return String.Format("{0}, {1}", uis, Name);
        }
    }
}
