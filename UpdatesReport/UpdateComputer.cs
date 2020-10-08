using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.UpdateServices.Administration;

namespace UpdatesReport
{
    class UpdateComputer
    {
        public string Name;
        public int NotInstalled;
        public int Failed;
        public int Unknown;
        public bool GoodState;
        public DateTime LastReportedStatus;
        public string OSDescription;

        private ArrayList Updates;

        public UpdateComputer(IComputerTarget target, IUpdateSummary summary)
        {
            Updates = new ArrayList();

            Name = target.FullDomainName;
            NotInstalled = summary.NotInstalledCount;
            Failed = summary.FailedCount;
            Unknown = summary.UnknownCount;
            LastReportedStatus = target.LastReportedStatusTime;
            OSDescription = target.OSDescription + " " + target.OSArchitecture + " (" + target.ClientVersion + ")";            

            if (NotInstalled > 0 || Failed > 0 || Unknown > 0)
            {
                // Look more closely into whether they are etc.
                UpdateScope updateProblems = new UpdateScope();
                updateProblems.UpdateApprovalActions = UpdateApprovalActions.Install;
                updateProblems.ExcludedInstallationStates = UpdateInstallationStates.Installed | UpdateInstallationStates.NotApplicable;
                updateProblems.UpdateSources = UpdateSources.MicrosoftUpdate;

                UpdateInstallationInfoCollection updcol = target.GetUpdateInstallationInfoPerUpdate(updateProblems);
                //ArrayList updates = new ArrayList();
                for (int j = 0; j < updcol.Count; j++)
                {
                    IUpdateInstallationInfo updinfo = updcol[j];
                    IUpdate upd = updinfo.GetUpdate();

                    if (updinfo.UpdateApprovalAction == UpdateApprovalAction.NotApproved)
                    {
                        // update whether this update is actually a problem or not
                        // we've asked for all approved updates and unapproved updates are getting through :o
                        switch (updinfo.UpdateInstallationState)
                        {
                            case UpdateInstallationState.Failed: this.Failed--; break;
                            case UpdateInstallationState.NotInstalled: this.NotInstalled--; break;
                            case UpdateInstallationState.Unknown: this.Unknown--; break;
                        }

                        continue;
                    }

                    // Add this update to the list.
                    Updates.Add(new Update(upd.Title, updinfo.UpdateInstallationState, upd.Id));
                }
            }

            // check if this is a good state.
            GoodState = (NotInstalled > 0 || Failed > 0 || Unknown > 0) ? false : true;
        }

        public void setUpdateErrorMessage(UpdateRevisionId updateID, string errorMsg)
        {
            foreach (Update u in Updates)
            {
                if (u.Id.RevisionNumber == updateID.RevisionNumber && u.Id.UpdateId == updateID.UpdateId)
                {
                    u.ErrorMessage = errorMsg;
                    //Console.WriteLine("set errormessage " + errorMsg);
                    break;
                }
            }
        }

        public string toString()
        {
            return Name;
        }

        public string detailedReport(bool bOnlyFailed = false)
        {
            string lastReportedText = LastReportedStatus.ToLongTimeString() + " " + LastReportedStatus.ToLongDateString();
            if (LastReportedStatus < DateTime.Today.AddDays(-14))
            {
                lastReportedText = lastReportedText.Insert(0, "<span style='color: red'>");
                lastReportedText += "</span>";
            }

            string report = "<u>" + Name + "</u> <span style='color: grey'>Last Reported Status - " + lastReportedText + "</span><br />";
            report += "<span style='color: blue'>OS = " + OSDescription + "</span><br />";
            
            // output information about the updates.
            foreach (Update u in Updates)
            {
                if (bOnlyFailed)
                {
                    if (u.State == UpdateInstallationState.Failed)
                    {
                        report += u.toString() + "<br />";

                        // output event error message
                        if (u.ErrorMessage != "")
                        {
                            report += "<span style='color: grey'>" + u.ErrorMessage + "</span><br />";
                            //Console.WriteLine("writing errormessage to email");
                        }
                    }
                }
                else
                {
                    if (u.State == UpdateInstallationState.Failed)
                    {
                        report += "<span style='color: red'>" + u.toString() + "</span><br />";
                    }
                    else
                    {
                        report += u.toString() + "<br />";
                    }
                }
            }

            return report;
        }
    }
}
