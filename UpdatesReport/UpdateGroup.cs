using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.UpdateServices.Administration;

namespace UpdatesReport
{
    class UpdateGroup
    {
        public string Name;
        public IComputerTargetGroup UpdateServicesGroupObject;

        public int Total;
        public int Good;
        public int NotInstalled;
        public int Failed;
        public int Unknown;

        private ArrayList Computers;

        public int ComputerCount()
        {
            return Computers.Count;
        }

        public UpdateGroup()
        {
            Computers = new ArrayList();
        }

        public void getGroupInformation(IUpdateServer srv)
        {
            UpdateScope targetApprovedProblems = new UpdateScope();
            targetApprovedProblems.ApprovedStates = ApprovedStates.LatestRevisionApproved;
            targetApprovedProblems.UpdateApprovalActions = UpdateApprovalActions.Install;
            targetApprovedProblems.IncludedInstallationStates = UpdateInstallationStates.All;
            targetApprovedProblems.ExcludedInstallationStates = UpdateInstallationStates.Installed;
            targetApprovedProblems.UpdateSources = UpdateSources.MicrosoftUpdate;

            ComputerTargetScope scope = new ComputerTargetScope();
            scope.ComputerTargetGroups.Add(this.UpdateServicesGroupObject);
            ComputerTargetCollection targets = srv.GetComputerTargets(scope);

            for (int k = 0; k < targets.Count; k++)
            {
                // Look at this computer
                IComputerTarget comp = targets[k];
                IUpdateSummary sum = comp.GetUpdateInstallationSummary(targetApprovedProblems);
                UpdateComputer computer = new UpdateComputer(comp, sum);

                if (computer.Failed > 0) {
                    UpdateEventCollection col = srv.GetUpdateEventHistory(DateTime.MinValue, DateTime.MaxValue, comp);
                    foreach (IUpdateEvent eve in col)
                    {
                        if (eve.Status == InstallationStatus.Failed)
                        {
                            try
                            {
                                IUpdate upd = srv.GetUpdate(eve.UpdateId);  // have to do this to make sure each event relates to an actual update.
                                computer.setUpdateErrorMessage(eve.UpdateId, eve.Message);
                            }
                            catch (Exception e)
                            {
                            }
                        }
                    }
                }

                // Add computer to the arraylist
                Computers.Add(computer);

                Total++;
                if (computer.NotInstalled > 0)
                    NotInstalled++;
                if (computer.Failed > 0)
                    Failed++;
                if (computer.Unknown > 0)
                    Unknown++;
                if (computer.NotInstalled == 0 && computer.Failed == 0 && computer.Unknown == 0)
                    Good++;
            }
        }

        public string toReport()
        {
            string red = "style='background-color: #ff0000'";
            if (Failed == 0)
                red = "";
            return String.Format("<tr><td>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td><td {6}>{4}</td><td>{5}</td></tr>", Name, Total, Good, NotInstalled, Failed, Unknown, red);
        }

        public string detailedReport()
        {
            // present detailed info 
            string report = "<h4><A href='javascript:'; onclick=\"if (this.parentNode.nextSibling.style.display == 'none') { this.parentNode.nextSibling.style.display = ''; } else { this.parentNode.nextSibling.style.display = 'none'; }\">" + Name + " - " + getNonGoodComputerCount() + " Computers</A></h4><div style='display: none'>";
            
            // send to computers.
            foreach (UpdateComputer uc in Computers)
            {
                if (uc.GoodState == false)
                    report += uc.detailedReport() + "<br />";
            }

            report += "</div>";
            return report;
        }

        public void addFailedComputers(ref ArrayList list)
        {
            foreach (UpdateComputer uc in Computers)
            {
                if (uc.Failed > 0)
                    list.Add(uc);
            }
        }

        private int getNonGoodComputerCount()
        {
            int count = 0;
            foreach (UpdateComputer uc in Computers)
            {
                if (uc.GoodState == false)
                    count++;
            }

            return count;
        }
    }
}
