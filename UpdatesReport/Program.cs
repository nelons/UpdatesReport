using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.UpdateServices.Administration;
using System.Net;
using System.Net.Mail;
using System.IO;

namespace UpdatesReport
{
    class Program
    {
        private class groupCompareClass : IComparer
        {
            int IComparer.Compare(Object x, Object y)
            {
                UpdateGroup a = (UpdateGroup) x;
                UpdateGroup b = (UpdateGroup) y;

                return ((new CaseInsensitiveComparer()).Compare(a.Name, b.Name));
            }
        }

        private class computerCompareClass : IComparer
        {
            int IComparer.Compare(Object x, Object y)
            {
                UpdateComputer a = (UpdateComputer)x;
                UpdateComputer b = (UpdateComputer)y;

                return ((new CaseInsensitiveComparer()).Compare(a.Name, b.Name));
            }
        }

        static void Main(string[] args)
        {
            string updateServer = "";
            int updateServerPort = 80;
            string smtpHost = "";
            string fromAddress = "";
            string toAddress = "";
            string reportfolder = "";
            
            Arguments cmdLine = new Arguments(args);
            if (args.Length == 0 || cmdLine["help"] != null)
            {
                // Output help.
                System.Console.WriteLine("Updates Reporting Help - Command Line options.");
                System.Console.WriteLine("server - the WSUS server to connect to. Required.");
                System.Console.WriteLine("folder - where to save the reports. Required.");
                System.Console.WriteLine("port - the port the WSUS server uses. Default is 80.");
                System.Console.WriteLine("smtphost - the outgoing mail server.");
                System.Console.WriteLine("from - the address the email is being sent from.");
                System.Console.WriteLine("to - the address the email is being sent to.\n\n");
                System.Console.WriteLine("Example: updatesreport.exe -server wsus.dds.net -port 8540 -smtphost mail.dds.net -from wsus@dds.net -to serverops@dds.net");
                return;
            }

            if (cmdLine["server"] != null)
                updateServer = cmdLine["server"];
            else
            {
                System.Console.WriteLine("The parameter \"server\" which specifies the WSUS server is missing.");
                return;
            }

            if (cmdLine["folder"] != null)
                reportfolder = cmdLine["folder"];
            else
            {
                System.Console.WriteLine("The location in which to save reports is missing.");
                return;
            }

            if (cmdLine["smtphost"] != null)
                smtpHost = cmdLine["smtphost"];

            if (cmdLine["port"] != null)
                updateServerPort = Int32.Parse(cmdLine["port"]);

            if (cmdLine["from"] != null)
                fromAddress = cmdLine["from"];

            if (cmdLine["to"] != null)
                toAddress = cmdLine["to"];

            IUpdateServer srv;
            try
            {
                srv = AdminProxy.GetUpdateServer(updateServer, false, updateServerPort);
            }
            catch (Exception ex)
            {
                System.Console.WriteLine(ex.Message);
                return;
            }

            ArrayList computerFailures = new ArrayList();       // List of UpdateComputer objects which have failures
            ArrayList updateGroups = new ArrayList();           // List of UpdateGroups
            UpdateGroup allComputers = new UpdateGroup();
            allComputers.Name = "All Computers";
            ComputerTargetGroupCollection groupCol = srv.GetComputerTargetGroups();
            for (int i = 0; i < groupCol.Count; i++)
            {
                if (groupCol[i].Name == "All Computers")
                    continue;

                // Create group and set name
                UpdateGroup group = new UpdateGroup();
                group.Name = groupCol[i].Name;
                group.UpdateServicesGroupObject = groupCol[i];

                // now get information for this group.
                group.getGroupInformation(srv);

                // add totals to allComputers updateGroup
                allComputers.Total += group.Total;
                allComputers.Good += group.Good;
                allComputers.Failed += group.Failed;
                allComputers.NotInstalled += group.NotInstalled;
                allComputers.Unknown += group.Unknown;

                if (group.Failed > 0)
                {
                    // add all the computers to the list
                    group.addFailedComputers(ref computerFailures);
                }

                updateGroups.Add(group);
            }

            // add allComputer to list.
            updateGroups.Insert(0, allComputers);

            // sort list alphabetically
            groupCompareClass comp = new groupCompareClass();
            updateGroups.Sort(comp);

            // Assemble the message body
            string body = "<html><body style='font-family: Verdana; font-size: 12px'><p style='font-weight: bold'>WSUS Update report</p><p>This displays only the status of approved updates - a computer can have updates with different states and therefore appear in more than one of the three right-most columns. Attached is the latest 'detailed' report.</p><p><table border=1 cellpadding=3 cellspacing=0 style='font-size: 12px'><tr><td>Group</td><td>Total</td><td>Good</td><td>Needed</td><td>Failed</td><td>Unknown</td></tr>";
            string detailed = "<html><body style='font-family: Verdana; font-size: 12px'><p>Detailed report for WSUS Computer errors.</p>";

            foreach (UpdateGroup ug in updateGroups)
            {
                if (ug.Name != "Testing" && ug.Name != "ZZZ Updates on Hold")
                {
                    body += ug.toReport();

                    if (ug.ComputerCount() > 0 && (ug.NotInstalled > 0 || ug.Failed > 0 || ug.Unknown > 0))
                        detailed += ug.detailedReport();
                }
            }

            // clean up body
            body += "</table>";

            // Sort computer failures alphabetically
            // *****************
            computerCompareClass ccc = new computerCompareClass();
            computerFailures.Sort(ccc);

            // Now list all computers with failures
            if (computerFailures.Count > 0)
            {
                body += "<P style='font-weight: bold'>Computers with failed updates.</P><P>The error message for the update failure are underneath the name of the update and in grey. If there is no message then WSUS has not been informed of why the update failed to install. Last Reported Status times are highlighted in red if the computer has not reported within the last 14 days.</P>";
                foreach (UpdateComputer uc in computerFailures)
                {
                    body += uc.detailedReport(true);
                    body += "<br />";
                }
            }

            // Now list all computers with communication failures.
            DateTime thirty = DateTime.Today;
            thirty = thirty.AddDays(-30);

            ComputerTargetScope scope = new ComputerTargetScope();
            scope.ComputerTargetGroups.Add(srv.GetComputerTargetGroup(ComputerTargetGroupId.AllComputers));
            scope.ToLastReportedStatusTime = thirty;

            ComputerTargetCollection col = srv.GetComputerTargets(scope);
            if (col.Count > 0)
            {
                body += "<b>Computers not communicated status in last 30 days</b><br /><br />";
                foreach (IComputerTarget item in col)
                {
                    body += item.FullDomainName + ", " + item.OSDescription + ", last reported " + item.LastReportedStatusTime.ToShortDateString() + "<br />";
                }
            }

            // Tidy up report.
            body += "<P>The report was generated with the following parameters:</P><table  style='font-size: 12px'><tr><td>WSUS Server</td><td>" + updateServer + "</td></tr><tr><td>WSUS Server Port</td><td>" + updateServerPort + "</td></tr><tr><td>SMTP Host</td><td>" + smtpHost + "</td></tr><tr><td>From Address</td><td>" + fromAddress + "</td></tr><tr><td>To Address</td><td>" + toAddress + "</td></tr><tr><td>Report Folder</td><td>" + reportfolder + "</td></tr></table>";
            body += "</body></html>";
            detailed += "</body></html>";

            // save file & open with iexplore
            TextWriter tw = new StreamWriter(reportfolder + "\\report.html");            tw.Write(body);
            tw.Close();

            // now output the detailed report
            tw = new StreamWriter(reportfolder + "\\detailed.html");
            tw.Write(detailed);
            tw.Close();

            if (smtpHost != "" && fromAddress != "" && toAddress != "")
            {
                SmtpClient smtp = new SmtpClient
                {
                    Host = smtpHost,
                    Port = 25
                };

                // Create message
                MailMessage msg = new MailMessage(fromAddress, toAddress);
                msg.Subject = "WSUS Computer Reports";
                msg.Body = body;
                msg.IsBodyHtml = true;

                // Attach file
                if (File.Exists(reportfolder + "\\detailed.html"))
                {
                    Attachment detailedReport = new Attachment(reportfolder + "\\detailed.html");
                    msg.Attachments.Add(detailedReport);
                }

                try
                {
                    smtp.Send(msg);
                }
                catch (Exception e)
                {
                    System.Console.WriteLine("Failed to send email.\n" + e.InnerException);
                }
            }
        }
    }
}
