using Microsoft.SharePoint;
using System;
using System.Runtime.InteropServices;
using Microsoft.SharePoint.Administration;
using TimerJobTest;
using System.IO;

namespace CYL_Project.Features.Feature2
{
    /// <summary>
    /// This class handles events raised during feature activation, deactivation, installation, uninstallation, and upgrade.
    /// This serves as an event receiver for the timer job feature.
    /// </summary>
    /// <remarks>
    /// The GUID attached to this class may be used during packaging and should not be modified.
    /// </remarks>

    [Guid("063dcfc5-b7f8-4ce1-ae91-05a0810bee90")]
    public class Feature2EventReceiver : SPFeatureReceiver
    {

        /// <summary>
        /// Installs the timer job when the feature is activated.
        /// </summary>
        /// <param name="properties"></param>
        public override void FeatureActivated(SPFeatureReceiverProperties properties)
        {   
            String path = @"C:\Users\Administrator\Documents\Logs\timer_job_activated.txt"; // Create a text file for logging and debugging purposes.

            if (!File.Exists(path))
            {
                using (StreamWriter sw = File.CreateText(path))
                {
                    sw.WriteLine("-----New File-----");
                }
            }

            try
            {
                SPSecurity.RunWithElevatedPrivileges(delegate ()
                {
                    SPWebApplication oWebApplication = (SPWebApplication)properties.Feature.Parent;
                    foreach (SPJobDefinition jobDefinition in oWebApplication.JobDefinitions)
                    {
                        // If a job with the same name already exists, delete it.
                        if (jobDefinition.Name == "TimerJob")
                        {
                            jobDefinition.Delete();
                            break;
                        }
                    }
                    // Install the job  
                    TimerJob timerJobObject = new TimerJob("TimerJob", oWebApplication);

                    // Create and update the timer schedule values. Modify this schedule as desired.
                    SPMinuteSchedule Timerschedule = new SPMinuteSchedule
                    {
                        BeginSecond = 0,
                        EndSecond = 59,
                        Interval = 10
                    };
                    timerJobObject.Schedule = Timerschedule;
                    timerJobObject.Update();

                    oWebApplication.JobDefinitions.Add(timerJobObject);
                });

                using (StreamWriter sw = File.AppendText(path))
                {
                    sw.WriteLine("TimerJob activated on " + DateTime.Now.ToString("dddd, dd MMMM yyyy HH:mm:ss"));
                }
            }
            catch (Exception ex) 
            {
                using (StreamWriter sw = File.AppendText(path))
                {
                    sw.WriteLine("Something went wrong after activating TimerJob: " + ex.Message + "\n");
                }
            }
        }

        /// <summary>
        /// Deletes the timer job when the feature is being deactivated.
        /// </summary>
        /// <param name="properties"></param>
        public override void FeatureDeactivating(SPFeatureReceiverProperties properties)
        {
            String path = @"C:\Users\Administrator\Documents\Logs\timer_job_deactivating.txt";
            if (!File.Exists(path))
            {
                using (StreamWriter sw = File.CreateText(path))
                {
                    sw.WriteLine("-----New File-----");
                }
            }

            try
            {
                SPSecurity.RunWithElevatedPrivileges(delegate ()
                {
                    SPWebApplication oWebApplication = (SPWebApplication)properties.Feature.Parent;
                 
                    foreach (SPJobDefinition jobDefinition in oWebApplication.JobDefinitions)
                    {
                        if (jobDefinition.Name == "TimerJob")
                        {
                            jobDefinition.Delete();
                        }
                    }
                });

                using (StreamWriter sw = File.AppendText(path))
                {
                    sw.WriteLine("TimerJob deleting on " + DateTime.Now.ToString("dddd, dd MMMM yyyy HH:mm:ss"));
                }
            }
            catch (Exception ex) 
            {
                using (StreamWriter sw = File.AppendText(path))
                {
                    sw.WriteLine("Something went wrong when deleting TimerJob: " + ex.Message + "\n");
                }
            }
        }


        // Uncomment the method below to handle the event raised after a feature has been installed.

        //public override void FeatureInstalled(SPFeatureReceiverProperties properties)
        //{
        //}


        // Uncomment the method below to handle the event raised before a feature is uninstalled.

        //public override void FeatureUninstalling(SPFeatureReceiverProperties properties)
        //{
        //}

        // Uncomment the method below to handle the event raised when a feature is upgrading.

        //public override void FeatureUpgrading(SPFeatureReceiverProperties properties, string upgradeActionName, System.Collections.Generic.IDictionary<string, string> parameters)
        //{
        //}
    }
}
