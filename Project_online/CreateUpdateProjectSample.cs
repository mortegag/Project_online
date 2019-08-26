using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using csom = Microsoft.ProjectServer.Client;

namespace CreateUpdateProjectSample
{
    partial class CreateUpdateProjectSample
    {

        // https://aigpanama.sharepoint.com/sites/Proyectos-TI/
        public static string pwaInstanceUrl = "https://aigpanama.sharepoint.com/sites/Proyectos-TI";         // your pwa url
        private static csom.ProjectContext context;
        const int DEFAULTTIMEOUTSECONDS = 300;

        private static string projectName = "DTR-2018005-AIG-KanbanMaestrodeProyecto";
        private static string localResourceName = "Javier Garrido";
        private static string taskName = "Confeccionar TDR";

        private static string projectCFName = "DTR-2018005-AIG-KanbanMaestrodeProyecto";
        private static string resourceCFName = "Javier Garrido";
        private static string taskCFName = "Confeccionar TDR";

        static void Main(string[] args)
        {
            //CreateProjectWithTaskAndAssignment();
            //context = GetContext(pwaInstanceUrl);
            ReadAndUpdateProject();
            UpdateCustomFieldValues();
        }

        #region Utility functions
        /// <summary>
        /// Log to Console the job state for queued jobs
        /// </summary>
        /// <param name="jobState">csom jobstate</param>
        /// <param name="jobDescription">job description</param>
        private static void JobStateLog(csom.JobState jobState, string jobDescription)
        {
            switch (jobState)
            {
                case csom.JobState.Success:
                    Console.WriteLine(jobDescription + " is successfully done.");
                    break;
                case csom.JobState.ReadyForProcessing:
                case csom.JobState.Processing:
                case csom.JobState.ProcessingDeferred:
                    Console.WriteLine(jobDescription + " is taking longer than usual.");
                    break;
                case csom.JobState.Failed:
                case csom.JobState.FailedNotBlocking:
                case csom.JobState.CorrelationBlocked:
                    Console.WriteLine(jobDescription + " failed. The job is in state: " + jobState);
                    break;
                default:
                    Console.WriteLine("Unkown error, job is in state " + jobState);
                    break;
            }
        }

        /// <summary>
        /// Get Publish project by name
        /// </summary>
        /// <param name="name">the name of the project</param>
        /// <param name="context">csom context</param>
        /// <returns></returns>
        private static csom.PublishedProject GetProjectByName(string name, csom.ProjectContext context)
        {
            IEnumerable<csom.PublishedProject> projs = context.LoadQuery(context.Projects.Where(p => p.Name == name));
            context.ExecuteQuery();

            if (!projs.Any())       // no project found
            {
                return null;
            }
            return projs.FirstOrDefault();
        }

        /// <summary>
        /// Get csom ProjectContext by letting user type in username and password
        /// </summary>
        /// <param name="url">pwa website url string</param>
        /// <returns></returns>
        private static csom.ProjectContext GetContext(string url)
        {
            csom.ProjectContext context = new csom.ProjectContext(url);
            string userName, passWord;

            //Console.WriteLine("Please enter your username for PWA");
            userName = "mortega@innovacion.gob.pa"; //Console.ReadLine();
            //Console.WriteLine("Please enter your password for PWA");
            passWord = "Coco.1961";//Console.ReadLine();

            NetworkCredential netcred = new NetworkCredential(userName, passWord);
            SharePointOnlineCredentials orgIDCredential = new SharePointOnlineCredentials(netcred.UserName, netcred.SecurePassword);
            context.Credentials = orgIDCredential;

            return context;
        }

        #endregion
    }
}
