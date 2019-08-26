using Microsoft.ProjectServer.Client;
using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Project_online
{
    class update
    {

        private static void UpdateProjectCustomField(Guid ProjectId)
        {
            DraftProject projCheckedOut = null;
            try
            {
                Dictionary<string, object> projDict = new Dictionary<string, object>();
        

                using (ProjectContext projContext = new ProjectContext(PWAUrl))
                {

                    projContext.ExecutingWebRequest += claimsHelper.clientContext_ExecutingWebRequest;

                    var PrjList = projContext.LoadQuery(projContext.Projects.Where(proj => proj.Name == ""));
                    
                    projContext.ExecuteQuery();
                    Guid pGuid = PrjList.First().Id;

                    PublishedProject proj2Edit = PrjList.First();
                    projCheckedOut = proj2Edit.CheckOut().IncludeCustomFields;
                    projContext.Load(projCheckedOut);
                    projContext.ExecuteQuery();

                    var cflist = projContext.LoadQuery(projContext.CustomFields.Where(cf => cf.Name == "Testcol"));
                    projContext.ExecuteQuery();
                    projCheckedOut.SetCustomFieldValue(cflist.FirstOrDefault().InternalName, "Entry_c8f0abff70f5e51180cc00155dd45b0a");
                
                    QueueJob qJob = projCheckedOut.Publish(true);
                    JobState jobState = projContext.WaitForQueue(qJob, 70);

                }
            }
            catch (Exception ex)
            {

            }
        }

    }
}
