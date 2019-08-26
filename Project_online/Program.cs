using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Project.Server.ClientOM;
using System.Security;
using Microsoft.ProjectServer.Client;
using csom=Microsoft.ProjectServer.Client;
using Microsoft.SharePoint.Client;
using System.IO;

namespace Project_online
{
    class Program
    {
        static void Main(string[] args)
        {

            using (csom.ProjectContext ProjectCont1 = new csom.ProjectContext("https://aigpanama.sharepoint.com/sites/Proyectos-TI/"))
            {

                SecureString password = new SecureString();
                foreach (char c in "Coco.1961".ToCharArray()) password.AppendChar(c);

                ProjectCont1.Credentials = new SharePointOnlineCredentials("mortega@innovacion.gob.pa", password);
                // var projects = ProjectCont1.Projects;
                // ProjectCont1.Load(projects);
                // ProjectCont1.ExecuteQuery();

                var projCollection = ProjectCont1.LoadQuery(
                ProjectCont1.Projects
                 .Where(p => p.Name == "DTR-2018005-AIG-KanbanMaestrodeProyecto"));
                ProjectCont1.ExecuteQuery();
                csom.PublishedProject proj2Edit = projCollection.First();
                csom.DraftProject projCheckedOut = proj2Edit.CheckOut();
                ProjectCont1.Load(projCheckedOut.Tasks);
                ProjectCont1.ExecuteQuery();
                csom.DraftTaskCollection tskcoll = projCheckedOut.Tasks;
                foreach (csom.DraftTask Task in tskcoll)
                {
                    if ((Task.Id != null) && (Task.Name == "Procedimientos Electrónico para SIRCEL"))
                    {

                        ProjectCont1.Load(Task.CustomFields);
                        ProjectCont1.ExecuteQuery();
                        Task.Name = "Firmas Electrónica para SIRCEL";// txtTache.Text;
                        Task.Start = DateTime.Today;
                        Task.PercentComplete = 10;
                        csom.AssignmentCreationInformation r = new csom.AssignmentCreationInformation();
                        r.Id = Guid.NewGuid();
                        r.TaskId = Task.Id;
                        Task.Assignments.Add(r);
                    }
                }
                projCheckedOut.Publish(true);
                csom.QueueJob qJob = ProjectCont1.Projects.Update();
                csom.JobState jobState = ProjectCont1.WaitForQueue(qJob, 20);

                //Añadir checkin para protejer el documento
            }




        }

        public void actualiza_tareas( string np, string nt, string nnt )
        {

            using (csom.ProjectContext ProjectCont1 = new csom.ProjectContext("https://aigpanama.sharepoint.com/sites/Proyectos-TI/"))
            {

                SecureString password = new SecureString();
                foreach (char c in "Coco.1961".ToCharArray()) password.AppendChar(c);

                ProjectCont1.Credentials = new SharePointOnlineCredentials("mortega@innovacion.gob.pa", password);
         
                var projCollection = ProjectCont1.LoadQuery(
                ProjectCont1.Projects
                 .Where(p => p.Name == np));
                ProjectCont1.ExecuteQuery();
                csom.PublishedProject proj2Edit = projCollection.First();
                csom.DraftProject projCheckedOut = proj2Edit.CheckOut();
                ProjectCont1.Load(projCheckedOut.Tasks);
                ProjectCont1.ExecuteQuery();
                csom.DraftTaskCollection tskcoll = projCheckedOut.Tasks;
                foreach (csom.DraftTask Task in tskcoll)
                {
                    if ((Task.Id != null) && (Task.Name == nt))
                    {

                        ProjectCont1.Load(Task.CustomFields);
                        ProjectCont1.ExecuteQuery();
                        Task.Name = nnt;// txtTache.Text;
                        Task.Start = DateTime.Today;
                        Task.PercentComplete = 100;
                        csom.AssignmentCreationInformation r = new csom.AssignmentCreationInformation();
                        r.Id = Guid.NewGuid();
                        r.TaskId = Task.Id;
                        Task.Assignments.Add(r);
                    }
                }
                projCheckedOut.Publish(true);
                csom.QueueJob qJob = ProjectCont1.Projects.Update();
                csom.JobState jobState = ProjectCont1.WaitForQueue(qJob, 20);
            }





        }


    }


    }

