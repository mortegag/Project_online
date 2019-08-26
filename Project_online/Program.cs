﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Project.Server.ClientOM;
using System.Security;
using MySql.Data.MySqlClient;
using Microsoft.ProjectServer.Client;
using csom=Microsoft.ProjectServer.Client;
using Microsoft.SharePoint.Client;
using System.IO;
using System.Data;

namespace Project_online
{


    class bd {


    }
   

    class Project {


        public void  actualiza_tareas(string np, string nt, string fi, string ff, int pa2)
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
                        Task.ActualStart = Convert.ToDateTime(fi); //DateTime.Today;
                        Task.ActualFinish = Convert.ToDateTime(ff);//DateTime.Today;
                        Task.PercentComplete = pa2;
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


    class Program
    {
        MySqlConnection connect = new MySqlConnection();
        string exePath = System.Reflection.Assembly.GetEntryAssembly().Location;


        static void Main(string[] args)
        {

            var classProject = new Project();
            classProject.actualiza_tareas("DTR-2018005-AIG-KanbanMaestrodeProyecto", "Revisión por Legal", "2019-08-19","2019-08-21",10);

       

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

