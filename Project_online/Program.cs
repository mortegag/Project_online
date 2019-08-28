using System;
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

        const string pwaPath = "https://aigpanama.sharepoint.com/sites/Proyectos-TI/";
        const string userName = "mortega@innovacion.gob.pa";
        const string passWord = "Coco.1961";
        static csom.ProjectContext ProjectCont1;
        MySqlConnection connect = new MySqlConnection();
        string exePath = System.Reflection.Assembly.GetEntryAssembly().Location;
        public PublishedProject pubProj;


        static void Main(string[] args)
        {


            //var classProject = new Project();
            //classProject.actualiza_tareas("DTR-2018005-AIG-KanbanMaestrodeProyecto", "Revisión por Legal", "2019-08-19","2019-08-21",10);

            Program connection = new Program();
            connection.conn();
           // connection.actualiza_tareas("DTR-2018005-AIG-KanbanMaestrodeProyecto", "Revisión por Legal", "2019-08-10", "2019-08-21", 11);
            // connection.leerbd();
            connection.listProject();
        }

        private void listProject() {
           int j = 1;

            using (ProjectCont1)
            {
                // Get the list of projects in Project Web App.
                Guid ProjectGuid = new Guid("5a32731f-a750-e911-ae73-34f39add815e");
                Guid TaskGuid = new Guid("ef32731f-a750-e911-ae73-34f39add815e");



                var projCollection = ProjectCont1.LoadQuery(
                    ProjectCont1.Projects
                     .Where(p => p.Id == ProjectGuid));
                ProjectCont1.ExecuteQuery();



                foreach (PublishedProject pubProj in projCollection)
                {
                  
                    Console.WriteLine("\n{0}. {1}   {2} \t{3} \n", j++, pubProj.Id, pubProj.Name, pubProj.CreatedDate);

                    PublishedTaskCollection collTask = pubProj.Tasks;
                    ProjectCont1.Load(collTask,
                        tsk => tsk.IncludeWithDefaultProperties(
                            t => t.Id, t => t.Name,
                            t => t.Assignments));
                    ProjectCont1.Load(collTask);
                    ProjectCont1.ExecuteQuery();
                    if (collTask.Count==1) { }
                    Console.WriteLine("Task collection count: {0}", collTask.Count.ToString());
                    if (collTask.Count > 0 )
                    {
                        int k = 1;    //Task counter.
                                      //Processing task list for current project
                        foreach (PublishedTask t in collTask)
                        {
                            if (t.Id!=null && t.Id==TaskGuid){ 

                            Console.WriteLine("{0}. Id:{1} \tName:{2}", k, t.Id, t.Name);
                            k++;
                            //Define and retrieve Assignments for current task
                            PublishedAssignmentCollection collAssgns = t.Assignments;
                            ProjectCont1.Load(collAssgns);
                            ProjectCont1.ExecuteQuery();
                            Console.WriteLine("    Assignment collection count: {0}", collAssgns.Count);
                            if (collAssgns.Count > 0)
                            {
                                //Output string for resources assigned to task
                                StringBuilder output = new StringBuilder();
                                output.AppendFormat("\t Assignments: ");
                                foreach (PublishedAssignment a in collAssgns)
                                {
                                    //Define and retrieve resource name for current assignment 
                                    //(an object)
                                    ProjectCont1.Load(a,
                                        b => b.Resource.Name);
                                    ProjectCont1.ExecuteQuery();
                                    output.AppendFormat("{0}, ", a.Resource.Name);
                                }
                                Console.WriteLine(output);
                            }
                            else
                            {
                                Console.WriteLine("\t Assignments: None");
                            }
                            }
                        }
                       
                    }   // endif
                }
            }
        }

      


        private void conn() {

            ProjectCont1 = new csom.ProjectContext(pwaPath);
            SecureString securePassword = new SecureString();
            foreach (char c in passWord.ToCharArray())
            {
                securePassword.AppendChar(c);
            }
            ProjectCont1.Credentials = new SharePointOnlineCredentials(userName, securePassword);
            
        }

        private void leerbd()
        {

            string ip = "10.252.70.131";
            string user = "mortega";
            string passw = "Panama2019";
            string db = "AIGDB_SSEC";

            try
            {

                string connectionString = "server=" + ip + ";uid=" + user + ";pwd=" + passw + " ;database=" + db + ";";
                connect = new MySqlConnection(connectionString);
                string sql = "select project_id,task_id, start_date, end_date,progress,updated_at  from projects where id =166";

                if (connect.State != ConnectionState.Open)
                {
                    connect.Open();
                }
                using (MySqlDataAdapter da = new MySqlDataAdapter(sql, connect))
                {

                    DataTable dt = new DataTable();
                    da.Fill(dt);
                    if (dt != null)
                    {

 
                        foreach (DataRow row in dt.Rows)
                        {
                           
                            actualiza_tareas(row[0].ToString(), row[1].ToString(), row[2].ToString(), row[3].ToString(), Convert.ToInt16(row[4]));
                        }

              
                    }
                    connect.Close();
                }

            }
            catch (Exception ex)
            {
                ex.ToString();

            }
        }

        public void  actualiza_tareas(string np, string nt, string fi, string ff, int pa2)
        {
           
            using(ProjectCont1)
            {

                Guid ProjectGuid = new Guid("2f7e6899-d9c8-e911-b070-00155db42408");
                Guid TaskGuid = new Guid("0a0e6daa-d9c8-e911-b07b-00155db45101");

                var projCollection = ProjectCont1.LoadQuery(
                    ProjectCont1.Projects
                     .Where(p => p.Id  == ProjectGuid));
                ProjectCont1.ExecuteQuery();
           
                csom.PublishedProject proj2Edit = projCollection.First();
                csom.DraftProject projCheckedOut = proj2Edit.CheckOut();
                ProjectCont1.Load(projCheckedOut.Tasks);
                ProjectCont1.ExecuteQuery();
                csom.DraftTaskCollection tskcoll = projCheckedOut.Tasks;
               
                foreach (csom.DraftTask Task in tskcoll)
                {
                    if ((Task.Id != null) && (Task.Id == TaskGuid) )
                    {

                        ProjectCont1.Load(Task.CustomFields);
                        ProjectCont1.ExecuteQuery();
                        Task.Name = "Moises was HERE ";
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

                projCheckedOut.CheckIn(true);

        
            }

        }
    }
    
    }

