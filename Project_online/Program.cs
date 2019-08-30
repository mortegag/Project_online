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

         
    class Program
    {

        const string pwaPath = "https://aigpanama.sharepoint.com/sites/Proyectos-TI/";
        const string userName = "mortega@innovacion.gob.pa";
        const string passWord = "Coco.1961";
        static csom.ProjectContext ProjectCont1;
        MySqlConnection connect = new MySqlConnection();
        string exePath = System.Reflection.Assembly.GetEntryAssembly().Location;
 


        static void Main(string[] args)
        {          
     
           Program connection = new Program();
            connection.conn();
            // connection.actualiza_tareas("be62971f-cfc9-e911-ab58-34f39add823a", "Avances", "03/10/2018", "30/10/2018", 55);
             connection.leerbd();
            // connection.listProject();
           // connection.UddateTask("","","",12 );
        }

        private void UddateTask(string gui, string fi, string ff,  int porcent)
        {

            using (ProjectCont1)
            {

                Guid ProjectGuid = new Guid(gui);
                var projCollection = ProjectCont1.LoadQuery(
                 ProjectCont1.Projects
                   .Where(p => p.Id == ProjectGuid));
                ProjectCont1.ExecuteQuery();
                csom.PublishedProject proj2Edit = projCollection.First();
                DraftProject draft2Edit = proj2Edit.CheckOut();
                ProjectCont1.Load(draft2Edit);
                ProjectCont1.Load(draft2Edit.Tasks);
                ProjectCont1.ExecuteQuery();
                //
                var tareas = draft2Edit.Tasks;
                foreach (DraftTask tsk in tareas)
                {
                 tsk.Start = Convert.ToDateTime(fi);
                 tsk.Finish = Convert.ToDateTime(ff);
                 //tsk.Duration = duracion;
                 tsk.PercentComplete = porcent;
                }

                draft2Edit.Publish(true);
                csom.QueueJob qJob = ProjectCont1.Projects.Update();
                csom.JobState jobState = ProjectCont1.WaitForQueue(qJob, 200);
                //
                qJob = ProjectCont1.Projects.Update();
                jobState = ProjectCont1.WaitForQueue(qJob, 20);

                if (jobState == JobState.Success)
                {
                    Console.WriteLine("\nSuccess!");
                }

            }
        }

        private void addTask() {

            using (ProjectCont1)
            {

                Guid ProjectGuid = new Guid("e4707c63-cbc9-e911-ab58-34f39add823a");
          


                var projCollection = ProjectCont1.LoadQuery(
                 ProjectCont1.Projects
                   .Where(p => p.Id == ProjectGuid));
                        ProjectCont1.ExecuteQuery();

                csom.PublishedProject proj2Edit = projCollection.First();

                DraftProject draft2Edit = proj2Edit.CheckOut();

                ProjectCont1.Load(draft2Edit);
                ProjectCont1.Load(draft2Edit.Tasks);
                ProjectCont1.ExecuteQuery();

                TaskCreationInformation newTask = new TaskCreationInformation();
                newTask.Name = "Prueba";
                newTask.IsManual = false;
                newTask.Start = Convert.ToDateTime("30/08/2019"); 
                newTask.Id = Guid.NewGuid();
                newTask.Finish = Convert.ToDateTime("10/09/2019"); 

                DraftTask draftTask = draft2Edit.Tasks.Add(newTask);

                draft2Edit.Publish(true);             
                csom.QueueJob qJob = ProjectCont1.Projects.Update();
                csom.JobState jobState = ProjectCont1.WaitForQueue(qJob, 20);
                //
                qJob = ProjectCont1.Projects.Update();
                jobState = ProjectCont1.WaitForQueue(qJob, 20);

                if (jobState == JobState.Success)
                {
                    Console.WriteLine("\nSuccess!");
                }

            }
        }

        private void listProject() {
           int j = 1;

            using (ProjectCont1)
            {
                // Get the list of projects in Project Web App.
                Guid ProjectGuid = new Guid("2f7e6899-d9c8-e911-b070-00155db42408");
                Guid TaskGuid = new Guid("0a0e6daa-d9c8-e911-b07b-00155db45101");



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

            string ip = "localhost";
            string user = "root";
            string passw = "";
            string db = "AIGDB_SSEC";

            try
            {

                string connectionString = "server=" + ip + ";uid=" + user + ";pwd=" + passw + " ;database=" + db + ";";
                connect = new MySqlConnection(connectionString);

                                   //(string gui, string fi, string ff, string duracion, int porcent)
                string sql = "SELECT project_id, start_date, end_date,progress,updated_at  from projects";
                sql += " where updated_at >= DATE_FORMAT((SYSDATE() - INTERVAL 6 DAY), '%Y-%m-%d')";
                sql += " OR  created_at >= DATE_FORMAT((SYSDATE() - INTERVAL 6 DAY), '%Y-%m-%d')";
                sql += " OR created_at >= DATE_FORMAT((SYSDATE() - INTERVAL 6 DAY), '%Y-%m-%d')";
                sql += " OR updated_at >= DATE_FORMAT((SYSDATE() - INTERVAL 6 DAY), '%Y-%m-%d')";
                sql += " ORDER BY id";


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

                            UddateTask(row[0].ToString(), row[1].ToString(), row[2].ToString(),  Convert.ToInt16(row[3]));
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

                Guid ProjectGuid = new Guid(np);
             //   Guid TaskGuid = new Guid(nt);


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
                projCheckedOut.CheckIn(true);


            }

        }
    }
    
    }

