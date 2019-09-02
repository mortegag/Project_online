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
        /*
        listado dep Proyectos nuevos del lado de Project Online de los 
        @project_New para insertar proyectos a MySql
        ultimos 5 dias actuales
        */
        List<string> project_New;
        /* Listado indices GUI de proyectos del lado de MYSQL
         */
        List<string> ssec_gui;
        /*Listado de indices GUI de proyectos del lado de Project Online
         */
        List<string> project_gui;
        
        static void Main(string[] args)
        {

            //
            var mysql = new List<string> { "carro","gato" };
            var project = new List<string> { "carro","gato","piñata"};
            var result = mysql.Except(project.ToList()); //pulsa , perro 
            //

            Program connection = new Program();
            connection.conn();
            // connection.actualiza_tareas("be62971f-cfc9-e911-ab58-34f39add823a", "Avances", "03/10/2018", "30/10/2018", 55);
            // connection.leerbd();
            // connection.listProject();
            // connection.UddateTask("","","",12 );
           connection.Project();

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
   
        private void Project() {
            int j = 1;

            using (ProjectCont1)
            {
               
                DateTime dia = DateTime.Today.AddDays(0);
                DateTime hoy = DateTime.Today;
                    DateTime ayer = hoy.AddDays(-4);            

                var projCollection = ProjectCont1.LoadQuery(
                    ProjectCont1.Projects
                     .Where(p => p.CreatedDate >= ayer));
                ProjectCont1.ExecuteQuery();

                foreach (PublishedProject pubProj in projCollection)
                {
                    string Guid = pubProj.Id.ToString();
                    //Lista de Proyectos del lado de Project para insertar
                    project_New = new List<string>(); 
                    project_New.Add(Guid);//Project_id IdDelProyecto
                    project_New.Add(pubProj.Name);//name NombreDeProyecto
                    project_New.Add(pubProj.Description);//description DescripciónDelProyecto
                    //grouper AgrupadordeProyecto
                    //compromise Compromiso                   
                    project_New.Add(pubProj.StartDate.ToShortDateString());//start_date ComienzoAnticipadoDelProyecto
                     project_New.Add(pubProj.FinishDate.ToShortDateString());//end_date FechaDeFinalizaciónDelProyecto
                    //institution InstitucióndelEstado
                    //action_line LíneadeAcción
                    //responsable SeguimientodeProyecto
                    project_New.Add(pubProj.CreatedDate.ToShortDateString());


                    //lista de GUI del lado de Project para comparar y borrar
                    ssec_gui = new List<string>();
                     ssec_gui.Add(Guid);

                    Console.WriteLine("\n{0}. {1}   {2} \t{3} \n", j++, pubProj.Id, pubProj.Name, pubProj.CreatedDate);
                    //intento de comparar dos matrices multidimencionales //comparar dos listas y buscar las diferencas
                    //var project = new List<string> { Guid + "," + pubProj.Name + "," + pubProj.CreatedDate };
                    //var projectGUI = new List<string> { Guid };
                    //var ssec = new List<string> { "datos del Mysql "};
                    //var projectFaltan = projectGUI.Except(ssec.ToList()); //list3 contains only 1, 2

                }
            }
            insertProject();
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
                            //Funcion de actualizacion de resitros en el Project
                            UddateTask(row[0].ToString(), row[1].ToString(), row[2].ToString(),  Convert.ToInt16(row[3]));
                            //Lista de GUI lado Mysql para Borrar
                            ssec_gui = new List<string>();
                            ssec_gui.Add(row[0].ToString());

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
        
        private void insertProject()
        {

            string ip = "localhost";
            string user = "root";
            string passw = "";
            string db = "AIGDB_SSEC";

            try
            {

                string connectionString = "server=" + ip + ";uid=" + user + ";pwd=" + passw + " ;database=" + db + ";";
                connect = new MySqlConnection(connectionString);
                var gui = project_New[0];
                var nombre = project_New[1];
                var fecha = project_New[2];

                string sql = " INSERT INTO projects ("+ gui+ "," + nombre + "," + fecha + ")";
                sql += " SELECT gui, nombre, fecha ";
                sql += " WHERE gui <> "+gui ;
                       

                if (connect.State != ConnectionState.Open)
                {
                    connect.Open();
                }


                MySqlCommand cmd = new MySqlCommand(sql, connect);
                cmd.ExecuteNonQuery();

                connect.Close(); 


            }
            catch (Exception ex)
            {
                ex.ToString();

            }
        }

        private void DeleteProject()
        {

            string ip = "localhost";
            string user = "root";
            string passw = "";
            string db = "AIGDB_SSEC";

            try
            {

                string connectionString = "server=" + ip + ";uid=" + user + ";pwd=" + passw + " ;database=" + db + ";";
                connect = new MySqlConnection(connectionString);

                if (connect.State != ConnectionState.Open){ connect.Open();}

                var projectFaltan = project_gui.Except(ssec_gui.ToList());

                foreach (string gui in projectFaltan)
                {
                string sql = " Delete projects ";
                sql += " WHERE gui ="+ gui;         
                MySqlCommand cmd = new MySqlCommand(sql, connect);
                cmd.ExecuteNonQuery();
                }

                connect.Close();


            }
            catch (Exception ex)
            {
                ex.ToString();

            }
        }

    }
    
    }

