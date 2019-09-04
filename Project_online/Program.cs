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
        const int PROJECT_BLOCK_SIZE = 20;
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
        static Dictionary<String, CustomField> pwaECF = new Dictionary<string, CustomField>();

        static void Main(string[] args)
        {

            //
            var mysql = new List<string> { "auto", "moto" };
            var project = new List<string> { "auto", "moto", "bici" };
            var result = mysql.Except(project.ToList());
            //

            Program connection = new Program();
            //connection.conn();
            // connection.actualiza_tareas("be62971f-cfc9-e911-ab58-34f39add823a", "Avances", "03/10/2018", "30/10/2018", 55);
            // connection.leerbd();
            // connection.listProject();
            // connection.UddateTask("","","",12 );
            // connection.Project();
            //connection.customfield();

            ListPWACustomFields();


        }

        public static void ReadCustomFieldValue()
        {
            /*
            using (ProjectContext ProjectCont = new ProjectContext(Credentials.SiteURL))
            {

                SecureString password2 = new SecureString();

                foreach (char c in Credentials.Password.ToCharArray()) password2.AppendChar(c);
                ProjectCont.Credentials = new SharePointOnlineCredentials(Credentials.UserName, passWord2);
                ProjectCont.Load(ProjectCont.Projects,
                    c => c.IncludeWithDefaultProperties
                    (
                    pr => pr.CustomFields,
                    pr => pr.IncludeCustomFields, 
                    pr => pr.IncludeCustomFields.CustomFields
                    ));
                ProjectCont.ExecuteQuery();
                foreach (PublishedProject item in ProjectCont.Projects)
                {
                    if (item.Name == "Project_00A")
                    {
                        foreach (var cust in item.IncludeCustomFields.FieldValues)
                        {
                            string customfieldID = cust.Key;
                            string CsutomfieldValue = cust.Value.ToString();
                        }
                    }
                }
            }*/
        }


        private static void ListPWACustomFields()
        {

            ProjectCont1 = new csom.ProjectContext(pwaPath);
            SecureString securePassword = new SecureString();
            foreach (char c in passWord.ToCharArray())
            {
                securePassword.AppendChar(c);
            }
            ProjectCont1.Credentials = new SharePointOnlineCredentials(userName, securePassword);
            var creds = new SharePointOnlineCredentials(userName, securePassword);

            //************************************
            // 2. Get project list with minimal information
            ProjectCont1.Load(ProjectCont1.Projects, qp => qp.Include(qr => qr.Id));
            ProjectCont1.ExecuteQuery();

            var allIds = ProjectCont1.Projects.Select(p => p.Id).ToArray();

            int numBlocks = allIds.Length / PROJECT_BLOCK_SIZE + 1;

            // Query all the child objects in blocks of PROJECT_BLOCK_SIZE
            for (int i = 0; i < numBlocks; i++)
            {
                var idBlock = allIds.Skip(i * PROJECT_BLOCK_SIZE).Take(PROJECT_BLOCK_SIZE);
                Guid[] block = new Guid[PROJECT_BLOCK_SIZE];
                Array.Copy(idBlock.ToArray(), block, idBlock.Count());

                //
          
                DateTime hoy = DateTime.Today;
                DateTime ayer = hoy.AddDays(0);
                string last = ayer.ToShortDateString();
                // 2. Retrieve and save project basic and custom field properties in an IEnumerable collection.
                var projBlk = ProjectCont1.LoadQuery(
                     ProjectCont1.Projects
                    .Where(p =>   // some elements will be Zero'd guids at the end
                        p.Id == block[0] || p.Id == block[1] ||
                        p.Id == block[2] || p.Id == block[3] ||
                        p.Id == block[4] || p.Id == block[5] ||
                        p.Id == block[6] || p.Id == block[7] ||
                        p.Id == block[8] || p.Id == block[9] ||
                        p.Id == block[10] || p.Id == block[11] ||
                        p.Id == block[12] || p.Id == block[13] ||
                        p.Id == block[14] || p.Id == block[15] ||
                        p.Id == block[16] || p.Id == block[17] ||
                        p.Id == block[18] || p.Id == block[19] &&  
                        p.CreatedDate >= ayer
                    )
                    .Include(p => p.Id,
                        p => p.Name,
                        p => p.Description,
                        p => p.StartDate,
                        p => p.FinishDate,
                        p => p.CreatedDate,
                        p => p.IncludeCustomFields,
                        p => p.IncludeCustomFields.CustomFields,
                        P => P.IncludeCustomFields.CustomFields.IncludeWithDefaultProperties(
                            lu => lu.LookupTable,
                            lu => lu.LookupEntries
                        )
                    )
                );

                ProjectCont1.ExecuteQuery();

                foreach (PublishedProject pubProj in projBlk)
                {

                    // Set up access to custom field collection of published project
                    var projECFs = pubProj.IncludeCustomFields.CustomFields;

                    // Set up access to custom field values of published project
                    Dictionary<string, object> ECFValues = pubProj.IncludeCustomFields.FieldValues;

                    Console.WriteLine("Name:\t{0}", pubProj.Name);
                    Console.WriteLine("Id:\t{0}", pubProj.Id);
                    Console.WriteLine("Description:\t{0}", pubProj.Description);
                    Console.WriteLine("Create Date:\t{0}", pubProj.CreatedDate);
                    Console.WriteLine("ToDay:\t{0}", ayer);
                    Console.WriteLine("ECF count: {0}\n", ECFValues.Count);

                    Console.WriteLine("\n\tType\t   Name\t\t       L.UP   Value                  Description");
                    Console.WriteLine("\t--------   ----------------    ----   --------------------   -----------");

                    foreach (CustomField cf in projECFs)
                    {

                        // 3A. Distinguish CF values that are simple from those that use entries in lookup tables.
                        if (!cf.LookupTable.ServerObjectIsNull.HasValue ||
                                            (cf.LookupTable.ServerObjectIsNull.HasValue && cf.LookupTable.ServerObjectIsNull.Value))
                        {
                            if (ECFValues[cf.InternalName] == null)
                            {   // 3B. Partial implementation. Not usable.
                                String textValue = "is not set";
                                Console.WriteLine("\t{0, -8}   {1, -20}        ***{2}",
                                    cf.FieldType, cf.Name, textValue);
                            }
                            else   // 3C. Simple, friendly value for the user
                            {
                                // CustomFieldType is a CSOM enumeration of ECF types.
                                switch (cf.FieldType)
                                {

                                    case CustomFieldType.COST:
                                        decimal costValue = (decimal)ECFValues[cf.InternalName];
                                        Console.WriteLine("\t{0, -8}   {1, -20}        {2, -22}",
                                            cf.FieldType, cf.Name, costValue.ToString("C"));
                                        break;

                                    case CustomFieldType.DATE:
                                    case CustomFieldType.FINISHDATE:
                                    case CustomFieldType.DURATION:
                                    case CustomFieldType.FLAG:
                                    case CustomFieldType.NUMBER:
                                    case CustomFieldType.TEXT:
                                        Console.WriteLine("\t{0, -8}   {1, -20}        {2, -22}",
                                            cf.FieldType, cf.Name, ECFValues[cf.InternalName]);
                                        break;

                                }

                            }
                        }
                        else         //3D. The ECF uses a Lookup table to store the values.
                        {
                            Console.Write("\t{0, -8}   {1, -20}", cf.FieldType, cf.Name);

                            String[] entries = (String[])ECFValues[cf.InternalName];

                            foreach (String entry in entries)
                            {
                                var luEntry = ProjectCont1.LoadQuery(cf.LookupTable.Entries
                                        .Where(e => e.InternalName == entry));

                                ProjectCont1.ExecuteQuery();

                                Console.WriteLine(" Yes    {0, -22}  {1}", luEntry.First().FullValue, luEntry.First().Description);
                            }
                        }


                    }

                    Console.WriteLine("     ------------------------------------------------------------------------\n");

                }

            }

                    Console.Write("\nPress any key to exit: ");
                        Console.ReadKey(false);

        }    //end of using






     


        private void ya() {


            ProjectCont1 = new csom.ProjectContext(pwaPath);
            SecureString securePassword = new SecureString();
            foreach (char c in passWord.ToCharArray())
            {
                securePassword.AppendChar(c);
            }
            ProjectCont1.Credentials = new SharePointOnlineCredentials(userName, securePassword);
            var creds = new SharePointOnlineCredentials(userName, securePassword);
            int j = 0;
            ProjectCont1.Load(ProjectCont1.Projects,
                   c=>c.IncludeWithDefaultProperties
                   (
                   pr => pr.CustomFields,
                   pr => pr.IncludeCustomFields,
                   pr => pr.IncludeCustomFields.CustomFields
                   ));
            ProjectCont1.ExecuteQuery();
            foreach (PublishedProject item in ProjectCont1.Projects)
            {
                if (item.Name == "PROYECTO TEST MOISES")
                {
                    foreach (var cust in item.IncludeCustomFields.FieldValues)
                    {
                        string customfieldID = cust.Key;                        
                        string CsutomfieldValue = cust.Value.ToString();
                        Console.WriteLine("\n{0}. {1} \t{2} \n", j++, customfieldID, CsutomfieldValue);
                        

                       
                    }
                }
            }


        }
                     

        private void conn()
        {

            ProjectCont1 = new csom.ProjectContext(pwaPath);
            SecureString securePassword = new SecureString();
            foreach (char c in passWord.ToCharArray())
            {
                securePassword.AppendChar(c);
            }
            ProjectCont1.Credentials = new SharePointOnlineCredentials(userName, securePassword);
            var creds = new SharePointOnlineCredentials(userName, securePassword);

        }

        private void customfield() {


            ProjectCont1 = new csom.ProjectContext(pwaPath);
            SecureString securePassword = new SecureString();
            foreach (char c in passWord.ToCharArray())
            {
                securePassword.AppendChar(c);
            }
            var creds = new SharePointOnlineCredentials(userName, securePassword);


            var fieldId = new Guid("012e39a5-c4cd-e911-b075-00155d8c9a02");
            var resourceId = new Guid("012e39a5-c4cd-e911-b075-00155d8c9a02");

            using (var ctx = new ProjectContext(pwaPath))
            {

                ctx.Credentials = creds;

                // Retrieve Enterprise Custom Field
                var field = ctx.CustomFields.GetByGuid(fieldId);

                // Load InernalName property, we will use it to get the value
                ctx.Load(field,
                    x => x.InternalName);

                // Execture prepared query on server side
                ctx.ExecuteQuery();

                var fieldInternalName = field.InternalName;

                // Retrieve recource by its Id
                var resource = ctx.EnterpriseResources.GetByGuid(resourceId);

                // !
                // Load custom field value
                ctx.Load(resource,
                    x => x[fieldInternalName]);
                ctx.ExecuteQuery();
                // 
            }

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

