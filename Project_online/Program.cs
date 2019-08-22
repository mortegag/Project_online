using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Project.Server.ClientOM;
using System.Security;
using Microsoft.ProjectServer.Client;
using Microsoft.SharePoint.Client;
using System.IO;

namespace Project_online
{
    class Program
    {
        static void Main(string[] args)
        {
            int id = 0;
            string nombre = null;
            string fecha = null;

            using (ProjectContext projContext = new ProjectContext("https://aigpanama.sharepoint.com/sites/Proyectos-TI/"))
            {
                SecureString password = new SecureString();
                foreach (char c in "Coco.1961".ToCharArray()) password.AppendChar(c);
                //Using SharePoint method to load Credentials
                projContext.Credentials = new SharePointOnlineCredentials("mortega@innovacion.gob.pa", password);


                var projects = projContext.Projects;
                projContext.Load(projects);
                int j = 1;
                projContext.ExecuteQuery();

                string mensaje = "";
                string filePath = @"c:\temp" + @"\proyectos.txt";

                foreach (PublishedProject pubProj in projContext.Projects)
                {
                   // Console.WriteLine("\n{0}. {1}   {2} \t{3} \n", j++, pubProj.Id, pubProj.Name, pubProj.CreatedDate);

                    string[] proyecto = new string[4];
                    proyecto[0] = j++.ToString();
                    proyecto[1] = pubProj.Id.ToString();
                    proyecto[2] = pubProj.Name;
                    proyecto[3] = pubProj.CreatedDate.ToString();

                    using (System.IO.StreamWriter writer = new StreamWriter(filePath, true))
                    {
                        writer.WriteLine("Message :" + proyecto + "<br/>" + Environment.NewLine + "ver archivo csv" +
                         "" + Environment.NewLine + "Date :" + DateTime.Now.ToString());
                        writer.WriteLine(Environment.NewLine + "-----------------------------------------------------------------------------" + Environment.NewLine);
                    }


                }

                
            }


           
        }

        public void texto() {

            string mensaje = "Proceso de lectura de Microsoft Project Online";
            string filePath = "c:\temp" + @"\proyectos.txt";
            using (System.IO.StreamWriter writer = new StreamWriter(filePath, true))
            {
                writer.WriteLine("Message :" + mensaje + "<br/>" + Environment.NewLine + "ver archivo csv" +
                 "" + Environment.NewLine + "Date :" + DateTime.Now.ToString());
                writer.WriteLine(Environment.NewLine + "-----------------------------------------------------------------------------" + Environment.NewLine);
            }


        }


    }
}
