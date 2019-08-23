using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.ProjectServer.Client;
using Microsoft.SharePoint.Client;
using System.Security;

class Program{

        const string pwaPath = "https://ServerName/pwa"; //Change the path to your
        const string userName = "LoginName@DomainName.com"; //PWA Username

        const string passWord = "password"; //PWA Password
        static ProjectContext projectContext; //Initialize the ProjectContext Object

        static void  project()//Main(string[] args)
        {
                 projectContext = new ProjectContext(pwaPath);
                 SecureString securePassword = new SecureString();
                  foreach (char c in passWord.ToCharArray())
                  { securePassword.AppendChar(c); }

                projectContext.Credentials = new SharePointOnlineCredentials(userName,securePassword);

        }
        
        private static void ListPublishedProjects() {

            projectContext.Load(projectContext.Projects); //Load Project from PWA
            projectContext.ExecuteQuery();
            Console.WriteLine("nList of all Published Projects:n");
            foreach (PublishedProject pubProj in projectContext.Projects)
            { Console.WriteLine("nt{0}", pubProj.Name);}

            Console.ReadLine();
        }


        private static void CreateNewProject()
        {

            ProjectCreationInformation newProj = new ProjectCreationInformation();
            newProj.Id = Guid.NewGuid();
            newProj.Name = "Test Project";
            newProj.Description = "This Project is created via CSOM";
            newProj.Start = DateTime.Today.Date;
            newProj.EnterpriseProjectTypeId = new Guid("09fa52b4-059b-4527-926e-99f9be96437a");
            PublishedProject newPublishedProj = projectContext.Projects.Add(newProj);
            projectContext.Projects.Update();
            projectContext.ExecuteQuery();

        }


        private static void GetTasksOfProject()
        {
            projectContext.Load(projectContext.Projects); //Load Projects from PWA
            projectContext.ExecuteQuery();
            Guid ProjectGuid = new Guid("b7dfde50-7a2d-e611-9bf8-681729bb2204");
            var project = projectContext.Projects.First(proj => proj.Id == ProjectGuid);

            //Here you can also use the property proj.Name
            projectContext.Load(project.Tasks); //Load Tasks of Project from PWA
            
            projectContext.ExecuteQuery();
            Console.WriteLine("Tasks:n");
            
            foreach (PublishedTask task in project.Tasks)
            { Console.WriteLine("nt{0}", task.Name); }
            Console.ReadLine();
        }

        private static void CreateNewTask()
        {

            Guid ProjectGuid = new Guid("b7dfde50-7a2d-e611-9bf8-681729bb2204");
            var Project = projectContext.Projects.GetByGuid(ProjectGuid);
            var draftProject = Project.CheckOut();
            TaskCreationInformation task = new TaskCreationInformation();
            task.Id = Guid.NewGuid();
            task.Name = "New Task";
            task.Start = DateTime.Today.Date;
            task.IsManual = false;
            DraftTask draftTask = draftProject.Tasks.Add(task);
            draftProject.Update();
            draftProject.Publish(true); //Publish and check-in the Project
            projectContext.ExecuteQuery();

        }


        private static void UpdateProjectCF(){

            Guid CustomFieldGuid =new Guid("99072634-db70-e611-80cb-00155da46e22");
            Guid ProjectGuid = new Guid("3531cb1f-51ce-4a27-a45a-590339b99aed");
            var proj = projectContext.Projects.GetByGuid(ProjectGuid);
            var draftProj = proj.CheckOut();
            projectContext.Load(projectContext.CustomFields); //Load Custom Fields from PWA
            projectContext.ExecuteQuery();
            var CustomField = projectContext.CustomFields.First(CF => CF.Id == CustomFieldGuid);

            //Here you can also use the property CF.Name
            var CustomFieldValue= "New Value";
            string internalName = CustomField.InternalName.ToString(); //Get internal name of      custom field
            draftProj.SetCustomFieldValue(internalName, CustomFieldValue); //Set custom field value
            draftProj.Publish(true); //Publish and check-in the Project         
            projectContext.ExecuteQuery();

        }



        private static void UpdateTaskCF(){

            projectContext.Load(projectContext.Projects);//Load Projects from PWA
            projectContext.ExecuteQuery();
            projectContext.Load(projectContext.CustomFields);//Load Custom Fields from PWA
            projectContext.ExecuteQuery();
            Guid ProjectGuid = new Guid("3531cb1f-51ce-4a27-a45a-590339b99aed");
            Guid TaskGuid = new Guid("56b0ba40-da78-e611-80cb-00155da4672c");
            Guid CustomFieldGuid = new Guid("3c3dfd8b-da78-e611-80cb-00155da4672c");
            var project = projectContext.Projects.GetByGuid(ProjectGuid);
            DraftProject draftProject = project.CheckOut(); //CheckOut Project to make it  editable
            projectContext.Load(draftProject.Tasks);//Load tasks of Project
            projectContext.ExecuteQuery();
            var taskToEdit = draftProject.Tasks.First(task => task.Id == TaskGuid);//Get the task to be updated, Here you can also use the property task.Name
            var CustomField = projectContext.CustomFields.First(CF => CF.Id == CustomFieldGuid);
            var CustomFieldValue = "New Value";
            string internalName = CustomField.InternalName.ToString();//Get internal name of custom field
            taskToEdit[internalName] = CustomFieldValue;//Set custom field value
            draftProject.Publish(true);//Publish and Check-in the project
            projectContext.ExecuteQuery();

}
}
