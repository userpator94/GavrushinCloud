using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v3;
using Google.Apis.Drive.v3.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace DriveQuickstart
{
    class Program
    {
        // If modifying these scopes, delete your previously saved credentials
        // at ~/.credentials/drive-dotnet-quickstart.json
        //static string[] Scopes = { DriveService.Scope.DriveFile };
        static string[] Scopes = { DriveService.Scope.Drive,
            "https://www.googleapis.com/auth/dfatrafficking",
            "https://www.googleapis.com/auth/dfareporting" };
        static string ApplicationName = "Drive API .NET Quickstart";
         

        static void Main(string[] args)
        {
            загрузитьНаДиск();
            backgroundWorker1_DoWork(); //получает список файлов с GD 
        }

        private static void backgroundWorker1_DoWork()//подключение к диску и загрузка файлов 
        { // object sender, DoWorkEventArgs e
            UserCredential credential;
            try
            {
                //download = new List<string>();//коллеция ссылок для загрузки
                FileStream str = new FileStream("client_secret.json", FileMode.Open, FileAccess.Read);
                using (FileStream stream =
                    new FileStream("client_secret.json", FileMode.Open, FileAccess.Read))// создание потока для чтения client_secret.json
                {
                    string credPath = System.Environment.GetFolderPath(
                        System.Environment.SpecialFolder.Personal);
                    credPath = Path.Combine(credPath, ".credentials/drive-dotnet-quickstart.json");//путь к файлу управления подключением

                    credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                              GoogleClientSecrets.Load(stream).Secrets,
                              Scopes,
                              GoogleClientSecrets.Load(str).Secrets.ClientId,
                              CancellationToken.None,
                              new FileDataStore(credPath, true)).Result;
                    var service = new Google.Apis.Drive.v3.DriveService(new BaseClientService.Initializer()
                    {
                        HttpClientInitializer = credential,
                        ApplicationName = ApplicationName,
                    });

                    str.Close();

                    // Define parameters of request.
                    Google.Apis.Drive.v3.FilesResource.ListRequest listRequest = service.Files.List();
                    listRequest.PageSize = 400;
                    listRequest.Fields = "nextPageToken, files(id, webViewLink, webContentLink, name, size, mimeType)";//требуемые свойства загружаемых файлов(можно убрать лишнее или добавить требуемое)
                    
                    // List files.                                                                                         
                    IList<Google.Apis.Drive.v3.Data.File> files = listRequest.Execute()
                        .Files;
                    Console.WriteLine("Files:");
                    if (files != null && files.Count > 0)
                    {
                        foreach (var file in files)
                        {
                            if (file.Name == "базы" || file.Name == "нормативы") continue;
                            Console.WriteLine("  {0} (Size: {1} КБ)", file.Name, file.Size/1000.0);
                        }
                    }
                    else
                    {
                        Console.WriteLine("No files found.");
                    }
                    Console.ReadLine();
                }
            }
            catch (Exception x)
            { Console.WriteLine(x.Message); Console.Read(); }
            Console.ReadLine();
        }

        static void загрузитьНаДиск()
        {//object sender, EventArgs e
            UserCredential credential;
            try
            {
                FileStream str = new FileStream("client_secret.json", FileMode.Open, FileAccess.Read);
                using (var stream =
                    new FileStream("client_secret.json", FileMode.Open, FileAccess.Read))// создание потока для чтения client_secret.json
                {
                    string credPath = System.Environment.GetFolderPath(
                        System.Environment.SpecialFolder.Personal);
                    credPath = Path.Combine(credPath, ".credentials/drive-dotnet-quickstart.json");//путь к файлу управления подключением

                    credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                              GoogleClientSecrets.Load(stream).Secrets,
                              Scopes,
                              GoogleClientSecrets.Load(str).Secrets.ClientId,//посоветоваться с Лазаревым
                              CancellationToken.None,
                              new FileDataStore(credPath, true)).Result;
                    var service = new Google.Apis.Drive.v3.DriveService(new BaseClientService.Initializer()
                    {
                        HttpClientInitializer = credential,
                        ApplicationName = ApplicationName,
                    });
                    str.Close();

                    string exeDir = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
                    int firstTime = 0;
                    string folder = null;
                    do
                    {
                        if (firstTime == 0) folder = @"\базы";
                        else folder = @"\нормативы";
                        string path = exeDir + folder;
                        string[] file_list = System.IO.Directory.GetFiles(path, "*.xls");                        

                        //создаём папку
                        var fileMetadata1 = new Google.Apis.Drive.v3.Data.File()
                        {
                            Name = DateTime.Now.ToString() + "_" + System.Environment.MachineName.ToString()+"_"+folder,
                            MimeType = "application/vnd.google-apps.folder"
                        };
                        var request1 = service.Files.Create(fileMetadata1);
                        request1.Fields = "id";
                        var google_folder = request1.Execute();
                        Console.WriteLine("Файлы будут помещены в папку: " + fileMetadata1.Name); //+google_folder.id
                        Console.WriteLine("{0} | {1} файлов", folder, file_list.Length);


                        if (file_list.GetLength(0) > 0)
                        {
                            //Google.Apis.Drive.v3.Data.File body = new Google.Apis.Drive.v3.Data.File();

                            Google.Apis.Drive.v3.Data.File body = new Google.Apis.Drive.v3.Data.File()
                            {
                                Name = "tablesheet",
                                Parents = new List<string>
                                {
                                    google_folder.Id
                                }
                            };

                            //поочереди загружаем файлы
                            for (int i = 0; i < file_list.Length; i++)
                            {
                                body.Name = System.IO.Path.GetFileName(file_list[i]);
                                body.Description = "Описание";
                                if (Path.GetExtension(file_list[i]) == ".xls")
                                    body.MimeType = "application/vnd.ms-excel";
                                if (Path.GetExtension(file_list[i]) == ".xlsx")
                                    body.MimeType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                                if (Path.GetExtension(file_list[i]) != ".xls" && Path.GetExtension(file_list[i]) != ".xlsx") continue;

                                Console.WriteLine("Name:{0} , Type:{1}", body.Name, body.MimeType);
                                byte[] byteArray = System.IO.File.ReadAllBytes(file_list[i]);
                                System.IO.MemoryStream stream2 = new System.IO.MemoryStream(byteArray);
                                Google.Apis.Drive.v3.FilesResource.CreateMediaUpload request = service.Files.Create(body, stream2, body.MimeType);
                                if (request.Upload().Exception != null)
                                {
                                    Console.Write(request.Upload().Exception.Message);
                                    Console.Read();
                                }
                                else { Console.WriteLine("Файл успешно загружен"); }
                            }
                            Console.WriteLine();
                        }                        
                        firstTime++;
                    }
                    while (firstTime < 2);
                                       
                }
            }

            catch (Exception x)
            { Console.WriteLine(x.Message); }
        } //the end


    }
}