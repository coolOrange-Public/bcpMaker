using bcpDevKit;
using bcpDevKit.Entities;
using bcpDevKit.Entities.Vault;
using Inventor;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace bcpMaker
{
    class Program
    {
        static void LogError(string message)
        {
            string msg = String.Format("{0}: {1}",DateTime.Now, message);
            System.IO.File.AppendAllLines("bcpMaker.log", new string[] { msg });
        }

        static void GetAllFilesFromFolderRecursively(string folder, List<string> files, string ignoreFiles, string ignoreFolders)
        {
            foreach(string ignoreFolder in ignoreFolders.Split(','))
                if (folder.ToLower().EndsWith(ignoreFolder.ToLower())) return;
            try
            {
                files.AddRange(System.IO.Directory.GetFiles(folder).Where(f => !ignoreFiles.ToLower().Split(',').Contains(System.IO.Path.GetExtension(f).ToLower())));
                Console.Write(String.Format("cleccted files {1}. Scanning folder '{0}'\r", folder, files.Count));
                foreach (var f in System.IO.Directory.GetDirectories(folder))
                    GetAllFilesFromFolderRecursively(f, files, ignoreFiles, ignoreFolders);
            }
            catch
            {
                Console.WriteLine(String.Format("ERROR: Folder {0} could not be scanned", folder));
                LogError(String.Format("ERROR: Folder {0} could not be scanned",folder));
            }
        }

        
        static void Main(string[] args)
        {
            Console.BufferWidth = 300;
            Console.WindowWidth=140;
            string msg = String.Format("coolOrange bcpMaker v{0}", System.Reflection.Assembly.GetExecutingAssembly().GetName().Version);
            Console.WriteLine(msg);
            Console.WriteLine("".PadRight(msg.Length, '*'));
            if (System.IO.File.Exists("bcpMaker.log")) System.IO.File.Delete("bcpMaker.log");
            System.IO.File.AppendAllLines("bcpMaker.log", new string[] { msg });
            if (!System.IO.File.Exists("config.xml"))
            {
                Console.WriteLine("config.xml file could not be found nearby the bcpMaker.exe");
                Console.ReadKey();
                return;
            }
            XmlDocument config = new XmlDocument();
            try
            {
                config.Load("config.xml");
            }
            catch (Exception ex)
            {
                Console.WriteLine(String.Format("Error loading config.xml: {0}",ex));
                Console.ReadKey();
                return;
            }
            
            var parameters = config.SelectSingleNode("//PARAMETERS");
            string sourceFolder = parameters.Attributes["SourceFilesFolder"].Value;
            string targetFolder = parameters.Attributes["TargetBCPFolder"].Value;
            var ignore = config.SelectSingleNode("//IGNORE");
            string ignoreFolder = ignore.Attributes["Folders"].Value;
            string ignoreFiles = ignore.Attributes["FileExtensions"].Value;
            var categoryRules = config.SelectNodes("//CATEGORYRULE");
            System.IO.File.AppendAllLines("bcpMaker.log", new string[] { String.Format("Source folder: {0} ",sourceFolder) });
            System.IO.File.AppendAllLines("bcpMaker.log", new string[] { String.Format("target folder: {0} ", sourceFolder) });
            System.IO.File.AppendAllLines("bcpMaker.log", new string[] { String.Format("ignore file extensions: {0} ", ignoreFiles) });
            System.IO.File.AppendAllLines("bcpMaker.log", new string[] { String.Format("ignore folder: {0} ", ignoreFolder) });
            if (!System.IO.Directory.Exists(sourceFolder))
            {
                Console.WriteLine(String.Format("source folder '{0}' does not exist", sourceFolder));
                Console.ReadKey();
                return;
            }
            try
            {
                System.IO.Directory.CreateDirectory(targetFolder);
            }
            catch
            {
                Console.WriteLine(String.Format("target folder '{0}' could not be created", targetFolder));
                Console.ReadKey();
                return;
            }

            #region apprentice server initialization
            ApprenticeServerComponent _invApp = new ApprenticeServerComponent();
            DesignProject InventorProject = _invApp.DesignProjectManager.ActiveDesignProject;
            List<string> libraryFolders = new List<string>();
            string ccPath = InventorProject.ContentCenterPath;
            libraryFolders.Add(ccPath.Replace(sourceFolder, ""));
            string workSpace = InventorProject.WorkspacePath;
            workSpace = workSpace.Replace(sourceFolder, "");
            var libraryPaths = InventorProject.LibraryPaths;
            msg = String.Format("Using Inventor Project File '{0}'", InventorProject.FullFileName);
            Console.WriteLine(msg);
            System.IO.File.AppendAllLines("bcpMaker.log", new string[] { msg });
            ProjectPaths libPaths = InventorProject.LibraryPaths;
            foreach (ProjectPath libpath in libPaths)
            {
                libraryFolders.Add(libpath.Path.TrimStart('.').Replace("\\", "/"));
            }
            #endregion

            List<string> files = new List<string>();
            Console.WriteLine("Scanning folders...");
            GetAllFilesFromFolderRecursively(sourceFolder, files,ignoreFiles, ignoreFolder);
            Console.WriteLine(String.Format("\rCollected files {0}", files.Count).PadRight(Console.BufferWidth, ' '));

            #region adding files to BCP package
            Console.WriteLine("Adding files to BCP package...");
            var bcpSvcBuilder = new BcpServiceBuilder {Version = BcpVersion._2016};
            bcpSvcBuilder.SetPackageLocation(targetFolder);
            var bcpSvc = bcpSvcBuilder.Build();

            int counter = 1;
            Dictionary<string, FileObject> bcpInventorFiles = new Dictionary<string, FileObject>(StringComparer.OrdinalIgnoreCase);
            
            foreach (string file in files)
            {
                string vaultTarget = file.Replace(sourceFolder,"$").Replace("\\","/");
                bool isLibrary = libraryFolders.Any(lf => vaultTarget.Contains(lf.Replace("\\", "/")));
                var bcpFile = bcpSvc.FileService.AddFile(vaultTarget,file,isLibrary);
                string extension = System.IO.Path.GetExtension(file);
                foreach (XmlNode categoryRule in categoryRules)
                {
                    var fileExtensions = categoryRule.SelectSingleNode("//FILEEXTENSION");
                    var category = categoryRule.SelectSingleNode("//CATEGORY");
                    var lifecycleDefinition = categoryRule.SelectSingleNode("//LIFECYCLEDEFINITON");
                    var state = categoryRule.SelectSingleNode("//STATE");
                    var revision = categoryRule.SelectSingleNode("//REVISION");
                    var revisionDefinition = categoryRule.SelectSingleNode("//REVISIONDEFINITION");
                    if (fileExtensions != null && (fileExtensions.InnerText == "" || fileExtensions.InnerText.Split(',').Contains(extension)))
                    {
                        if (category != null && category.InnerText != "") bcpFile.Category = category.InnerText;
                        string stateName = state != null ? state.InnerText : "";
                        if (lifecycleDefinition != null && lifecycleDefinition.InnerText != "") bcpFile.LatestIteration.Setstate(lifecycleDefinition.InnerText,stateName);
                        bcpFile.LatestRevision.SetRevisionDefinition(revisionDefinition != null ? revisionDefinition.InnerText : "", revision != null ? revision.InnerText : "");
                        break;
                    }
                }
                if (file.ToLower().EndsWith("iam") || file.ToLower().EndsWith("ipt") || file.ToLower().EndsWith("idw") || file.ToLower().EndsWith("ipn") || file.ToLower().EndsWith("dwg"))
                {
                    bcpInventorFiles.Add(file, bcpFile);
                    System.IO.File.AppendAllLines("bcpMaker.log", new string[] { String.Format("Adding to reference check list: {0}", file) });
                }
                msg = String.Format("\rAdding file {0} of {1} to BCP package", counter++, files.Count());
                Console.Write(msg);
            }
            Console.WriteLine(String.Format("\rAdded files {0}", bcpSvc.EntitiesTable.Vault.Statistics.Totfiles).PadRight(Console.BufferWidth,' '));
            #endregion


            var inventorFiles = files.Where(file=>file.ToLower().EndsWith("iam") || file.ToLower().EndsWith("ipt") || file.ToLower().EndsWith("idw") || file.ToLower().EndsWith("ipn") || file.ToLower().EndsWith("dwg")).ToList();
            Console.WriteLine("Building references for Inventor files...");
            counter = 0;
            foreach (string iFile in inventorFiles)
            {
                msg = String.Format("\r{1}/{2}: Building references and properties for {0}",iFile,counter++,inventorFiles.Count);
                Console.Write(msg.PadRight(Console.BufferWidth, ' '));
                ApprenticeServerDocument doc = null;
                try
                {
                    doc = _invApp.Open(iFile);
                }
                catch (Exception ex)
                {
                    msg = String.Format("\r\nOpen ERROR!File {0} could not be opened", iFile);
                    Console.WriteLine(msg);
                    System.IO.File.AppendAllLines("bcpMaker.log", new string[] { String.Format("{0}\r\n{1}", msg, ex) });
                    continue;
                }
                var bcpParent = bcpInventorFiles[iFile];
                string databaseRevisionID = "";
                string lastSavedLocation = "";
                object indices = null;
                object oldPaths = null;
                object currentPaths = null;
                try { 
                    doc._GetReferenceInfo(out databaseRevisionID, out lastSavedLocation, out indices, out oldPaths, out currentPaths, true);
                }
                catch (Exception ex)
                {
                    msg = String.Format("\r\nRead ERROR!References for file {0} could not retrieved", iFile);
                    Console.WriteLine(msg);
                    System.IO.File.AppendAllLines("bcpMaker.log", new string[] { String.Format("{0}\r\n{1}", msg, ex) });
                    continue;
                }

                string[] refs = currentPaths as string[];
                int[] idx = indices as int[];
                for (int i = 0; i < refs.Count(); i++)
                {
                    string child = refs[i];
                    if (child == null)
                    {
                        msg = String.Format("\r\nReference Warning!Reference {0} not found for assembly {1}", (oldPaths as string[])[i], iFile);
                        Console.WriteLine(msg);
                        System.IO.File.AppendAllLines("bcpMaker.log", new string[] { msg });
                        continue;
                    }
                    System.IO.File.AppendAllLines("bcpMaker.log", new string[] { String.Format("Reference {0}",child ) });

                    if (bcpInventorFiles.ContainsKey(child))
                    {
                        var bcpChild = bcpInventorFiles[child];
                        var assoc = bcpParent.LatestIteration.AddAssociation(bcpChild.LatestIteration, AssociationObject.AssocType.Dependency);
                        assoc.refId = idx[i].ToString();
                        assoc.needsresolution = true;
                    }
                    else
                    {
                        msg = String.Format("\rPackage ERROR!Child {0} not in bcp package for assembly {1}", (oldPaths as string[])[i], iFile);
                        Console.WriteLine(msg);
                        System.IO.File.AppendAllLines("bcpMaker.log", new string[] { msg });
                    }
                }
                string propName = "";
                try
                { 
                    foreach (PropertySet propSet in doc.PropertySets)
                    {
                        foreach (Property prop in propSet)
                        {
                            propName = prop.Name;
                            if(!propName.Equals("Thumbnail") && !propName.Equals("Part Icon") && prop.Value != null)
                                bcpParent.LatestIteration.AddProperty(prop.DisplayName, prop.Value.ToString());
                        }
                    }
                }
                catch(Exception ex)
                {
                    msg = String.Format("\r\nProperty ERROR!Property {1} for file {0} could not be retrieved", iFile,propName);
                    Console.WriteLine(msg);
                    System.IO.File.AppendAllLines("bcpMaker.log", new string[] { String.Format("{0}\r\n{1}", msg,ex) });
                }
                doc.Close();
            }
            _invApp.Close();

            bcpSvc.Flush();
            System.Diagnostics.Process.Start(targetFolder);
        }
    }
}
