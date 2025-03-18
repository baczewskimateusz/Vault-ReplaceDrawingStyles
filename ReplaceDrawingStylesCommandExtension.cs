using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Autodesk.Connectivity.Explorer.Extensibility;
using Autodesk.Connectivity.WebServices;
using Autodesk.DataManagement.Client.Framework.Vault.Currency.Connections;

namespace ReplaceDrawingStyles
{
    internal class ReplaceDrawingStylesCommandExtension : IExplorerExtension
    {
        public IEnumerable<CommandSite> CommandSites()
        {
            CommandSite fileContextCmdSite = new CommandSite("ReplaceDrawingStylesCommand.FileContextMenu", "Podmiana styli rysunków")
            {
                Location = CommandSiteLocation.FileContextMenu,
                DeployAsPulldownMenu = true
            };

            CommandItem changePropertiesCmdItemIam = new CommandItem("ReplaceDrawingStylesCommandIam", "Wszystkie pliki")
            {
                NavigationTypes = new SelectionTypeId[] { SelectionTypeId.File, SelectionTypeId.FileVersion },
                MultiSelectEnabled = false
            };

            CommandItem changePropertiesCmdItemMultiple = new CommandItem("ChangePropertiesCommandIpt", "Zaznaczone pliki")
            {
                NavigationTypes = new SelectionTypeId[] { SelectionTypeId.File, SelectionTypeId.FileVersion },
                MultiSelectEnabled = true
            };

            fileContextCmdSite.AddCommand(changePropertiesCmdItemIam);
            fileContextCmdSite.AddCommand(changePropertiesCmdItemMultiple);
            changePropertiesCmdItemIam.Execute += ReplaceDrawingStylesCommandHandler;
            changePropertiesCmdItemMultiple.Execute += ReplaceDrawingStylesCommandHandler;
            List<CommandSite> sites = new List<CommandSite>();
            sites.Add(fileContextCmdSite);
            return sites;
        }

        private void ReplaceDrawingStylesCommandHandler(object sender, CommandItemEventArgs e)
        {
            Connection vaultConn = e.Context.Application.Connection;
            

            DialogResult answer = MessageBox.Show("Czy na pewno chcesz podmienić style w rysunkach?", "Uwaga!", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (answer == DialogResult.Yes)
            {
                List<File> files = GetFiles(vaultConn, sender, e);

                foreach (File file in files)
                {
                    AddJob(vaultConn, file);
                }
                MessageBox.Show("Zmiana styli została dodana do kolejki zadań.", "Informacja" ,MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void AddJob(Connection vaultConn, File file )
        {
            List<string> nameOfNotAdded  = new List<string>();  
            JobParam[] jobParams = new JobParam[]
            {
                new JobParam()
                {
                    Name = "FileId",
                    Val = file.Id.ToString()
                }
            };
            try
            {
                vaultConn.WebServiceManager.JobService.AddJob("KRATKI.ReplaceDrawingStyles", $"KRATKI.ReplaceDrawingStyles: {file.Name}", jobParams, 50);
            }
            catch 
            {
                nameOfNotAdded.Add(file.Name); 
            }

            if(nameOfNotAdded.Count > 0)
            {
                string result = string.Join(", ", nameOfNotAdded);
                MessageBox.Show($"Wystąpił błąd podaczas dodawania zadania zamiany styli dla modeli: {result}. Skontaktuj się z administratorem.");
            }
            
        }

        private List<File> GetFiles(Connection vaultConn, object sender, CommandItemEventArgs e)
        {
            long[] selectionId = e.Context.CurrentSelectionSet.Select(n => n.Id).ToArray();
            File[] selectionFiles = vaultConn.WebServiceManager.DocumentService.GetLatestFilesByMasterIds(selectionId);
            CommandItem commandItem = (CommandItem)sender;
            List<File> files = new List<File>();

            try
            {
                if (commandItem.Label == "Wszystkie pliki")
                {
                    FileAssocLite[] associationFiles = vaultConn.WebServiceManager.DocumentService.GetFileAssociationLitesByIds(
                    new long[] { selectionFiles[0].Id },
                    FileAssocAlg.LatestTip,
                    FileAssociationTypeEnum.Dependency,
                    false,
                    FileAssociationTypeEnum.Dependency,
                    true,
                    false,
                    false,
                    false
                    );

                    long[] assocDirectFiles = associationFiles
                        .Where(file => !file.ExpectedVaultPath.Contains("Biblioteka") & !file.ExpectedVaultPath.Contains("Content Center"))
                        .Select(n => n.CldFileId).ToArray();

                    File[] assocFilesArray = vaultConn.WebServiceManager.DocumentService.GetFilesByIds(assocDirectFiles);

                    files.Add(selectionFiles[0]);

                    foreach (File file in assocFilesArray)
                    {
                        if (!files.Any(f => f.Name == file.Name))
                        {
                            files.Add(file);
                        }
                    }
                }
                else
                {
                    files = selectionFiles.ToList();
                }
            }
            catch (Exception ex) 
            {
                MessageBox.Show($"Wystąpił błąd: {ex.Message}. Skontaktuj się z administratorem.");
            }
            

            return files;
        }

        public IEnumerable<CustomEntityHandler> CustomEntityHandlers()
        {
            return null;
        }
        public IEnumerable<DetailPaneTab> DetailTabs()
        {
            return null;
        }
        public IEnumerable<string> HiddenCommands()
        {
            return null;
        }
        public void OnLogOff(IApplication application)
        {
        }
        public void OnLogOn(IApplication application)
        {

        }
        public void OnShutdown(IApplication application)
        {
        }
        public void OnStartup(IApplication application)
        {
        }
    }
}
