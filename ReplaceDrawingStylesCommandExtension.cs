using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Forms;
using System.Xml.Linq;
using Autodesk.Connectivity.Explorer.Extensibility;
using Autodesk.Connectivity.Extensibility.Framework;
using Autodesk.Connectivity.WebServices;
using Autodesk.Connectivity.WebServicesTools;
using Autodesk.DataManagement.Client.Framework.Vault.Currency.Connections;
using ACW = Autodesk.Connectivity.WebServices;

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

                MessageBox.Show("Dodaje zadanie do job processora....");
            }

        }

        private void AddJob(Connection vaultConn, File file )
        {

            JobParam[] jobParams = new JobParam[1];

                jobParams[0] = new JobParam()
                {
                    Name = "FileId",
                    Val = file.Id.ToString()
                };

            vaultConn.WebServiceManager.JobService.AddJob("KRATKI.ReplaceDrawingStyles", $"KRATKI.ReplaceDrawingStyles: {file.Name}", jobParams, 1);
        }

        private List<File> GetFiles(Connection vaultConn, object sender, CommandItemEventArgs e)
        {
            long[] selectionId = e.Context.CurrentSelectionSet.Select(n => n.Id).ToArray();
            File[] selectionFiles = vaultConn.WebServiceManager.DocumentService.GetLatestFilesByMasterIds(selectionId);

            CommandItem commandItem = (CommandItem)sender;

            List<File> files = new List<File>();
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

                files = files.Concat(assocFilesArray.Where(f => !files.Any(existing => existing.Name == f.Name)))
                    .ToList();
            }
            else
            {
                files = selectionFiles.ToList();
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
