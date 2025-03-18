using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Autodesk.Connectivity.Extensibility.Framework;
using Autodesk.Connectivity.JobProcessor.Extensibility;
using Autodesk.Connectivity.WebServices;
using Autodesk.Connectivity.WebServicesTools;
using Autodesk.DataManagement.Client.Framework.Vault.Currency.Connections;
using Autodesk.DataManagement.Client.Framework.Vault.Currency.Entities;
using Autodesk.DataManagement.Client.Framework.Vault.Results;
using Autodesk.DataManagement.Client.Framework.Vault.Settings;
using Inventor;
using ACW = Autodesk.Connectivity.WebServices;
using VDF = Autodesk.DataManagement.Client.Framework;

[assembly: ApiVersion("15.0")]
[assembly: ExtensionId("45ec1c53-3ab3-471e-82df-0f4d145a8469")]


namespace ReplaceDrawingStyles
{
    public class JobExtension : IJobHandler
    {
        private static string JOB_TYPE = "Kratki.ReplaceDrawingStyles";

        #region IJobHandler Implementation
        public bool CanProcess(string jobType)
        {
            return jobType == JOB_TYPE;
        }

        public JobOutcome Execute(IJobProcessorServices context, IJob job)
        {
            try
            {
                ReplaceStyles(context, job);
                return JobOutcome.Success;
            }
            catch (Exception ex)
            {
                context.Log(ex, "B³¹d podczas aktualizowania styli: " + ex.ToString() + " ");
                context.Connection.WebServiceManager.Dispose();
                return JobOutcome.Failure;
            }
        }

        private void ReplaceStyles(IJobProcessorServices context, IJob job)
        {
            Connection vaultConn = context.Connection;
            InventorServer mInv = context.InventorObject as InventorServer;

         
            // Znajdz plik 
            ACW.File file = GetFileById(vaultConn, job.Params["FileId"]);
           
            try
            {
               
                // Ustaw projekt InventorServer
                SetProject(context, mInv);
                
                // Znajdz plik rysunek do zmiany
                ACW.File drawingToChangeFile = GetDrawingToChangeFile(vaultConn, file);

                // Pobranie pliku rysunku dla wskazanej czeœci/z³o¿enia
                FileAcquisitionResult drawingFileResult = DownloadFile(drawingToChangeFile, vaultConn, true);

                //Otwórz rysunek
                DrawingDocument drawingDocumentToChange = OpenDocument(mInv, drawingFileResult);


                //znajdŸ standardowy rysunek (rysunek z którego bierzemy style/tabelkê/ramkê
                ACW.File standardDrawFile = GetStandardDrawingFile(vaultConn, file);

                // SprawdŸ ramki i bloki tytu³owe
                Dictionary<Sheet, bool> hasDocumentBorder = HasBorder(drawingDocumentToChange);
                Dictionary<Sheet, bool> hasDocumentTitleBlock = HasTitleBlock(drawingDocumentToChange);

                // Wyczyœæ wszystkie elementy rysunku
                ClearDrawingElements(drawingDocumentToChange);

                // pobierz stanadrdowy rysunek
                FileAcquisitionResult standardDraw = DownloadFile(standardDrawFile, vaultConn);

                //Otwórz rysunek standardowy
                DrawingDocument oStandardDrawDocument = OpenDocument(mInv, standardDraw);

                //string standardDrawLocalPath = GetLocalPath(standardDraw);
                //DrawingDocument oStandardDrawDocument = (DrawingDocument)mInv.Documents.Open(standardDrawLocalPath, false);

                //dodaj tabelkê rysunkowa i ramkê
                AddTitleBLockAndBorder(drawingDocumentToChange, hasDocumentBorder, hasDocumentTitleBlock, oStandardDrawDocument);

                // zmieñ style
                ChangeStyles(mInv, drawingDocumentToChange);

                //zmieñ pozycjê listy czêœci
                ChangeTablePosition(drawingDocumentToChange, hasDocumentBorder, hasDocumentTitleBlock, mInv);

                //zmien pozycjê opisów
                ChangeNotePosition(drawingDocumentToChange, hasDocumentTitleBlock, mInv);

                //zmiana jêzyka na polski
                //ChangeBendNoteLanguage(drawingDocumentToChange);

                //zapisz i uaktualinj plik w vaulcie
                drawingDocumentToChange.Update();
                drawingDocumentToChange.Save();
                UpdateFile(vaultConn, drawingDocumentToChange, drawingToChangeFile, drawingFileResult);

                //dodaj zadanie aktualizacji po³¹czeñ miedzy plikem a itemem
                AddUpadteItemLinksJob(drawingToChangeFile, vaultConn);

                // dodaj zadanie eksportowanie PDF
                AddExportPDFJob(drawingToChangeFile, vaultConn);
            }
            catch
            {
                if (file.CheckedOut)
                {
                    vaultConn.WebServiceManager.DocumentService.UndoCheckoutFile(file.MasterId, out ByteArray downloadTicket);
                }
                throw;
            }
            finally
            {
                vaultConn.WebServiceManager.Dispose();
                mInv = null;
                vaultConn = null;
            }
        }

        private void ChangeBendNoteLanguage(DrawingDocument doc)
        {
            string bendNoteText;
            foreach(Sheet sheet in doc.Sheets)
            {   
                BendNotes bendNotes = sheet.DrawingNotes.BendNotes;

                if (bendNotes.Count > 0)
                {
                    foreach(BendNote bendNote in bendNotes)
                    {
                        bendNoteText = bendNote.Text;
                        if (bendNote.Text.Contains("UP"))
                        {
                            bendNoteText = bendNoteText.Replace("DOWN", "W DÓ£");
                            bendNote.Text = bendNoteText;
                        }
                        else if (bendNote.Text.Contains("DOWN"))
                        {
                            bendNoteText = bendNoteText.Replace("DOWN", "W DÓ£");
                            bendNote.Text = bendNoteText;
                        }
                    }
                }
            }
        }

        private void AddTitleBLockAndBorder(DrawingDocument drawingDocumentToChange, Dictionary<Sheet, bool> hasDocumentBorder, Dictionary<Sheet, bool> hasDocumentTitleBlock, DrawingDocument oStandardDrawDocument)
        {
            try
            {
                TitleBlockDefinition standardTitleBLock = GetTitleBLockDefinitionByName("kratki", oStandardDrawDocument);
                TitleBlockDefinition newTitleBLock = standardTitleBLock.CopyTo((_DrawingDocument)drawingDocumentToChange, true);

                BorderDefinition standardBorder = GetBorderDefinitionByName("kratki", oStandardDrawDocument);
                BorderDefinition newBorderDefinition = standardBorder.CopyTo((_DrawingDocument)drawingDocumentToChange, true);

                Sheets sourceDocumentSheets = drawingDocumentToChange.Sheets;
                foreach (Sheet sheet in sourceDocumentSheets)
                {
                    sheet.Activate();
                    if (hasDocumentTitleBlock[sheet] == true)
                    {
                        sheet.AddTitleBlock(newTitleBLock);
                    }
                    if (hasDocumentBorder[sheet] == true)
                    {
                        sheet.AddBorder(newBorderDefinition);
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Wyst¹pi³ b³¹d podczas dodawania tabeli lub ramki: {drawingDocumentToChange.DisplayName}. Metoda: {nameof(AddTitleBLockAndBorder)}. Szczegó³y: {ex.Message}", ex);
            }
           
        }

        private void ClearDrawingElements(DrawingDocument drawingDocumentToChange)
        {
            try
            {
                ClearBorder(drawingDocumentToChange);
                ClearTitleBLock(drawingDocumentToChange);
                ClearReferencedBorder(drawingDocumentToChange);
                ClearReferencedTitleBLock(drawingDocumentToChange);
            }
            catch(Exception ex)
            {
                throw new Exception($"Wyst¹pi³ b³¹d podczas otwierania rysunku dla pliku: {drawingDocumentToChange.DisplayName}. Metoda: {nameof(ClearDrawingElements)}. Szczegó³y: {ex.Message}", ex);
            }
            
        }

        private ACW.File GetStandardDrawingFile(Connection vaultConn, ACW.File file)
        {
            string fileCatName = file.Cat.CatName;

            ACW.File standardDrawFile = null;

            if (fileCatName == "Zespó³" || fileCatName == "Zespó³ spawany")
            {
                standardDrawFile = FindFile("Standard monta¿.idw", vaultConn);
            }
            else if (fileCatName == "Czêœæ" || fileCatName == "Blacha" || fileCatName == "Kszta³townik z ContentCenter")
            {
                standardDrawFile = FindFile("Standard czêœæ.idw", vaultConn);
            }
            else { throw new Exception("Nieznana kategoria pliku"); }

            return standardDrawFile;
        }

        private DrawingDocument OpenDocument(InventorServer mInv, FileAcquisitionResult drawingFileResult)
        {
            string drawingFileLocalPath;
            DrawingDocument drawingDocument;
            try
            {
                drawingFileLocalPath = GetLocalPath(drawingFileResult);

                drawingDocument = (DrawingDocument)mInv.Documents.Open(drawingFileLocalPath, false);
            }
            catch(Exception ex)
            {
                throw new Exception($"Wyst¹pi³ b³¹d podczas otwierania rysunku dla pliku: {drawingFileResult.File.EntityName}. Metoda: {nameof(OpenDocument)}. Szczegó³y: {ex.Message}", ex);
            }

            if (drawingFileLocalPath == null)
            {
                throw new Exception($"Nie znaleziono rysunku dla pliku: {drawingFileResult.File.EntityName}");
            }


            return drawingDocument;
        }

        private static ACW.File GetDrawingToChangeFile(Connection vaultConn, ACW.File file)
        {
            ACW.File drawingToChangeFile;
            try
            {
                FileAssocArray associations = vaultConn.WebServiceManager.DocumentService
                     .GetLatestFileAssociationsByMasterIds(new long[] { file.MasterId }, FileAssociationTypeEnum.Dependency, false, FileAssociationTypeEnum.None, false, false, false, false).First();


                drawingToChangeFile = associations.FileAssocs
                    .FirstOrDefault(fa => fa.ParFile.Name.Contains(".idw") &&
                                          System.IO.Path.GetFileNameWithoutExtension(fa.ParFile.Name) == System.IO.Path.GetFileNameWithoutExtension(file.Name))
                    ?.ParFile;
            }
            catch (Exception ex)
            {
                throw new Exception($"Wyst¹pi³ b³¹d podczas szukania rysunku dla pliku: {file.Name}. Metoda: {nameof(GetDrawingToChangeFile)}. Szczegó³y: {ex.Message}", ex);
            }
            
            if (drawingToChangeFile == null)
            {
                throw new Exception($"Nie znaleziono rysunku dla pliku: {file}");
            }

            return drawingToChangeFile;
        }

        private void ChangeTablePosition(DrawingDocument drawDocument, Dictionary<Sheet, bool> hasDocumentBorder, Dictionary<Sheet, bool> hasDocumentTitleBlock,InventorServer mInv)
        {
            PartsList oPartsList =null;
            TransientGeometry oTg = mInv.TransientGeometry;
            double oTableNewPostitionX;
            double oTableNewPostitionY;
            foreach (Sheet sheet in drawDocument.Sheets)
            {
                sheet.Activate();
                if(sheet.PartsLists.Count > 0) 
                {
                    oPartsList = sheet.PartsLists[1];    
                }
                else { return; }

                if ((hasDocumentTitleBlock[sheet] == true && hasDocumentBorder[sheet] == true) || hasDocumentTitleBlock[sheet])
                {
                    TitleBlock oTitleBlock = sheet.TitleBlock;
                    oTableNewPostitionX = sheet.Width - 0.6;
                    oTableNewPostitionY = oTitleBlock.RangeBox.MaxPoint.Y + oPartsList.RangeBox.MaxPoint.Y - oPartsList.RangeBox.MinPoint.Y;

                    Point2d newTablePostion = oTg.CreatePoint2d(oTableNewPostitionX, oTableNewPostitionY);
                    oPartsList.Position = newTablePostion;
                }
                else
                {
                    oTableNewPostitionX = sheet.Width-0.6;
                    oTableNewPostitionY = 0.6 + oPartsList.RangeBox.MaxPoint.Y - oPartsList.RangeBox.MinPoint.Y;

                    Point2d newTablePostion = oTg.CreatePoint2d(oTableNewPostitionX, oTableNewPostitionY);
                    oPartsList.Position = newTablePostion;
                }
            }
        }

        private void ChangeNotePosition(DrawingDocument drawDocument,Dictionary<Sheet, bool> hasDocumentTitleBlock, InventorServer mInv)
        {
            PartsList oPartsList = null;
            TransientGeometry oTg = mInv.TransientGeometry;
            Point2d positionPoint=null;
            GeneralNote generalNote = null;

            foreach (Sheet sheet in drawDocument.Sheets)
            {
                sheet.Activate();
                PartsLists oPartsLists = sheet.PartsLists;

                if (sheet.DrawingNotes.Count == 0)
                {
                    return;
                }
               
                foreach (DrawingNote drawingNote in sheet.DrawingNotes)
                {

                    if (drawingNote.Type == ObjectTypeEnum.kGeneralNoteObject && (drawingNote.Text.ToLower().Contains("uwaga") || drawingNote.Text.ToLower().Contains("uwagi")))
                    {
                        generalNote = (GeneralNote)drawingNote;
                        
                       
                        if (hasDocumentTitleBlock[sheet] == true)
                        {
                            TitleBlock titleBlock = sheet.TitleBlock;
                            if (titleBlock.RangeBox.MinPoint.X - 1 > generalNote.FittedTextWidth)
                            {

                                positionPoint = oTg.CreatePoint2d(titleBlock.RangeBox.MinPoint.X - generalNote.FittedTextWidth, generalNote.FittedTextHeight + 0.6);
                            }
                            else
                            {
                                positionPoint = oTg.CreatePoint2d(sheet.Width - 0.8 - generalNote.FittedTextWidth, titleBlock.RangeBox.MaxPoint.Y + generalNote.FittedTextHeight);
                            }

                        }
                        else if (oPartsLists.Count > 0)
                        {
                            oPartsList = oPartsLists[1];

                            if (oPartsList.RangeBox.MinPoint.X - 1 > generalNote.FittedTextWidth)
                            {

                                positionPoint = oTg.CreatePoint2d(oPartsList.RangeBox.MinPoint.X - generalNote.FittedTextWidth, generalNote.FittedTextHeight + 0.6);
                            }
                            else
                            {
                                positionPoint = oTg.CreatePoint2d(sheet.Width - 0.8 - generalNote.FittedTextWidth, oPartsList.RangeBox.MaxPoint.Y + generalNote.FittedTextHeight);
                            }
                        }
                        else
                        {
                            positionPoint = oTg.CreatePoint2d(sheet.Width - 0.8 - generalNote.FittedTextWidth, generalNote.FittedTextHeight + 0.6);
                        }
                    }
                    else if (drawingNote.Type == ObjectTypeEnum.kGeneralNoteObject)
                    {
                        generalNote = (GeneralNote)drawingNote;
                        if (oPartsLists.Count > 0)
                        {
                            if(generalNote.RangeBox.MaxPoint.X > oPartsList.RangeBox.MinPoint.X && generalNote.RangeBox.MinPoint.Y < oPartsList.RangeBox.MaxPoint.Y)
                            {
                                positionPoint = oTg.CreatePoint2d(generalNote.Position.X, generalNote.FittedTextHeight + oPartsList.RangeBox.MaxPoint.Y);
                            }
                        }
                        else if (hasDocumentTitleBlock[sheet] == true)
                        {
                            TitleBlock titleBlock = sheet.TitleBlock;
                            if (generalNote.RangeBox.MaxPoint.X > titleBlock.RangeBox.MinPoint.X && generalNote.RangeBox.MinPoint.Y < titleBlock.RangeBox.MaxPoint.Y)
                            {
                                positionPoint = oTg.CreatePoint2d(generalNote.Position.X, generalNote.FittedTextHeight + titleBlock.RangeBox.MaxPoint.Y);
                            }
                        }
                    }
                    if (positionPoint != null) 
                    {
                        generalNote.Position = positionPoint;
                    }

                    
                }
            }
        }
        private Dictionary<Sheet, bool> HasTitleBlock(DrawingDocument drawingDocument)
        {
            Sheets documentSheets = drawingDocument.Sheets;
            Dictionary<Sheet, bool> titleBlockDict = new Dictionary<Sheet, bool>();
            foreach (Sheet sheet in documentSheets)
            {
                TitleBlock titleBlock = sheet.TitleBlock;
                if (titleBlock == null) { titleBlockDict.Add(sheet, false); }
                else { titleBlockDict.Add(sheet, true); }
            }
            return titleBlockDict;  
        }

        private Dictionary<Sheet, bool> HasBorder(DrawingDocument drawingDocument)
        {
            Sheets documentSheets = drawingDocument.Sheets;
            Dictionary<Sheet, bool> borderDict = new Dictionary<Sheet, bool>();
            foreach (Sheet sheet in documentSheets)
            {
                Border border = sheet.Border;
                if (border == null) { borderDict.Add(sheet, false);}
                else { borderDict.Add(sheet, true); }
            }

            return borderDict;
        }

        public void UpdateFile(Connection vaultConn, DrawingDocument drawingDocument, ACW.File drawingFile, FileAcquisitionResult fileResult)
        {
            Stream fileCont = new FileStream(drawingDocument.FullFileName, FileMode.Open, FileAccess.Read);
            byte[] fileContents = new byte[fileCont.Length];
            //ByteArray uploadTicket = UploadFileResource(vaultConn.WebServiceManager, drawingDocument.FullFileName, fileContents);

            FileAssocArray[] drawAssoc = vaultConn.WebServiceManager.DocumentService
                .GetLatestFileAssociationsByMasterIds(new long[] { drawingFile.MasterId }, FileAssociationTypeEnum.None, false, FileAssociationTypeEnum.Dependency, false, false, false, false);
         
            List<ACW.FileAssocParam> mFileAssocParams = new List<ACW.FileAssocParam>();
            if (drawAssoc.FirstOrDefault().FileAssocs != null)
            {
                foreach (ACW.FileAssoc item in drawAssoc.FirstOrDefault().FileAssocs)
                {
                    ACW.FileAssocParam mFileAssocParam = new ACW.FileAssocParam();
                    mFileAssocParam.CldFileId = item.CldFile.Id;
                    mFileAssocParam.ExpectedVaultPath = item.ExpectedVaultPath;
                    mFileAssocParam.RefId = item.RefId;
                    mFileAssocParam.Source = item.Source;
                    mFileAssocParam.Typ = item.Typ;

                    mFileAssocParams.Add(mFileAssocParam);
                }
            }
            string localPath = drawingDocument.FullDocumentName;
            drawingDocument.Close(true);
            
            FileIteration mUploadedFile = vaultConn.FileManager.CheckinFile(fileResult.File,
                "Created by Job Processor",
                false, mFileAssocParams.ToArray(),
                null,
                false,
                null,
                ACW.FileClassification.DesignDocument,
                false,
                fileResult.LocalPath);

            fileCont.Dispose();
            fileCont.Close();

            if (System.IO.File.Exists(localPath))
            {
                System.IO.FileInfo fileInfo = new FileInfo(localPath);
                fileInfo.IsReadOnly = false;
                fileInfo.Delete();
            }
        }

        public void ChangeStyles(InventorServer mInv, DrawingDocument drawingDocument)
        {
            TransientObjects transientObjects = mInv.TransientObjects;
            DrawingStylesManager stylesManager = drawingDocument.StylesManager;
            DrawingStandardStylesEnumerator drawingStandardStyles = stylesManager.StandardStyles;

            foreach (Style style in drawingStandardStyles)
            {
                if (style.Name.Contains("Vault"))
                {
                    stylesManager.ActiveStandardStyle = (DrawingStandardStyle)style;
                }
            }
            
            Styles drawingStyles = stylesManager.Styles;
            string styleName;
            foreach (Style replacingStyle in drawingStyles)
            {
                string name = replacingStyle.Name;

                if (replacingStyle.Name.Contains("Vault") && replacingStyle.StyleType != StyleTypeEnum.kStandardStyleType)
                {

                    ObjectCollection objects = transientObjects.CreateObjectCollection();

                    styleName = replacingStyle.Name;
                    styleName = styleName.Replace(" VaultAddin", "");

                    try
                    {
                        foreach (Style styleToReplace in drawingStyles)
                        {
                            if (styleToReplace.Name == styleName && styleToReplace.Type == replacingStyle.Type && styleToReplace.InUse)
                            {
                                objects.Add(styleToReplace);
                                stylesManager.ReplaceStyles(objects, replacingStyle, false);
                                break;
                            }
                        }
                    }
                    catch
                    {
                    }
                    
                }
  
            }
        }
        private void ClearReferencedTitleBLock(DrawingDocument oStandardDrawDocument)
        {
            TitleBlockDefinitions titleBlockDefinitions = oStandardDrawDocument.TitleBlockDefinitions;
            foreach (TitleBlockDefinition titleBlock in titleBlockDefinitions)
            {
                if (titleBlock.IsReferenced == false)
                {
                    titleBlock.Delete();
                }
            }
        }
        private void ClearReferencedBorder(DrawingDocument oStandardDrawDocument)
        {
            BorderDefinitions borderDefiniotions = oStandardDrawDocument.BorderDefinitions;
            foreach (BorderDefinition border in borderDefiniotions)
            {
                if (border.IsReferenced == false)
                {
                    border.Delete();
                }
            }
        }
        public void ClearBorder(DrawingDocument drawingDocument)
        {
            Sheets documentSheets = drawingDocument.Sheets; 
            foreach (Sheet sheet in documentSheets)
            {
                Border border = sheet.Border;
                if (border != null)
                {
                    sheet.Border.Delete();
                }
            }
        }
        public void ClearTitleBLock(DrawingDocument drawingDocument)
        {
            Sheets documentSheets = drawingDocument.Sheets;
            foreach (Sheet sheet in documentSheets)
            {   TitleBlock titleBlock = sheet.TitleBlock;
                if (titleBlock != null)
                {
                    sheet.TitleBlock.Delete();
                }
            }
        }

        public BorderDefinition GetBorderDefinitionByName(string name, DrawingDocument drawingDocument)
        {
            BorderDefinitions borderDefinitions = drawingDocument.BorderDefinitions;

            BorderDefinition standardBorder = null;

            foreach (BorderDefinition borderDefinition in borderDefinitions)
            {
                if (borderDefinition.Name.ToLower().Contains(name))
                {
                    standardBorder = borderDefinition;
                }
            }

            return standardBorder;
        }

        public TitleBlockDefinition GetTitleBLockDefinitionByName(string name, DrawingDocument drawingDocument)
        {
            TitleBlockDefinitions titleBlockDefinitions = drawingDocument.TitleBlockDefinitions;
            TitleBlockDefinition standardTitleBLock = null;

            foreach (TitleBlockDefinition titleBlockDefinition in titleBlockDefinitions)
            {
                if (titleBlockDefinition.Name.ToLower().Contains(name))
                {
                    standardTitleBLock = titleBlockDefinition;
                }
            }

            return standardTitleBLock;
        }

        private void AddUpadteItemLinksJob(ACW.File file, Connection vaultCon)
        {
            
            Item item = vaultCon.WebServiceManager.ItemService.GetItemsByFileId(file.Id).FirstOrDefault();

            if (item == null) { return; }

            JobParam[] jobParams = new JobParam[]
            { new JobParam()
                {
                Name = "FileId",
                Val = file.Id.ToString()
                }
            };

            try
            {
                vaultCon.WebServiceManager.JobService.AddJob("KRATKI.UpdateItemLinks", $"KRATKI.UpdateItemLinks: {file.Name}", jobParams, 10);
            }
            catch (Exception ex)
            {
                throw new Exception($"Wyst¹pi³ b³¹d podczas dodawania zadania do job processora. {ex.InnerException}.");
            }
        }

        private void AddExportPDFJob(ACW.File file, Connection vaultCon)
        {
            JobParam[] jobParams = new JobParam[]
            { 
                new JobParam()
                {
                    Name = "FileMasterId",
                    Val = file.MasterId.ToString()
                },
                new JobParam()
                {
                    Name = "FileName",
                    Val = file.Name
                },
                new JobParam()
                {
                    Name = "EntID",
                    Val = file.MasterId.ToString()
                },
                new JobParam()
                {
                    Name = "EntityClassId",
                    Val = "FILE" 
                },
                new JobParam()
                {
                    Name = "ExportFomats",
                    Val = "PDF"
                }
            };

            try
            {
                vaultCon.WebServiceManager.JobService.AddJob("ASP.ExportJob.PDF", $"ASP.ExportJob.PDF - {file.Name}", jobParams, 10);
            }
            catch (Exception ex)
            {
                throw new Exception($"Wyst¹pi³ b³¹d podczas dodawania zadania do job processora. {ex.InnerException}.");
            }
        }

        public static ByteArray UploadFileResource(WebServiceManager svcmgr, string filename, byte[] fileContents)
        {
            svcmgr.FilestoreService.FileTransferHeaderValue = new FileTransferHeader();
            svcmgr.FilestoreService.FileTransferHeaderValue.Identity = Guid.NewGuid();
            svcmgr.FilestoreService.FileTransferHeaderValue.Extension = System.IO.Path.GetExtension(filename);
            svcmgr.FilestoreService.FileTransferHeaderValue.Vault = svcmgr.WebServiceCredentials.VaultName;

            int MAX_FILE_PART_SIZE = 49 * 1024 * 1024;

            ByteArray uploadTicket = new ByteArray();
            int bytesTotal = (fileContents != null ? fileContents.Length : 0);
            int bytesTransferred = 0;
            do
            {
                int bufferSize = (bytesTotal - bytesTransferred) % MAX_FILE_PART_SIZE;
                byte[] buffer = null;
                if (bufferSize == bytesTotal)
                {
                    buffer = fileContents;
                }
                else
                {
                    buffer = new byte[bufferSize];
                    Array.Copy(fileContents, (long)bytesTransferred, buffer, 0, (long)bufferSize);
                }

                svcmgr.FilestoreService.FileTransferHeaderValue.Compression = Compression.None;
                svcmgr.FilestoreService.FileTransferHeaderValue.IsComplete = (bytesTransferred + bufferSize) == bytesTotal ? true : false;
                svcmgr.FilestoreService.FileTransferHeaderValue.UncompressedSize = bufferSize;

                using (var fileContentsStream = new MemoryStream(fileContents))
                    uploadTicket.Bytes = svcmgr.FilestoreService.UploadFilePart(fileContentsStream);
                bytesTransferred += bufferSize;

            } while (bytesTransferred < bytesTotal);

            return uploadTicket;
        }



        public ACW.File FindFile(string templateName, Connection vaultConn)
        {
            ACW.SrchCond mSrchCond = new ACW.SrchCond()
            {
                SrchOper = 3,
                PropTyp = PropertySearchType.AllProperties,
                SrchTxt = templateName,
                SrchRule = SearchRuleType.Must
                
            };

            string bookmark = string.Empty;
            ACW.SrchStatus status = null;
            List<ACW.File> totalResults = new List<ACW.File>();

            Autodesk.Connectivity.WebServicesTools.WebServiceManager mWsMgr = vaultConn.WebServiceManager;

            try
            {

                ACW.File result = mWsMgr.DocumentService.FindFilesBySearchConditions(new ACW.SrchCond[] { mSrchCond },
                           null, null, false, true, ref bookmark, out status).First(n => !n.Name.Contains(".dwf") && !n.Name.Contains("BH"));

                if (result != null) { return result; }
            }
            catch (Exception ex)
            {
                throw new Exception($"Wyst¹pi³ b³¹d podczas szukania pliku. {ex.InnerException}. Nazwa pliku: {templateName}");
            }


            return null;
        }

        public string GetLocalPath(FileAcquisitionResult file)
        {
            string localPath = file.LocalPath.FullPath;
            return localPath;
        }
        public FileAcquisitionResult DownloadFile(ACW.File downloadFile, Connection vaultConn, bool checkOut = false)
        {
            FileAcquisitionResult file = null;
            try
            {
                VDF.Vault.Currency.Entities.FileIteration fileIteration = new VDF.Vault.Currency.Entities.FileIteration(vaultConn, downloadFile);
                VDF.Vault.Settings.AcquireFilesSettings settings = new VDF.Vault.Settings.AcquireFilesSettings(vaultConn);
                settings.OptionsRelationshipGathering.FileRelationshipSettings.VersionGatheringOption = VDF.Vault.Currency.VersionGatheringOption.Latest;
                AcquireFilesSettings.AcquisitionOption options;
                
                if (checkOut)
                {
                    options = settings.DefaultAcquisitionOption = VDF.Vault.Settings.AcquireFilesSettings.AcquisitionOption.Checkout |
                                                VDF.Vault.Settings.AcquireFilesSettings.AcquisitionOption.Download;
                }
                else
                {
                    options = settings.DefaultAcquisitionOption = VDF.Vault.Settings.AcquireFilesSettings.AcquisitionOption.Download;
                }
                settings.OptionsResolution.OverwriteOption = AcquireFilesSettings.AcquireFileResolutionOptions.OverwriteOptions.ForceOverwriteAll;

                settings.AddFileToAcquire(fileIteration, options);

                AcquireFilesResults results = vaultConn.FileManager.AcquireFiles(settings);

                file = results.FileResults.FirstOrDefault(f => f.LocalPath.FileName.Contains(fileIteration.EntityName));

            }
            catch (Exception ex)
            {
                throw new Exception($"Wyst¹pi³ b³¹d podczas pobierania pliku: {downloadFile.Name}. Metoda: {nameof(DownloadFile)}. {ex.InnerException}");
            }

            return file;
        }


        private ACW.File GetFileById(Connection vaultConn, string fileId)
        {
            ACW.File file = null;
            if (!long.TryParse(fileId, out long id)) 
            { 
                throw new Exception("Nieprawid³owy typ FileId."); 
            }

            try
            {
                WebServiceManager webServiceManager = vaultConn.WebServiceManager;
                DocumentService documentService = webServiceManager.DocumentService;

                file = documentService.GetFileById(id);
               
            }
            catch (Exception ex) 
            {
                throw new Exception($"Wyst¹pi³ b³¹d podczas szukania pliku po ID. Metoda: {nameof(GetFileById)}. Szczegó³y: {ex.Message}", ex);
            }

            if (file == null)
            {
                throw new Exception($"Nie znaleziono pliku dla ID: {id}");
            }

            return file;
        }

        private void SetDesignData(IJobProcessorServices context, InventorServer mInv)
        {
            ACW.Folder mDesignDataFolder;
            Connection connection = context.Connection;
            Autodesk.Connectivity.WebServicesTools.WebServiceManager mWsMgr = connection.WebServiceManager;

            try
            {
               //get the projects drawingFile object for download
                ACW.PropDef[] filePropDefs = mWsMgr.PropertyService.GetPropertyDefinitionsByEntityClassId("FILE");
                ACW.PropDef mNamePropDef = filePropDefs.Single(n => n.SysName == "Name");
                ACW.SrchCond mSrchCond = new ACW.SrchCond()
                {
                    PropDefId = mNamePropDef.Id,
                    PropTyp = ACW.PropertySearchType.SingleProperty,
                    SrchOper = 3, // is equal
                    SrchRule = ACW.SearchRuleType.Must,
                    SrchTxt = "Design Data"
                };
                string bookmark = string.Empty;
                ACW.SrchStatus status = null;
                List<ACW.Folder> totalResults = new List<ACW.Folder>();
                while (status == null || totalResults.Count < status.TotalHits)
                {
                    ACW.Folder[] results = mWsMgr.DocumentService.FindFoldersBySearchConditions(new ACW.SrchCond[] { mSrchCond },
                        null, null, false, ref bookmark, out status);
                    if (results != null)
                        totalResults.AddRange(results);
                    else
                        break;
                }
                if (totalResults.Count == 1)
                {
                    mDesignDataFolder = totalResults[0];
                }
                else
                {
                    throw new Exception("Job execution stopped due to ambigous project drawingFile definitions; single project drawingFile per Vault expected");
                }

                string localPath = mDesignDataFolder.FullName.Replace(@"$", @"C:\Vault").Replace(@"/",@"\");

                if(System.IO.Directory.Exists(localPath))
                {
                    mInv.FileOptions.DesignDataPath = localPath;
                }
                else 
                {
                    throw new Exception("Nie uda³o siê znaleŸæ lokalnie folderu Design Data");
                }
             

            }
            catch (Exception ex)
            {
                throw new Exception("Zadanie nie mog³o aktywowaæ pliku projektu Inventor – Uwaga: Plik .ipj nie mo¿e byæ wypo¿yczony przez innego u¿ytkownika.", ex.InnerException);
            }
        }
        private void SetProject(IJobProcessorServices context, InventorServer mInv)
        {

            Inventor.DesignProjectManager projectManager;
            Inventor.DesignProject mSaveProject = null, mProject = null;
            String mIpjPath = "";
            String mWfPath = "";
            String mIpjLocalPath = "";
            ACW.File mProjFile;
            VDF.Vault.Currency.Entities.FileIteration mIpjFileIter = null;
            Connection connection = context.Connection;
            Autodesk.Connectivity.WebServicesTools.WebServiceManager mWsMgr = connection.WebServiceManager;

            //download and activate the Inventor Project drawingFile in VaultInventorServer
            try
            {
                //Download enforced ipj drawingFile
                if (mWsMgr.DocumentService.GetEnforceWorkingFolder() && mWsMgr.DocumentService.GetEnforceInventorProjectFile())
                {
                    mIpjPath = mWsMgr.DocumentService.GetInventorProjectFileLocation();
                    mWfPath = mWsMgr.DocumentService.GetRequiredWorkingFolderLocation();
                }
                else
                {
                    throw new Exception("Zadanie wymaga w³¹czenia obu ustawieñ: 'Enforce Workingfolder' i 'Enforce Inventor Project'.");
                }

                String[] mIpjFullFileName = mIpjPath.Split(new string[] { "/" }, StringSplitOptions.None);
                String mIpjFileName = mIpjFullFileName.LastOrDefault();

                //get the projects drawingFile object for download
                ACW.PropDef[] filePropDefs = mWsMgr.PropertyService.GetPropertyDefinitionsByEntityClassId("FILE");
                ACW.PropDef mNamePropDef = filePropDefs.Single(n => n.SysName == "ClientFileName");
                ACW.SrchCond mSrchCond = new ACW.SrchCond()
                {
                    PropDefId = mNamePropDef.Id,
                    PropTyp = ACW.PropertySearchType.SingleProperty,
                    SrchOper = 3, // is equal
                    SrchRule = ACW.SearchRuleType.Must,
                    SrchTxt = mIpjFileName
                };
                string bookmark = string.Empty;
                ACW.SrchStatus status = null;
                List<ACW.File> totalResults = new List<ACW.File>();
                while (status == null || totalResults.Count < status.TotalHits)
                {
                    ACW.File[] results = mWsMgr.DocumentService.FindFilesBySearchConditions(new ACW.SrchCond[] { mSrchCond },
                        null, null, false, true, ref bookmark, out status);
                    if (results != null)
                        totalResults.AddRange(results);
                    else
                        break;
                }
                if (totalResults.Count == 1)
                {
                    mProjFile = totalResults[0];
                }
                else
                {
                    throw new Exception("Job execution stopped due to ambigous project drawingFile definitions; single project drawingFile per Vault expected");
                }

                VDF.Vault.Settings.AcquireFilesSettings mDownloadSettings = new VDF.Vault.Settings.AcquireFilesSettings(connection);
                mDownloadSettings.LocalPath = new VDF.Currency.FolderPathAbsolute(mWfPath);
                mIpjFileIter = new VDF.Vault.Currency.Entities.FileIteration(connection, mProjFile);
                mDownloadSettings.AddFileToAcquire(mIpjFileIter, VDF.Vault.Settings.AcquireFilesSettings.AcquisitionOption.Download);

                VDF.Vault.Results.AcquireFilesResults mDownLoadResult;
                VDF.Vault.Results.FileAcquisitionResult fileAcquisitionResult;
                mDownLoadResult = connection.FileManager.AcquireFiles(mDownloadSettings);
                fileAcquisitionResult = mDownLoadResult.FileResults.FirstOrDefault();
                mIpjLocalPath = fileAcquisitionResult.LocalPath.FullPath;

                projectManager = mInv.DesignProjectManager;
                try
                {
                    if (projectManager.ActiveDesignProject != null && projectManager.ActiveDesignProject.FullFileName != mIpjLocalPath)
                    {
                        mSaveProject = projectManager.ActiveDesignProject;
                    }
                }
                catch (Exception)
                {
                }
                mProject = projectManager.DesignProjects.AddExisting(mIpjLocalPath);

                mProject.Activate();

            }
            catch (Exception ex)
            {
                throw new Exception("Zadanie nie mog³o aktywowaæ pliku projektu Inventor – Uwaga: Plik .ipj nie mo¿e byæ wypo¿yczony przez innego u¿ytkownika.", ex.InnerException);
            }
        }
        public void OnJobProcessorShutdown(IJobProcessorServices context)
        {
            //throw new NotImplementedException();
        }
        public void OnJobProcessorSleep(IJobProcessorServices context)
        {
            //throw new NotImplementedException();
        }
        public void OnJobProcessorStartup(IJobProcessorServices context)
        {
            //throw new NotImplementedException();
        }
        public void OnJobProcessorWake(IJobProcessorServices context)
        {
            //throw new NotImplementedException();
        }

        #endregion IJobHandler Implementation

       


    }
}
