using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace SolidWorksAddIn
{
    internal static class SwRename
    {
        public static void Execute(ISldWorks app, IModelDoc2 activeDoc, string newBaseName)
        {
            var docType = SwHelpers.GetDocType(activeDoc.GetPathName());
            string activePath = SwHelpers.GetPathSafe(activeDoc);

            if (string.IsNullOrEmpty(activePath))
            {
                MessageBox.Show("Документ ещё не сохранён. Сначала сохраните файл.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (docType == swDocumentTypes_e.swDocDRAWING)
            {
                RenameFromDrawing(app, activeDoc, newBaseName);
            }
            else if (docType == swDocumentTypes_e.swDocPART || docType == swDocumentTypes_e.swDocASSEMBLY)
            {
                RenameFromModel(app, activeDoc, newBaseName);
            } 
            else if (docType == swDocumentTypes_e.swDocASSEMBLY)
            {
                RenameAssembly(app, activeDoc, newBaseName);
            }
            else
            {
                MessageBox.Show("Тип документа не поддерживается.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private static void RenameFromModel(ISldWorks app, IModelDoc2 modelDoc, string newBaseName)
        {
            string modelOldPath = modelDoc.GetPathName();
            string modelNewPath = SwFileSearch.ComposeRenamedPath(modelOldPath, newBaseName);
            var modelDocType = SwHelpers.GetDocType(modelOldPath);

            // Найти все чертежи ссылающиеся на модель
            var siblingDrawings = SwFileSearch.FindDrawingsSibling(modelOldPath).ToList();

            // Закрываем модель
            app.CloseDoc(modelOldPath);

            // Переименовываем модель
            try
            {
                File.Move(modelOldPath, modelNewPath);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Критическая ошибка: не удалось переименовать файл модели.\n{ex.Message}", "Ошибка файла", MessageBoxButtons.OK, MessageBoxIcon.Error);
                int errs = 0, warns = 0;
                app.OpenDoc6(modelOldPath, (int)modelDocType, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", ref errs, ref warns);
                return;
            }

            // Обновляем ссылки в чертежах
            foreach (var drwOldPath in siblingDrawings)
            {
                app.CloseDoc(drwOldPath);

                bool success = app.ReplaceReferencedDocument(drwOldPath, modelOldPath, modelNewPath);

                if (success)
                {
                    string drwNewPath = SwFileSearch.ComposeRenamedPath(drwOldPath, newBaseName);
                    try
                    {
                        File.Move(drwOldPath, drwNewPath);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Ссылка в чертеже '{Path.GetFileName(drwOldPath)}' обновлена, но не удалось переименовать файл:\n{ex.Message}", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
            }

            // Обновляем ссылки в сборках
            string searchRoot = Path.GetDirectoryName(modelNewPath);
            UpdateReferencesInAssemblies(app, searchRoot, modelOldPath, modelNewPath);

            // Переоткрываем модель
            int finalErrs = 0, finalWarns = 0;
            app.OpenDoc6(modelNewPath, (int)modelDocType, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", ref finalErrs, ref finalWarns);

            MessageBox.Show("Переименование завершено.", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private static void RenameFromDrawing(ISldWorks app, IModelDoc2 drawingDoc, string newBaseName)
        {
            string drwOldPath = drawingDoc.GetPathName();
            string drwNewPath = SwFileSearch.ComposeRenamedPath(drwOldPath, newBaseName);

            var modelPaths = new HashSet<string>(StringComparer.InvariantCultureIgnoreCase);
            SwHelpers.ForEachDrawingView((DrawingDoc)drawingDoc, v =>
            {
                try
                {
                    string refPath = v.GetReferencedModelName();
                    if (!string.IsNullOrEmpty(refPath) && File.Exists(refPath))
                        modelPaths.Add(refPath);
                }
                catch { }
            });

            app.CloseDoc(drwOldPath);

            foreach (var modelOldPath in modelPaths)
            {
                string modelNewPath = SwFileSearch.ComposeRenamedPath(modelOldPath, newBaseName);
                app.CloseDoc(modelOldPath);

                try
                {
                    File.Move(modelOldPath, modelNewPath);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Не удалось переименовать файл модели '{Path.GetFileName(modelOldPath)}'.\n{ex.Message}", "Ошибка файла", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    continue;
                }

                app.ReplaceReferencedDocument(drwOldPath, modelOldPath, modelNewPath);

                // Обновляем ссылки в сборках
                string searchRoot = Path.GetDirectoryName(modelNewPath);
                UpdateReferencesInAssemblies(app, searchRoot, modelOldPath, modelNewPath);
            }

            try
            {
                File.Move(drwOldPath, drwNewPath);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Модели были успешно переименованы, но не удалось переименовать файл чертежа:\n{ex.Message}", "Ошибка файла", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            int errs = 0, warns = 0;
            app.OpenDoc6(drwNewPath, (int)swDocumentTypes_e.swDocDRAWING, (int)swOpenDocOptions_e.swOpenDocOptions_LoadModel, "", ref errs, ref warns);

            MessageBox.Show("Переименование завершено.", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private static void RenameAssembly(ISldWorks app, IModelDoc2 assemblyDoc, string newBaseName)
        {
            if (assemblyDoc == null)
                return;

            string asmOldPath = assemblyDoc.GetPathName();
            string asmNewPath = SwFileSearch.ComposeRenamedPath(asmOldPath, newBaseName);
            var asmDocType = SwHelpers.GetDocType(asmOldPath);

            IAssemblyDoc asm = assemblyDoc as IAssemblyDoc;
            if (asm == null)
            {
                MessageBox.Show("Активный документ не является сборкой.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // 1. Переименовываем компоненты и обновляем ссылки в открытой сборке
            object[] comps = asm.GetComponents(false);
            if (comps != null)
            {
                foreach (var c in comps)
                {
                    IComponent2 comp = c as IComponent2;
                    if (comp == null) continue;

                    string compPath = comp.GetPathName();
                    if (string.IsNullOrEmpty(compPath) || !File.Exists(compPath))
                        continue;

                    string compNewPath = SwFileSearch.ComposeRenamedPath(compPath, newBaseName);

                    // Закрываем компонент перед переименованием
                    app.CloseDoc(compPath);

                    int waitCompMs = 0;
                    while (IsFileLocked(compPath) && waitCompMs < 5000)
                    {
                        System.Threading.Thread.Sleep(100);
                        waitCompMs += 100;
                    }

                    try
                    {
                        File.Move(compPath, compNewPath);
                        // Обновляем ссылку на компонент внутри сборки
                        app.ReplaceReferencedDocument(asmOldPath, compPath, compNewPath);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Не удалось переименовать компонент '{Path.GetFileName(compPath)}': {ex.Message}",
                            "Ошибка файла", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        continue;
                    }
                }
            }

            // 2. Сохраняем сборку под новым именем через API
            int errs = 0, warns = 0;
            bool saved = assemblyDoc.SaveAs4(asmNewPath,
                                             (int)swSaveAsOptions_e.swSaveAsOptions_Silent,
                                             (int)swSaveAsOptions_e.swSaveAsOptions_Copy,
                                             ref errs, ref warns);

            if (!saved || errs != 0)
            {
                MessageBox.Show($"Не удалось сохранить сборку под новым именем '{asmNewPath}'. Ошибки: {errs}, Предупреждения: {warns}",
                    "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // 3. Закрываем старый документ
            app.CloseDoc(asmOldPath);

            // 4. Переоткрываем сборку с новым именем
            errs = 0; warns = 0;
            app.OpenDoc6(asmNewPath, (int)asmDocType, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", ref errs, ref warns);

            MessageBox.Show("Сборка и её компоненты успешно переименованы.", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        // Проверка, занят ли файл
        private static bool IsFileLocked(string path)
        {
            try
            {
                using (FileStream stream = File.Open(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
                {
                    return false;
                }
            }
            catch
            {
                return true;
            }
        }

        /// <summary>
        /// Метод для обновления ссылок в сборке 
        /// </summary>
        /// <param name="app"></param>
        /// <param name="searchRoot"></param>
        /// <param name="oldModelPath"></param>
        /// <param name="newModelPath"></param>
        private static void UpdateReferencesInAssemblies(ISldWorks app, string searchRoot, string oldModelPath, string newModelPath)
        {
            IEnumerable<string> asmFiles;
            try
            {
                asmFiles = Directory.EnumerateFiles(searchRoot, "*.sldasm", SearchOption.TopDirectoryOnly);
            }
            catch
            {
                return;
            }

            foreach (var asmPath in asmFiles)
            {
                if (string.Equals(asmPath, newModelPath, StringComparison.InvariantCultureIgnoreCase))
                    continue;

                try
                {
                    app.CloseDoc(asmPath);
                    app.ReplaceReferencedDocument(asmPath, oldModelPath, newModelPath);
                }
                catch { }
            }
        }
    }
}
