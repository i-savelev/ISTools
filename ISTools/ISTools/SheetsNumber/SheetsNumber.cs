using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using System;
using System.Linq;
using System.Data;
using ISTools.Forms;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Collections;
using System.Text.RegularExpressions;

namespace ISTools
{
    [Autodesk.Revit.Attributes.TransactionAttribute(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class SheetsNumber : IExternalCommand
    {
        //---PluginsManager---//
        public static string IS_TAB_NAME => "ISTools";
        public static string IS_NAME => "Номера листов";
        public static string IS_IMAGE => "ISTools.Resources.Sheet_number32.png";
        public static string IS_DESCRIPTION => "Инструкция: \nПлагин позволяет автоматически перенумеровать выбранные наборы листов";
        //---PluginsManager---//
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            List<ObjSheet> objSheetList = new List<ObjSheet>();

            List<string> selectedLists = new List<string>(); 

            List<TreeNode> selectedNode = new List<TreeNode>();

            List<string> parametersName = new List<string>();

            var shets = new FilteredElementCollector(doc).
                OfCategory(BuiltInCategory.OST_Sheets).
                WhereElementIsNotElementType().
                Cast<ViewSheet>().
                ToList(); 

            DataTable dtSheets = new DataTable();

            List<string> tempSheets = new List<string>();

            dtSheets.Columns.Add("Продпросмотр");

            SheetNumberForm window = new SheetNumberForm();

            BrowserOrganization org = BrowserOrganization.GetCurrentBrowserOrganizationForSheets(doc);

            foreach (var sheet in shets)
            {
                ObjSheet sheet_obj = new ObjSheet(sheet.Name, sheet.SheetNumber);
                sheet_obj.Elem = sheet;

                List<FolderItemInfo> folderfields = org.GetFolderItems(sheet.Id).ToList();

                foreach (FolderItemInfo info in folderfields)
                {
                    string groupheader = info.Name;
                    sheet_obj.GroupList.Add(groupheader);
                }
                objSheetList.Add(sheet_obj);
            }

            var sheetParameters = shets[0].Parameters;

            foreach (Parameter param in sheetParameters)
            {
                parametersName.Add(param.Definition.Name);
            }

            foreach (ObjSheet objSheet in objSheetList)
            {
                AddToTree(objSheet, window.treeView1);
            }
            parametersName.Sort();


            window.Text = "Нумерация листов";
            window.comboBox1.DataSource = parametersName;
            window.groupBox4.Text = "Список листов";
            window.groupBox3.Text = "Настройка";
            window.dataGridView3.DataSource = dtSheets;
            window.treeView1.AfterCheck += TreeViewAfterCheck;
            window.button1.Click += (s, e) => { SheetSelect(); };
            window.button1.Text = "Предпросмотр";
            window.button2.Click += (s, e) => { SheetChangeNumber(); };
            window.button2.Text = "Перезаписать номера";
            window.textBox4.Text = "Префикс";
            window.textBox5.Text = "Начальное значение";
            window.textBox6.Text = "Суффикс";
            window.checkBox1.Text = "Дублировать номер в параметр";
            window.checkBox2.Text = "Дублировать номер в параметр без префикса и суффикса";
            window.textBox10.Text = "Параметр";
            window.textBox8.Text = "Длина номера листа";
            window.textBox9.Text = "К номеру листа можно добавить начальные нули. Для этого задайте минимальную длину строки, представляющей номер листа. Например, если номер листа 15 и задать длину З, в результате получится 015. Чтобы не добавлять начальных нулей, задайте длину 1";
            window.ShowDialog();

            void SheetSelect()
            {
                selectedLists.Clear();
                selectedNode.Clear();
                dtSheets.Clear();
                tempSheets.Clear();

                foreach (TreeNode parent_Node in window.treeView1.Nodes)
                {
                    if (parent_Node.Nodes.Count > 0)
                    {
                        NodeChecker(parent_Node, selectedLists, selectedNode);
                    }
                    else
                    {
                        if (parent_Node.Checked)
                        {
                            selectedLists.Add(parent_Node.Text);
                            selectedNode.Add(parent_Node);
                        }
                    }
                }

                int firstNumber = (int)window.numericUpDown1.Value;
                int zerosCount = (int)window.numericUpDown2.Value;
                string format = new string('0', zerosCount);
                var uniqueNameList = selectedLists.Select(obj => obj.Split(':')[1].TrimStart(' ')).ToList();
                foreach (var s in selectedLists)
                {
                    string formattedNumber = firstNumber.ToString(format);
                    var preview = $"{s.Split(':')[0]} → {window.textBox2.Text}{formattedNumber}{window.textBox3.Text}: {s.Split(':')[1].TrimStart(' ')}";
                    var new_number = $"{window.textBox2.Text}{formattedNumber}{window.textBox3.Text}";
                    
                    if (objSheetList.Any(p => p.Number == new_number))
                    {
                        var list = objSheetList.FirstOrDefault(p => p.Number == new_number);
                        if (!uniqueNameList.Contains(list.Name))
                        {
                            preview = "!УЖЕ ИСПОЛЬЗУЕТСЯ-" + preview;
                        }
                    }
                    dtSheets.Rows.Add
                    (
                        preview
                    );
                    tempSheets.Add(preview);
                    firstNumber++;
                }
                window.dataGridView3.DataSource = dtSheets;
            }

            void SheetChangeNumber()
            {
                if (!tempSheets.Any(p => p.Contains("!УЖЕ ИСПОЛЬЗУЕТСЯ")))
                {
                    using (Transaction tx = new Transaction(doc))
                    {
                        tx.Start("Перенумерация - подготовка");
                        foreach (var s in selectedNode)
                        {
                            var list = objSheetList.FirstOrDefault(p => p.Number == s.Text.Split(':')[0]);
                            var tempNumber = $"{s.Text.Split(':')[0]}%TEMP%";
                            list.Elem.SheetNumber = tempNumber;
                            list.Number = tempNumber;
                        }
                        tx.Commit();

                        tx.Start("Перенумерация");
                        int firstNumber = (int)window.numericUpDown1.Value;
                        int zerosCount = (int)window.numericUpDown2.Value;
                        string format = new string('0', zerosCount);
                        window.progressBar1.Value = 0;
                        window.progressBar1.Maximum = selectedNode.Count;
                        window.progressBar1.Step = 1;
                        foreach (var s in selectedNode)
                        {
                            window.progressBar1.PerformStep();
                            string formattedNumber = firstNumber.ToString(format);
                            var newNumber = $"{window.textBox2.Text}{formattedNumber}{window.textBox3.Text}";
                            var list = objSheetList.FirstOrDefault(p => p.Number.Replace("%TEMP%", "") == s.Text.Split(':')[0]);
                            list.Elem.SheetNumber = newNumber;
                            var newName = $"{window.textBox2.Text}{formattedNumber}{window.textBox3.Text}: {s.Text.Split(':')[1].TrimStart(' ')}";
                            s.Text = newName;

                            if (window.checkBox1.Checked)
                            {
                                try
                                {
                                    list.Elem.LookupParameter(window.comboBox1.Text).Set(newNumber);
                                }
                                catch { }
                            }
                            else if (window.checkBox2.Checked)
                            {
                                try
                                {
                                    list.Elem.LookupParameter(window.comboBox1.Text).Set(formattedNumber.ToString());
                                }
                                catch { }
                            }
                            firstNumber++;
                        }
                        tx.Commit();
                        window.progressBar1.Value = selectedNode.Count;
                    }

                    var shetsNew = new FilteredElementCollector(doc).
                        OfCategory(BuiltInCategory.OST_Sheets).
                        WhereElementIsNotElementType().
                        Cast<ViewSheet>().
                        ToList();

                    objSheetList.Clear();

                    foreach (var sheet in shetsNew)
                    {
                        ObjSheet sheetObj = new ObjSheet(sheet.Name, sheet.SheetNumber);
                        sheetObj.Elem = sheet;

                        List<FolderItemInfo> folderFields = org.GetFolderItems(sheet.Id).ToList();

                        foreach (FolderItemInfo info in folderFields)
                        {
                            string groupheader = info.Name;
                            sheetObj.GroupList.Add(groupheader);
                        }

                        objSheetList.Add(sheetObj);
                    }
                }
                else TaskDialog.Show("Предупреждение", "Лист с таким имененем уже есть");
            }

            void NodeChecker(TreeNode parentNode, List<string> selectedList, List<TreeNode> selectedNodes)
            {
                foreach (TreeNode childNode in parentNode.Nodes)
                {
                    if (childNode.Nodes.Count > 0)
                    {
                        NodeChecker(childNode, selectedList, selectedNodes);
                    }
                    else
                    {
                        if (childNode.Checked)
                        {
                            selectedList.Add(childNode.Text);
                            selectedNodes.Add(childNode);
                        }
                    }   
                }
            }

            void AddToTree(ObjSheet sheet, TreeView treeView)
            {
                // Проверяем наличие хотя бы одного элемента в списке Group_list
                if (sheet.GroupList.Any())
                {
                    TreeNode rootNode = null;
                    var existingRoot = treeView.Nodes.Cast<TreeNode>().FirstOrDefault(n => n.Text == sheet.GroupList[0]);
                    if (existingRoot != null)
                    {
                        rootNode = existingRoot;
                    }

                    else
                    {
                        rootNode = new TreeNode(sheet.GroupList[0]);
                        treeView.Nodes.Add(rootNode);
                    }

                    // Переменная для хранения текущего узла
                    TreeNode currentNode = rootNode;

                    // Проходим по остальным элементам списка Group_list
                    for (int i = 1; i < sheet.GroupList.Count; i++)
                    {
                        string group = sheet.GroupList[i];

                        // Проверяем, существует ли уже такой узел среди детей текущего узла
                        var existingChild = currentNode.Nodes.Cast<TreeNode>().FirstOrDefault(n => n.Text == group);

                        if (existingChild != null)
                        {
                            // Если такой узел найден, делаем его текущим
                            currentNode = existingChild;
                        }

                        else
                        {
                            // Если нет, создаем новый узел и добавляем его
                            var newNode = new TreeNode(group);
                            currentNode.Nodes.Add(newNode);
                            currentNode = newNode;
                        }
                    }
                    // Последний узел - это Name
                    currentNode.Nodes.Add(new TreeNode($"{sheet.Number}: {sheet.Name}"));
                }

                else
                {
                    // Если список Group_list пуст, просто добавляем узел с именем
                    treeView.Nodes.Add(new TreeNode($"{sheet.Number}: {sheet.Name}"));
                }

                treeView.TreeViewNodeSorter = new NodeSorter();
                treeView.Sort();
            }

            void TreeViewAfterCheck(object sender, TreeViewEventArgs e)
            {
                // Отключить повторное вхождение в обработчик при массовом изменении состояния
                window.treeView1.AfterCheck -= TreeViewAfterCheck;

                try
                {
                    UpdateChildNodes(e.Node, e.Node.Checked);
                }
                finally
                {
                    window.treeView1.AfterCheck += TreeViewAfterCheck;
                }
            }

            void UpdateChildNodes(TreeNode parentNode, bool isChecked)
            {
                foreach (TreeNode childNode in parentNode.Nodes)
                {
                    childNode.Checked = isChecked;

                    // Рекурсивное обновление вложенных узлов
                    if (childNode.Nodes.Count > 0)
                    {
                        UpdateChildNodes(childNode, isChecked);
                    }
                }
            }
            return Result.Succeeded;
        }
    }
    public class NodeSorter : IComparer
    {
        public NodeSorter() { }

        public int Compare(object _left, object _right)
        {
            // Преобразуем входные объекты в TreeNode
            TreeNode left = _left as TreeNode;
            TreeNode right = _right as TreeNode;

            if (left == null || right == null)
                throw new ArgumentException("Оба объекта должны быть типа TreeNode.");

            // Извлекаем текст из TreeNode
            string leftText = left.Text.Split(':')[0];
            string rightText = right.Text.Split(':')[0];

            // Разбиваем строки на текстовую и числовую части
            string leftPrefix = ExtractPrefix(leftText);
            string rightPrefix = ExtractPrefix(rightText);

            // Сравниваем текстовые префиксы
            int prefixComparison = string.Compare(leftPrefix, rightPrefix, StringComparison.Ordinal);
            if (prefixComparison != 0)
                return prefixComparison;

            // Если префиксы равны, сравниваем числовые части
            string[] leftParts = ExtractNumberParts(leftText);
            string[] rightParts = ExtractNumberParts(rightText);

            int maxLength = Math.Max(leftParts.Length, rightParts.Length);
            for (int i = 0; i < maxLength; i++)
            {
                // Получаем числовые значения для текущей части
                int leftNumber = i < leftParts.Length && IsNumeric(leftParts[i]) ? int.Parse(leftParts[i]) : 0;
                int rightNumber = i < rightParts.Length && IsNumeric(rightParts[i]) ? int.Parse(rightParts[i]) : 0;

                // Сравниваем числа
                if (leftNumber != rightNumber)
                    return leftNumber - rightNumber;
            }

            // Если все части равны, строки считаются равными
            return 0;
        }

        // Вспомогательный метод для извлечения текстового префикса
        private string ExtractPrefix(string text)
        {
            // Удаляем все числа и точки, оставляя только текстовый префикс
            return Regex.Replace(text, @"[\d\.]+", "").Trim();
        }

        // Вспомогательный метод для извлечения числовых частей
        private string[] ExtractNumberParts(string text)
        {
            // Удаляем текстовый префикс и разбиваем оставшуюся строку по точкам
            string numericPart = Regex.Replace(text, @"[^\d\.]", "").Trim('.');
            return numericPart.Split('.');
        }

        // Вспомогательный метод для проверки, является ли строка числом
        private bool IsNumeric(string value)
        {
            return int.TryParse(value, out _);
        }
    }
}

