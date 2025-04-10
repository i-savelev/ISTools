using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using System;
using System.Linq;
using System.Data;
using Plugin.Forms;
using System.Collections.Generic;
using System.Windows.Forms;

namespace Plugin
{
    [Autodesk.Revit.Attributes.TransactionAttribute(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class SheetsNumber : IExternalCommand
    {
        //-********-//
        public static string IS_TAB_NAME => "01.Общее";
        public static string IS_NAME => "Номера листов";
        public static string IS_IMAGE => "Plugin.Resources.Sheet_number32.png";
        public static string IS_DESCRIPTION => "Плагин автоматически перенумеровать выбранные наборы листов";
        //-********-//
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            List<Obj_Sheet> obj_sheet_list = new List<Obj_Sheet>();

            List<string> selected_lists = new List<string>(); 

            List<TreeNode> selected_node = new List<TreeNode>();

            List<string> parameters_name = new List<string>();

            var shets = new FilteredElementCollector(doc).
                OfCategory(BuiltInCategory.OST_Sheets).
                WhereElementIsNotElementType().
                Cast<ViewSheet>().
                ToList(); 

            DataTable dt_sheets = new DataTable();

            List<string> temp_sheets = new List<string>();

            dt_sheets.Columns.Add("Продпросмотр");

            SheetNumberForm window = new SheetNumberForm();

            BrowserOrganization org = BrowserOrganization.GetCurrentBrowserOrganizationForSheets(doc);

            foreach (var sheet in shets)
            {
                Obj_Sheet sheet_obj = new Obj_Sheet(sheet.Name, sheet.SheetNumber);
                sheet_obj.Elem = sheet;

                List<FolderItemInfo> folderfields = org.GetFolderItems(sheet.Id).ToList();

                foreach (FolderItemInfo info in folderfields)
                {
                    string groupheader = info.Name;
                    sheet_obj.Group_list.Add(groupheader);
                }
                obj_sheet_list.Add(sheet_obj);
            }

            var sheet_parameters = shets[0].Parameters;

            foreach (Parameter param in sheet_parameters)
            {
                parameters_name.Add(param.Definition.Name);
            }

            foreach (Obj_Sheet obj_sheet in obj_sheet_list)
            {
                AddToTree(obj_sheet, window.treeView1);
            }
            parameters_name.Sort();


            window.Text = "Нумерация листов";
            window.comboBox1.DataSource = parameters_name;
            window.groupBox4.Text = "Список листов";
            window.groupBox3.Text = "Настройка";
            window.dataGridView3.DataSource = dt_sheets;
            window.treeView1.AfterCheck += TreeView1_AfterCheck;
            window.button1.Click += (s, e) => { Sheet_select(); };
            window.button1.Text = "Предпросмотр";
            window.button2.Click += (s, e) => { Sheet_change_number(); };
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

            void Sheet_select()
            {
                selected_lists.Clear();
                selected_node.Clear();
                dt_sheets.Clear();
                temp_sheets.Clear();

                foreach (TreeNode parent_Node in window.treeView1.Nodes)
                {
                    if (parent_Node.Nodes.Count > 0)
                    {
                        Node_Checker(parent_Node, selected_lists, selected_node);
                    }
                    else
                    {
                        if (parent_Node.Checked)
                        {
                            selected_lists.Add(parent_Node.Text);
                            selected_node.Add(parent_Node);
                        }
                    }
                }

                int first_number = (int)window.numericUpDown1.Value;
                int zeros_count = (int)window.numericUpDown2.Value;
                string format = new string('0', zeros_count);
                var unique_name_list = selected_lists.Select(obj => obj.Split(':')[1].TrimStart(' ')).ToList();
                foreach (var s in selected_lists)
                {

                    string formattedNumber = first_number.ToString(format);
                    var a = $"{s.Split(':')[0]} → {window.textBox2.Text}{formattedNumber}{window.textBox3.Text}: {s.Split(':')[1].TrimStart(' ')}";
                    var new_number = $"{window.textBox2.Text}{formattedNumber}{window.textBox3.Text}";
                    
                    if (obj_sheet_list.Any(p => p.Number == new_number))
                    {
                        var list = obj_sheet_list.FirstOrDefault(p => p.Number == new_number);
                        if (!unique_name_list.Contains(list.Name))
                        {
                            a = "!УЖЕ ИСПОЛЬЗУЕТСЯ-" + a;
                        }
                        
                    }
                    dt_sheets.Rows.Add
                    (
                        a
                    );
                    temp_sheets.Add(a);

                    first_number++;
                }
                window.dataGridView3.DataSource = dt_sheets;
            }

            void Sheet_change_number()
            {
                if (!temp_sheets.Any(p => p.Contains("!УЖЕ ИСПОЛЬЗУЕТСЯ")))
                {
                    //try
                    //{
                        using (Transaction tx = new Transaction(doc))
                        {
                            tx.Start("Перенумерация - подготовка");
                            foreach (var s in selected_node)
                            {
                                var list = obj_sheet_list.FirstOrDefault(p => p.Number == s.Text.Split(':')[0]);
                                var temp_number = $"{s.Text.Split(':')[0]}%TEMP%";
                                list.Elem.SheetNumber = temp_number;
                                list.Number = temp_number;
                            }
                            tx.Commit();

                            tx.Start("Перенумерация");
                            int first_number = (int)window.numericUpDown1.Value;
                            int zeros_count = (int)window.numericUpDown2.Value;
                            string format = new string('0', zeros_count);
                            window.progressBar1.Value = 0;
                            window.progressBar1.Maximum = selected_node.Count;
                            window.progressBar1.Step = 1;
                            foreach (var s in selected_node)
                            {
                                window.progressBar1.PerformStep();
                                string formattedNumber = first_number.ToString(format);
                                var new_number = $"{window.textBox2.Text}{formattedNumber}{window.textBox3.Text}";
                                var list = obj_sheet_list.FirstOrDefault(p => p.Number.Replace("%TEMP%", "") == s.Text.Split(':')[0]);
                                list.Elem.SheetNumber = new_number;
                                var new_name = $"{window.textBox2.Text}{formattedNumber}{window.textBox3.Text}: {s.Text.Split(':')[1].TrimStart(' ')}";
                                s.Text = new_name;

                                if (window.checkBox1.Checked)
                                {
                                    try
                                    {
                                        list.Elem.LookupParameter(window.comboBox1.Text).Set(new_number);
                                    }
                                    catch
                                    {
                                    }
                                }
                                else if (window.checkBox2.Checked)
                                {
                                    try
                                    {
                                        list.Elem.LookupParameter(window.comboBox1.Text).Set(formattedNumber.ToString());
                                    }
                                    catch
                                    {
                                    }
                                }
                                first_number++;
                            }
                            tx.Commit();
                            window.progressBar1.Value = selected_node.Count;
                        }

                        var shets_new = new FilteredElementCollector(doc).
                            OfCategory(BuiltInCategory.OST_Sheets).
                            WhereElementIsNotElementType().
                            Cast<ViewSheet>().
                            ToList();

                        obj_sheet_list.Clear();

                        foreach (var sheet in shets_new)
                        {
                            Obj_Sheet sheet_obj = new Obj_Sheet(sheet.Name, sheet.SheetNumber);
                            sheet_obj.Elem = sheet;

                            List<FolderItemInfo> folderfields = org.GetFolderItems(sheet.Id).ToList();

                            foreach (FolderItemInfo info in folderfields)
                            {
                                string groupheader = info.Name;
                                sheet_obj.Group_list.Add(groupheader);
                            }

                            obj_sheet_list.Add(sheet_obj);
                        }
                    //}
                    //catch (Exception ex)
                    //{
                    //    TaskDialog.Show("a", ex.Message);

                    //}
                }
                else TaskDialog.Show("Предупреждение", "Лист с таким имененем уже есть");
                
            }

            void Node_Checker(TreeNode parent_Node, List<string> selected_list, List<TreeNode> selected_nodes)
            {
                foreach (TreeNode child_Node in parent_Node.Nodes)
                {
                    if (child_Node.Nodes.Count > 0)
                    {
                        Node_Checker(child_Node, selected_list, selected_nodes);
                    }
                    else
                    {
                        if (child_Node.Checked)
                        {
                            selected_list.Add(child_Node.Text);
                            selected_nodes.Add(child_Node);
                        }
                    }   
                }
            }

            void AddToTree(Obj_Sheet sheet, TreeView treeView)
            {
                // Проверяем наличие хотя бы одного элемента в списке Group_list
                if (sheet.Group_list.Any())
                {
                    TreeNode rootNode = null;
                    var existingRoot = treeView.Nodes.Cast<TreeNode>().FirstOrDefault(n => n.Text == sheet.Group_list[0]);
                    if (existingRoot != null)
                    {
                        rootNode = existingRoot;
                    }

                    else
                    {
                        rootNode = new TreeNode(sheet.Group_list[0]);
                        treeView.Nodes.Add(rootNode);
                    }

                    // Переменная для хранения текущего узла
                    TreeNode currentNode = rootNode;

                    // Проходим по остальным элементам списка Group_list
                    for (int i = 1; i < sheet.Group_list.Count; i++)
                    {
                        string group = sheet.Group_list[i];

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

            void TreeView1_AfterCheck(object sender, TreeViewEventArgs e)
            {
                // Отключить повторное вхождение в обработчик при массовом изменении состояния
                window.treeView1.AfterCheck -= TreeView1_AfterCheck;

                try
                {
                    Update_Child_Nodes(e.Node, e.Node.Checked);
                }
                finally
                {
                    window.treeView1.AfterCheck += TreeView1_AfterCheck;
                }
            }

            void Update_Child_Nodes(TreeNode parentNode, bool isChecked)
            {
                foreach (TreeNode childNode in parentNode.Nodes)
                {
                    childNode.Checked = isChecked;

                    // Рекурсивное обновление вложенных узлов
                    if (childNode.Nodes.Count > 0)
                    {
                        Update_Child_Nodes(childNode, isChecked);
                    }
                }
            }
            return Result.Succeeded;
        }
    }
    public class NodeSorter : System.Collections.IComparer
    {
        public NodeSorter() { }

        public int Compare(object _left, object _right)
        {
            TreeNode left = _left as TreeNode;
            TreeNode right = _right as TreeNode;
            int leftNumber = 0;
            int rightNumber = 0;

            System.Text.RegularExpressions.Regex leftRegex = new System.Text.RegularExpressions.Regex(@"\d+");
            System.Text.RegularExpressions.Match leftMatch = leftRegex.Match(left.Text);
            System.Text.RegularExpressions.Regex rightRegex = new System.Text.RegularExpressions.Regex(@"\d+");
            System.Text.RegularExpressions.Match rightMatch = leftRegex.Match(right.Text);

            if (leftMatch.Success && rightMatch.Success)
            {
                leftNumber = Convert.ToInt32(leftMatch.Value);
                rightNumber = Convert.ToInt32(rightMatch.Value);
                if (leftNumber != rightNumber)
                    return leftNumber - rightNumber;
                return 0;
            }
            return 0;
        }
    }
}

