using ISTools.Forms;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace ISTools.ISTools.SheetsCopy
{
    internal class SheetsCopyWindowManager
    {
        public SheetCopyForm Window {  get; set; }
        public SheetManager SheetManager { get; set; }
        public SheetsCopyWindowManager(SheetManager sheetManager)
        {
            this.SheetManager = sheetManager;
            SheetCopyForm window = new SheetCopyForm();
            Window = window;
            window.treeView1.AfterCheck += TreeViewAfterCheck;
            window.button1.Click += (s, e) => { SheetCopy(); };
            window.textBox3.Text = "В Списке листов нужно выбрать листы, которые необходимо копировать. После нажатия кнопки \"Копировать листы\" выбранные листы будут скопированы. Номера новых листов будут начинаться с указанного префикса. Для размещения на новых листах все исходные виды будут скопированы с указанным префиксом в названии.";
            CreateTreeViewFromSheets();
            window.ShowDialog();
        }

        void CreateTreeViewFromSheets()
        {
            foreach (ObjSheet objSheet in SheetManager.ObjSheetList)
            {
                AddToTree(objSheet, Window.treeView1);
            }
        }

        void SheetCopy()
        {
            var sheetsList = CollectSelectedSheets();
            SheetManager.CopySheets(sheetsList, this);
            Window.treeView1.Nodes.Clear();
            CreateTreeViewFromSheets();

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

        List<string> CollectSelectedSheets()
        {
            List<string> selectedNode = new List<string>();;
            selectedNode.Clear();

            foreach (TreeNode parent_Node in Window.treeView1.Nodes)
            {
                if (parent_Node.Nodes.Count > 0)
                {
                    NodeChecker(parent_Node, selectedNode);
                }
                else
                {
                    if (parent_Node.Checked)
                    {
                        selectedNode.Add(parent_Node.Text);
                    }
                }
            }
            return selectedNode;
        }

        void TreeViewAfterCheck(object sender, TreeViewEventArgs e)
        {
            // Отключить повторное вхождение в обработчик при массовом изменении состояния
            Window.treeView1.AfterCheck -= TreeViewAfterCheck;

            try
            {
                UpdateChildNodes(e.Node, e.Node.Checked);
            }
            finally
            {
                Window.treeView1.AfterCheck += TreeViewAfterCheck;
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

        void NodeChecker(TreeNode parentNode, List<string> selectedNodes)
        {
            foreach (TreeNode childNode in parentNode.Nodes)
            {
                if (childNode.Nodes.Count > 0)
                {
                    NodeChecker(childNode, selectedNodes);
                }
                else
                {
                    if (childNode.Checked)
                    {
                        selectedNodes.Add(childNode.Text);
                    }
                }
            }
        }
    }
}
