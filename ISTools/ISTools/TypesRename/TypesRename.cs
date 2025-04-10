using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Data;
using System.Windows.Forms;

namespace ISTools
{
    [Autodesk.Revit.Attributes.TransactionAttribute(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class TypesRename : IExternalCommand
    {
        //---PluginsManager---//
        public static string IS_TAB_NAME => "ISTools";
        public static string IS_NAME => "Многослойные конструкции";
        public static string IS_IMAGE => "ISTools.Resources.TypesRename32.png";
        public static string IS_DESCRIPTION => "Команда предоставляет удобный интерфейс для переименования многослойных коснтуркций";
        //---PluginsManager---//
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;
            string docName = doc.Title.Replace(".rvt", "");
            List<Type> types = new List<Type>() {
                typeof(FloorType),
                typeof(WallFoundationType),
                typeof(WallType),
                typeof(RoofType),
            };

            DataTable dtElems = new DataTable();
            dtElems.Columns.Add("Категория", typeof(string));
            dtElems.Columns.Add("Название", typeof(string));
            dtElems.Columns.Add("id", typeof(string));

            List<Element> elems = new List<Element>();
            TypeElementCollector(elems, types);

            TypesRenameForm window = new TypesRenameForm();
            window.dataGridView2.DataSource = dtElems;
            window.Text = "Менеджер многослойных конструкций";

            window.textBox4.Text = "Плагин отображает структуру многослойных элементов из диспетчера проекта и позволет переименовать эти элементы.";

            window.dataGridView2.CellClick += SetLayersToTable;
            window.button2.Click += (s, e) => { Rename(); };
            window.button2.Text = "Переименовать элемент";
            window.groupBox1.Text = "Список многослойных конструкций";
            window.groupBox3.Text = "Список материалов";
            window.groupBox5.Text = "Информация";
            window.textBox3.Text = "Введите имя:";
            window.ShowDialog();

            void TypeElementCollector(List<Element> elemList, List<Type> classes)
            { 
                foreach (Type typ in classes)
                {
                    var els = new FilteredElementCollector(doc).
                        WhereElementIsElementType().
                        OfClass(typ).
                        Cast<Element>().
                        ToList();
                    if (typ == typeof(WallType))
                    {
                        foreach (Element el in els)
                        {
                            try
                            {
                                if ((el as WallType).Kind.ToString() == "Basic")
                                {
                                    elemList.Add(el);
                                    dtElems.Rows.Add(el.Category.Name,
                                        el.Name,
                                        el.Id
                                        );
                                }
                            }
                            catch { }
                        }
                    }
                    else if (typ == typeof(RoofType))
                    {
                        foreach (Element el in els)
                        {
                            try
                            {
                                if ((el as RoofType).FamilyName.ToString() == "Базовая крыша")
                                {
                                    elemList.Add(el);
                                    dtElems.Rows.Add(el.Category.Name,
                                        el.Name,
                                        el.Id
                                        );
                                }
                            }
                            catch { }
                        }
                    }
                    else
                    {
                        foreach (Element el in els)
                        {
                            try
                            {
                                elemList.Add(el);
                                dtElems.Rows.Add(el.Category.Name,
                                    el.Name,
                                    el.Id
                                    );
                            }
                            catch { }
                        }
                    }
                }
            }

            void SetLayersToTable(object sender, DataGridViewCellEventArgs e)
            {
                DataTable dtLayers = new DataTable();
                dtLayers.Columns.Add("Материал", typeof(string));
                dtLayers.Columns.Add("Функция", typeof(string));
                dtLayers.Columns.Add("Порядок", typeof(string));
                dtLayers.Columns.Add("Толщина", typeof(string));

                if (e.RowIndex < 0 ) return;

                int id = (int)Convert.ToDouble(window.dataGridView2["id", e.RowIndex].Value);
                Element el = doc.GetElement(new ElementId(id));
                
                switch (el.GetType().ToString())
                {
                    case "Autodesk.Revit.DB.FloorType":
                        var floortype = el as FloorType;
                        var layers = floortype.GetCompoundStructure().GetLayers();
                        foreach (var layer in layers)
                        {
                            var materialId = layer.MaterialId;
                            dtLayers.Rows.Add(
                                GetMaterialName(materialId),
                                layer.Function.ToString(),
                                layer.LayerId.ToString(),
                                layer.Width*304.8
                                );
                            window.textBox2.Text = window.dataGridView2[1, e.RowIndex].Value.ToString();
                        }
                        break;

                    case "Autodesk.Revit.DB.RoofType":
                        var rooftype = el as RoofType;
                        var rooflayers = rooftype.GetCompoundStructure().GetLayers();
                        foreach (var layer in rooflayers)
                        {
                            var materialId = layer.MaterialId;
                            dtLayers.Rows.Add(
                                GetMaterialName(materialId),
                                layer.Function.ToString(),
                                layer.LayerId.ToString(),
                                layer.Width * 304.8
                                );
                            window.textBox2.Text = window.dataGridView2[1, e.RowIndex].Value.ToString();
                        }
                        break;

                    case "Autodesk.Revit.DB.WallType":
                        var walltype = el as WallType;
                        var walllayers = walltype.GetCompoundStructure().GetLayers();
                        foreach (var layer in walllayers)
                        {
                            var materialId = layer.MaterialId;
                            dtLayers.Rows.Add(
                                GetMaterialName(materialId),
                                layer.Function.ToString(),
                                layer.LayerId.ToString(),
                                layer.Width*304.8
                                );
                            window.textBox2.Text = window.dataGridView2[1, e.RowIndex].Value.ToString();
                        }
                        break;
                }
                window.dataGridView1.DataSource = dtLayers;
                window.dataGridView1.Columns[0].Width = 400;
            }

            void Rename()
            {
                using (Transaction tx = new Transaction(doc))
                {
                    if (window.textBox2.Text != "")
                    {
                        if (window.dataGridView2.SelectedRows.Count == 1)
                        {
                            tx.Start("Переименование элементов");
                            foreach (DataGridViewRow cell in window.dataGridView2.SelectedRows)
                            {
                                var colindex = window.dataGridView2.Columns["id"].Index;
                                var rowindex = cell.Index;
                                int id = (int)Convert.ToDouble(window.dataGridView2["id", rowindex].Value);
                                Element el = doc.GetElement(new ElementId(id));
                                if (rowindex < 0) return;
                                window.dataGridView2["Название", rowindex].Value = window.textBox2.Text;
                                el.Name = window.textBox2.Text;
                            }
                            tx.Commit();
                        }
                        else
                        {
                            TaskDialog.Show("Предупреждение", "Выберете только один элемент");
                        }
                    }
                    else
                    {
                        TaskDialog.Show("Предупреждение", "Заполните поле с наименвоанием");
                    }
                }
            }

            string GetMaterialName(ElementId materialId)
            {
                if (materialId.ToString() == "-1")
                {
                    return "По категории";
                }
                else
                {
                    Material mat = doc.GetElement(materialId) as Material;
                    return mat.Name;
                }
            }
            return Result.Succeeded;
            }
        }
    }