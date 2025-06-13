using Autodesk.Revit.DB.Events;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;


namespace ISTools
{
    internal class App : IExternalApplication
    {
        static AddInId addinId = new AddInId(new Guid("89905C29-797A-4702-8F3C-401ACAFAC71B"));
        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }

        public Result OnStartup(UIControlledApplication application)
        {
            // Событие OnApplicationInitialized произойдет только после загрузки всех плагинов
            application.ControlledApplication.ApplicationInitialized += OnApplicationInitialized;
            return Result.Succeeded;
        }

        private void OnApplicationInitialized(object sender, ApplicationInitializedEventArgs e)
        {
            Autodesk.Revit.ApplicationServices.Application m_app = sender as Autodesk.Revit.ApplicationServices.Application;
            UIApplication uiApp = new UIApplication(m_app);
            AddButtonToExistingTab(uiApp);
        }

        private void AddButtonToExistingTab(UIApplication uiApp)
        {
            try
            {
                IsUtils.AddButtonToExistTab(
                    uiApp,
                    "Каталог\nсемейств",
                    "ISTools.FamilyCatalog",
                    "ISTools",
                    "Параметры",
                    @"ISTools.Resources.FamilyCatalog32.png",
                    @"ISTools.Resources.FamilyCatalog32.png",
                    "Данная команда имеет три функции:\n- создание каталога семейств и типоразмеров системных элементов модели\n- дополнение существующего каталога\n- запись значений параметров из каталога в параметры элементов модели",
                    @"https://github.com/i-savelev/ISTools/wiki/Каталог-семейств"
                    );

                IsUtils.AddButtonToExistTab(
                    uiApp,
                    "Каталог\nмарок",
                    "ISTools.MarksCatalog",
                    "ISTools",
                    "Параметры",
                    @"ISTools.Resources.MarksCatalog32.png",
                    @"ISTools.Resources.MarksCatalog32.png",
                    "Получение списка уникальных значений выбранного параметра и выгрузка списка в формат excel или дополнение новыми значениями существующего списка",
                    @"https://github.com/i-savelev/ISTools/wiki/Каталог-марок"
                    );

                IsUtils.AddButtonToExistTab(
                    uiApp,
                    "Параметры\nпо категории",
                    "ISTools.ParamByCat",
                    "ISTools",
                    "Параметры",
                    @"ISTools.Resources.ParamByCat32.png",
                    @"ISTools.Resources.ParamByCat32.png",
                    " ",
                    @"https://github.com/i-savelev/ISTools/wiki/Параметры-по-категории"
                    );

                IsUtils.AddButtonToExistTab(
                    uiApp,
                    "Комбинация\nпараметров",
                    "ISTools.ParamCombine",
                    "ISTools",
                    "Параметры",
                    @"ISTools.Resources.Combine32.png",
                    @"ISTools.Resources.Combine32.png",
                    "Плагин позволяет комбинировать значения параметров и записывать результат в указанный параметр. При комбинировании моно задавать любое форматирования (пробелы, прочерки, скобки и т.д.)",
                    @"https://github.com/i-savelev/ISTools/wiki/Комбинация-параметров"
                    );

                IsUtils.AddButtonToExistTab(
                    uiApp,
                    "Параметры\nиз помещений",
                    "ISTools.ParamFromRoom",
                    "ISTools",
                    "Параметры",
                    @"ISTools.Resources.Rooms32.png",
                    @"ISTools.Resources.Rooms32.png",
                    "Копирование значений параметров из помещений в элементы, которые в нем расположены. Также есть возможность отстраивать геометрию помещений.",
                    @"https://github.com/i-savelev/ISTools/wiki/Параметры-из-помещений"
                    );

                IsUtils.AddButtonToExistTab(
                    uiApp,
                    "Маппинг\nпараметров",
                    "ISTools.ParamMapping",
                    "ISTools",
                    "Параметры",
                    @"ISTools.Resources.ParamMapping32.png",
                    @"ISTools.Resources.ParamMapping32.png",
                    "Копирование значений параметров из помещений в элементы, которые в нем расположены. Также есть возможность отстраивать геометрию помещений.",
                    @"https://github.com/i-savelev/ISTools/wiki/Параметры-из-помещений"
                    );

                IsUtils.AddButtonToExistTab(
                    uiApp,
                    "Параметры\nматериалов",
                    "ISTools.Materials",
                    "ISTools",
                    "Параметры",
                    @"ISTools.Resources.Materials32.png",
                    @"ISTools.Resources.Materials32.png",
                    "Перенос значения из параметра материала \"Модель\" в любой выбранный параметр элемента.",
                    @"https://github.com/i-savelev/ISTools/wiki/Параметры-материалов"
                    );

                IsUtils.AddButtonToExistTab(
                    uiApp,
                    "Многослойные\nконструкции",
                    "ISTools.TypesRename",
                    "ISTools",
                    "Параметры",
                    @"ISTools.Resources.TypesRename32.png",
                    @"ISTools.Resources.TypesRename32.png",
                    "Плагин позволяет быстро просматривать состав многослойных конструкций и менять название типоразмера.",
                    @"https://github.com/i-savelev/ISTools/wiki/Многослойные-конструкции"
                    );

                IsUtils.AddButtonToExistTab(
                    uiApp,
                    "Номера\nлистов",
                    "ISTools.SheetsNumber",
                    "ISTools",
                    "Общее",
                    @"ISTools.Resources.Sheet_number32.png",
                    @"ISTools.Resources.Sheet_number32.png",
                    "Плагин позволяет изменять номера выбранных групп листов",
                    @"https://github.com/i-savelev/ISTools/wiki/Номера-листов"
                    );

                IsUtils.AddButtonToExistTab(
                    uiApp,
                    "Копирование\nлистов",
                    "ISTools.SheetsCopy",
                    "ISTools",
                    "Общее",
                    @"ISTools.Resources.SheetsCopy32.png",
                    @"ISTools.Resources.SheetsCopy32.png",
                    "Плагин позволяет копировать листы вместе с видами, аннатциями, текстом и параметрами",
                    @"https://github.com/i-savelev/ISTools/wiki/Копирование-листов"
                    );

                IsUtils.AddButtonToExistTab(
                    uiApp,
                    "Фильтры\nвидов",
                    "ISTools.SetFilters",
                    "ISTools",
                    "Общее",
                    @"ISTools.Resources.SetFilters32.png",
                    @"ISTools.Resources.SetFilters32.png",
                    "Плагин позволяет массово перенести фильтры на выбранные виды и шаблоны",
                    @"https://github.com/i-savelev/ISTools/wiki/Фильтры-видов"
                    );

                IsUtils.AddButtonToExistTab(
                    uiApp,
                    "Рабочие\nнаборы",
                    "ISTools.SetWorksets",
                    "ISTools",
                    "Общее",
                    @"ISTools.Resources.Worksets32.png",
                    @"ISTools.Resources.Worksets32.png",
                    "Плагин позволяет автоматически распределить элементы по рабочим наборам",
                    @"https://github.com/i-savelev/ISTools/wiki/Рабочие-наборы"
                    );

                IsUtils.AddButtonToExistTab(
                    uiApp,
                    "Раскраска\nэлементов",
                    "ISTools.SetColor",
                    "ISTools",
                    "Общее",
                    @"ISTools.Resources.Set_color_32.png",
                    @"ISTools.Resources.Set_color_32.png",
                    "Плагин позволяет раскрасить элементы выбраной категории по цветам в зависимости от значения выбранного параметра",
                    @"https://github.com/i-savelev/ISTools/wiki/Раскраска-элементов"
                    );
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Ошибка", ex.Message);
            }
        }
    }
}
