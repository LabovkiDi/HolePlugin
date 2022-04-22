using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HolePlugin
{
    [Transaction(TransactionMode.Manual)]

    public class AddHole : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {
            //получение файла АР
            Document arDoc = commandData.Application.ActiveUIDocument.Document;
            //получение связанного файл ОВ, через файл АР 
            Document ovDoc = arDoc.Application.Documents.OfType<Document>().Where(x => x.Title.Contains("ОВ")).FirstOrDefault(); //ищем файл по суффиксу ОВ
            if (ovDoc == null)
            {
                //если такого файла нет, сообщаем об этом
                TaskDialog.Show("Ошибка", "Не найден ОВ файл");
                return Result.Cancelled;
            }
            // проверяем загружено ли в файл нужное семейство
            FamilySymbol familySymbol = new FilteredElementCollector(arDoc)
                .OfClass(typeof(FamilySymbol))
                .OfCategory(BuiltInCategory.OST_GenericModel)
                .OfType<FamilySymbol>()
                .Where(x => x.FamilyName.Equals("Отверстие"))
                .FirstOrDefault();
            if (familySymbol == null)
            {
                //если такого семейства нет, сообщаем об этом
                TaskDialog.Show("Ошибка", "Не найден семейство \"Отверстие\"");
                return Result.Cancelled;
            }

            //через фильтр выполняем поиск воздуховодов
            List<Duct> ducts = new FilteredElementCollector(ovDoc)
                .OfClass(typeof(Duct))
                .OfType<Duct>()
                .ToList();
            //через фильтр выполняем поиск труб
            List<Pipe> pipes = new FilteredElementCollector(ovDoc)
               .OfClass(typeof(Pipe))
               .OfType<Pipe>()
               .ToList();
            //поиск 3д вида
            View3D view3D = new FilteredElementCollector(arDoc)
                .OfClass(typeof(View3D))
                .OfType<View3D>()
                .Where(x => !x.IsTemplate) //3д вид не является шаблоном вида
                .FirstOrDefault();
            if (view3D == null)
            {
                //если такого вида нет, сообщаем об этом
                TaskDialog.Show("Ошибка", "Не найден 3D вид");
                return Result.Cancelled;
            }

            ReferenceIntersector referenceIntersector = new ReferenceIntersector(new ElementClassFilter(typeof(Wall)), FindReferenceTarget.Element, view3D);

            Transaction transaction0 = new Transaction(arDoc);
            transaction0.Start("Расстановка отверстий");
            if (!familySymbol.IsActive)
            {
                familySymbol.Activate();
            }
            transaction0.Commit();

            using (var ts = new Transaction(arDoc, "Расстановка отверстий для воздуховодов"))
            {
                ts.Start();

                foreach (Duct d in ducts)
                {
                    Line curve = (d.Location as LocationCurve).Curve as Line;
                    XYZ point = curve.GetEndPoint(0); // начальная точка воздуховода
                    XYZ direction = curve.Direction; // направление воздуховода
                                                     //получаем набор всех пересечений
                    List<ReferenceWithContext> intersections = referenceIntersector.Find(point, direction)
                         .Where(x => x.Proximity <= curve.Length) // расстояния Proximity не превышает длину воздуховода
                         .Distinct(new ReferenceWithContextElementEqualityComparer()) // оставляет одно пересечение из двух идентичных
                         .ToList();
                    foreach (ReferenceWithContext refer in intersections)
                    {
                        //определим точку вставки и добавим экземпляр семейства отверстие
                        double proximity = refer.Proximity; // расстояние до объекта
                        Reference reference = refer.GetReference(); // ссылка
                        Wall wall = arDoc.GetElement(reference.ElementId) as Wall; // через ссылку, получаем элемент
                        Level level = arDoc.GetElement(wall.LevelId) as Level; // через стену, получаем уровень
                        XYZ pointHole = point + (direction * proximity); // получаем точку вставки отверстия

                        FamilyInstance hole = arDoc.Create.NewFamilyInstance(pointHole, familySymbol, wall, level, StructuralType.NonStructural);
                        Parameter width = hole.LookupParameter("Ширина");
                        Parameter height = hole.LookupParameter("Высота");
                        width.Set(d.Diameter);
                        height.Set(d.Diameter);
                    }
                }
                ts.Commit();
            }

            using (var ts = new Transaction(arDoc, "Расстановка отверстий для труб"))
            {
                ts.Start();

                foreach (Pipe p in pipes)
                {
                    Line curve = (p.Location as LocationCurve).Curve as Line;
                    XYZ point = curve.GetEndPoint(0); // начальная точка трубы
                    XYZ direction = curve.Direction; // направление трубы
                                                     //получаем набор всех пересечений
                    List<ReferenceWithContext> intersections = referenceIntersector.Find(point, direction)
                         .Where(x => x.Proximity <= curve.Length) // расстояния Proximity не превышает длину трубы
                         .Distinct(new ReferenceWithContextElementEqualityComparer()) // оставляет одно пересечение из двух идентичных
                         .ToList();
                    foreach (ReferenceWithContext refer in intersections)
                    {
                        //определим точку вставки и добавим экземпляр семейства отверстие
                        double proximity = refer.Proximity; // расстояние до объекта
                        Reference reference = refer.GetReference(); // ссылка
                        Wall wall = arDoc.GetElement(reference.ElementId) as Wall; // через ссылку, получаем элемент
                        Level level = arDoc.GetElement(wall.LevelId) as Level; // через стену, получаем уровень
                        XYZ pointHole = point + (direction * proximity); // получаем точку вставки отверстия

                        FamilyInstance hole = arDoc.Create.NewFamilyInstance(pointHole, familySymbol, wall, level, StructuralType.NonStructural);
                        Parameter width = hole.LookupParameter("Ширина");
                        Parameter height = hole.LookupParameter("Высота");
                        width.Set(p.Diameter);
                        height.Set(p.Diameter);
                    }
                }
                ts.Commit();
            }
            return Result.Succeeded;

        }
        public class ReferenceWithContextElementEqualityComparer : IEqualityComparer<ReferenceWithContext>
        {
            //метод Equals определяет будут ли 2 заданных объекта одинаковыми
            public bool Equals(ReferenceWithContext x, ReferenceWithContext y)
            {
                if (ReferenceEquals(x, y)) return true;
                if (ReferenceEquals(null, x)) return false;
                if (ReferenceEquals(null, y)) return false;

                var xReference = x.GetReference();

                var yReference = y.GetReference();
                //если у обоих объектов одинаковые ElementId/LinkedElementId мы получим точки на одной стене, вернется true
                return xReference.LinkedElementId == yReference.LinkedElementId
                           && xReference.ElementId == yReference.ElementId;
            }
            //метод GetHashCode должен возвращать HashCode объекта
            public int GetHashCode(ReferenceWithContext obj)
            {
                var reference = obj.GetReference();

                unchecked
                {
                    return (reference.LinkedElementId.GetHashCode() * 397) ^ reference.ElementId.GetHashCode();
                }
            }
        }
    }
}
