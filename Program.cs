using System.Data;
using DataStructures;
using OfficeOpenXml;

namespace DataImportExport
{
    class Program
    {

        static String BaseDirectory = "";
        static void Main(string[] args)
        {

            // Инициализация таблицы
            var table = GetTestDataTable();

            ExportTable(table,"Users", BaseDirectory + "exporttable.xlsx");

            // Инициализация списка пользователей
            var users = GetTestUserList();

            // Вызов метода экспорта данных из списка пользователей в файл
            ExportData(users, BaseDirectory + "exportlist.xlsx");

            // Вызов метода импорта данных из файла
            var persons = ImportData<User>("users",BaseDirectory + "importlist.xlsx");
            
        }

        // Сохранение таблицы 'table' в лист 'worksheet' файла 'filename'
        public static void ExportTable(DataTable table, string worksheet, string filename){
            // Лицензия
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                         
            using(var package = new ExcelPackage()){
                var sheet = package.Workbook.Worksheets.Add(worksheet);
                var filledRange = sheet.Cells["A1"].LoadFromDataTable(table, c=>c.PrintHeaders=true);
                package.SaveAs(new FileInfo(filename));
            }
        }

        // Экпорт из списка List
        public static void ExportData<T>(List<T> list, string filename){

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            
            using(ExcelPackage excelPackage = new ExcelPackage()){
                excelPackage.Workbook.Worksheets.Add("User").Cells[1,1].LoadFromCollection(list, true);
                excelPackage.SaveAs(new FileInfo(filename));
            }
        }

        // Импорт в список
        public static List<T> ImportData<T>(string worksheets, string filename){

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            
            using(ExcelPackage excelPackage = new ExcelPackage(new FileInfo(filename))){

                List<T> list = new List<T>();
                
                var sheet = excelPackage.Workbook.Worksheets[worksheets];

                var columnInfo = Enumerable
                    .Range(1,sheet.Dimension.Columns)
                    .ToList()
                    .Select( n => new {Index=n,ColumntName=sheet.Cells[1,n].Value.ToString()} );

                for(int row=2; row < sheet.Dimension.Rows; row++){
                    T obj = (T)Activator.CreateInstance(typeof(T));
                    foreach(var prop in typeof(T).GetProperties()){
                        int column = columnInfo.SingleOrDefault(c=>c.ColumntName == prop.Name).Index;
                        var value = sheet.Cells[row,column].Value;
                        var propType = prop.PropertyType;
                        prop.SetValue(obj,Convert.ChangeType(value,propType));
                    }
                    list.Add(obj);
                }
                return list;
            }
        }

        public static DataTable GetTestDataTable(){
            var table = new DataTable("Users");
            table.Columns.Add("Name", typeof(string));
            table.Columns.Add("Age",typeof(int));
            table.Columns["Name"].Caption="Name";
            table.Columns["Age"].Caption="Age";

            table.Rows.Add("Ivan", 25);
            table.Rows.Add("Mary", 18);
            return table;
        }

        public static List<User> GetTestUserList(){
            return new List<User> {
                new User{Name="Tom",Age=5},
                new User{Name="Jerry",Age=3}
            };
        }
    }
}