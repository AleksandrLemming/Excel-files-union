using Microsoft.Office.Interop.Excel;
using System;

namespace Excel_files_union_for_MMW_app
{
    class Excel
    {
        Application excelapp = new Application();

        private string pathfolder = null;
        public string[] filelists = null;
        public int rowsinheader = 1;
        public int sheetnumber = 1;

        public void PathFolder(string s)
        {
            string datetime = DateTime.Now.Hour.ToString() + "_" + DateTime.Now.Minute.ToString()
                    + "_" + DateTime.Now.Second.ToString(); // добавляем уникальности - время создания

            pathfolder = s + "\\UnionFile_Creation Time " + datetime + ".xlsb";
        }

        public void UnionFiles() //Объединяем файлы
        {
            #region Создаем эксель приложение и новую книгу в которой будут все данные из других книг
                
                Workbook wb = excelapp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
                Worksheet ws = wb.Worksheets[1];

                ws.Name = "UnionData";
                    wb.SaveAs(pathfolder, XlFileFormat.xlExcel12, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing);
            #endregion

            //открываем по очереди наши эксель файлы
            for (int i = 0; i < filelists.Length; i++)
            {
                Workbook wbFrom = excelapp.Workbooks.Open(filelists[i]);
                Worksheet wsFrom = wbFrom.Worksheets[sheetnumber];
                //копируем диапазон
                if (i == 0) //если книга первая, то надо скопировать с шапкой таблицы, если нет - копируем лишь содержимое
                {
                    Range areaFrom = wsFrom.Range[wsFrom.Cells[1, 1], wsFrom.Cells.SpecialCells(XlCellType.xlCellTypeLastCell)];
                    Range areaTo = ws.Cells[1, 1];
                    areaFrom.Copy(areaTo);
                }
                else
                {
                    int rowcount = ws.Cells.SpecialCells(XlCellType.xlCellTypeLastCell).Row; //надо определить строку, чтобы не копировать на одно и тоже место
                    Range areaFrom = wsFrom.Range[wsFrom.Cells[rowsinheader + 1, 1], wsFrom.Cells.SpecialCells(XlCellType.xlCellTypeLastCell)];
                    Range areaTo = ws.Cells[rowcount + 1, 1];
                    areaFrom.Copy(areaTo);
                }
                wbFrom.Close(false);
            }

            wb.Close(true);
            excelapp.Quit();
        }
    }
}
