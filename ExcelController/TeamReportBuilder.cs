using System;
using System.Collections.Generic;
using DataObjects;
using Microsoft.Office.Interop.Excel;

namespace ExcelController
{
    public class TeamReportBuilder : IDisposable
    {
        private ExcelHelper _excelController;

        public TeamReportBuilder(string filename)
        {
            FileName = filename;
            _excelController = new ExcelHelper();
            _excelController.CreateOpenExcelFile(filename);
        }

        public string FileName { get; set; }

        public void WriteTeamData(string projectname, List<Team> teams)
        {
            //With Project Name, Create Tab
            var workSheet = _excelController.CreateWorksheet(projectname);
            SetUpWorksheetHeaders(workSheet);
            var currRow = 2;
            foreach (var team in teams)
            {
                ((Range) workSheet.Cells[currRow, 1]).Value2 = team.TeamName;
                foreach (var member in team.TeamMembers)
                {
                    currRow++;
                    ((Range) workSheet.Cells[currRow, 3]).Value2 = member.DisplayName;
                    ((Range) workSheet.Cells[currRow, 4]).Value2 = member.UniqueName;
                    if (member.Email == string.Empty)
                    {
                        var names = member.DisplayName.Split(' ');
                        var email = string.Join(".", names);
                        ((Range) workSheet.Cells[currRow, 5]).Value2 = email + "@aon.com";
                    }
                    else
                    {
                        ((Range) workSheet.Cells[currRow, 5]).Value2 = member.Email;
                    }
                }

                currRow++;
            }

            _excelController.SaveWorkBook();
        }

        private void SetUpWorksheetHeaders(Worksheet workSheet)
        {
            ((Range) workSheet.Cells[1, 1]).Value2 = "Team Name";
            ((Range) workSheet.Cells[1, 2]).Value2 = "Scrum Team Name";
            ((Range) workSheet.Cells[1, 3]).Value2 = "Member Display Name";
            ((Range) workSheet.Cells[1, 4]).Value2 = "Member Unique Name";
            ((Range) workSheet.Cells[1, 5]).Value2 = "Member EMail";
            ((Range) workSheet.Cells[1, 6]).Value2 = "MSDN License Level";
        }

        #region IDisposable Support

        private bool disposedValue; // To detect redundant calls

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                    if (_excelController != null)
                    {
                        _excelController.DisposeExcel();
                        _excelController = null;
                    }

                // TODO: free unmanaged resources (unmanaged objects) and override a finalizer below.
                // TODO: set large fields to null.

                disposedValue = true;
            }
        }

        // TODO: override a finalizer only if Dispose(bool disposing) above has code to free unmanaged resources.
        // ~WorkItemParsert() {
        //   // Do not change this code. Put cleanup code in Dispose(bool disposing) above.
        //   Dispose(false);
        // }

        // This code added to correctly implement the disposable pattern.
        public void Dispose()
        {
            // Do not change this code. Put cleanup code in Dispose(bool disposing) above.
            Dispose(true);
            // TODO: uncomment the following line if the finalizer is overridden above.
            GC.SuppressFinalize(this);
        }

        #endregion
    }
}