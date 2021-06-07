using System;
using System.Collections.Generic;
//using DataObjects;
using ExcelController.Model;
using Microsoft.Office.Interop.Excel;

namespace ExcelController
{
    public class TeamParser : IDisposable
    {
        private ExcelHelper _excelController;

        public TeamParser(string filename)
        {
            FileName = filename;
            _excelController = new ExcelHelper();
            _excelController.CreateOpenExcelFile(filename);
        }

        public TeamParser(ColumnConfiguration columnConfig, string fileName)
        {
            FileName = fileName;
            ColumnConfig = columnConfig;
            _excelController = new ExcelHelper();
            _excelController.CreateOpenExcelFile(fileName);
        }

        public string FileName { get; set; }
        public ColumnConfiguration ColumnConfig { get; set; }


        public List<Team> ParseTeamData(string workSheetName, int startingRow = 2,
            int startingCol = 1)
        {
            var sheet = _excelController.GetWorkSheet(workSheetName);
            var teamDict = new Dictionary<string, Team>();
            var teams = new List<Team>();
            var currRow = startingRow;
            while (((Range) sheet.Cells[currRow, 1]).Value2 != null || ((Range) sheet.Cells[currRow, 2]).Value2 != null)
            {
                if (((Range) sheet.Cells[currRow, 2]).Value2 != null)
                {
                    string teamName = ((Range) sheet.Cells[currRow, 2]).Value2;
                    if (!teamDict.ContainsKey(teamName))
                        teamDict.Add(teamName, new Team(teamName, new List<TeamMember>()));
                    var member = new TeamMember(((Range) sheet.Cells[currRow, 3]).Value2,
                        ((Range) sheet.Cells[currRow, 5]).Value2, ((Range) sheet.Cells[currRow, 6]).Value2);
                    teamDict[teamName].TeamMembers.Add(member);
                }

                currRow++;
            }

            foreach (var key in teamDict.Keys) teams.Add(teamDict[key]);

            return teams;
        }

        public Dictionary<string, FunctionalTeam> ParseScrumTeams(string worksheetname, int startingrow = 2,
            int startingcol = 1)
        {
            var sheet = _excelController.GetWorkSheet(worksheetname);


            var values = new Dictionary<string, FunctionalTeam>();

            var currRow = startingrow;
            while (((Range) sheet.Cells[currRow, 1]).Value2 != null)
            {
                var member = new Member();
                if (((Range) sheet.Cells[currRow, 2]).Value2 != null)
                {
                    member.DisplayName = ((Range) sheet.Cells[currRow, ColumnConfig.DisplayName]).Value2.ToString();
                    member.UserId = ((Range) sheet.Cells[currRow, ColumnConfig.UserId]).Value2.ToString();
                }

                // var User = ((Range)sheet.Cells[currRow, 2]).Value2.ToString();
                string scrumteam = ((Range) sheet.Cells[currRow, ColumnConfig.ProjectTeam]).Value2.ToString();
                string functionalteam = ((Range) sheet.Cells[currRow, ColumnConfig.FunctionalTeam]).Value2.ToString();

                //string TeamName = string.Format("{0}-{1}", functionalteam, scrumteam);

                //Check if Functional Team Exists
                if (values.ContainsKey(functionalteam))
                {
                    //Check if Scrum team exists
                    if (values[functionalteam].ScrumTeams.ContainsKey(scrumteam))
                    {
                        values[functionalteam].ScrumTeams[scrumteam].Add(member);
                    }
                    else
                    {
                        //Add new Scrum Team
                        var members = new List<Member>();
                        members.Add(member);
                        values[functionalteam].ScrumTeams.Add(scrumteam, members);
                    }

                    //values[TeamName].ScrumTeams(User);
                }
                else
                {
                    //New Functional team
                    var t = new FunctionalTeam();
                    //New Scrum team
                    t.ScrumTeams = new Dictionary<string, List<Member>>();
                    //New Team Member to team list
                    var members = new List<Member>();
                    members.Add(member);

                    //Add members to scrum team
                    t.ScrumTeams.Add(scrumteam, members);
                    //Add Scrum team to Functional team
                    values.Add(functionalteam, t);
                }

                currRow++;
            }

            return values;
        }

        public Dictionary<string, FunctionalTeam> ParseFunctionalTeamData(string worksheetname,
            Dictionary<string, FunctionalTeam> teams, int startingrow = 2, int startingcol = 1)
        {
            var sheet = _excelController.GetWorkSheet(worksheetname);
            var currRow = startingrow;
            while (((Range) sheet.Cells[currRow, 1]).Value2 != null)
            {
                string functionalteam = ((Range) sheet.Cells[currRow, 1]).Value2.ToString();
                //Check if Functional Team Exists
                if (teams.ContainsKey(functionalteam))
                {
                    teams[functionalteam].TeamName = ((Range) sheet.Cells[currRow, 1]).Value2 != null
                        ? ((Range) sheet.Cells[currRow, 1]).Value2.ToString()
                        : string.Empty;
                    teams[functionalteam].CurrentTfs = ((Range) sheet.Cells[currRow, 2]).Value2 != null
                        ? ((Range) sheet.Cells[currRow, 2]).Value2.ToString()
                        : string.Empty;
                    teams[functionalteam].CurrentProjectName = ((Range) sheet.Cells[currRow, 3]).Value2 != null
                        ? ((Range) sheet.Cells[currRow, 3]).Value2.ToString()
                        : string.Empty;
                    teams[functionalteam].SharedQuery = ((Range) sheet.Cells[currRow, 4]).Value2 != null
                        ? ((Range) sheet.Cells[currRow, 4]).Value2.ToString()
                        : string.Empty;
                    teams[functionalteam].AreaRoot = ((Range) sheet.Cells[currRow, 5]).Value2 != null
                        ? ((Range) sheet.Cells[currRow, 5]).Value2.ToString()
                        : string.Empty;
                }
                else
                {
                    //New Functional team
                    var t = new FunctionalTeam();
                    t.TeamName = ((Range) sheet.Cells[currRow, 1]).Value2 != null
                        ? ((Range) sheet.Cells[currRow, 1]).Value2.ToString()
                        : string.Empty;
                    t.CurrentTfs = ((Range) sheet.Cells[currRow, 2]).Value2 != null
                        ? ((Range) sheet.Cells[currRow, 2]).Value2.ToString()
                        : string.Empty;
                    t.CurrentProjectName = ((Range) sheet.Cells[currRow, 3]).Value2 != null
                        ? ((Range) sheet.Cells[currRow, 3]).Value2.ToString()
                        : string.Empty;
                    t.SharedQuery = ((Range) sheet.Cells[currRow, 4]).Value2 != null
                        ? ((Range) sheet.Cells[currRow, 4]).Value2.ToString()
                        : string.Empty;
                    t.AreaRoot = ((Range) sheet.Cells[currRow, 5]).Value2 != null
                        ? ((Range) sheet.Cells[currRow, 5]).Value2.ToString()
                        : string.Empty;

                    //New Scrum team
                    t.ScrumTeams = new Dictionary<string, List<Member>>();
                    //Add Scrum team to Functional team
                    teams.Add(functionalteam, t);
                }

                currRow++;
            }

            return teams;
        }

        public void ParseWorkItems(string worksheetname, Dictionary<int, List<int>> wItems, int startingrow = 2,
            int startingcol = 1)
        {
            var sheet = _excelController.GetWorkSheet(worksheetname);
            var currRow = startingrow;
            while (((Range) sheet.Cells[currRow, 1]).Value2 != null)
            {
            }
        }

        #region Implementation of IDisposable

        /// <summary>
        ///     Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged resources.
        /// </summary>
        /// <filterpriority>2</filterpriority>
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }


        protected virtual void Dispose(bool disposing)
        {
            if (disposing)
                if (_excelController != null)
                {
                    _excelController.DisposeExcel();
                    _excelController = null;
                }
        }

        #endregion
    }
}