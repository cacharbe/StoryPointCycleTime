using System.Collections.Generic;

// ReSharper disable once CheckNamespace
namespace ExcelController
{
    public class FunctionalTeam
    {
        public string TeamName { get; set; }
        public string CurrentTfs { get; set; }
        public string CurrentProjectName { get; set; }
        public string SharedQuery { get; set; }
        public string AreaRoot { get; set; }
        public Dictionary<string, List<Member>> ScrumTeams { get; set; }
    }
}