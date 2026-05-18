using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OneNoteStudyPlanner.Models
{
    public class StudyDay
    {
        public int Day { get; set; }

        public string Title { get; set; } = string.Empty;

        public List<string> Topics { get; set; } = new();
    }
}
