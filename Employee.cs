using System.Collections.Generic;

namespace consolereadjsonfile
{
    public class Employee
    {
        public long Id { get; set; }
        public string Name { get; set; }
        public string dept { get; set; }
        public List<Punch> Punchs { get; set; }
        // public ICollection<Punch> Punchs { get; set; }
        // public Punch Punchs { get; set; }

    }
}