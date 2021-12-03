using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApp
{
     public class Attendance
    {
        public long Id { get; set; }
        public string Code { get; set; }
        public long HrConfigId { get; set; }
        public long BranchId { get; set; }
        public string Note { get; set; }
        public DateTime CreatedDate { get; set; }
        public DateTime? ModifiedDate { get; set; }
        public DateTime StartDate { get; set; }
        public DateTime? EndDate { get; set; }

        public IList<EmployeeAttendance> EmployeeAttendances { get; set; }

        public Attendance()
        {
            EmployeeAttendances = new List<EmployeeAttendance>();
        }
    }
}
