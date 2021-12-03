using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApp
{
    public class EmployeeAttendance
    {
        public long Id { get; set; }
        public long AttendanceId { get; set; }
        public long EmployeeId { get; set; }
        public string Note { get; set; }
        public DateTime CreatedDate { get; set; }
        public DateTime? ModifiedDate { get; set; }
        public DateTime? StartDate { get; set; }
        public DateTime? EndDate { get; set; }
        public long PresenceId { get; set; }
        public long StartStatusId { get; set; }
        public long EndStatusId { get; set; }
	}
}
