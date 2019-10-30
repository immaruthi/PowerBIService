using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerBIExcelService.DataModels
{
    public class ServiceDataModel
    {
        public DataTable dataTables { get; set; }

        public ProgramFilters programFilters { get; set; }
    }
}
