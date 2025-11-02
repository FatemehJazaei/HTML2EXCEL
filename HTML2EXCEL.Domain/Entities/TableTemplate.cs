using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HTML2EXCEL.Domain.Entities
{
    public class TableTemplate
    {
        public int Id { get; set; }
        public int Rows { get; set; }       
        public int Cols { get; set; }  


        public TableTemplate() { }

        public TableTemplate(int id, int rows, int cols)
        {
            Id = id;
            Rows = rows;
            Cols = cols;
        }

    }
}
