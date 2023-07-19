
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

namespace UploadExcel.Models
{
    public partial class Employer
    {
        [Key]
        public int Id { get; set; }
        public string? Name { get; set; }
        public string? Surname { get; set; }
        public string? Tel { get; set; }
        public string? Mail { get; set; }
    }
}

