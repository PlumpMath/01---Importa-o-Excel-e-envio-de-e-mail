using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ImportExcel.Models
{
    public class Planilha
    {
        public string nrDemanda { get; set; }
        public string job { get; set; }
        public profissional gerente { get; set; }
        public string caminhoArquivo { get; set; }
    }

    public class profissional
    {
        public string nome { get; set; }
        public string email { get; set; }

    }
}