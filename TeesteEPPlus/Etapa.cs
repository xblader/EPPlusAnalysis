using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TeesteEPPlus
{
    public class Etapa
    {
        public DateTime Inicio { get; set; }
        public DateTime Fim { get; set; }
        public string  Localizacao { get; set; }
        public string Tipo { get; set; }
        public string Descricao { get; set; }
        public string HoraInicio { get; set; }
        public string HoraFim { get; set; }
    }
}
