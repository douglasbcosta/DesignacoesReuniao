using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DesignacoesReuniao.Domain.Models
{
    public class ReuniaoDesignacoes
    {
        public Pessoa Designado { get; set; }
        public string Tipo { get; set; } // Ancião, Servo Ministerial, Estudante, etc.
        public string Parte { get; set; }
        public int Semana { get; set; }
        public string Cor { get; set; } // Cor especificada na planilha
    }
}
