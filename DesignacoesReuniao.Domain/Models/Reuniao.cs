namespace DesignacoesReuniao.Domain.Models
{
    public class Reuniao
    {
        public string Semana { get; set; }
        public DateOnly InicioSemana { get; set; }
        public string LeituraDaSemana { get; set; }
        public Pessoa Presidente { get; set; } 
        public Pessoa OracaoInicial { get; set; } 
        public Pessoa OracaoFinal { get; set; } 
        public List<string> Canticos { get; set; } = new List<string>();
        public List<Sessao> Sessoes { get; set; }
        public Reuniao() { 
            Sessoes = new List<Sessao>(); 
            Presidente = new Pessoa();
            OracaoInicial = new Pessoa();
            OracaoFinal = new Pessoa();
        }
        public void AdicionarSessao(Sessao sessao) { 
            Sessoes.Add(sessao); 
        }
        public void AdicionarCantico(string cantico)
        {
            Canticos.Add(cantico);
        }

        public static string[] GetSessoesReunioes()
        {
            return new string[] { "TESOUROS DA PALAVRA DE DEUS", "FAÇA SEU MELHOR NO MINISTÉRIO", "NOSSA VIDA CRISTÃ" };
        }

        public static List<Reuniao> PreencherReunioes(List<Reuniao> reunioesProgramacao, List<Reuniao> reunioesImportadas)
        {
            foreach (var reuniaoProgramada in reunioesProgramacao)
            {
                var reuniaoImportada = reunioesImportadas.FirstOrDefault(r => r.Semana == reuniaoProgramada.Semana);
                if (reuniaoImportada != null)
                {
                    reuniaoProgramada.Presidente = reuniaoImportada.Presidente;
                    reuniaoProgramada.OracaoInicial = reuniaoImportada.OracaoInicial;
                    reuniaoProgramada.OracaoFinal = reuniaoImportada.OracaoFinal;

                    foreach (var sessaoProgramada in reuniaoProgramada.Sessoes)
                    {
                        var sessaoImportada = reuniaoImportada.Sessoes.FirstOrDefault(s => s.TituloSessao == sessaoProgramada.TituloSessao);
                        if (sessaoImportada != null)
                        {
                            foreach (var parteProgramada in sessaoProgramada.Partes)
                            {
                                var parteImportada = sessaoImportada.Partes.FirstOrDefault(p => p.TituloParte.Trim() == parteProgramada.TituloParte.Trim() && p.IndiceParte == parteProgramada.IndiceParte);
                                if (parteImportada != null)
                                {
                                    if (parteImportada.ContemDesignado())
                                    {
                                        parteProgramada.AdicionarDesignado(parteImportada.Designado);
                                    }
                                    if (parteImportada.ContemAjudante())
                                    {
                                        parteProgramada.AdicionarDesignado(parteImportada.Ajudante);
                                    }
                                    parteProgramada.TempoMinutos = parteImportada.TempoMinutos;
                                }
                            }
                        }
                    }
                }
            }
            return reunioesProgramacao;
        }

        public static List<Reuniao> PreencherReunioes(List<Reuniao> reunioesProgramacao, List<ReuniaoDesignacoes> designacoesImportadas)
        {
            int semana = 0;
            foreach (var reuniaoProgramada in reunioesProgramacao)
            {
                semana++;
                var designacoesSemanaImportadas = designacoesImportadas.Where(r => semana == r.Semana).ToList();

                reuniaoProgramada.Presidente = designacoesSemanaImportadas.FirstOrDefault(d => d.Parte == "P").Designado;

                foreach (var sessaoProgramada in reuniaoProgramada.Sessoes)
                {
                    if (sessaoProgramada.TituloSessao == Reuniao.GetSessoesReunioes()[0])
                    {
                        sessaoProgramada.Partes[0].AdicionarDesignado(designacoesSemanaImportadas.FirstOrDefault(d => d.Parte == "T").Designado);
                        sessaoProgramada.Partes[1].AdicionarDesignado(designacoesSemanaImportadas.FirstOrDefault(d => d.Parte == "J").Designado);
                        sessaoProgramada.Partes[2].AdicionarDesignado(designacoesSemanaImportadas.FirstOrDefault(d => d.Parte == "L").Designado);
                    }

                    if (sessaoProgramada.TituloSessao == Reuniao.GetSessoesReunioes()[1])
                    {
                        int ordemParte = 1;
                        foreach (var parte in sessaoProgramada.Partes)
                        {

                            if (parte.TituloParte.Contains("Discurso"))
                            {
                                parte.AdicionarDesignado(designacoesSemanaImportadas.FirstOrDefault(d => d.Parte == "D" && d.Tipo.Contains("Estudante")).Designado);
                                continue;
                            }
                            if (parte.TempoMinutos > 5 && ordemParte == 1)
                            {
                                parte.AdicionarDesignado(designacoesSemanaImportadas.FirstOrDefault(d => d.Parte == "D." && d.Tipo.Contains("Servos")).Designado);
                            }
                            if (parte.TempoMinutos > 5 && ordemParte == 2)
                            {
                                parte.AdicionarDesignado(designacoesSemanaImportadas.FirstOrDefault(d => d.Parte == "D.." && d.Tipo.Contains("Servos")).Designado);
                            }

                            var pessoa = designacoesSemanaImportadas.FirstOrDefault(d => d.Parte == ordemParte.ToString() && d.Tipo.Contains("Estudante"));

                            if (pessoa != null)
                            {
                                parte.AdicionarDesignado(pessoa.Designado);
                                var ajudante = designacoesSemanaImportadas.FirstOrDefault(d => d.Cor == pessoa.Cor && d.Tipo.Contains("Estudante") && d.Parte == "A");
                                if (ajudante != null)
                                {
                                    parte.AdicionarDesignado(ajudante.Designado);
                                }
                            }

                            ordemParte++;
                        }
                    }
                    if (sessaoProgramada.TituloSessao == Reuniao.GetSessoesReunioes()[2])
                    {
                        int ordemParte = 1;
                        foreach (var parte in sessaoProgramada.Partes)
                        {
                            if (parte.TituloParte.Contains("Necessidades locais"))
                            {
                                parte.AdicionarDesignado(designacoesSemanaImportadas.FirstOrDefault(d => d.Parte == "N").Designado);
                                continue;
                            }

                            if (parte.TituloParte.Contains("Estudo bíblico"))
                            {
                                parte.AdicionarDesignado(designacoesSemanaImportadas.FirstOrDefault(d => d.Parte == "E").Designado);
                                continue;
                            }

                            var pessoa = designacoesSemanaImportadas.FirstOrDefault(d => d.Parte == ordemParte.ToString() && d.Tipo.Contains("Anciãos"));
                            if (pessoa != null)
                            {
                                parte.AdicionarDesignado(pessoa.Designado);
                            }
                            ordemParte++;
                        }
                    }
                }
            }
            return reunioesProgramacao;
        }
    }
}
