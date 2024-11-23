﻿using DesignacoesReuniao.Domain.Models;
using DesignacoesReuniao.Infra.Repostories.Interface;

namespace DesignacoesReuniao.Infra.Repostories
{
    public class PessoaRepository : IPessoaRepository
    {
        private List<Pessoa> pessoas;
        public PessoaRepository()
        {
            pessoas = InstanciarListaPessoas();
        }
        public Pessoa? BuscarPessoa(string nome)
        {
            if (string.IsNullOrEmpty(nome))
            {
                return null;
            }
            var palavras = nome.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

            var pessoa = pessoas.FirstOrDefault(p =>
                palavras.All(palavra => p.NomeCompleto.IndexOf(palavra, StringComparison.OrdinalIgnoreCase) >= 0));
            if (pessoa != null) 
            {
                return pessoa;
            }
            return new Pessoa(nome, nome);
        }

        private List<Pessoa> InstanciarListaPessoas()
        {
            return new List<Pessoa>
            {
                // Anciãos
                new Pessoa("Sebastiao Valdir Santana", "Valdir Santana"),
                new Pessoa("Marcos Alves de Souza", "Marcos Alves"),
                new Pessoa("Claudeísio Coelho", "Claudeísio Coelho"),
                new Pessoa("Lídio Ferreira", "Lídio Ferreira"),
                new Pessoa("Vigilato Andrade da Silva", "Vigilato Andrade"),
                new Pessoa("Guilherme Bonni", "Guilherme Bonni"),
                new Pessoa("Jessé Rodrigues da Silva", "Jessé Rodrigues"),
                new Pessoa("Moisés Alves Pereira", "Moisés Pereira"),
                new Pessoa("Rodrigo da Silva Cruz", "Rodrigo Cruz"),

                // Servos
                new Pessoa("Cristiano Batista Pessoa", "Cristiano Batista"),
                new Pessoa("Marcos Rodrigues de Souza", "Marcos Rodrigues"),
                new Pessoa("Marcos Vinícius Gargliardi", "Marcos Vinícius"),
                new Pessoa("Claudineto Ferraz e Silva", "Claudineto Ferraz"),
                new Pessoa("João Paulo", "João Paulo"),
                new Pessoa("William Pereira", "William Pereira"),
                new Pessoa("Claudinei Ferraz", "Claudinei Ferraz"),
                new Pessoa("Lucas Nunes dos Santos", "Lucas Santos"),
                new Pessoa("Marcos Rossin", "Marcos Rossin"),
                new Pessoa("Douglas Brisola da Costa", "Douglas Costa"),

                // Estudantes
                new Pessoa("Doralice Santos Santana", "Doralice Santana"),
                new Pessoa("Elizabeth Gomes da Silva", "Elizabeth Silva"),
                new Pessoa("Luana Rodrigues de Sá", "Luana Sá"),
                new Pessoa("Maria das Neves Patrício da", "Maria Patrício"),
                new Pessoa("Maria José Gomes da Silva", "Maria José Silva"),
                new Pessoa("Flávia Granado Viana Meira", "Flávia Meira"),
                new Pessoa("BEATRIZ VILELA OLIVEIRA", "Beatriz Oliveira"),
                new Pessoa("Joana Santos Lessa", "Joana Lessa"),
                new Pessoa("Maria Helena Oliveira", "Maria Helena"),
                new Pessoa("Maria José de Morais", "Maria Morais"),
                new Pessoa("Susana Granado Viana Meira", "Susana Meira"),
                new Pessoa("Noemy do Carmo S.", "Noemy do Carmo"),
                new Pessoa("Patricia Rossin", "Patricia Rossin"),
                new Pessoa("Laídes Borges de Souza", "Laídes Souza"),
                new Pessoa("Ana Paula Procópio", "Ana Procópio"),
                new Pessoa("Carina Silva Rogeri Alves de", "Carina Alves"),
                new Pessoa("Larissa Silva Borges", "Larissa Borges"),
                new Pessoa("Ivanilde Souza e Silva", "Ivanilde Silva"),
                new Pessoa("Juraci de Almeida Lisboa", "Juraci Lisboa"),
                new Pessoa("Maria Aparecida Viera", "Maria Viera"),
                new Pessoa("Maria Marta de Faria", "Marta Faria"),
                new Pessoa("Munike Ferraz", "Munike Ferraz"),
                new Pessoa("Simone Morais Ferraz", "Simone Ferraz"),
                new Pessoa("Rosilene Lemes Leal", "Rosilene Leal"),
                new Pessoa("Milena Almeida Alves", "Milena Alves"),
                new Pessoa("Adriana da Silva Coelho", "Adriana Coelho"),
                new Pessoa("Caroline Coelho de Souza", "Caroline Coelho"),
                new Pessoa("Ellen Abreu", "Ellen Abreu"),
                new Pessoa("Luciene Brito de Almeida", "Luciene Almeida"),
                new Pessoa("Maristela Cunha dos Santos", "Maristela Santos"),
                new Pessoa("Zaine Cruz Almeida", "Zaine Almeida"),
                new Pessoa("Laurinda N. Souza", "Laurinda Souza"),
                new Pessoa("Maria da Graça Almeida", "Graça Almeida"),
                new Pessoa("Simone Cunha", "Simone Cunha"),
                new Pessoa("Geovana Santos Barreto", "Geovana Barreto"),
                new Pessoa("Conceição de Maria Souza", "Conceição Souza"),
                new Pessoa("Inácia Coelho de Souza", "Inácia Souza"),
                new Pessoa("Júlia Oliveira", "Júlia Oliveira"),
                new Pessoa("Natália Pontes Pereira", "Natália Pereira"),
                new Pessoa("Bárbara Renata Teodoro", "Bárbara Teodoro"),
                new Pessoa("Tamara Ferreira Costa Brisola", "Tamara Brisola"),
                new Pessoa("Lucy Azevedo da Silva", "Lucy Silva"),
                new Pessoa("Vera Lúcia Bastocellis Ruiz", "Vera Ruiz"),
                new Pessoa("Josiane de Oliveira Pereira", "Josiane Pereira"),
                new Pessoa("Regina Aparecida Cunha", "Regina Cunha"),
                new Pessoa("Claudenísia Coelho de Souza", "Claudenísia Souza"),
                new Pessoa("Elisete Timoteo Jesus da", "Elisete Jesus"),
                new Pessoa("Isabelli Vasconcelos de", "Isabelli Vasconcelos"),
                new Pessoa("Katia Albuquerque A.", "Katia Albuquerque"),
                new Pessoa("Elizier Moura", "Elizier Moura"),
                new Pessoa("Lucimar Cardoso Menezes", "Lucimar Menezes"),
                new Pessoa("Maria de Nazaré Gomes", "Maria Gomes"),
                new Pessoa("Neuza Maria Bento Silva", "Neuza Silva")
            };


        }
    }
}