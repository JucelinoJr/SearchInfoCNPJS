using ClosedXML.Excel;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using searchInfoCNPJS.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading;
using System.Web.Helpers;
using System.Web.UI;


namespace searchInfoCNPJS
{
    
    public partial class _Default : Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            
        }
        protected void btnConsultar_Click(object sender, EventArgs e)
        {
            var xls = new XLWorkbook(@"C:\searchInfoCNPJS\cnpjporcidade.xlsx");
            var qtdPlanilha = xls.Worksheets.Count;

            for (int p = 3; p <= qtdPlanilha; p++) { 
                var planilha = xls.Worksheet(p);
                var totalLinhas = planilha.LastRowUsed().RowNumber();

                for (int l = 2; l <= totalLinhas; l++)
                {
                    var B = planilha.Cell($"B{l}").Value;
                    var C = planilha.Cell($"C{l}").Value;
                    var D = planilha.Cell($"D{l}").Value;
                    var E = planilha.Cell($"E{l}").Value;
                    var F = planilha.Cell($"F{l}").Value;

                if(B != "" && C != "" && D != "" && E != "" && F != "")
                    { continue; }
                    else { 

                        string codigo = planilha.Cell($"A{l}").GetHyperlink().ExternalAddress.AbsolutePath;
                        //consultar API de CNPJ
                        var valores = GetCNPJ(codigo);

                        //telefones
                        if (valores.estabelecimento.ddd1 == null && valores.estabelecimento.telefone1 == null)
                        {planilha.Cell($"B{l}").SetValue("");}
                        else
                        {planilha.Cell($"B{l}").SetValue(valores.estabelecimento.ddd1 + valores.estabelecimento.telefone1 + " , " + valores.estabelecimento.ddd2 + valores.estabelecimento.telefone2);}
                    
                        //email
                        if (valores.estabelecimento.email == null)
                        {planilha.Cell($"C{l}").SetValue("");}
                        else
                        { planilha.Cell($"C{l}").SetValue(valores.estabelecimento.email);}

                        //endereço
                        planilha.Cell($"D{l}").SetValue(valores.estabelecimento.logradouro +" , "+ valores.estabelecimento.numero +" , "+ valores.estabelecimento.bairro);

                        //socios
                        if (valores.socios.Length == 0)
                        {planilha.Cell($"E{l}").SetValue(valores.socios);}
                        else
                        {
                            int qtdSocios = valores.socios.Length;
                            var list = new List<string>();
                            var socios = "";
                            for (int i = 0; i < qtdSocios; i++)
                            {
                                list.Add(valores.socios[i].nome);
                                socios = String.Join(",", list);
                            }
                            planilha.Cell($"E{l}").SetValue(socios);
                        }
                        planilha.Cell($"F{l}").SetValue(valores.estabelecimento.atividade_principal.descricao);
                        xls.Save();
                        Thread.Sleep(25000);
                    }
                }
            }


        }

        public dynamic GetCNPJ(string cnpj)
        {
            var PJ = new PessoaJuridica();

            HttpClient httpClient = new HttpClient();
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Ssl3;
            ServicePointManager.ServerCertificateValidationCallback += (sender, certificate, chain, sslPolicyErrors) => true;

            //httpClient.BaseAddress = new Uri("https://www.receitaws.com.br");
            httpClient.BaseAddress = new Uri("https://publica.cnpj.ws/cnpj");
            HttpResponseMessage resposta = httpClient.GetAsync("/cnpj"+cnpj).Result;
            
            resposta.EnsureSuccessStatusCode();
            var model = resposta.Content.ReadAsStringAsync().Result;
            var jsonfinal = Json.Decode(model);

            return jsonfinal;

        }

    }


}
