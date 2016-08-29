using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.IO;
using System.Net;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PG_EnvioNF.Logic
{
    public class OrderBusinessLogic
    {
        IOrganizationService service;

        /// <summary>
        /// Construtor da classe 
        /// </summary>
        /// <param name="_service">provedor de serviço da organização</param>
        public OrderBusinessLogic(IOrganizationService _service)
        {
            service = _service;
        }

        /// <summary>
        /// Método principal de envio de informações da NF para o serviço da integração 
        /// </summary>
        public void EnviarNF(Entity order)
        {
            /// Declaração das variáveis 
            String notaFiscal = String.Empty;
            EntityCollection orderCollection;

            /// Validação: se o novo status é igual a "completo", prossegue.
            if (((OptionSetValue)order["statuscode"]).Value == 1)
            {
                orderCollection = RetrieveOrderDetails(order.Id);
                /// Validação dos resultados 
                if (orderCollection != null && orderCollection.Entities.Count != 0)
                {
                    notaFiscal = preencherNF(order, orderCollection).ToString();


                    // preencher no format notaFiscal

                    

                    /// TODO: integração com o serviço responsável por receber os dados  
                    /// 

                  var ret =  RequisicaoPost(notaFiscal);





                }
            }
        }

        /// <summary>
        /// Busca de itens do pedido
        /// </summary>
        /// <param name="orderID">ID do pedido</param>
        /// <returns>coleção com todos itens do pedido solicitado</returns>
        public EntityCollection RetrieveOrderDetails(Guid orderID)
        {
            /// Declaração de variáveis 
            EntityCollection resultsCollection;

            /// Construção da consulta
            QueryExpression query = new QueryExpression("salesorderdetail");
            query.ColumnSet = new ColumnSet(true);
            query.Criteria.AddCondition("salesorderid", ConditionOperator.Equal, orderID);

            /// Consulta à API
            resultsCollection = service.RetrieveMultiple(query);

            /// Retorno dos resultados 
            return resultsCollection;
        }

        /// <summary>
        /// Preenchimento da NF 
        /// </summary>
        /// <param name="pedido">objeto do pedido</param>
        /// <param name="itensPedido">coleção com os itens do pedido</param>
        /// <returns></returns>
        public StringBuilder preencherNF(Entity pedido, EntityCollection itensPedido)
        {
            /// Declaração das variáveis 
            //String notaFiscal1 = String.Empty;
            var notaFiscal = new StringBuilder();
                        
            notaFiscal.AppendLine($"formato=notaFiscal{pedido.Attributes["new_numeropedidomaxweb"]}");
            notaFiscal.AppendLine($"numLote=0");
            notaFiscal.AppendLine($"grupo=edoc");
            //notaFiscal.AppendLine($"cnpj=03939020000110");
            notaFiscal.AppendLine($"INCLUIR");
            //notaFiscal.AppendLine($"Id_A03=0{vazio}");
            notaFiscal.AppendLine($"versao_A02=3.10");
            notaFiscal.AppendLine($"cUF_B02=41");
            //notaFiscal.AppendLine($"cNF_B03=01{vazio}");
            notaFiscal.AppendLine($"natOp_B04=VENDA DE PRODUCAO FORA DO ESTADO");
            //notaFiscal.AppendLine($"indPag_B05=1{vazio}");
            //notaFiscal.AppendLine($"mod_B06=55{vazio}");
            //notaFiscal.AppendLine($"serie_B07=666{vazio}");
            //notaFiscal.AppendLine($"nNF_B08=1400{vazio}");
            //notaFiscal.AppendLine($"dhEmi_B09={DateTime.Now.ToString("yyyy-MM-ddTHH:mm:sszzz")}");
            //notaFiscal.AppendLine($"dhSaiEnt_B10={DateTime.Now.ToString("yyyy-MM-ddTHH:mm:sszzz")}");
            //notaFiscal.AppendLine($"tpNF_B11=1{vazio}");
            //notaFiscal.AppendLine($"idDest_B11a=2{vazio}");
            //notaFiscal.AppendLine($"cMunFG_B12=4115200{vazio}");
            //notaFiscal.AppendLine($"tpImp_B21=1{vazio}");
            //notaFiscal.AppendLine($"tpEmis_B22=1{vazio}");
            //notaFiscal.AppendLine($"cDV_B23=0{vazio}");
            //notaFiscal.AppendLine($"tpAmb_B24=2{vazio}");
            //notaFiscal.AppendLine($"finNFe_B25=1{vazio}");
            //notaFiscal.AppendLine($"indFinal_B25a=0{vazio}");
            //notaFiscal.AppendLine($"indPres_B25b=9{vazio}");
            //notaFiscal.AppendLine($"procEmi_B26=0{vazio}");
            //notaFiscal.AppendLine($"verProc_B27=000{vazio}");

            notaFiscal.AppendLine($"CRT_C21=1");

            notaFiscal.AppendLine($"CNPJ_C02=03939020000110");
            notaFiscal.AppendLine($"xNome_C03=NUTRYERVAS DO BRASIL LTDA - MATRIZ");
            notaFiscal.AppendLine($"xFant_C04=NUTRYERVAS DO BRASIL");
            notaFiscal.AppendLine($"xLgr_C06=RUA ROCHA POMBO, 184");
            notaFiscal.AppendLine($"nro_C07=184");
            notaFiscal.AppendLine($"xBairro_C09=Centro");
            notaFiscal.AppendLine($"cMun_C10=4114203");
            notaFiscal.AppendLine($"xMun_C11=MANDAGUARI");
            notaFiscal.AppendLine($"UF_C12=PR");
            notaFiscal.AppendLine($"CEP_C13=86975000");
            notaFiscal.AppendLine($"cPais_C14=1058");
            notaFiscal.AppendLine($"xPais_C15=Brasil");

            notaFiscal.AppendLine($"IE_C17=9038099229");


            notaFiscal.AppendLine($"CNPJ_E02=23389116000160");
            notaFiscal.AppendLine($"xNome_E04=LISANDRO MAYER 01911800922");
            notaFiscal.AppendLine($"xLgr_E06=AV. ANTONIO ANDRE MAGGI, 1140 SL 2");
            notaFiscal.AppendLine($"nro_E07=11402");
            notaFiscal.AppendLine($"xBairro_E09=CENTRO");
            notaFiscal.AppendLine($"cMun_E10=5107875");
            notaFiscal.AppendLine($"xMun_E11=SAPEZAL");
            notaFiscal.AppendLine($"UF_E12=MT");
            notaFiscal.AppendLine($"CEP_E13=78365000");
            notaFiscal.AppendLine($"cPais_E14=1058");
            notaFiscal.AppendLine($"xPais_E15=BRASIL");
            notaFiscal.AppendLine($"fone_E16=6533832374");
            notaFiscal.AppendLine($"indIEDest_E16a=9");
            //notaFiscal.AppendLine($"IE_E17=1570027061{vazio}");



            foreach (Entity en in itensPedido.Entities)
            {

                notaFiscal.AppendLine($"INCLUIRITEM");

                //notaFiscal.AppendLine($"nItem_H02={n++}");
                notaFiscal.AppendLine($"cProd_I02=00001");
                //notaFiscal.AppendLine($"CEAN_I03={vazio}");
                notaFiscal.AppendLine($"xProd_I04=LEVEDO 400 COMPRIMIDOS");
                notaFiscal.AppendLine($"NCM_I05=21022000");
                notaFiscal.AppendLine($"CFOP_I08=6101");
                //notaFiscal.AppendLine($"uCom_I09=UN");
                notaFiscal.AppendLine($"qCom_I10={en.Attributes["quantity"]}");
                //notaFiscal.AppendLine($"vUnCom_I10a=15.52{vazio}");
                //notaFiscal.AppendLine($"vProd_I11=155.20{vazio}");
                //notaFiscal.AppendLine($"cEANTrib_I12={vazio}");
                //notaFiscal.AppendLine($"uTrib_I13=Un{vazio}");
                //notaFiscal.AppendLine($"qTrib_I14=10.0000{vazio}");
                //notaFiscal.AppendLine($"vUnTrib_I14a=15.5200{vazio}");
                //notaFiscal.AppendLine($"indTot_I17b=1{vazio}");

                //notaFiscal.AppendLine($"orig_N11=0{vazio}");

                //notaFiscal.AppendLine($"CST_N12=00{vazio}");
                //notaFiscal.AppendLine($"MODBC_N13=3{vazio}");
                //notaFiscal.AppendLine($"VBC_N15=155.20{vazio}");
                //notaFiscal.AppendLine($"PICMS_N16=12.00{vazio}");
                //notaFiscal.AppendLine($"VICMS_N17=18.62{vazio}");
                //notaFiscal.AppendLine($"modBCST_N18=4{vazio}");
                //notaFiscal.AppendLine($"vBCST_N21=0.00{vazio}");
                //notaFiscal.AppendLine($"pICMSST_N22=0.00{vazio}");
                //notaFiscal.AppendLine($"vICMSST_N23=0.00{vazio}");

                //notaFiscal.AppendLine($"CST_Q06=99{vazio}");
                //notaFiscal.AppendLine($"VBC_Q07=155.20{vazio}");
                //notaFiscal.AppendLine($"PPIS_Q08=5.00{vazio}");
                //notaFiscal.AppendLine($"VPIS_Q09=7.76{vazio}");

                //notaFiscal.AppendLine($"CST_S06=99{vazio}");
                //notaFiscal.AppendLine($"VBC_S07=155.20{vazio}");
                //notaFiscal.AppendLine($"PCOFINS_S08=5.00{vazio}");
                //notaFiscal.AppendLine($"VCOFINS_S11={}");


                notaFiscal.AppendLine($"SALVARITEM");
            }

            //notaFiscal.AppendLine($"vBC_W03=155.20{vazio}");
            //notaFiscal.AppendLine($"vICMS_W04=18.62{vazio}");
            //notaFiscal.AppendLine($"vICMSDeson_W04a=0.00{vazio}");
            //notaFiscal.AppendLine($"vBCST_W05=0.00{vazio}");
            //notaFiscal.AppendLine($"vST_W06=0.00{vazio}");
            //notaFiscal.AppendLine($"vProd_W07=155.20{vazio}");
            //notaFiscal.AppendLine($"vFrete_W08=0.00{vazio}");
            //notaFiscal.AppendLine($"vSeg_W09=0.00{vazio}");
            //notaFiscal.AppendLine($"vDesc_W10=0.00{vazio}");
            //notaFiscal.AppendLine($"vII_W11=0.00{vazio}");
            //notaFiscal.AppendLine($"vIPI_W12=0.00{vazio}");
            //notaFiscal.AppendLine($"vPIS_W13=7.76{vazio}");
            //notaFiscal.AppendLine($"vCOFINS_W14=7.76{vazio}");
            //notaFiscal.AppendLine($"vOutro_W15=0.00{vazio}");
            //notaFiscal.AppendLine($"vNF_W16=155.20{vazio}");


            notaFiscal.AppendLine($"modFrete_X02=0");

            //notaFiscal.AppendLine($"INCLUIRCOBRANCA{vazio}");
            //notaFiscal.AppendLine($"nFat_Y03=2000{vazio}");
            //notaFiscal.AppendLine($"vOrig_Y04=500.00{vazio}");
            //notaFiscal.AppendLine($"vDesc_Y05=100.00{vazio}");
            //notaFiscal.AppendLine($"vLiq_Y06=400.00{vazio}");

            //notaFiscal.AppendLine($"nDup_Y08=1{vazio}");
            //notaFiscal.AppendLine($"dVenc_Y09=2009-04-25{vazio}");
            //notaFiscal.AppendLine($"vDup_Y10=100.00{vazio}");

            //notaFiscal.AppendLine($"nDup_Y08=2{vazio}");
            //notaFiscal.AppendLine($"dVenc_Y09=2009-04-25{vazio}");
            //notaFiscal.AppendLine($"vDup_Y10=100.00{vazio}");

            //notaFiscal.AppendLine($"nDup_Y08=3{vazio}");
            //notaFiscal.AppendLine($"dVenc_Y09=2009-04-25{vazio}");
            //notaFiscal.AppendLine($"vDup_Y10=100.00{vazio}");

            //notaFiscal.AppendLine($"nDup_Y08=4{vazio}");
            //notaFiscal.AppendLine($"dVenc_Y09=2009-04-25{vazio}");
            //notaFiscal.AppendLine($"vDup_Y10=100.00{vazio}");
            //notaFiscal.AppendLine($"SALVARCOBRANCA{vazio}");

            notaFiscal.AppendLine($"infAdFisco_Z02=OBSERVACAO TESTE DA DANFE - FISCO");
            notaFiscal.AppendLine($"infCpl_Z03=OBSERVACAO TESTE DA DANFE CONTRIBUINTE");
            notaFiscal.AppendLine($"EmailDestinatario={pedido.Attributes["new_numeropedidomaxweb"]}");
            notaFiscal.AppendLine($"SALVAR");


            /// Retorno da string de NF gerada
            //return notaFiscal1;
            return notaFiscal;
        }


        public string RequisicaoPost(string notaFiscal)
        {
            
            var url = " https://managersaashom.tecnospeed.com.br:7071/ManagerAPIWeb/nfe/envia?CNPJ=08187168000160&grupo=edoc";

            try
            {
                HttpWebRequest objRequest = (HttpWebRequest)WebRequest.Create(url + "&arquivo=" + notaFiscal);

                byte[] credentialBuffer = new UTF8Encoding().GetBytes("admin:123mudar");
                objRequest.Headers["Authorization"] = "Basic " + Convert.ToBase64String(credentialBuffer);

                objRequest.Method = "POST";

                HttpWebResponse objResponse = (HttpWebResponse)objRequest.GetResponse();

                Stream stream = objResponse.GetResponseStream();

                Encoding encoding = Encoding.Default;

                StreamReader response = new StreamReader(stream, encoding);

                return response.ReadToEnd();

               // Console.WriteLine(retorno);
            }
            catch (Exception erro)
            {
                return "Erro: "+erro.Message;
            }

        }

    }
}
