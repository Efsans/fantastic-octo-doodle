using Newtonsoft.Json;
using SolidWorks.Interop.sldworks;
using System;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Net.Http;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows;
using Xarial.XCad.Documents;
using Xarial.XCad.SolidWorks;

namespace FormsAndWpfControls
{
    public static class FuncoesExternas
    {
        public static void Acao1()
        {
            try
            {
                var swApp = (SldWorks)Marshal.GetActiveObject("SldWorks.Application");
                var model = swApp?.IActiveDoc2;

                if (model == null)
                {
                    MessageBox.Show("Nenhum documento aberto no SolidWorks.", "Aviso", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                string nomeArquivo = model.GetPathName();
                if (string.IsNullOrEmpty(nomeArquivo))
                {
                    MessageBox.Show("O arquivo ainda não foi salvo.", "Aviso",MessageBoxButton.OK , MessageBoxImage.Warning);
                    return;
                }

                string codigo = Path.GetFileNameWithoutExtension(nomeArquivo);

                // Use o CustomPropertyManager para garantir atualização
                var custPropMgr = model.Extension.get_CustomPropertyManager("");
                int result = custPropMgr.Add3("Codigo", (int)SolidWorks.Interop.swconst.swCustomInfoType_e.swCustomInfoText, codigo, (int)SolidWorks.Interop.swconst.swCustomPropertyAddOption_e.swCustomPropertyReplaceValue);
                // Se Add3 não funcionar, use Set para garantir
                custPropMgr.Set("Codigo", codigo);

                model.ForceRebuild3(false);

                MessageBox.Show($"Código '{codigo}' salvo com sucesso!", "Sucesso", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Erro", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        public static async void Acao2()
        {
            try
            {
                var swApp = (SldWorks)Marshal.GetActiveObject("SldWorks.Application");
                var model = swApp?.IActiveDoc2;

                if (model == null)
                {
                    MessageBox.Show("Nenhum documento aberto no SolidWorks.", "Aviso", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                var props = model.GetCustomInfoNames2("") as string[];
                if (props == null)
                {
                    MessageBox.Show("Nenhuma propriedade personalizada encontrada.", "Aviso", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                var dados = new System.Collections.Generic.Dictionary<string, string>();

                foreach (var prop in props)
                {
                    string valor = model.GetCustomInfoValue("", prop);
                    dados[prop] = valor ?? "";
                }

                var jsonPayload = JsonConvert.SerializeObject(new { dados });

                using (HttpClient client = new HttpClient())
                {
                    var content = new StringContent(jsonPayload, Encoding.UTF8, "application/json");

                    var response = await client.PostAsync("https://www.zohoapis.com/creator/custom/grupoaiz/SolidWorks?publickey=4WTWAfSnDWdjzatDCYr6gyJ4B", content);

                    if (response.IsSuccessStatusCode)
                    {
                        MessageBox.Show("Enviado com sucesso!", "Sucesso", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                    else
                    {
                        var erro = await response.Content.ReadAsStringAsync();
                        MessageBox.Show($"Erro ao enviar: {response.StatusCode}", "Aviso", MessageBoxButton.OK, MessageBoxImage.Warning);
                        MessageBox.Show(erro, "Erro", MessageBoxButton.OK, MessageBoxImage.Warning);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Erro", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        public static void Acao3()
        {
            try
            {
                var swApp = (SldWorks)Marshal.GetActiveObject("SldWorks.Application");
                var model = swApp?.IActiveDoc2;

                if (model == null)
                {
                    MessageBox.Show("Nenhum documento aberto no SolidWorks!", "Erro", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                var custPropMgr = model.Extension.get_CustomPropertyManager("");

                string codigo_key = "Codigo";
                string codigo = custPropMgr.Get(codigo_key);
                string filial = "09ALFA";

                if (string.IsNullOrEmpty(codigo))
                {
                    var resp = MessageBox.Show("Código não está preenchido! Deseja o auto preenchimento.", "Atenção", MessageBoxButton.YesNo, MessageBoxImage.Warning);
                    if (resp == MessageBoxResult.Yes)
                    {
                        AtualizarCodigo(model);
                        return;
                    }
                    return;
                    }

                using (var conexao = new SqlConnection(GetConnectionString()))
                {
                    conexao.Open();
                    string query = @"
                        SELECT 
                            B1_COD,
                            B1_TIPO,
                            B1_UM,
                            RTRIM(B1_DESC) AS B1_DESC,
                            B1_POSIPI,
                            B1_ORIGEM,
                            B1_FILIAL
                        FROM SB1010 SB1
                        WHERE SB1.D_E_L_E_T_ = '' AND SB1.B1_FILIAL = @filial AND SB1.B1_COD = @codigo
                    ";

                    using (var cmd = new SqlCommand(query, conexao))
                    {
                        cmd.Parameters.AddWithValue("@filial", filial);
                        cmd.Parameters.AddWithValue("@codigo", codigo);

                        using (var reader = cmd.ExecuteReader())
                        {
                            if (!reader.Read())
                            {
                                MessageBox.Show("Código não encontrado na base de dados!", "Atenção", MessageBoxButton.OK, MessageBoxImage.Warning);
                                return;
                            }

                            ProcessarResposta(model, reader);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Erro", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private static void ProcessarResposta(IModelDoc2 model, IDataRecord row)
        {
            if (row == null)
            {
                MessageBox.Show("Nenhum dado encontrado! Verifique a tabela ou os parâmetros.", "Aviso", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var mapeamento = new System.Collections.Generic.Dictionary<string, string>
            {
                {"B1_COD", "Codigo"},
                {"B1_TIPO", "Tipo"},
                {"B1_UM", "Unidade"},
                {"B1_DESC", "Descrição"},
                {"B1_POSIPI", "Pos.IPI/NCM"},
                {"B1_ORIGEM", "Origem"}
            };

            var custPropMgr = model.Extension.get_CustomPropertyManager("");
            int camposAlterados = 0;

            foreach (var kvp in mapeamento)
            {
                string coluna = kvp.Key;
                string nomeProp = kvp.Value;
                object valor = null;
                try { valor = row[coluna]; } catch { valor = null; }
                if (valor != null)
                {
                    string valorStr = valor.ToString();
                    try
                    {
                        // Sempre sobrescreve e garante atualização
                        custPropMgr.Add3(nomeProp, (int)SolidWorks.Interop.swconst.swCustomInfoType_e.swCustomInfoText, valorStr, (int)SolidWorks.Interop.swconst.swCustomPropertyAddOption_e.swCustomPropertyReplaceValue);
                        custPropMgr.Set(nomeProp, valorStr);
                        camposAlterados++;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Erro ao definir a propriedade {nomeProp}: {ex.Message}");
                    }
                }
            }

            if (camposAlterados > 0)
            {
                MessageBox.Show($"Atualização concluída! {camposAlterados} campos alterados.", "Sucesso", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            else
            {
                MessageBox.Show("Nenhum campo foi atualizado.", "Aviso", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private static string GetConnectionString()
        {
            // Retorna a string de conexão com o SQL Server
            return "Data Source=TOTVSAPL;Initial Catalog=protheus12_producao;User ID=consulta;Password=consulta;";
        }

        private static void AtualizarCodigo(IModelDoc2 model)
        {
            // Implemente a lógica de auto-preenchimento do código conforme necessário
            // Exemplo simples: pega o nome do arquivo
            try
            {
                string nomeArquivo = model.GetPathName();
                if (!string.IsNullOrEmpty(nomeArquivo))
                {
                    string codigo = Path.GetFileNameWithoutExtension(nomeArquivo);
                    model.AddCustomInfo3("", "Codigo", 30, codigo);
                    model.ForceRebuild3(false);
                    MessageBox.Show("Código foi preenchido automaticamente!", "Sucesso", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                else
                {
                    MessageBox.Show("O arquivo ainda não foi salvo, impossível preencher o código.", "Aviso", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erro ao preencher código automaticamente: " + ex.Message, "Erro", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}