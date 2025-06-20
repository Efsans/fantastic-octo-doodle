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
using System.Linq; // Para usar .Select e .ToArray
using SolidWorks.Interop.swconst;
using Xarial.XCad.SolidWorks.Services;
using System.Collections.Generic; // Para List<string>
using System.Diagnostics; // Para DebugMessage

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
                    MessageBox.Show("O arquivo ainda não foi salvo.", "Aviso", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                string codigo = Path.GetFileNameWithoutExtension(nomeArquivo);
                var resp = MessageBox.Show($"Deseja salvar o Código '{codigo}' ?", "Confirmação", MessageBoxButton.YesNo, MessageBoxImage.Information);
                if (resp == MessageBoxResult.No)
                {
                    return; // Usuário cancelou a ação
                }
                // Use o CustomPropertyManager para garantir atualização
                var materialName = model.Extension.get_CustomPropertyManager("");
                int result = materialName.Add3("Codigo", (int)SolidWorks.Interop.swconst.swCustomInfoType_e.swCustomInfoText, codigo, (int)SolidWorks.Interop.swconst.swCustomPropertyAddOption_e.swCustomPropertyReplaceValue);
                // Se Add3 não funcionar, use Set para garantir
                materialName.Set("Codigo", codigo);

                model.ForceRebuild3(false);

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

                var result = MessageBox.Show($"deseja enviar '{jsonPayload}' para o sistema?", "Confirmação", MessageBoxButton.YesNo, MessageBoxImage.Question);
                if (result != MessageBoxResult.Yes)
                {
                    return; // Usuário cancelou a ação
                }

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

                var materialName = model.Extension.get_CustomPropertyManager("");

                string codigo_key = "Codigo";
                string codigo = materialName.Get(codigo_key);
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

            var materialName = model.Extension.get_CustomPropertyManager("");
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
                        materialName.Add3(nomeProp, (int)SolidWorks.Interop.swconst.swCustomInfoType_e.swCustomInfoText, valorStr, (int)SolidWorks.Interop.swconst.swCustomPropertyAddOption_e.swCustomPropertyReplaceValue);
                        materialName.Set(nomeProp, valorStr);
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

        public static void Acao4()
        {
            SldWorks swApp = null;
            IModelDoc2 swModel = null;
            StringBuilder sb = new StringBuilder();

            try
            {
                // 1. Obter a instância ativa do SolidWorks
                try
                {
                    swApp = (SldWorks)Marshal.GetActiveObject("SldWorks.Application");
                    DebugMessage("SolidWorks Application obtido com sucesso.");
                }
                catch (COMException ex)
                {
                    DebugMessage($"Erro COM ao obter SolidWorks Application: {ex.Message}");
                    MessageBox.Show("SolidWorks não está em execução ou o Add-in não conseguiu se conectar.", "Erro de Conexão SW", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }
                catch (Exception ex)
                {
                    DebugMessage($"Erro geral ao obter SolidWorks Application: {ex.Message}");
                    MessageBox.Show($"Ocorreu um erro inesperado ao conectar ao SolidWorks: {ex.Message}", "Erro", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                // 2. Obter o documento ativo
                swModel = swApp?.IActiveDoc2;
                if (swModel == null)
                {
                    DebugMessage("Nenhum documento ativo encontrado. Abortando Acao4.");
                    MessageBox.Show("Nenhum documento aberto no SolidWorks. Por favor, abra um documento.", "Aviso", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
                DebugMessage($"Documento ativo encontrado: '{swModel.GetTitle() ?? "Não Salvo"}'");

                // Verificar se o documento é uma peça
                if (swModel.GetType() != (int)swDocumentTypes_e.swDocPART)
                {
                    DebugMessage($"Documento ativo não é uma peça. Tipo: {swModel.GetType()}. Abortando Acao4.");
                    MessageBox.Show("Esta função é relevante apenas para documentos de peça (.SLDPRT), pois lida com propriedades de material.", "Aviso", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                IPartDoc swPart = swModel as IPartDoc;
                string materialName = "";
                string materialId = "";

                if (swPart != null)
                {
                    // 3. Obter o nome do material aplicado à peça
                    materialName = swPart.GetMaterialPropertyName2("", out materialId);

                    if (!string.IsNullOrWhiteSpace(materialName))
                    {
                        sb.AppendLine($"Material Aplicado ao Documento: {materialName}");
                        DebugMessage($"Material aplicado: '{materialName}' (ID: {materialId})");

                        // 4. Montar o nome correto para o CustomPropertyManager
                        string customPropMgrName = $"Material@{materialName}";

                        // 5. Tentar obter o CustomPropertyManager
                        var materialCustPropMgr = swModel.Extension?.get_CustomPropertyManager(customPropMgrName);

                        // Se ainda assim for nulo, tente sem espaços extras ou caracteres especiais
                        if (materialCustPropMgr == null && materialName.Contains("|"))
                        {
                            // Alguns materiais podem vir com "nome|caminho", tente só o nome antes do pipe
                            string nomeLimpo = materialName.Split('|')[0].Trim();
                            customPropMgrName = $"Material@{nomeLimpo}";
                            materialCustPropMgr = swModel.Extension?.get_CustomPropertyManager(customPropMgrName);
                        }

                        sb.AppendLine("\n--- Propriedades Personalizadas do MATERIAL (Aba 'Personalizado' no diálogo de material) ---");

                        if (materialCustPropMgr == null)
                        {
                            sb.AppendLine("Não foi possível acessar as propriedades personalizadas do material. Certifique-se de que o material está corretamente aplicado e que possui propriedades personalizadas.");
                            DebugMessage($"materialCustPropMgr retornou nulo para '{customPropMgrName}'.");
                        }
                        else
                        {
                            var propNames = materialCustPropMgr.GetNames() as string[];

                            if (propNames != null && propNames.Length > 0)
                            {
                                foreach (var prop in propNames)
                                {
                                    string val = materialCustPropMgr.Get(prop);
                                    sb.AppendLine($"{prop}: {val ?? "[Vazio]"}");
                                }
                                DebugMessage($"Encontradas {propNames.Length} propriedades personalizadas do material.");
                            }
                            else
                            {
                                sb.AppendLine("Nenhuma propriedade personalizada encontrada na aba 'Personalizado' para este material.");
                                DebugMessage("Nenhuma propriedade personalizada do material encontrada.");
                            }
                        }
                    }
                    else
                    {
                        sb.AppendLine("Nenhum material explicitamente aplicado à peça.");
                        sb.AppendLine("Não foi possível buscar propriedades personalizadas do material.");
                        DebugMessage("Nenhum material aplicado à peça.");
                    }
                }
                else
                {
                    sb.AppendLine("Documento não é uma peça. Não há material para obter propriedades personalizadas.");
                    DebugMessage("Documento não é uma peça.");
                }

                // 6. Exibir a lista resultante
                MessageBox.Show(sb.ToString(), "Propriedades Personalizadas do Material", MessageBoxButton.OK, MessageBoxImage.Information);
                DebugMessage("Acao4 concluída com sucesso.");
            }
            catch (COMException comEx)
            {
                DebugMessage($"ERRO COM EXCEPTION: {comEx.Message} (HRESULT: {comEx.ErrorCode:X8})");
                MessageBox.Show($"Ocorreu um erro de comunicação com o SolidWorks:\n{comEx.Message}\n(Código: 0x{comEx.ErrorCode:X8})", "Erro de COM", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            catch (Exception ex)
            {
                DebugMessage($"ERRO GERAL EXCEPTION: {ex.Message}\nStackTrace: {ex.StackTrace}");
                MessageBox.Show($"Ocorreu um erro inesperado:\n{ex.Message}\n\nDetalhes para Depuração:\n{ex.StackTrace}", "Erro Inesperado", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// Função auxiliar para exibir mensagens de depuração no Output/Debug window (se um debugger estiver anexado).
        /// </summary>
        private static void DebugMessage(string message)
        {
            System.Diagnostics.Debug.WriteLine($"[Acao4 Debug] {DateTime.Now:HH:mm:ss.fff} - {message}");
        }
    }
}