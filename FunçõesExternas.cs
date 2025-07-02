using Newtonsoft.Json;          // Biblioteca para manipulação de JSON: Essencial para serializar e deserializar objetos C# para e de JSON, usado na comunicação com APIs RESTful.
using SolidWorks.Interop.sldworks; // SOLIDWORKS API principal: Permite a interação programática direta com o aplicativo SOLIDWORKS.
using SolidWorks.Interop.swconst; // Constantes da SOLIDWORKS API: Fornece enums e constantes para tipos de documentos, opções e comportamentos da API.
using System;                   // Core .NET: Tipos fundamentais do sistema, como DateTime, Exception e classes base para programação.
using System.Collections.Generic; // Coleções genéricas: Usado para estruturas de dados como List<T> e Dictionary<TKey, TValue>, essenciais para organização de dados.
using System.Data;              // Acesso a dados: Classes base para interação com fontes de dados, como ADO.NET.
using System.Data.SqlClient;    // SQL Server Client: Provedor de dados específico para Microsoft SQL Server, crucial para a conexão com o Protheus.
using System.Diagnostics;       // Diagnóstico e Debug: Oferece classes para depuração, rastreamento de eventos e temporizadores, útil para logs (ex: DebugMessage).
using System.IO;                // Operações de I/O de Arquivos: Permite ler, escrever e manipular arquivos e diretórios no sistema de arquivos.
using System.Linq;              // LINQ (Language Integrated Query): Facilita consultas e transformações de dados em coleções (ex: .Select, .FirstOrDefault, .Any).
using System.Net.Http;          // HTTP Client: Ferramenta moderna para enviar requisições HTTP e manipular respostas, ideal para integração com serviços web.
using System.Runtime.InteropServices; // Interoperabilidade COM: Necessário para que o código C# possa se comunicar com aplicações COM, como o SOLIDWORKS.
using System.Text;              // Manipulação de Texto: Classes para construção eficiente de strings (StringBuilder) e codificação de caracteres.
using System.Windows;           // WPF Core: Componentes base do Windows Presentation Foundation, incluindo a exibição de mensagens (MessageBox).
using System.Windows.Controls;  // Controles WPF: Elementos de interface de usuário como botões, painéis, caixas de texto (UserControl, StackPanel, TextBox).
using System.Windows.Media.Media3D; // Gráficos 3D WPF: Embora presente, pode não ser diretamente usado neste contexto, mas oferece funcionalidades para modelos 3D em UI.
using System.Xml;               // Processamento XML: Suporte básico para leitura e escrita de documentos XML.
using System.Xml.Linq;          // LINQ to XML: Uma API mais moderna e intuitiva para manipulação de XML usando LINQ, crucial para arquivos .sldmat.
using Xarial.XCad.Documents;   // XCAD Documents API: Abstrações de documentos para CAD, fornecendo uma interface unificada para diferentes plataformas.
using Xarial.XCad.SolidWorks;  // XCAD SOLIDWORKS API: Extensões da XCAD específicas para a plataforma SOLIDWORKS, simplificando interações complexas.
using Xarial.XCad.SolidWorks.Services; // XCAD SOLIDWORKS Services: Serviços de baixo nível e utilitários específicos da XCAD para SOLIDWORKS.

namespace FormsAndWpfControls
{
    /// <summary>
    /// <para>Esta classe estática `FuncoesExternas` atua como um hub para operações automatizadas
    /// e integrações cruciais para o departamento de engenharia da AIZ, diretamente do SolidWorks.</para>
    /// <para>Ela encapsula lógicas para:
    /// <list type="bullet">
    /// <item>Automatizar a inserção de códigos de peça.</item>
    /// <item>Sincronizar propriedades de modelos com sistemas externos (e.g., Zoho Creator).</item>
    /// <item>Atualizar propriedades de modelos com dados provenientes do ERP (Protheus).</item>
    /// <item>Validar e extrair informações de materiais de bancos de dados customizados do SolidWorks.</item>
    /// </list>
    /// </para>
    /// <para>Todas as funcionalidades são expostas como métodos estáticos para fácil acesso
    /// e não dependem de uma instância específica da classe.</para>
    /// </summary>
    /// <remarks>
    /// Embora a classe herde de `UserControl`, seus métodos principais são estáticos.
    /// Isso sugere que ela pode ter evoluído de um controle de UI, mas agora serve
    /// primariamente como uma biblioteca de utilitários de back-end.
    /// Os campos de instância `swApp`, `swModel` e `swPropMgr` não são utilizados pelos métodos estáticos e poderiam ser removidos
    /// para uma estrutura mais limpa se a classe for puramente estática.
    /// </remarks>
    public class FuncoesExternas : UserControl
    {
        // ============================================================================================
        // MEMBROS DE CLASSE (Instâncias da API do SolidWorks)
        // ATENÇÃO: Estes membros não são utilizados pelos métodos estáticos. Se a classe for puramente estática, podem ser removidos.
        // ============================================================================================
        public SldWorks swApp;                 // Referência à instância principal do aplicativo SolidWorks.
        private IModelDoc2 swModel;            // Referência ao documento ativo no SolidWorks (peça, montagem ou desenho).
        private CustomPropertyManager swPropMgr; // Gerenciador para acessar e modificar as propriedades personalizadas do documento.

        /// <summary>
        /// <para><b>Ação 1: Salvar Código da Peça no SolidWorks.</b></para>
        /// <para>Esta função automatiza o preenchimento da propriedade personalizada "Codigo"
        /// de um documento SolidWorks. Ela extrai o nome do arquivo (sem extensão)
        /// e o atribui como o código da peça, garantindo consistência.</para>
        /// </summary>
        /// <remarks>
        /// <para><b>Pré-requisitos:</b></para>
        /// <list type="bullet">
        /// <item>SolidWorks deve estar em execução e com um documento (peça, montagem ou desenho) ativo.</item>
        /// <item>O documento ativo deve ter sido salvo pelo menos uma vez.</item>
        /// </list>
        /// <para><b>Fluxo:</b></para>
        /// <list type="number">
        /// <item>Obtém a instância ativa do SolidWorks e o documento atual.</item>
        /// <item>Valida a existência e o estado salvo do documento.</item>
        /// <item>Extrai o nome do arquivo como o código.</item>
        /// <item>Solicita confirmação do usuário antes de proceder.</item>
        /// <item>Define/Atualiza a propriedade "Codigo" usando `CustomPropertyManager.Add3` e `Set` para robustez.</item>
        /// <item>Força a reconstrução do modelo para que a propriedade seja visualmente atualizada no SolidWorks.</item>
        /// </list>
        /// </remarks>
        public static void Acao1()
        {
            try
            {
                // Tenta obter a instância ativa do SolidWorks através da tabela de objetos em execução (ROT).
                var swApp = (SldWorks)Marshal.GetActiveObject("SldWorks.Application");
                // Obtém uma referência ao documento ativo no SolidWorks.
                var model = swApp?.IActiveDoc2;

                // --- Validação Inicial do Documento ---
                if (model == null)
                {
                    MessageBox.Show("Nenhum documento SolidWorks está aberto ou ativo. Por favor, abra um documento e tente novamente.", "Aviso: Documento Não Encontrado", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                string nomeArquivo = model.GetPathName();
                if (string.IsNullOrEmpty(nomeArquivo))
                {
                    MessageBox.Show("O arquivo SolidWorks ainda não foi salvo. Por favor, salve o documento antes de executar esta ação.", "Aviso: Arquivo Não Salvo", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                // Extrai o nome do arquivo sem a extensão, que será usado como o "Codigo".
                string codigo = Path.GetFileNameWithoutExtension(nomeArquivo);

                // --- Confirmação do Usuário ---
                var resp = MessageBox.Show($"Deseja realmente definir '{codigo}' como a propriedade 'Codigo' para este documento?", "Confirmação de Código", MessageBoxButton.YesNo, MessageBoxImage.Information);
                if (resp == MessageBoxResult.No)
                {
                    MessageBox.Show("Operação cancelada pelo usuário.", "Informação", MessageBoxButton.OK, MessageBoxImage.Information);
                    return; // Usuário cancelou a ação.
                }

                // --- Atualização da Propriedade Personalizada ---
                // Obtém o gerenciador de propriedades personalizadas do documento SolidWorks (a string vazia "" refere-se ao nível do documento).
                var customPropMgr = model.Extension.get_CustomPropertyManager("");

                // Adiciona ou atualiza a propriedade "Codigo".
                // swCustomInfoType_e.swCustomInfoText: Define o tipo da propriedade como texto.
                // swCustomPropertyAddOption_e.swCustomPropertyReplaceValue: Garante que, se a propriedade já existir, seu valor será sobrescrito.
                customPropMgr.Add3("Codigo", (int)SolidWorks.Interop.swconst.swCustomInfoType_e.swCustomInfoText, codigo, (int)SolidWorks.Interop.swconst.swCustomPropertyAddOption_e.swCustomPropertyReplaceValue);

                // Chama 'Set' adicionalmente para garantir que a propriedade seja atualizada em cenários onde 'Add3' pode ter comportamento inconsistente (raro, mas é uma camada de segurança).
                customPropMgr.Set("Codigo", codigo);

                // Força a reconstrução do modelo. Isso é essencial para que as propriedades personalizadas sejam atualizadas
                // e visíveis nas tabelas de lista de materiais, desenhos, etc.
                // O parâmetro 'false' indica uma reconstrução mais leve, não um "rebuild all".
                model.ForceRebuild3(false);

                MessageBox.Show($"Propriedade 'Codigo' definida com sucesso para '{codigo}'.", "Sucesso", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                // Captura e exibe qualquer exceção inesperada, fornecendo detalhes ao usuário.
                MessageBox.Show($"Ocorreu um erro ao definir o código da peça: {ex.Message}", "Erro Inesperado", MessageBoxButton.OK, MessageBoxImage.Error);
                Debug.WriteLine($"Erro em Acao1: {ex.ToString()}"); // Para log de depuração mais detalhado.
            }
        }

        /// <summary>
        /// <para><b>Ação 2: Sincronizar Propriedades com o Sistema Zoho Creator.</b></para>
        /// <para>Esta função lê todas as propriedades personalizadas do documento SolidWorks ativo,
        /// as empacota em um formato JSON e as envia para uma aplicação específica no Zoho Creator
        /// através de uma requisição HTTP POST.</para>
        /// </summary>
        /// <remarks>
        /// <para><b>Pré-requisitos:</b></para>
        /// <list type="bullet">
        /// <item>SolidWorks deve estar em execução e com um documento ativo.</item>
        /// <item>Conexão ativa com a internet para acessar o Zoho Creator.</item>
        /// <item>A URL e a chave pública do Zoho Creator devem estar corretas.</item>
        /// </list>
        /// <para><b>Fluxo:</b></para>
        /// <list type="number">
        /// <item>Obtém a instância ativa do SolidWorks e o documento atual.</item>
        /// <item>Extrai todos os nomes e valores das propriedades personalizadas do documento.</item>
        /// <item>Serializa os dados coletados em um objeto JSON.</item>
        /// <item>Solicita confirmação do usuário antes de enviar os dados.</item>
        /// <item>Utiliza `HttpClient` para enviar o JSON via POST para a API do Zoho Creator.</item>
        /// <item>Verifica o status da resposta HTTP e informa o usuário sobre o sucesso ou falha da operação.</item>
        /// </list>
        /// </remarks>
        public static async void Acao2()
        {
            try
            {
                // Obtém a instância ativa do SolidWorks.
                var swApp = (SldWorks)Marshal.GetActiveObject("SldWorks.Application");
                // Obtém uma referência ao documento ativo.
                var model = swApp?.IActiveDoc2;

                // --- Validação Inicial do Documento ---
                if (model == null)
                {
                    MessageBox.Show("Nenhum documento SolidWorks está aberto ou ativo. Por favor, abra um documento para coletar as propriedades.", "Aviso: Documento Não Encontrado", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                // Obtém um array com os nomes de todas as propriedades personalizadas do documento.
                var propNames = model.GetCustomInfoNames2("") as string[];
                if (propNames == null || propNames.Length == 0)
                {
                    MessageBox.Show("Este documento não possui propriedades personalizadas para serem enviadas. Verifique as propriedades e tente novamente.", "Aviso: Sem Propriedades", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                // Cria um dicionário para armazenar os dados das propriedades (Nome da Propriedade -> Valor).
                var dadosParaEnvio = new Dictionary<string, string>();

                // Itera sobre os nomes das propriedades e recupera seus valores.
                foreach (var propName in propNames)
                {
                    string propValue = model.GetCustomInfoValue("", propName);
                    // Adiciona a propriedade ao dicionário, tratando valores nulos como string vazia.
                    dadosParaEnvio[propName] = propValue ?? "";
                }

                // Serializa o dicionário de dados em uma string JSON.
                // O objeto anônimo `{ dados }` cria um JSON com uma chave "dados" contendo o dicionário.
                var jsonPayload = JsonConvert.SerializeObject(new { dados = dadosParaEnvio });

                // --- Confirmação do Usuário para Envio ---
                var result = MessageBox.Show($"Deseja enviar os seguintes dados para o sistema externo?\n\n{jsonPayload}\n\n(Verifique os dados antes de continuar)", "Confirmação de Envio de Dados", MessageBoxButton.YesNo, MessageBoxImage.Question);
                if (result != MessageBoxResult.Yes)
                {
                    MessageBox.Show("Envio de dados cancelado pelo usuário.", "Informação", MessageBoxButton.OK, MessageBoxImage.Information);
                    return; // Usuário cancelou a ação.
                }

                // --- Envio de Dados via HTTP POST ---
                using (HttpClient client = new HttpClient())
                {
                    // Define o conteúdo da requisição HTTP como JSON com codificação UTF-8.
                    var content = new StringContent(jsonPayload, Encoding.UTF8, "application/json");

                    // Envia a requisição POST de forma assíncrona para a URL do Zoho Creator.
                    // O 'await' permite que a UI permaneça responsiva enquanto a requisição é processada.
                    var response = await client.PostAsync("https://www.zohoapis.com/creator/custom/grupoaiz/SolidWorks?publickey=4WTWAfSnDWdjzatDCYr6gyJ4B", content);

                    // Verifica se o código de status da resposta indica sucesso (2xx).
                    if (response.IsSuccessStatusCode)
                    {
                        MessageBox.Show("Dados enviados com sucesso para o sistema externo!", "Envio Concluído", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                    else
                    {
                        // Se houver um erro, lê o corpo da resposta para obter mensagens de erro detalhadas do servidor.
                        var errorMessage = await response.Content.ReadAsStringAsync();
                        MessageBox.Show($"Falha ao enviar dados para o sistema externo.\nStatus: {response.StatusCode} - {response.ReasonPhrase}\nDetalhes: {errorMessage}", "Erro no Envio de Dados", MessageBoxButton.OK, MessageBoxImage.Error);
                        Debug.WriteLine($"Erro em Acao2 (HTTP Status: {response.StatusCode}): {errorMessage}"); // Para log de depuração.
                    }
                }
            }
            catch (Exception ex)
            {
                // Captura e exibe qualquer exceção inesperada que ocorra durante o processo de envio.
                MessageBox.Show($"Ocorreu um erro ao sincronizar as propriedades: {ex.Message}", "Erro Inesperado", MessageBoxButton.OK, MessageBoxImage.Error);
                Debug.WriteLine($"Erro em Acao2: {ex.ToString()}"); // Para log de depuração mais detalhado.
            }
        }

        /// <summary>
        /// <para><b>Ação 3: Atualizar Propriedades do SolidWorks com Dados do Protheus.</b></para>
        /// <para>Esta função consulta o banco de dados SQL Server (Protheus) usando o "Codigo"
        /// da peça SolidWorks e, se encontrar, preenche ou atualiza outras propriedades
        /// personalizadas do SolidWorks com as informações correspondentes do ERP.</para>
        /// </summary>
        /// <remarks>
        /// <para><b>Pré-requisitos:</b></para>
        /// <list type="bullet">
        /// <item>SolidWorks deve estar em execução e com um documento ativo.</item>
        /// <item>A propriedade personalizada "Codigo" deve estar preenchida no documento SolidWorks.</item>
        /// <item>Conectividade com o servidor SQL Server do Protheus.</item>
        /// <item>As credenciais de banco de dados definidas em <see cref="GetConnectionString"/> devem ser válidas.</item>
        /// </list>
        /// <para><b>Fluxo:</b></para>
        /// <list type="number">
        /// <item>Obtém a instância ativa do SolidWorks e o documento atual.</item>
        /// <item>Verifica a existência da propriedade "Codigo" e oferece auto-preenchimento se ausente.</item>
        /// <item>Abre uma conexão com o banco de dados SQL Server.</item>
        /// <item>Executa uma consulta parametrizada na tabela `SB1010` do Protheus.</item>
        /// <item>Se encontrar o código, chama <see cref="ProcessarResposta"/> para atualizar as propriedades do SolidWorks.</item>
        /// <item>Informa o usuário sobre o resultado da consulta e atualização.</item>
        /// </list>
        /// </remarks>
        public static void Acao3()
        {
            try
            {
                // Obtém a instância ativa do SolidWorks.
                var swApp = (SldWorks)Marshal.GetActiveObject("SldWorks.Application");
                // Obtém uma referência ao documento ativo.
                var model = swApp?.IActiveDoc2;

                // --- Validação Inicial do Documento ---
                if (model == null)
                {
                    MessageBox.Show("Nenhum documento SolidWorks está aberto ou ativo. Por favor, abra um documento para buscar dados do ERP.", "Erro: Documento Não Encontrado", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                // Obtém o gerenciador de propriedades personalizadas do documento.
                var customPropMgr = model.Extension.get_CustomPropertyManager("");

                const string CODIGO_PROPERTY_NAME = "Codigo"; // Constante para o nome da propriedade "Codigo".
                string codigo = customPropMgr.Get(CODIGO_PROPERTY_NAME); // Obtém o valor da propriedade "Codigo".
                string filialProtheus = "09ALFA"; // Filial padrão utilizada na consulta ao Protheus.

                // --- Verificação e Auto-Preenchimento do Código ---
                if (string.IsNullOrEmpty(codigo))
                {
                    var response = MessageBox.Show($"A propriedade '{CODIGO_PROPERTY_NAME}' não está preenchida no documento SolidWorks. Deseja tentar auto-preenchê-la com o nome do arquivo?", "Atenção: Código Ausente", MessageBoxButton.YesNo, MessageBoxImage.Warning);
                    if (response == MessageBoxResult.Yes)
                    {
                        AtualizarCodigo(model); // Tenta preencher automaticamente.
                        // Após o auto-preenchimento, o 'Codigo' pode ter sido definido, mas esta ação não irá continuar imediatamente
                        // a consulta ao banco de dados. O usuário precisaria executar Acao3 novamente se quisesse consultar.
                        // Se a intenção for continuar, seria necessário re-obter 'codigo' após 'AtualizarCodigo'.
                        return; // Sai da função após a tentativa de auto-preenchimento.
                    }
                    MessageBox.Show("Ação de atualização de propriedades cancelada devido à ausência do Código.", "Informação", MessageBoxButton.OK, MessageBoxImage.Information);
                    return; // Usuário não quis auto-preencher ou cancelou.
                }

                // --- Conexão e Consulta ao Banco de Dados (Protheus) ---
                using (var conexao = new SqlConnection(GetConnectionString()))
                {
                    conexao.Open(); // Abre a conexão com o banco de dados.

                    // Query SQL para buscar informações do item na tabela de produtos (SB1010) do Protheus.
                    // `RTRIM(B1_DESC)` é usado para remover espaços em branco à direita do campo de descrição.
                    string query = @"
                        SELECT  
                            B1_COD,    -- Código do Produto
                            B1_TIPO,   -- Tipo do Produto
                            B1_UM,     -- Unidade de Medida
                            RTRIM(B1_DESC) AS B1_DESC, -- Descrição do Produto (tratada para remover espaços extras)
                            B1_POSIPI, -- Posição Fiscal (NCM)
                            B1_ORIGEM, -- Origem do Produto
                            B1_FILIAL  -- Filial do Produto
                        FROM SB1010 SB1
                        WHERE SB1.D_E_L_E_T_ = ''    -- Garante que o registro não está marcado para deleção lógica
                          AND SB1.B1_FILIAL = @filial -- Filtra pela filial especificada
                          AND SB1.B1_COD = @codigo;   -- Filtra pelo código do produto
                    ";

                    using (var cmd = new SqlCommand(query, conexao))
                    {
                        // Adiciona parâmetros para a consulta SQL. Isso é crucial para segurança (prevenção de SQL Injection)
                        // e melhor performance.
                        cmd.Parameters.AddWithValue("@filial", filialProtheus);
                        cmd.Parameters.AddWithValue("@codigo", codigo);

                        using (var reader = cmd.ExecuteReader())
                        {
                            // Verifica se o leitor retornou alguma linha (se o código foi encontrado no Protheus).
                            if (!reader.Read())
                            {
                                MessageBox.Show($"O código '{codigo}' não foi encontrado na base de dados do Protheus para a filial '{filialProtheus}'.", "Atenção: Código Não Encontrado no ERP", MessageBoxButton.OK, MessageBoxImage.Warning);
                                return;
                            }

                            // Processa a linha de dados encontrada e atualiza as propriedades do SolidWorks.
                            ProcessarResposta(model, reader);
                        }
                    }
                }
            }
            catch (SqlException sqlEx)
            {
                // Captura e exibe exceções específicas de SQL, fornecendo detalhes do erro do banco de dados.
                MessageBox.Show($"Erro de Banco de Dados: Não foi possível conectar ou consultar o Protheus.\nDetalhes: {sqlEx.Message}", "Erro de Conexão/Consulta SQL", MessageBoxButton.OK, MessageBoxImage.Error);
                Debug.WriteLine($"Erro SQL em Acao3: {sqlEx.ToString()}"); // Log detalhado para depuração.
            }
            catch (Exception ex)
            {
                // Captura e exibe qualquer outra exceção inesperada.
                MessageBox.Show($"Ocorreu um erro inesperado ao atualizar as propriedades: {ex.Message}", "Erro Inesperado", MessageBoxButton.OK, MessageBoxImage.Error);
                Debug.WriteLine($"Erro geral em Acao3: {ex.ToString()}"); // Log detalhado para depuração.
            }
        }

        /// <summary>
        /// <para><b>Processar Resposta do Banco de Dados.</b></para>
        /// <para>Este método auxiliar recebe uma linha de dados (registro) de uma consulta SQL
        /// e mapeia os valores das colunas para as propriedades personalizadas correspondentes
        /// no documento SolidWorks ativo.</para>
        /// </summary>
        /// <param name="model">A instância <see cref="IModelDoc2"/> do documento SolidWorks para ser atualizado.</param>
        /// <param name="row">Um objeto <see cref="IDataRecord"/> contendo os dados de uma linha do banco de dados.</param>
        /// <remarks>
        /// <para><b>Mapeamento:</b></para>
        /// <para>O método utiliza um dicionário para definir a correspondência entre os nomes das colunas
        /// do Protheus (e.g., "B1_COD") e os nomes das propriedades personalizadas no SolidWorks (e.g., "Codigo").</para>
        /// <para><b>Atualização:</b></para>
        /// <para>Para cada par mapeado, ele tenta obter o valor do `IDataRecord` e, se disponível,
        /// atualiza a propriedade correspondente no SolidWorks. Utiliza `Add3` com a opção `swCustomPropertyReplaceValue`
        /// e `Set` para garantir que a propriedade seja criada ou seu valor substituído.</para>
        /// </remarks>
        private static void ProcessarResposta(IModelDoc2 model, IDataRecord row)
        {
            // Validação de entrada: Garante que há dados para processar.
            if (row == null)
            {
                MessageBox.Show("Dados nulos recebidos para processamento. Não foi possível atualizar as propriedades.", "Aviso: Dados Vazios", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // Dicionário de mapeamento: Chave = Nome da Coluna no Banco de Dados; Valor = Nome da Propriedade no SolidWorks.
            var columnPropertyMapping = new Dictionary<string, string>
            {
                {"B1_COD", "Codigo"},
                {"B1_TIPO", "Tipo"},
                {"B1_UM", "Unidade"},
                {"B1_DESC", "Descrição"},
                {"B1_POSIPI", "Pos.IPI/NCM"},
                {"B1_ORIGEM", "Origem"}
            };

            // Obtém o gerenciador de propriedades personalizadas para o documento.
            var customPropMgr = model.Extension.get_CustomPropertyManager("");
            int updatedFieldsCount = 0; // Contador para saber quantos campos foram efetivamente alterados.

            // Itera sobre cada entrada no dicionário de mapeamento.
            foreach (var mapping in columnPropertyMapping)
            {
                string columnName = mapping.Key;   // Nome da coluna no banco de dados (ex: "B1_COD").
                string propertyName = mapping.Value; // Nome da propriedade personalizada no SolidWorks (ex: "Codigo").
                object columnValue = null;

                try
                {
                    // Tenta obter o valor da coluna pelo seu nome no IDataRecord.
                    // Isso pode falhar se a coluna não existir no conjunto de resultados da consulta.
                    columnValue = row[columnName];
                }
                catch (IndexOutOfRangeException)
                {
                    // Se a coluna não for encontrada (ex: nome da coluna incorreto na query ou no mapeamento),
                    // apenas loga o aviso e continua para a próxima propriedade.
                    Console.WriteLine($"Aviso: Coluna '{columnName}' não encontrada nos dados do banco de dados. A propriedade '{propertyName}' não será atualizada.");
                    continue; // Pula para a próxima iteração.
                }
                catch (Exception ex)
                {
                    // Captura outras exceções que possam ocorrer ao tentar acessar a coluna.
                    Console.WriteLine($"Erro ao acessar a coluna '{columnName}': {ex.Message}");
                    continue; // Pula para a próxima iteração.
                }

                // Se o valor da coluna for DBNull (equivalente a NULL no banco de dados) ou null.
                if (columnValue == null || columnValue == DBNull.Value)
                {
                    // Opcional: Poderia-se decidir limpar a propriedade no SolidWorks se o valor for nulo no banco.
                    // customPropMgr.Delete(propertyName); // Exemplo de como deletar.
                    Console.WriteLine($"A propriedade '{propertyName}' não foi atualizada pois o valor da coluna '{columnName}' é nulo ou vazio.");
                    continue; // Pula para a próxima iteração.
                }

                string valueAsString = columnValue.ToString(); // Converte o valor do banco de dados para string.

                try
                {
                    // Define/Atualiza a propriedade personalizada no SolidWorks.
                    // Usamos swCustomInfoText (30) para o tipo da propriedade.
                    // swCustomPropertyReplaceValue garante que a propriedade será atualizada se já existir.
                    customPropMgr.Add3(propertyName, (int)SolidWorks.Interop.swconst.swCustomInfoType_e.swCustomInfoText, valueAsString, (int)SolidWorks.Interop.swconst.swCustomPropertyAddOption_e.swCustomPropertyReplaceValue);
                    customPropMgr.Set(propertyName, valueAsString); // Chamada redundante, mas que assegura a definição.
                    updatedFieldsCount++; // Incrementa o contador de campos atualizados.
                }
                catch (Exception ex)
                {
                    // Loga erros específicos que ocorrem ao tentar definir uma propriedade.
                    Console.WriteLine($"Erro ao definir a propriedade SolidWorks '{propertyName}' com valor '{valueAsString}': {ex.Message}");
                }
            }

            // Exibe o resultado final da atualização.
            if (updatedFieldsCount > 0)
            {
                MessageBox.Show($"Atualização de propriedades concluída! Um total de {updatedFieldsCount} campos foram atualizados no documento SolidWorks.", "Atualização Concluída", MessageBoxButton.OK, MessageBoxImage.Information);
                model.ForceRebuild3(false); // Força a reconstrução para refletir as novas propriedades.
            }
            else
            {
                MessageBox.Show("Nenhum campo de propriedade foi atualizado no documento SolidWorks.", "Aviso: Nenhuma Atualização", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        /// <summary>
        /// <para><b>Obter String de Conexão com o SQL Server.</b></para>
        /// <para>Este método retorna a string de conexão para o banco de dados SQL Server do Protheus.
        /// Ela é utilizada para estabelecer a comunicação com o ERP.</para>
        /// </summary>
        /// <returns>Uma <see cref="string"/> contendo os detalhes da conexão com o SQL Server.</returns>
        /// <remarks>
        /// <para><b>Segurança:</b></para>
        /// <para>**ATENÇÃO**: Manter strings de conexão com credenciais diretamente no código-fonte
        /// não é uma prática recomendada para ambientes de produção. Considere utilizar
        /// arquivos de configuração (e.g., `App.config`, `Web.config`) ou um sistema de
        /// gerenciamento de segredos para maior segurança e flexibilidade.</para>
        /// </remarks>
        private static string GetConnectionString()
        {
            // String de conexão para o SQL Server.
            // Data Source: Endereço do servidor SQL.
            // Initial Catalog: Nome do banco de dados.
            // User ID/Password: Credenciais de acesso.
            return "Data Source=TOTVSAPL;Initial Catalog=protheus12_producao;User ID=consulta;Password=consulta;";
        }

        /// <summary>
        /// <para><b>Atualizar a Propriedade 'Codigo' Automaticamente.</b></para>
        /// <para>Este método é um utilitário para preencher a propriedade personalizada "Codigo"
        /// de um documento SolidWorks usando o nome do arquivo atual (sem extensão).</para>
        /// </summary>
        /// <param name="model">A instância <see cref="IModelDoc2"/> do documento SolidWorks a ser atualizado.</param>
        /// <remarks>
        /// <para><b>Processo:</b></para>
        /// <list type="number">
        /// <item>Verifica se o documento SolidWorks está salvo.</item>
        /// <item>Extrai o nome do arquivo (sem extensão) para ser o código.</item>
        /// <item>Define a propriedade "Codigo" no SolidWorks, forçando uma reconstrução para aplicar a mudança.</item>
        /// </list>
        /// <para>É frequentemente chamada por outras funções (e.g., `Acao3`) quando o "Codigo" está ausente.</para>
        /// </remarks>
        private static void AtualizarCodigo(IModelDoc2 model)
        {
            try
            {
                string fileName = model.GetPathName();
                if (!string.IsNullOrEmpty(fileName))
                {
                    string code = Path.GetFileNameWithoutExtension(fileName); // Obtém o nome do arquivo sem a extensão.

                    // Adiciona ou atualiza a propriedade "Codigo". O '30' representa o tipo de propriedade de texto (swCustomInfoText).
                    model.AddCustomInfo3("", "Codigo", 30, code);
                    model.ForceRebuild3(false); // Força a reconstrução do modelo para que a propriedade seja atualizada.
                    MessageBox.Show($"A propriedade 'Codigo' foi preenchida automaticamente com '{code}'.", "Sucesso: Auto-preenchimento", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                else
                {
                    MessageBox.Show("O documento SolidWorks não foi salvo. Não é possível preencher o código automaticamente.", "Aviso: Arquivo Não Salvo", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
            catch (Exception ex)
            {
                // Captura e exibe erros que ocorram durante o auto-preenchimento.
                MessageBox.Show($"Erro ao tentar auto-preencher o código: {ex.Message}", "Erro: Auto-preenchimento", MessageBoxButton.OK, MessageBoxImage.Error);
                Debug.WriteLine($"Erro em AtualizarCodigo: {ex.ToString()}");
            }
        }

        /// <summary>
        /// <para><b>Ação 0: Método Placeholder.</b></para>
        /// <para>Este método é um placeholder e atualmente não possui implementação funcional.
        /// Ele retorna `null` e pode ser utilizado para testes ou como ponto de partida
        /// para futuras funcionalidades.</para>
        /// </summary>
        /// <returns>Sempre retorna <c>null</c>.</returns>
        public static void Acao0()
        {

            
        }

        /// <summary>
        /// <para><b>Verificar e Exibir Propriedades de Material.</b></para>
        /// <para>Esta função é responsável por identificar o material aplicado ao documento de peça ativo no SolidWorks,
        /// e então buscar as propriedades personalizadas desse material em uma lista de arquivos `.sldmat` especificados.
        /// As informações do material e suas propriedades customizadas são exibidas ao usuário.</para>
        /// </summary>
        /// <param name="filePaths">Uma coleção de <see cref="string"/>s, onde cada string é o caminho completo
        /// para um arquivo de banco de dados de materiais do SolidWorks (`.sldmat`).</param>
        /// <remarks>
        /// <para><b>Fluxo de Verificação:</b></para>
        /// <list type="number">
        /// <item>Obtém a instância ativa do SolidWorks e o documento atual.</item>
        /// <item>Verifica se o documento é uma peça (`swDocPART`) e se um material está aplicado.</item>
        /// <item>Extrai o nome do material aplicado (ex: "Aço 1020").</item>
        /// <item>Itera sobre os `filePaths` fornecidos, carregando e analisando cada arquivo `.sldmat` como XML.</item>
        /// <item>Utiliza a função <see cref="BuscarMaterialRecursivo"/> para localizar o material e sua categoria dentro do XML.</item>
        /// <item>Se o material for encontrado, extrai suas propriedades customizadas (se existirem).</item>
        /// <item>Apresenta um `MessageBox` detalhado com todas as informações coletadas do material.</item>
        /// </list>
        /// <para><b>Tratamento de Erros:</b></para>
        /// <para>Inclui tratamento robusto para `COMException` (erros de comunicação com SolidWorks)
        /// e `Exception` geral, fornecendo feedback claro ao usuário.</para>
        /// </remarks>
        public static void RunMaterialCheck(IEnumerable<string> filePaths)
        {
            SldWorks swApp = null;     // Instância do aplicativo SolidWorks.
            IModelDoc2 swModel = null; // O documento ativo no SolidWorks.

            try
            {
                // Tenta obter a instância ativa do SolidWorks.
                swApp = (SldWorks)Marshal.GetActiveObject("SldWorks.Application");
                swModel = swApp?.IActiveDoc2 as IModelDoc2;

                // --- Validações do Documento SolidWorks ---
                if (swModel == null)
                {
                    MessageBox.Show("Nenhum documento SolidWorks está aberto ou ativo. Por favor, abra uma peça.", "Erro: Documento SolidWorks Não Encontrado", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                // Verifica se o documento ativo é do tipo "Peça" (Part).
                if (swModel.GetType() != (int)swDocumentTypes_e.swDocPART)
                {
                    MessageBox.Show("O documento SolidWorks ativo não é uma peça (Part). Esta função só pode ser executada em peças.", "Aviso: Tipo de Documento Incorreto", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                // Obtém o nome completo do material aplicado à peça (formato "Biblioteca|NomeDoMaterial").
                string fullMaterialName = swModel.MaterialIdName;
                if (string.IsNullOrEmpty(fullMaterialName))
                {
                    MessageBox.Show("Nenhum material foi aplicado à peça SolidWorks. Por favor, aplique um material antes de executar esta verificação.", "Informação: Material Ausente", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                // Extrai apenas o nome do material (e.g., "Aço 1020" de "SolidWorks Materials|Aço 1020").
                string materialName = fullMaterialName.Split('|')[1];
                MessageBox.Show($"Material atualmente selecionado no SolidWorks: '{materialName}' (Caminho completo: '{fullMaterialName}')", "Material Identificado", MessageBoxButton.OK, MessageBoxImage.Information);

                XElement foundMaterialElement = null;   // Armazena o elemento XML do material encontrado.
                XElement foundCategoryElement = null;   // Armazena o elemento XML da categoria do material.
                string foundSldmatFilePath = null;      // O caminho do arquivo .sldmat onde o material foi localizado.

                // --- Busca o Material nos Arquivos .sldmat ---
                // Itera sobre a lista de caminhos de arquivos de materiais fornecidos.
                foreach (string path in filePaths)
                {
                    if (File.Exists(path))
                    {
                        try
                        {
                            // Carrega o arquivo .sldmat como um documento XML.
                            XDocument doc = XDocument.Load(path);
                            // Chama a função recursiva para buscar o material dentro deste documento XML.
                            var searchResult = BuscarMaterialRecursivo(doc.Root, materialName);

                            // Se o material foi encontrado neste arquivo .sldmat.
                            if (searchResult.material != null)
                            {
                                foundMaterialElement = searchResult.material;
                                foundCategoryElement = searchResult.categoria;
                                foundSldmatFilePath = path;
                                break; // Material encontrado, não há necessidade de verificar os outros arquivos.
                            }
                            else
                            {
                                Console.WriteLine($"Debug: Material '{materialName}' NÃO encontrado em '{Path.GetFileName(path)}'. Verificando o próximo arquivo...");
                            }
                        }
                        catch (XmlException xmlEx)
                        {
                            // Erro ao carregar ou analisar o XML de um arquivo .sldmat.
                            Console.WriteLine($"Aviso: Não foi possível carregar ou analisar o arquivo XML '{Path.GetFileName(path)}'. Detalhes: {xmlEx.Message}. Tentando o próximo caminho...");
                        }
                        catch (Exception fileEx)
                        {
                            // Outros erros durante o processamento do arquivo (ex: permissão negada).
                            Console.WriteLine($"Aviso: Erro ao processar o arquivo '{Path.GetFileName(path)}'. Detalhes: {fileEx.Message}. Tentando o próximo caminho...");
                        }
                    }
                    else
                    {
                        Console.WriteLine($"Debug: Arquivo '{Path.GetFileName(path)}' não encontrado no caminho especificado. Verificando o próximo caminho...");
                    }
                }

                // --- Resultado da Busca e Exibição ---
                if (foundMaterialElement == null)
                {
                    MessageBox.Show($"O material '{materialName}' não foi encontrado em nenhum dos arquivos .sldmat especificados. Verifique os bancos de dados de material.", "Aviso: Material Não Encontrado na Biblioteca", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                // Extrai o nome da categoria do material.
                string categoryName = foundCategoryElement.Attribute("name")?.Value ?? "Categoria Desconhecida";
                // Tenta obter o elemento 'custom' que contém as propriedades personalizadas do material.
                var customPropertiesElement = foundMaterialElement.Element("custom");
                // Dicionário para armazenar as propriedades customizadas para exibição ou uso futuro.
                var materialCustomProps = new Dictionary<string, string>();

                // Constrói a mensagem final para o usuário.
                StringBuilder messageBuilder = new StringBuilder();
                messageBuilder.AppendLine("--- Detalhes do Material SolidWorks ---");
                messageBuilder.AppendLine($"Arquivo .sldmat: {Path.GetFileName(foundSldmatFilePath) ?? "N/A"}");
                messageBuilder.AppendLine($"Classificação: {categoryName}");
                messageBuilder.AppendLine($"Material Aplicado: {materialName}");
                messageBuilder.AppendLine("\n--- Propriedades Customizadas do Material ---");

                if (customPropertiesElement != null && customPropertiesElement.Elements("prop").Any())
                {
                    foreach (var prop in customPropertiesElement.Elements("prop"))
                    {
                        string propName = prop.Attribute("name")?.Value ?? "[Nome Vazio]";
                        string propDescription = prop.Attribute("description")?.Value ?? "[Descrição Vazia]";
                        string propUnits = prop.Attribute("units")?.Value ?? "[Unidades Vazia]";
                        string propValue = prop.Attribute("value")?.Value ?? "[Valor Vazio]";

                        materialCustomProps[propName] = propValue; // Armazena a propriedade.
                        messageBuilder.AppendLine($"- {propName} ({propDescription}): {propValue} {propUnits}".Trim());
                    }
                }
                else
                {
                    messageBuilder.AppendLine("Nenhuma propriedade customizada foi definida para este material na biblioteca.");
                }

                MessageBox.Show(messageBuilder.ToString(), $"Propriedades do Material: {materialName}", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (COMException comEx)
            {
                // Erro específico da API COM do SolidWorks (ex: SolidWorks não está rodando, problema de marshaling).
                MessageBox.Show($"Um erro de comunicação com o SolidWorks ocorreu: {comEx.Message} (HRESULT: {comEx.ErrorCode}). Certifique-se de que o SolidWorks está aberto e funcionando corretamente.", "Erro de COM", MessageBoxButton.OK, MessageBoxImage.Error);
                Debug.WriteLine($"COM Exception in RunMaterialCheck: {comEx.ToString()}");
            }
            catch (Exception ex)
            {
                // Captura qualquer outra exceção geral.
                MessageBox.Show($"Ocorreu um erro inesperado durante a verificação do material: {ex.Message}\n\nDetalhes Técnicos:\n{ex.StackTrace}", "Erro Crítico", MessageBoxButton.OK, MessageBoxImage.Error);
                Debug.WriteLine($"General Exception in RunMaterialCheck: {ex.ToString()}");
            }
        }

        /// <summary>
        /// <para><b>Ação 4: Iniciar Verificação de Material com Caminhos Padrão.</b></para>
        /// <para>Este método é o ponto de entrada para a funcionalidade de verificação de material.
        /// Ele define uma lista predefinida de caminhos para arquivos `.sldmat` (incluindo caminhos de rede da AIZ)
        /// e chama a função principal <see cref="RunMaterialCheck"/> para iniciar o processo.</para>
        /// </summary>
        /// <remarks>
        /// Os caminhos listados devem ser configurados para refletir a estrutura de instalação do SolidWorks
        /// e a localização dos bancos de dados de materiais customizados da AIZ na rede.
        /// </remarks>
        public static void Acao4()
        {
            // Define uma lista de caminhos potenciais para os arquivos de banco de dados de materiais do SolidWorks (.sldmat).
            // Esta lista inclui caminhos padrão de instalação do SolidWorks e caminhos específicos da rede AIZ.
            var sldmatFilePaths = new List<string>
            {
                // Caminhos padrão de instalação do SolidWorks (idioma português-brasileiro).
                @"C:\Program Files\SOLIDWORKS Corp\SOLIDWORKS\lang\portuguese-brazilian\sldmaterials\solidworks materials.sldmat",
                @"C:\Program Files\SOLIDWORKS Corp\SOLIDWORKS\lang\portuguese-brazilian\sldmaterials\sustainability extras.sldmat",
                @"C:\Program Files\SOLIDWORKS Corp\SOLIDWORKS\lang\portuguese-brazilian\sldmaterials\SolidWorks DIN Materials.sldmat",
                
                // Caminhos de rede específicos da AIZ para materiais personalizados.
                // VERIFIQUE E ATUALIZE ESTES CAMINHOS CONFORME A INFRAESTRUTURA DA REDE AIZ.
                @"\\toronto\AIZI\TEMPLATES AIZI\MATERIAIS CADASTRADOS AIZ.sldmat",
                @"\\toronto\AIZ IMPLEMENTOS\TEMPLATES AIZI\Bancos de Dados de Material\Materiais personalizados.sldmat",
            };

            // Inicia o processo de verificação de material, passando a lista de caminhos para busca.
            RunMaterialCheck(sldmatFilePaths);
        }

        /// <summary>
        /// <para><b>Busca Recursiva por Material em Estrutura XML (`.sldmat`).</b></para>
        /// <para>Este método auxiliar percorre recursivamente os elementos XML de um arquivo `.sldmat`
        /// para localizar um material específico com base em seu nome. Ele procura tanto em elementos
        /// `material` diretos quanto dentro de `classification` aninhadas.</para>
        /// </summary>
        /// <param name="currentElement">O <see cref="XElement"/> atual a ser pesquisado (pode ser a raiz do documento ou um elemento `classification`).</param>
        /// <param name="materialToFind">O nome do material a ser encontrado (a busca é case-insensitive).</param>
        /// <returns>
        /// Uma tupla (<see cref="XElement"/> categoria, <see cref="XElement"/> material) contendo:
        /// <list type="bullet">
        /// <item><c>categoria</c>: O elemento <see cref="XElement"/> da `classification` onde o material foi encontrado.</item>
        /// <item><c>material</c>: O elemento <see cref="XElement"/> do `material` correspondente.</item>
        /// </list>
        /// Se o material não for encontrado, retorna <c>(null, null)</c>.
        /// </returns>
        /// <remarks>
        /// A estrutura esperada dos arquivos `.sldmat` é hierárquica, com elementos `<classification>`
        /// contendo outros `<classification>` ou `<material>`.
        /// </remarks>
        private static (XElement categoria, XElement material) BuscarMaterialRecursivo(XElement currentElement, string materialToFind)
        {
            // 1. Tenta encontrar o material diretamente entre os filhos 'material' do elemento atual.
            var foundMaterial = currentElement.Elements("material")
                // Compara o atributo 'name' do material de forma case-insensitive.
                .FirstOrDefault(m => string.Equals(m.Attribute("name")?.Value, materialToFind, StringComparison.OrdinalIgnoreCase));

            // Se o material for encontrado no nível atual da recursão.
            if (foundMaterial != null)
            {
                return (currentElement, foundMaterial); // Retorna o elemento pai (categoria) e o material.
            }

            // 2. Se o material não foi encontrado diretamente, procura recursivamente em sub-classificações.
            foreach (var subClassification in currentElement.Elements("classification"))
            {
                // Chama a função recursivamente para cada sub-classificação.
                var recursiveResult = BuscarMaterialRecursivo(subClassification, materialToFind);

                // Se o material foi encontrado em uma das sub-classificações, propaga o resultado.
                if (recursiveResult.material != null)
                {
                    return recursiveResult; // Retorna a categoria e o material encontrados na sub-chamada.
                }
            }

            // Se o material não foi encontrado no elemento atual nem em suas sub-classificações, retorna nulo.
            return (null, null);
        }

        public static void Proth1()
        {
                      
        }
    }
}