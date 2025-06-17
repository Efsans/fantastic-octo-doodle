using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace FormsAndWpfControls
{
    /// <summary>
    /// Controle WPF dedicado a exibir e gerenciar propriedades personalizadas
    /// de documentos do tipo 'Peça' (Part) no SolidWorks.
    /// </summary>
    public class PartSpecificPropertiesControl : UserControl
    {
        // =====================================================================
        // MEMBROS DE CLASSE
        // =====================================================================

        private StackPanel panel;               // Painel principal para organizar os controles de UI
        private IModelDoc2 swModel;             // Referência ao documento SolidWorks (Peça)
        private CustomPropertyManager swPropMgr; // Gerenciador de propriedades personalizadas do SolidWorks

        // Define os campos de propriedades personalizadas que são considerados "fixos"
        // para documentos de 'Peça'.
        private string[] camposFixosPart = new string[]
        {
            "Material", "Description", "Codigo" , "Tipo" , "Origem"
        };

        // =====================================================================
        // CONSTRUTOR
        // =====================================================================

        /// <summary>
        /// Inicializa uma nova instância de <see cref="PartSpecificPropertiesControl"/>.
        /// </summary>
        /// <param name="model">A instância do documento de Peça do SolidWorks.</param>
        public PartSpecificPropertiesControl(IModelDoc2 model)
        {
            swModel = model;
            panel = new StackPanel { Margin = new Thickness(0, 0, 0, 13) };
            this.Content = panel;

            // Carrega e exibe os campos de propriedades ao inicializar o controle.
            AtualizarCampos();
        }

        // =====================================================================
        // MÉTODOS PÚBLICOS
        // =====================================================================

        /// <summary>
        /// Limpa o painel e recarrega todos os campos de propriedades personalizadas
        /// com base no documento SolidWorks atual.
        /// Este método é a lógica central para a renderização da UI.
        /// </summary>
        public void AtualizarCampos()
        {
            panel.Children.Clear(); // Limpa controles existentes.

            if (swModel == null)
            {
                panel.Children.Add(new TextBlock { Text = "Nenhum documento de Peça selecionado." });
                return;
            }

            // Obtém o CustomPropertyManager para o documento ativo.
            // É importante re-obter para garantir que esteja sempre sincronizado.
            swPropMgr = swModel.Extension.CustomPropertyManager[""];

            // --- Seção 1: Adiciona campos "fixos" (editáveis ou variáveis) ---
            foreach (string nomeVisivel in camposFixosPart)
            {
                string valorRes;
                string valorBruto = ObterValorPropriedade(nomeVisivel, out valorRes) ?? "";

                TextBlock label = new TextBlock { Text = nomeVisivel };
                TextBox caixa = new TextBox
                {
                    Margin = new Thickness(0, 0, 0, 13),
                    Text = valorRes,
                    Height = 20,
                    FontSize = 13,
                    HorizontalAlignment = HorizontalAlignment.Stretch
                };
                caixa.IsReadOnly = false;
                // Associa um evento para atualizar a propriedade no SolidWorks quando o texto muda.
                caixa.TextChanged += CriarAtualizadorPropriedade(nomeVisivel);
                panel.Children.Add(label);
                panel.Children.Add(caixa);
            }

            // --- Seção 2: Adiciona todas as outras propriedades que contêm variáveis ---
            object propNames = null;
            object propTypes = null;
            object propValues = null;

            // Recupera todas as propriedades personalizadas do documento.
            swPropMgr.GetAll(ref propNames, ref propTypes, ref propValues);

            string[] todasPropriedades = propNames as string[];

            // Cria uma lista normalizada dos campos já tratados para evitar duplicação.
            List<string> camposJaTratados = new List<string>();
            foreach (string f in camposFixosPart)
                camposJaTratados.Add(f.Trim().ToLower());

            if (todasPropriedades != null)
            {
                foreach (string nomePropriedade in todasPropriedades)
                {
                    // Ignora propriedades que já foram exibidas como "fixas".
                    if (camposJaTratados.Contains(nomePropriedade.Trim().ToLower()))
                        continue;

                    string valorRes;
                    string valorBruto = ObterValorPropriedade(nomePropriedade, out valorRes) ?? "";

                    // Verifica se a propriedade é uma variável do SolidWorks.
                    bool isVar = ContemVariavel(valorBruto) || (!string.IsNullOrEmpty(valorBruto) && valorBruto != valorRes);

                    if (isVar)
                    {
                        // Se for variável, adiciona um ComboBox somente leitura.
                        AdicionarComboBoxVariavel(panel, nomePropriedade, valorRes);
                    }
                    // Campos que não são "Material" (já tratado) e não são variáveis não são adicionados por esta lógica.
                }
            }
        }

        // =====================================================================
        // MÉTODOS AUXILIARES
        // =====================================================================

        /// <summary>
        /// Cria um manipulador de evento <see cref="TextChangedEventHandler"/> que,
        /// ao ser acionado, atualiza a propriedade personalizada correspondente no SolidWorks.
        /// </summary>
        /// <param name="nomePropriedade">O nome da propriedade a ser atualizada.</param>
        private TextChangedEventHandler CriarAtualizadorPropriedade(string nomePropriedade)
        {
            return delegate (object sender, TextChangedEventArgs e)
            {
                TextBox caixa = sender as TextBox;
                if (caixa != null && swPropMgr != null)
                {
                    // Usa Add3 para adicionar ou substituir o valor da propriedade.
                    swPropMgr.Add3(nomePropriedade, (int)swCustomInfoType_e.swCustomInfoText, caixa.Text, (int)swCustomPropertyAddOption_e.swCustomPropertyReplaceValue);
                }
            };
        }

        /// <summary>
        /// Obtém o valor bruto (expressão) e o valor resolvido (final) de uma propriedade personalizada.
        /// </summary>
        /// <param name="nomePropriedade">O nome da propriedade.</param>
        /// <param name="valorResolvido">Saída: o valor da propriedade após a resolução de variáveis.</param>
        /// <returns>O valor bruto da propriedade ou null se não encontrada.</returns>
        private string ObterValorPropriedade(string nomePropriedade, out string valorResolvido)
        {
            string valOut = ""; // Valor da expressão (ex: "${D1@Sketch1}")
            string valRes = ""; // Valor resolvido (ex: "50mm")

            // Garante que o swPropMgr esteja disponível.
            if (swPropMgr == null && swModel != null)
            {
                swPropMgr = swModel.Extension.CustomPropertyManager[""];
            }

            if (swPropMgr != null && swPropMgr.Get4(nomePropriedade, false, out valOut, out valRes))
            {
                valorResolvido = valRes;
                return valOut;
            }
            valorResolvido = "";
            return null;
        }

        /// <summary>
        /// Verifica se uma string contém o padrão de uma variável do SolidWorks (e.g., "${nome_variavel}").
        /// </summary>
        /// <param name="valor">A string a ser verificada.</param>
        /// <returns>True se contiver um padrão de variável; caso contrário, False.</returns>
        private bool ContemVariavel(string valor)
        {
            if (string.IsNullOrEmpty(valor)) return false;
            return valor.StartsWith("${") && valor.EndsWith("}");
        }

        /// <summary>
        /// Adiciona um <see cref="ComboBox"/> ao painel para exibir propriedades que são variáveis.
        /// Este ComboBox é somente leitura, mostrando apenas o valor resolvido da variável.
        /// </summary>
        /// <param name="destino">O <see cref="StackPanel"/> onde o ComboBox será adicionado.</param>
        /// <param name="nomePropriedade">O nome da propriedade para o rótulo.</param>
        /// <param name="valorRes">O valor resolvido da variável a ser exibido.</param>
        private void AdicionarComboBoxVariavel(StackPanel destino, string nomePropriedade, string valorRes)
        {
            TextBlock label = new TextBlock { Text = nomePropriedade };
            ComboBox combo = new ComboBox
            {
                Margin = new Thickness(0, 0, 0, 13),
                Height = 20,
                FontSize = 13,
                IsReadOnly = true,    // Impede edição direta
                IsEditable = false,   // Impede que o usuário digite novos itens
                ItemsSource = new List<string> { valorRes }, // Exibe apenas o valor resolvido
                SelectedIndex = 0,    // Seleciona o primeiro (e único) item
                HorizontalAlignment = HorizontalAlignment.Stretch
            };
            destino.Children.Add(label);
            destino.Children.Add(combo);
        }
    }
}