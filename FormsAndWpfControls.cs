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
            "Material", "Description", "Codigo", "Tipo", "Origem", "Linha" // Adicionado "Linha" aqui se for fixo para peças
        };

        // Lista de opções para o campo "Linha" - DEVE SER A MESMA DO DynamicWpfTaskPaneControl
        

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
            panel = new StackPanel { Margin = new Thickness(16, 12, 16, 12) };
            this.Content = panel;

            AtualizarCampos();
        }

        // =====================================================================
        // MÉTODOS PÚBLICOS
        // =====================================================================

        /// <summary>
        /// Limpa o painel e recarrega todos os campos de propriedades personalizadas
        /// com base no documento SolidWorks atual.
        /// </summary>
        public void AtualizarCampos()
        {
            panel.Children.Clear();

            if (swModel == null)
            {
                panel.Children.Add(new TextBlock { Text = "Nenhum documento de Peça selecionado." });
                return;
            }

            swPropMgr = swModel.Extension.CustomPropertyManager[""];

            // --- Seção 1: Adiciona campos "fixos" (editáveis ou variáveis) ---
            foreach (string nomeVisivel in camposFixosPart)
            {
                string valorRes;
                string valorBruto = ObterValorPropriedade(nomeVisivel, out valorRes) ?? "";

                // Verifica se é uma variável do SolidWorks
                bool isVar = ContemVariavel(valorBruto) || (!string.IsNullOrEmpty(valorBruto) && valorBruto != valorRes);

                if (isVar)
                {
                    AdicionarComboBoxVariavel(panel, nomeVisivel, valorRes);
                }
                else if (nomeVisivel.Trim().ToLower() == "linha") // Trata "Linha" como ComboBox
                {
                    AdicionarComboBoxComOpcoes(panel, nomeVisivel,null, valorRes);
                }
                else
                {
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
                    caixa.TextChanged += CriarAtualizadorPropriedade(nomeVisivel);
                    panel.Children.Add(label);
                    panel.Children.Add(caixa);
                }
            }

            // --- Seção 2: Adiciona todas as outras propriedades que contêm variáveis ---
            object propNames = null;
            object propTypes = null;
            object propValues = null;

            if (swPropMgr != null)
            {
                swPropMgr.GetAll(ref propNames, ref propTypes, ref propValues);
            }

            string[] todasPropriedades = propNames as string[];

            List<string> camposJaTratados = new List<string>();
            foreach (string f in camposFixosPart)
                camposJaTratados.Add(f.Trim().ToLower());

            if (todasPropriedades != null)
            {
                foreach (string nomePropriedade in todasPropriedades)
                {
                    if (camposJaTratados.Contains(nomePropriedade.Trim().ToLower()))
                        continue;

                    string valorRes;
                    string valorBruto = ObterValorPropriedade(nomePropriedade, out valorRes) ?? "";

                    bool isVar = ContemVariavel(valorBruto) || (!string.IsNullOrEmpty(valorBruto) && valorBruto != valorRes);

                    if (isVar)
                    {
                        AdicionarComboBoxVariavel(panel, nomePropriedade, valorRes);
                    }
                    // Campos que não são "fixos" (já tratados), não são "Linha" (já tratado),
                    // e não são variáveis não são adicionados por esta lógica.
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
        private TextChangedEventHandler CriarAtualizadorPropriedade(string nomePropriedade)
        {
            return delegate (object sender, TextChangedEventArgs e)
            {
                TextBox caixa = sender as TextBox;
                if (caixa != null && swPropMgr != null)
                {
                    swPropMgr.Add3(nomePropriedade, (int)swCustomInfoType_e.swCustomInfoText, caixa.Text, (int)swCustomPropertyAddOption_e.swCustomPropertyReplaceValue);
                }
            };
        }

        /// <summary>
        /// Cria um manipulador para a mudança de seleção em uma ComboBox, atualizando a propriedade no SolidWorks.
        /// </summary>
        private SelectionChangedEventHandler CriarAtualizadorPropriedadeComboBox(string nomePropriedade)
        {
            return delegate (object sender, SelectionChangedEventArgs e)
            {
                ComboBox combo = sender as ComboBox;
                if (combo != null && swPropMgr != null)
                {
                    if (combo.SelectedItem != null)
                    {
                        swPropMgr.Add3(nomePropriedade, (int)swCustomInfoType_e.swCustomInfoText, combo.SelectedItem.ToString(), (int)swCustomPropertyAddOption_e.swCustomPropertyReplaceValue);
                    }
                    else
                    {
                        swPropMgr.Add3(nomePropriedade, (int)swCustomInfoType_e.swCustomInfoText, "", (int)swCustomPropertyAddOption_e.swCustomPropertyReplaceValue);
                    }
                }
            };
        }

        /// <summary>
        /// Obtém o valor bruto (expressão) e o valor resolvido (final) de uma propriedade personalizada.
        /// </summary>
        private string ObterValorPropriedade(string nome, out string valorResolvido)
        {
            string valOut = "";
            string valRes = "";
            if (swModel != null && swPropMgr == null)
            {
                swPropMgr = swModel.Extension.CustomPropertyManager[""];
            }

            if (swPropMgr != null && swPropMgr.Get4(nome, false, out valOut, out valRes))
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
        private bool ContemVariavel(string valor)
        {
            if (string.IsNullOrEmpty(valor)) return false;
            return valor.StartsWith("${") && valor.EndsWith("}");
        }

        /// <summary>
        /// Adiciona um <see cref="ComboBox"/> ao painel para exibir propriedades que são variáveis.
        /// Este ComboBox é somente leitura, mostrando apenas o valor resolvido da variável.
        /// </summary>
        private void AdicionarComboBoxVariavel(StackPanel destino, string nomePropriedade, string valorRes)
        {
            TextBlock label = new TextBlock { Text = nomePropriedade };
            ComboBox combo = new ComboBox
            {
                Margin = new Thickness(0, 0, 0, 13),
                Height = 20,
                FontSize = 13,
                IsReadOnly = true,
                IsEditable = false,
                ItemsSource = new List<string> { valorRes },
                SelectedIndex = 0,
                HorizontalAlignment = HorizontalAlignment.Stretch
            };
            destino.Children.Add(label);
            destino.Children.Add(combo);
        }

        /// <summary>
        /// Adiciona um ComboBox com opções predefinidas a um painel.
        /// </summary>
        private void AdicionarComboBoxComOpcoes(StackPanel destino, string nomePropriedade, List<string> opcoes, string valorAtual)
        {
            TextBlock label = new TextBlock { Text = nomePropriedade };
            ComboBox combo = new ComboBox
            {
                Margin = new Thickness(0, 0, 0, 13),
                Height = 20,
                FontSize = 13,
                ItemsSource = opcoes,
                HorizontalAlignment = HorizontalAlignment.Stretch
            };

            if (!string.IsNullOrEmpty(valorAtual))
            {
                combo.SelectedItem = opcoes.FirstOrDefault(o => o.Equals(valorAtual, StringComparison.OrdinalIgnoreCase));
            }

            combo.SelectionChanged += CriarAtualizadorPropriedadeComboBox(nomePropriedade);
            destino.Children.Add(label);
            destino.Children.Add(combo);
        }
    }
}