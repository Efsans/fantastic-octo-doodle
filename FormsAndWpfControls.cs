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

                if (nomeVisivel == "Material") // Lógica específica para o campo "Material"
                {
                    // Obtém o nome do material (o valor resolvido da propriedade)
                    string materialName = valorRes;

                    // Cria um DockPanel para organizar o rótulo, o campo de texto do material e o botão na mesma linha.
                    var materialRowPanel = new DockPanel { Margin = new Thickness(0, 0, 0, 13) };

                    // Rótulo "Material" alinhado à esquerda.
                    var labelMaterial = new TextBlock
                    {
                        Margin = new Thickness(0, 0, 0, 13),
                        Text = nomeVisivel,
                        VerticalAlignment = VerticalAlignment.Center,
                        HorizontalAlignment = HorizontalAlignment.Left,
                        Width = 25 // Largura fixa para o rótulo para alinhamento consistente.
                    };
                    DockPanel.SetDock(labelMaterial, Dock.Left);
                    panel.Children.Add(labelMaterial);

                    // Botão que chama a função Acao0 para exibir propriedades do material, alinhado à direita.
                    var btnAcao0 = new Button
                    {
                        Content = "...", // Texto ou ícone para o botão (ex: "Ver", "Detalhes")
                        Width = 25,
                        Height = 20,
                        Margin = new Thickness(4, 0, 0, 0), // Pequena margem à esquerda.
                        ToolTip = "Ver propriedades do material" // Dica de ferramenta.
                    };
                    btnAcao0.Click += (s, e) => FuncoesExternas.Acao0(); // Associa o clique à função Acao0.
                    DockPanel.SetDock(btnAcao0, Dock.Right);
                    materialRowPanel.Children.Add(btnAcao0);

                    // TextBox para exibir o nome do material (somente leitura), preenchendo o espaço restante.
                    var txtMaterial = new TextBox
                    {
                        Margin = new Thickness(0, 0, 0, 13), // Ajuste a margem se necessário
                        Text = materialName,
                        Height = 20,
                        FontSize = 13,
                        IsReadOnly = true, // Torna o TextBox somente leitura.
                        HorizontalAlignment = HorizontalAlignment.Stretch,
                        VerticalAlignment = VerticalAlignment.Center
                    };
                    materialRowPanel.Children.Add(txtMaterial); // Adiciona por último para que preencha o espaço disponível.

                    panel.Children.Add(materialRowPanel); // Adiciona o DockPanel ao painel principal.
                }
                else if (isVar)
                {
                    AdicionarComboBoxVariavel(panel, nomeVisivel, valorRes);
                }
                else if (nomeVisivel.Trim().ToLower() == "linha") // Trata "Linha" como ComboBox
                {
                    // Aqui você pode preencher 'opcoesLinha' com os valores desejados para o ComboBox "Linha".
                    // Exemplo: List<string> opcoesLinha = new List<string> { "Linha A", "Linha B", "Linha C" };
                    // AdicionarComboBoxComOpcoes(panel, nomeVisivel, opcoesLinha, valorRes);
                    AdicionarComboBoxComOpcoes(panel, nomeVisivel, null, valorRes); // Mantido null conforme o código original, caso as opções venham de outra fonte.
                }
                else // Para outros campos "fixos", cria um TextBox editável.
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

            // --- Seção 2: Adiciona todas as outras propriedades que contêm variáveis (e que não são "fixas") ---
            object propNames = null;
            object propTypes = null;
            object propValues = null;

            if (swPropMgr != null)
            {
                swPropMgr.GetAll(ref propNames, ref propTypes, ref propValues); // Obtém todas as propriedades personalizadas do documento.
            }

            string[] todasPropriedades = propNames as string[];

            // Cria uma lista de campos já tratados (convertidos para minúsculas para comparação).
            List<string> camposJaTratados = new List<string>();
            foreach (string f in camposFixosPart)
                camposJaTratados.Add(f.Trim().ToLower());

            if (todasPropriedades != null)
            {
                foreach (string nomePropriedade in todasPropriedades)
                {
                    // Ignora propriedades que já foram adicionadas como campos "fixos".
                    if (camposJaTratados.Contains(nomePropriedade.Trim().ToLower()))
                        continue;

                    string valorRes;
                    string valorBruto = ObterValorPropriedade(nomePropriedade, out valorRes) ?? "";

                    // Verifica se a propriedade é uma variável do SolidWorks.
                    bool isVar = ContemVariavel(valorBruto) || (!string.IsNullOrEmpty(valorBruto) && valorBruto != valorRes);

                    if (isVar) // Se for uma variável, adiciona como um ComboBox somente leitura.
                    {
                        AdicionarComboBoxVariavel(panel, nomePropriedade, valorRes);
                    }
                    // Outras propriedades não variáveis e não "fixas" não são adicionadas por esta lógica,
                    // assumindo que não são relevantes para este controle específico de propriedades de Peças.
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
        /// <param name="nomePropriedade">O nome da propriedade SolidWorks a ser atualizada.</param>
        /// <returns>Um delegado <see cref="TextChangedEventHandler"/>.</returns>
        private TextChangedEventHandler CriarAtualizadorPropriedade(string nomePropriedade)
        {
            return delegate (object sender, TextChangedEventArgs e)
            {
                TextBox caixa = sender as TextBox;
                if (caixa != null && swPropMgr != null)
                {
                    // Atualiza ou adiciona a propriedade personalizada no SolidWorks.
                    swPropMgr.Add3(nomePropriedade, (int)swCustomInfoType_e.swCustomInfoText, caixa.Text, (int)swCustomPropertyAddOption_e.swCustomPropertyReplaceValue);
                }
            };
        }

        /// <summary>
        /// Cria um manipulador para a mudança de seleção em uma ComboBox, atualizando a propriedade no SolidWorks.
        /// </summary>
        /// <param name="nomePropriedade">O nome da propriedade SolidWorks a ser atualizada.</param>
        /// <returns>Um delegado <see cref="SelectionChangedEventHandler"/>.</returns>
        private SelectionChangedEventHandler CriarAtualizadorPropriedadeComboBox(string nomePropriedade)
        {
            return delegate (object sender, SelectionChangedEventArgs e)
            {
                ComboBox combo = sender as ComboBox;
                if (combo != null && swPropMgr != null)
                {
                    if (combo.SelectedItem != null)
                    {
                        // Atualiza a propriedade com o item selecionado do ComboBox.
                        swPropMgr.Add3(nomePropriedade, (int)swCustomInfoType_e.swCustomInfoText, combo.SelectedItem.ToString(), (int)swCustomPropertyAddOption_e.swCustomPropertyReplaceValue);
                    }
                    else
                    {
                        // Se nada for selecionado (e.g., seleção desfeita), a propriedade é definida como vazia.
                        swPropMgr.Add3(nomePropriedade, (int)swCustomInfoType_e.swCustomInfoText, "", (int)swCustomPropertyAddOption_e.swCustomPropertyReplaceValue);
                    }
                }
            };
        }

        /// <summary>
        /// Obtém o valor bruto (expressão) e o valor resolvido (final) de uma propriedade personalizada.
        /// </summary>
        /// <param name="nome">O nome da propriedade.</param>
        /// <param name="valorResolvido">Parâmetro de saída para o valor resolvido da propriedade.</param>
        /// <returns>O valor bruto da propriedade, ou null se não for encontrado.</returns>
        private string ObterValorPropriedade(string nome, out string valorResolvido)
        {
            string valOut = "";
            string valRes = "";
            if (swModel != null && swPropMgr == null)
            {
                swPropMgr = swModel.Extension.CustomPropertyManager[""]; // Garante que o gerenciador esteja inicializado.
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
        /// <param name="valor">A string a ser verificada.</param>
        /// <returns>True se a string contiver uma variável, caso contrário, false.</returns>
        private bool ContemVariavel(string valor)
        {
            if (string.IsNullOrEmpty(valor)) return false;
            return valor.StartsWith("${") && valor.EndsWith("}");
        }

        /// <summary>
        /// Adiciona um <see cref="ComboBox"/> ao painel para exibir propriedades que são variáveis.
        /// Este ComboBox é somente leitura, mostrando apenas o valor resolvido da variável.
        /// </summary>
        /// <param name="destino">O StackPanel onde o ComboBox será adicionado.</param>
        /// <param name="nomePropriedade">O nome da propriedade SolidWorks.</param>
        /// <param name="valorRes">O valor resolvido da propriedade (que será exibido no ComboBox).</param>
        private void AdicionarComboBoxVariavel(StackPanel destino, string nomePropriedade, string valorRes)
        {
            TextBlock label = new TextBlock { Text = nomePropriedade };
            ComboBox combo = new ComboBox
            {
                Margin = new Thickness(0, 0, 0, 13),
                Height = 20,
                FontSize = 13,
                IsReadOnly = true,      // ComboBox somente leitura.
                IsEditable = false,     // Não permite edição direta.
                ItemsSource = new List<string> { valorRes }, // Exibe o valor resolvido como a única opção.
                SelectedIndex = 0,      // Seleciona o primeiro (e único) item.
                HorizontalAlignment = HorizontalAlignment.Stretch
            };
            destino.Children.Add(label);
            destino.Children.Add(combo);
        }

        /// <summary>
        /// Adiciona um ComboBox com opções predefinidas a um painel.
        /// </summary>
        /// <param name="destino">O StackPanel onde o ComboBox será adicionado.</param>
        /// <param name="nomePropriedade">O nome da propriedade SolidWorks.</param>
        /// <param name="opcoes">Uma lista de strings para popular as opções do ComboBox. Pode ser null se as opções forem carregadas posteriormente.</param>
        /// <param name="valorAtual">O valor atual da propriedade, que será pré-selecionado no ComboBox se uma correspondência for encontrada nas opções.</param>
        private void AdicionarComboBoxComOpcoes(StackPanel destino, string nomePropriedade, List<string> opcoes, string valorAtual)
        {
            TextBlock label = new TextBlock { Text = nomePropriedade };
            ComboBox combo = new ComboBox
            {
                Margin = new Thickness(0, 0, 0, 13),
                Height = 20,
                FontSize = 13,
                ItemsSource = opcoes, // As opções são definidas aqui.
                HorizontalAlignment = HorizontalAlignment.Stretch
            };

            if (!string.IsNullOrEmpty(valorAtual))
            {
                // Tenta selecionar o item que corresponde ao valor atual, ignorando maiúsculas/minúsculas.
                combo.SelectedItem = opcoes?.FirstOrDefault(o => o.Equals(valorAtual, StringComparison.OrdinalIgnoreCase));
            }

            combo.SelectionChanged += CriarAtualizadorPropriedadeComboBox(nomePropriedade); // Associa o evento de mudança de seleção.
            destino.Children.Add(label);
            destino.Children.Add(combo);
        }
    }
}