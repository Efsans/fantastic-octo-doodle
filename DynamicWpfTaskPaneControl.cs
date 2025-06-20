using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace FormsAndWpfControls
{
    /// <summary>
    /// Controle WPF para o Painel de Tarefas do SolidWorks, adaptando a UI
    /// com base no tipo de documento ativo (Part, Assembly).
    /// </summary>
    public class DynamicWpfTaskPaneControl : UserControl
    {
        // Controles de UI para o layout do painel de tarefas
        private TabControl tabControl;
        private StackPanel tabFixosPanel;
        private StackPanel tabAdicionaisPanel;
        private Dictionary<string, TextBox> caixasTexto; // Mantido para campos de texto, mas não para Linha

        // Objetos da API do SolidWorks para interação
        private SldWorks swApp;
        private IModelDoc2 swModel;
        private CustomPropertyManager swPropMgr;

        // Controle WPF específico para exibir propriedades de documentos do tipo 'Part'
        private PartSpecificPropertiesControl partPropertiesControl;

        private Button btnMais;

        // Definição dos campos de propriedades comuns e específicos por tipo de documento
        private readonly string[] camposComuns =
        {
            "Codigo", "Description", "T.MATERIAL", "Linha", "Projetista", "Desenhista", "Material"
        };
        private readonly string[] camposAssembly =
        {
            "Codigo","Description", "T.MATERIAL", "Linha", "Projetista", "Desenhista",
        };

        // Lista de opções para o campo "Linha"
        

        /// <summary>
        /// Construtor: Inicializa a UI, os botões de ação e tenta conectar ao SolidWorks.
        /// </summary>
        public DynamicWpfTaskPaneControl()
        {
            caixasTexto = new Dictionary<string, TextBox>();
            var painelPrincipal = new DockPanel();

            // Configuração dos botões de ação no topo
            var painelBotoesFixos = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Left,
                Margin = new Thickness(0, 2, 8, 0)
            };

            Button btn1 = new Button { Content = "1", Width = 35, Height = 35, Margin = new Thickness(2), ToolTip = "Auto preencher código" };
            btn1.Click += Btn1_Click;
            painelBotoesFixos.Children.Add(btn1);

            Button btn2 = new Button { Content = "2", Width = 35, Height = 35, Margin = new Thickness(2), ToolTip = "Usar API do zoho" };
            btn2.Click += Btn2_Click;
            painelBotoesFixos.Children.Add(btn2);

            Button btn3 = new Button { Content = "3", Width = 35, Height = 35, Margin = new Thickness(2), ToolTip = "Auto preencher produto pelo sistema" };
            btn3.Click += Btn3_Click;
            painelBotoesFixos.Children.Add(btn3);

            Button btn4 = new Button { Content = "4", Width = 35, Height = 35, Margin = new Thickness(2), ToolTip = "lista de propriedades material" };
            btn4.Click += Btn4_Click;
            painelBotoesFixos.Children.Add(btn4);

            DockPanel.SetDock(painelBotoesFixos, Dock.Top);
            painelPrincipal.Children.Add(painelBotoesFixos);

            // Configuração do TabControl com abas verticais à esquerda
            tabControl = new TabControl
            {
                Margin = new Thickness(0),
                Background = null,
                BorderThickness = new Thickness(0),
                HorizontalAlignment = HorizontalAlignment.Stretch,
                VerticalAlignment = VerticalAlignment.Stretch,
                TabStripPlacement = Dock.Left
            };

            // Aba de campos "fixos"
            tabFixosPanel = new StackPanel { Margin = new Thickness(16, 12, 16, 12) };
            var tabFixos = new TabItem
            {
                Header = "🔧",
                Width = 35,
                Height = 35,
                Content = new ScrollViewer { VerticalScrollBarVisibility = ScrollBarVisibility.Auto, HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled, Background = null, BorderThickness = new Thickness(0), Content = tabFixosPanel }
            };
            tabControl.Items.Add(tabFixos);

            // Aba de campos "adicionais"
            tabAdicionaisPanel = new StackPanel { Margin = new Thickness(16, 12, 16, 12) };
            var tabAdicionais = new TabItem
            {
                Header = "+",
                Width = 35,
                Height = 35,
                Content = new ScrollViewer { VerticalScrollBarVisibility = ScrollBarVisibility.Auto, HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled, Background = null, BorderThickness = new Thickness(0), Content = tabAdicionaisPanel }
            };
            tabControl.Items.Add(tabAdicionais);

            // Adiciona botão "Mais" à aba de campos adicionais
            btnMais = new Button
            {
                Content = "+",
                Width = 60,
                Height = 28,
                FontSize = 20,
                Margin = new Thickness(0, 0, 0, 10),
                HorizontalAlignment = HorizontalAlignment.Left
            };
            btnMais.Click += BtnMais_Click;
            tabAdicionaisPanel.Children.Add(btnMais);

            painelPrincipal.Children.Add(tabControl);
            this.Content = painelPrincipal;

            // Tenta obter a instância ativa do SolidWorks e registra o evento de mudança de documento
            try
            {
                swApp = System.Runtime.InteropServices.Marshal.GetActiveObject("SldWorks.Application") as SldWorks;
                if (swApp == null)
                {
                    tabFixosPanel.Children.Add(new TextBlock { Text = "SolidWorks não encontrado." });
                    return;
                }
                swApp.ActiveDocChangeNotify += OnActiveDocChangeNotify;
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                tabFixosPanel.Children.Add(new TextBlock { Text = "SolidWorks não está em execução ou não pode ser acessado." });
                return;
            }

            AtualizarPainel(); // Atualização inicial do painel
        }

        /// <summary>
        /// Manipulador para o evento de mudança de documento ativo no SolidWorks.
        /// </summary>
        public int OnActiveDocChangeNotify()
        {
            Dispatcher.Invoke(AtualizarPainel); // Garante que a atualização da UI ocorra na thread correta.
            return 0;
        }

        /// <summary>
        /// Atualiza o conteúdo do painel com base no tipo de documento ativo do SolidWorks.
        /// </summary>
        private void AtualizarPainel()
        {
            swModel = swApp.ActiveDoc as IModelDoc2;

            // Limpa os painéis se nenhum documento estiver aberto
            if (swModel == null)
            {
                tabFixosPanel.Children.Clear();
                tabAdicionaisPanel.Children.Clear();
                // Remove o btnMais do tabAdicionaisPanel se não houver documento ativo
                if (tabAdicionaisPanel.Children.Contains(btnMais))
                {
                    tabAdicionaisPanel.Children.Remove(btnMais);
                }
                if (partPropertiesControl != null) // Remove o controle de Part se ele existir
                {
                    tabFixosPanel.Children.Remove(partPropertiesControl);
                    partPropertiesControl = null;
                }
                tabFixosPanel.Children.Add(new TextBlock { Text = "Nenhum documento aberto no SolidWorks." });
                return;
            }

            // Limpa painéis e dados para reconstrução da UI
            tabFixosPanel.Children.Clear();
            tabAdicionaisPanel.Children.Clear();
            tabAdicionaisPanel.Children.Add(btnMais); // Adiciona o botão 'Mais' de volta
            caixasTexto.Clear();

            // Garante que o controle de Part seja removido antes de redesenhar
            if (partPropertiesControl != null)
            {
                tabFixosPanel.Children.Remove(partPropertiesControl);
                partPropertiesControl = null;
            }

            int docType = swModel.GetType();

            // Lógica para exibir campos específicos para Part ou Assembly
            if (docType == (int)swDocumentTypes_e.swDocPART) // Se for uma Peça
            {
                // Carrega o controle específico para peças (Material e variáveis)
                partPropertiesControl = new PartSpecificPropertiesControl(swModel);
                tabFixosPanel.Children.Add(partPropertiesControl);
                // Preenche campos adicionais, excluindo "Material" se já tratado
                // Nota: se "Linha" estiver em camposComuns, ela será tratada aqui também.
                CriarCamposAdicionais(camposComuns.Append("Material").ToArray());
            }
            else if (docType == (int)swDocumentTypes_e.swDocASSEMBLY) // Se for uma Montagem
            {
                CriarCamposFixos(camposAssembly); // Campos específicos para Assembly
                CriarCamposAdicionais(camposAssembly); // Campos adicionais, excluindo os fixos de Assembly
            }
            else // Outros tipos de documento
            {
                tabFixosPanel.Children.Add(new TextBlock { Text = "Documento atual não é uma Peça nem uma Montagem." });
                CriarCamposAdicionais(new string[0]); // Sem campos adicionais para outros tipos
            }
        }

        // ====================================================================
        // MANIPULADORES DE EVENTOS DOS BOTÕES
        // ====================================================================

        private void Btn1_Click(object sender, RoutedEventArgs e)
        {
            FuncoesExternas.Acao1();
            AtualizarPainelComDelay();
        }

        private async void Btn2_Click(object sender, RoutedEventArgs e)
        {
            await System.Threading.Tasks.Task.Run(() => FuncoesExternas.Acao2());
            AtualizarPainelComDelay();
        }

        private void Btn3_Click(object sender, RoutedEventArgs e)
        {
            FuncoesExternas.Acao3();
            AtualizarPainelComDelay();
        }

        private void Btn4_Click(object sender, RoutedEventArgs e)
        {
            FuncoesExternas.Acao4();
            AtualizarPainelComDelay();
        }

        private void BtnEnviar_Click(object sender, RoutedEventArgs e)
        {
            // Lógica para o botão Enviar (a ser implementada)
        }

        private void BtnMais_Click(object sender, RoutedEventArgs e)
        {
            var win = new AdicionarCampoWindow { Owner = Window.GetWindow(this) };
            if (win.ShowDialog() == true)
            {
                string campoNormalizado = win.NomeCampo.Trim().ToLower();

                // Adiciona o campo conforme a escolha do usuário
                if (win.EhVariavel)
                {
                    AdicionarComboBoxVariavel(tabAdicionaisPanel, win.NomeCampo, "");
                }
                else if (campoNormalizado == "linha") // Campo "Linha" como ComboBox
                {
                    AdicionarComboBoxComOpcoes(tabAdicionaisPanel, win.NomeCampo, null, "");
                }
                else // Outros campos como TextBox
                {
                    TextBlock label = new TextBlock { Text = win.NomeCampo };
                    TextBox caixa = new TextBox
                    {
                        Margin = new Thickness(0, 0, 0, 13),
                        Height = 20,
                        FontSize = 13,
                        HorizontalAlignment = HorizontalAlignment.Stretch
                    };
                    caixa.IsReadOnly = false;
                    caixa.TextChanged += CriarAtualizadorPropriedade(win.NomeCampo);
                    tabAdicionaisPanel.Children.Add(label);
                    tabAdicionaisPanel.Children.Add(caixa);
                }
            }
        }

        // ====================================================================
        // MÉTODOS AUXILIARES DE UI E DADOS
        // ====================================================================

        /// <summary>
        /// Atualiza o painel após um pequeno atraso para permitir que o SolidWorks processe as ações.
        /// </summary>
        private void AtualizarPainelComDelay()
        {
            Dispatcher.Invoke(() =>
            {
                DispatcherTimer timer = new DispatcherTimer();
                timer.Interval = TimeSpan.FromMilliseconds(350);
                timer.Tick += (s, e) =>
                {
                    timer.Stop();
                    AtualizarPainel();
                };
                timer.Start();
            });
        }

        /// <summary>
        /// Cria e popula os controles para campos de propriedades "fixos".
        /// </summary>
        private void CriarCamposFixos(string[] camposParaExibir)
        {
            if (swModel == null) return;
            swPropMgr = swModel.Extension.CustomPropertyManager[""]; // Obtém o PropertyManager

            foreach (string nomeVisivel in camposParaExibir)
            {
                string nomeChave = nomeVisivel.Replace(".", "").Replace(" ", "").ToLower();
                string valorRes;
                string valorBruto = ObterValorPropriedade(nomeVisivel, out valorRes) ?? "";

                bool isVar = ContemVariavel(valorBruto) || (!string.IsNullOrEmpty(valorBruto) && valorBruto != valorRes);

                if (isVar)
                {
                    AdicionarComboBoxVariavel(tabFixosPanel, nomeVisivel, valorRes);
                }
                else if (nomeChave == "linha") // Trata "Linha" como ComboBox
                {
                    AdicionarComboBoxComOpcoes(tabFixosPanel, nomeVisivel, null, valorRes);
                }
                else // Outros campos como TextBox
                {
                    TextBlock label = new TextBlock { Text = nomeVisivel };
                    TextBox caixa = new TextBox { Margin = new Thickness(0, 0, 0, 13), Text = valorRes, Height = 20, FontSize = 13, HorizontalAlignment = HorizontalAlignment.Stretch };
                    caixa.IsReadOnly = false;
                    caixa.TextChanged += CriarAtualizadorPropriedade(nomeVisivel);
                    caixasTexto[nomeChave] = caixa;
                    tabFixosPanel.Children.Add(label);
                    tabFixosPanel.Children.Add(caixa);
                }
            }
        }

        /// <summary>
        /// Cria e popula os controles para campos de propriedades "adicionais",
        /// excluindo aqueles já exibidos como fixos.
        /// </summary>
        private void CriarCamposAdicionais(string[] camposExcluir)
        {
            if (swModel == null) return;
            swPropMgr = swModel.Extension.CustomPropertyManager[""]; // Obtém o PropertyManager

            object propNames = null, propTypes = null, propValues = null;
            if (swPropMgr != null)
            {
                swPropMgr.GetAll(ref propNames, ref propTypes, ref propValues); // Obtém todas as propriedades
            }

            string[] todasPropriedades = propNames as string[];
            List<string> camposExcluirNormalizados = camposExcluir.Select(f => f.Trim().ToLower()).ToList();

            if (todasPropriedades != null)
            {
                foreach (string nomePropriedade in todasPropriedades)
                {
                    if (camposExcluirNormalizados.Contains(nomePropriedade.Trim().ToLower()))
                        continue; // Pula propriedades já tratadas

                    string valorRes;
                    string valorBruto = ObterValorPropriedade(nomePropriedade, out valorRes) ?? "";

                    bool isVar = ContemVariavel(valorBruto) || (!string.IsNullOrEmpty(valorBruto) && valorBruto != valorRes);

                    if (isVar)
                    {
                        AdicionarComboBoxVariavel(tabAdicionaisPanel, nomePropriedade, valorRes);
                    }
                    else if (nomePropriedade.Trim().ToLower() == "linha") // Trata "Linha" como ComboBox nos adicionais
                    {
                        AdicionarComboBoxComOpcoes(tabAdicionaisPanel, nomePropriedade, null, valorRes);
                    }
                    else // Outros campos como TextBox
                    {
                        TextBlock label = new TextBlock { Text = nomePropriedade };
                        TextBox caixa = new TextBox { Margin = new Thickness(0, 0, 0, 13), Text = valorRes, Height = 20, FontSize = 13, HorizontalAlignment = HorizontalAlignment.Stretch };
                        caixa.IsReadOnly = false;
                        caixa.TextChanged += CriarAtualizadorPropriedade(nomePropriedade);
                        tabAdicionaisPanel.Children.Add(label);
                        tabAdicionaisPanel.Children.Add(caixa);
                    }
                }
            }
        }

        /// <summary>
        /// Cria um manipulador para a mudança de texto em uma TextBox, atualizando a propriedade no SolidWorks.
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
                    // Garante que haja um item selecionado antes de tentar acessar SelectedItem
                    if (combo.SelectedItem != null)
                    {
                        swPropMgr.Add3(nomePropriedade, (int)swCustomInfoType_e.swCustomInfoText, combo.SelectedItem.ToString(), (int)swCustomPropertyAddOption_e.swCustomPropertyReplaceValue);
                    }
                    else
                    {
                        // Se nada for selecionado, você pode optar por limpar a propriedade ou definir um valor padrão
                        swPropMgr.Add3(nomePropriedade, (int)swCustomInfoType_e.swCustomInfoText, "", (int)swCustomPropertyAddOption_e.swCustomPropertyReplaceValue);
                    }
                }
            };
        }

        /// <summary>
        /// Obtém o valor bruto e resolvido de uma propriedade personalizada.
        /// </summary>
        private string ObterValorPropriedade(string nomePropriedade, out string valorResolvido)
        {
            string valOut = "", valRes = "";
            if (swModel != null)
            {
                swPropMgr = swModel.Extension.CustomPropertyManager[""];
                // Adicionado verificação para swPropMgr ser não nulo antes de usar
                if (swPropMgr != null && swPropMgr.Get4(nomePropriedade, false, out valOut, out valRes))
                {
                    valorResolvido = valRes;
                    return valOut;
                }
            }
            valorResolvido = "";
            return null;
        }

        /// <summary>
        /// Verifica se a string de valor de uma propriedade contém uma variável do SolidWorks (e.g., "${D1@Sketch1}").
        /// </summary>
        private bool ContemVariavel(string valor)
        {
            if (string.IsNullOrEmpty(valor)) return false;
            return valor.StartsWith("${") && valor.EndsWith("}");
        }

        /// <summary>
        /// Adiciona um ComboBox somente leitura para exibir o valor resolvido de uma propriedade variável.
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

            // Tenta selecionar o valor atual da propriedade na ComboBox
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