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
    /// Este controle WPF atua como o Painel de Tarefas principal dentro do SolidWorks.
    /// Sua função é adaptar dinamicamente a interface do usuário (UI) e os campos exibidos
    /// com base no tipo de documento SolidWorks que está ativo no momento,
    /// como uma Peça (Part) ou uma Montagem (Assembly). Isso garante que o usuário
    /// veja apenas as propriedades e opções relevantes para o contexto atual de trabalho.
    /// </summary>
    public class DynamicWpfTaskPaneControl : UserControl
    {
        // ====================================================================
        // MEMBROS DE CLASSE (Variáveis e Controles)
        // ====================================================================

        // Controles de UI para o layout do painel de tarefas.
        // O `TabControl` gerencia as abas "fixos" e "adicionais".
        private TabControl tabControl;
        // O `tabFixosPanel` contém campos de propriedades padrão e frequentemente usados.
        private StackPanel tabFixosPanel;
        // O `tabAdicionaisPanel` é para propriedades extras ou adicionadas dinamicamente.
        private StackPanel tabAdicionaisPanel;
        // Dicionário para armazenar referências a `TextBox` criadas, 
        // permitindo acesso e manipulação fácil dos dados de entrada.
        private Dictionary<string, TextBox> caixasTexto;

        // Objetos da API do SolidWorks para interação direta com o software.
        // `swApp` representa a instância principal do aplicativo SolidWorks.
        private SldWorks swApp;
        // `swModel` representa o documento ativo (peça, montagem ou desenho).
        private IModelDoc2 swModel;
        // `swPropMgr` permite gerenciar as propriedades personalizadas do documento ativo.
        private CustomPropertyManager swPropMgr;

        // Referência a um controle WPF específico (`PartSpecificPropertiesControl`)
        // que é carregado apenas quando o documento ativo é uma Peça.
        private PartSpecificPropertiesControl partPropertiesControl;

        // Botão para adicionar novos campos de propriedade dinamicamente.
        private Button btnMais;

        // Definição dos nomes dos campos de propriedades que são considerados
        // "comuns" a todos os tipos de documentos, ou "específicos" para Montagens.
        private readonly string[] camposComuns =
        {
            "Codigo", "Description", "T.MATERIAL", "Linha", "Projetista", "Desenhista", "Material"
        };
        private readonly string[] camposAssembly =
        {
            "Codigo","Description", "T.MATERIAL", "Linha", "Projetista", "Desenhista",
        };

        // ====================================================================
        // CONSTRUTOR
        // ====================================================================

        /// <summary>
        /// Construtor da classe `DynamicWpfTaskPaneControl`.
        /// Responsável por inicializar todos os componentes da interface do usuário,
        /// configurar os botões de ação e tentar estabelecer uma conexão com a instância
        /// em execução do SolidWorks, registrando-se para receber notificações
        /// de mudança de documento ativo.
        /// </summary>
        public DynamicWpfTaskPaneControl()
        {
            caixasTexto = new Dictionary<string, TextBox>();
            var painelPrincipal = new DockPanel();

            // Configuração dos botões de ação na parte superior do painel de tarefas.
            // Cada botão é associado a uma função externa para automação de tarefas.
            var painelBotoesFixos = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Left,
                Margin = new Thickness(0, 2, 8, 0)
            };

            // Exemplo de configuração de um botão e seu evento de clique.
            // Os "ToolTips" fornecem dicas úteis ao usuário.
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

            Button btn5 = new Button { Content = "5", Width = 35, Height = 35, Margin = new Thickness(2), ToolTip = "lista de propriedades material" };
            btn5.Click += Btn5_Click;
            painelBotoesFixos.Children.Add(btn5);

            DockPanel.SetDock(painelBotoesFixos, Dock.Top);
            painelPrincipal.Children.Add(painelBotoesFixos);

            // Configuração do TabControl para organizar as abas na lateral esquerda.
            // Isso economiza espaço e oferece uma navegação intuitiva.
            tabControl = new TabControl
            {
                Margin = new Thickness(0),
                Background = null,
                BorderThickness = new Thickness(0),
                HorizontalAlignment = HorizontalAlignment.Stretch,
                VerticalAlignment = VerticalAlignment.Stretch,
                TabStripPlacement = Dock.Left // Abas dispostas verticalmente
            };

            // Cria a aba de campos "fixos", que contém as propriedades essenciais.
            tabFixosPanel = new StackPanel { Margin = new Thickness(16, 12, 16, 12) };
            var tabFixos = new TabItem
            {
                Header = "🔧", // Ícone para a aba
                Width = 35,
                Height = 35,
                Content = new ScrollViewer { VerticalScrollBarVisibility = ScrollBarVisibility.Auto, HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled, Background = null, BorderThickness = new Thickness(0), Content = tabFixosPanel }
            };
            tabControl.Items.Add(tabFixos);

            // Cria a aba de campos "adicionais", para propriedades que podem ser menos comuns
            // ou adicionadas pelo usuário.
            tabAdicionaisPanel = new StackPanel { Margin = new Thickness(16, 12, 16, 12) };
            var tabAdicionais = new TabItem
            {
                Header = "+", // Ícone para a aba
                Width = 35,
                Height = 35,
                Content = new ScrollViewer { VerticalScrollBarVisibility = ScrollBarVisibility.Auto, HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled, Background = null, BorderThickness = new Thickness(0), Content = tabAdicionaisPanel }
            };
            tabControl.Items.Add(tabAdicionais);

            // Adiciona o botão "Mais" à aba de campos adicionais, permitindo ao usuário
            // adicionar novas propriedades personalizadas.
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
            this.Content = painelPrincipal; // Define o conteúdo principal do controle WPF

            // Tenta obter a instância ativa do SolidWorks para interagir com ela.
            // Se o SolidWorks não estiver aberto ou não puder ser acessado, exibe uma mensagem de erro.
            try
            {
                swApp = System.Runtime.InteropServices.Marshal.GetActiveObject("SldWorks.Application") as SldWorks;
                if (swApp == null)
                {
                    tabFixosPanel.Children.Add(new TextBlock { Text = "SolidWorks não encontrado." });
                    return;
                }
                // Registra o evento para ser notificado quando o documento ativo do SolidWorks muda.
                // Isso é crucial para manter a UI do painel de tarefas sincronizada com o SolidWorks.
                swApp.ActiveDocChangeNotify += OnActiveDocChangeNotify;
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                tabFixosPanel.Children.Add(new TextBlock { Text = "SolidWorks não está em execução ou não pode ser acessado." });
                return;
            }

            AtualizarPainel(); // Chama a atualização inicial do painel ao carregar.
        }

        // ====================================================================
        // MANIPULADORES DE EVENTOS DO SOLIDWORKS
        // ====================================================================

        /// <summary>
        /// Este método é o manipulador para o evento `ActiveDocChangeNotify` do SolidWorks.
        /// Ele é invocado sempre que o documento ativo no SolidWorks é alterado.
        /// O `Dispatcher.Invoke` é usado para garantir que a atualização da UI (`AtualizarPainel`)
        /// ocorra na thread correta da UI, evitando erros de cross-threading.
        /// </summary>
        public int OnActiveDocChangeNotify()
        {
            Dispatcher.Invoke(AtualizarPainel);
            return 0;
        }

        // ====================================================================
        // LÓGICA PRINCIPAL DE ATUALIZAÇÃO DO PAINEL
        // ====================================================================

        /// <summary>
        /// Atualiza o conteúdo do painel de tarefas com base no documento SolidWorks ativo.
        /// Esta é uma função central que determina quais campos de propriedade são exibidos
        /// (e como) com base no tipo de documento (Peça, Montagem ou outro).
        /// </summary>
        private void AtualizarPainel()
        {
            // Obtém o documento ativo do SolidWorks.
            swModel = swApp.ActiveDoc as IModelDoc2;

            // Se nenhum documento estiver aberto no SolidWorks, limpa todos os painéis
            // e exibe uma mensagem indicando que não há documento ativo.
            if (swModel == null)
            {
                tabFixosPanel.Children.Clear();
                tabAdicionaisPanel.Children.Clear();
                // Remove o botão "+", pois não faz sentido adicionar campos sem um documento.
                if (tabAdicionaisPanel.Children.Contains(btnMais))
                {
                    tabAdicionaisPanel.Children.Remove(btnMais);
                }
                // Se o controle de propriedades de Peça estiver presente, ele é removido.
                if (partPropertiesControl != null)
                {
                    tabFixosPanel.Children.Remove(partPropertiesControl);
                    partPropertiesControl = null;
                }
                tabFixosPanel.Children.Add(new TextBlock { Text = "Nenhum documento aberto no SolidWorks." });
                return; // Sai da função, pois não há mais o que fazer.
            }

            // Limpa os painéis e o dicionário de caixas de texto para reconstruir a UI.
            tabFixosPanel.Children.Clear();
            tabAdicionaisPanel.Children.Clear();
            tabAdicionaisPanel.Children.Add(btnMais); // Adiciona o botão 'Mais' de volta para a aba de adicionais.
            caixasTexto.Clear();

            // Garante que o controle de propriedades de Peça seja removido antes de redesenhar
            // para evitar duplicatas ou comportamento incorreto ao mudar de tipo de documento.
            if (partPropertiesControl != null)
            {
                tabFixosPanel.Children.Remove(partPropertiesControl);
                partPropertiesControl = null;
            }

            int docType = swModel.GetType(); // Obtém o tipo de documento SolidWorks.

            // Lógica para exibir campos específicos para Peça ou Montagem.
            if (docType == (int)swDocumentTypes_e.swDocPART) // Se o documento é uma Peça
            {
                // Carrega e exibe um controle WPF específico para gerenciar as propriedades de Peças.
                partPropertiesControl = new PartSpecificPropertiesControl(swModel);
                tabFixosPanel.Children.Add(partPropertiesControl);
                // Preenche campos adicionais. O campo "Material" é tratado no controle específico de Peças.
                CriarCamposAdicionais(camposComuns.Append("Material").ToArray());
            }
            else if (docType == (int)swDocumentTypes_e.swDocASSEMBLY) // Se o documento é uma Montagem
            {
                // Cria e exibe os campos fixos e adicionais relevantes para Montagens.
                CriarCamposFixos(camposAssembly);
                CriarCamposAdicionais(camposAssembly);
            }
            else // Para outros tipos de documento (Ex: Desenho), exibe uma mensagem simples.
            {
                tabFixosPanel.Children.Add(new TextBlock { Text = "Documento atual não é uma Peça nem uma Montagem." });
                CriarCamposAdicionais(new string[0]); // Sem campos adicionais para outros tipos.
            }
        }

        // ====================================================================
        // MANIPULADORES DE EVENTOS DOS BOTÕES (UI)
        // ====================================================================

        // Os métodos abaixo são manipuladores de eventos `Click` para os botões da UI.
        // Eles chamam funções externas (`FuncoesExternas`) para realizar ações
        // específicas e, em seguida, forçam uma atualização do painel de tarefas
        // após um pequeno atraso para refletir quaisquer mudanças no modelo do SolidWorks.

        private void Btn1_Click(object sender, RoutedEventArgs e)
        {
            FuncoesExternas.Acao1();
            AtualizarPainelComDelay();
        }

        private async void Btn2_Click(object sender, RoutedEventArgs e)
        {
            // Executa a ação em uma thread separada para evitar congelar a UI.
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
            // Obtém o nome do material do modelo ativo e chama a ação externa.
            string materialName = ObterNomeMaterialAtivo();
            if (!string.IsNullOrEmpty(materialName))
            {
                MessageBox.Show($"Material ativo: {materialName}", "Material Ativo", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            else
            {
                MessageBox.Show("Nenhum material encontrado no modelo ativo.", "Erro", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
        private void Btn5_Click(object sender, RoutedEventArgs e)
        {
            FuncoesExternas.Acao4();
        }

        private void BtnEnviar_Click(object sender, RoutedEventArgs e)
        {
            // Lógica para o botão Enviar (a ser implementada futuramente).
        }

        /// <summary>
        /// Manipulador para o clique do botão "Mais". Abre uma nova janela
        /// (`AdicionarCampoWindow`) para permitir ao usuário definir e adicionar
        /// um novo campo de propriedade personalizada ao painel de tarefas.
        /// </summary>
        private void BtnMais_Click(object sender, RoutedEventArgs e)
        {
            var win = new AdicionarCampoWindow { Owner = Window.GetWindow(this) };
            if (win.ShowDialog() == true) // Se o usuário confirmar na janela
            {
                string campoNormalizado = win.NomeCampo.Trim().ToLower();

                // Adiciona o campo de propriedade com base no tipo escolhido pelo usuário.
                if (win.EhVariavel) // Se for uma variável (e.g., "${D1@Sketch1}")
                {
                    AdicionarComboBoxVariavel(tabAdicionaisPanel, win.NomeCampo, "");
                }
                else if (campoNormalizado == "linha") // Se for o campo "Linha", usa um ComboBox com opções predefinidas.
                {
                    AdicionarComboBoxComOpcoes(tabAdicionaisPanel, win.NomeCampo, null, "");
                }
                else // Para outros campos, cria um TextBox padrão.
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
                    // Associa um manipulador de evento para atualizar a propriedade no SolidWorks
                    // sempre que o texto na caixa for alterado.
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
        /// Este método introduz um pequeno atraso antes de chamar `AtualizarPainel()`.
        /// Isso é útil após certas operações do SolidWorks (como manipulação de API)
        /// para dar tempo ao SolidWorks para processar as mudanças antes que a UI seja atualizada,
        /// garantindo que os dados exibidos estejam corretos.
        /// </summary>
        private void AtualizarPainelComDelay()
        {
            Dispatcher.Invoke(() =>
            {
                DispatcherTimer timer = new DispatcherTimer();
                timer.Interval = TimeSpan.FromMilliseconds(350); // Atraso de 350 milissegundos
                timer.Tick += (s, e) =>
                {
                    timer.Stop();
                    AtualizarPainel();
                };
                timer.Start();
            });
        }

        /// <summary>
        /// Percorre uma lista de nomes de campos e cria os controles de UI (TextBox ou ComboBox)
        /// correspondentes no painel de campos "fixos". Ele também lida com a leitura
        /// dos valores existentes das propriedades personalizadas do SolidWorks.
        /// </summary>
        private void CriarCamposFixos(string[] camposParaExibir)
        {
            if (swModel == null) return;
            swPropMgr = swModel.Extension.CustomPropertyManager[""]; // Obtém o PropertyManager padrão

            foreach (string nomeVisivel in camposParaExibir)
            {
                string nomeChave = nomeVisivel.Replace(".", "").Replace(" ", "").ToLower();
                string valorRes;
                string valorBruto = ObterValorPropriedade(nomeVisivel, out valorRes) ?? "";

                // Verifica se a propriedade é uma variável do SolidWorks ou um valor já resolvido.
                bool isVar = ContemVariavel(valorBruto) || (!string.IsNullOrEmpty(valorBruto) && valorBruto != valorRes);

                if (isVar) // Se for uma variável, exibe em um ComboBox somente leitura.
                {
                    AdicionarComboBoxVariavel(tabFixosPanel, nomeVisivel, valorRes);
                }
                else if (nomeChave == "linha") // Se o campo for "Linha", usa um ComboBox.
                {
                    AdicionarComboBoxComOpcoes(tabFixosPanel, nomeVisivel, null, valorRes);
                }
                else // Para outros campos, cria um TextBox editável.
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
        /// Popula o painel de campos "adicionais" com propriedades existentes no documento
        /// do SolidWorks que ainda não foram exibidas nos campos "fixos". Isso garante
        /// que todas as propriedades personalizadas sejam visíveis e editáveis.
        /// </summary>
        private void CriarCamposAdicionais(string[] camposExcluir)
        {
            if (swModel == null) return;
            swPropMgr = swModel.Extension.CustomPropertyManager[""];

            object propNames = null, propTypes = null, propValues = null;
            if (swPropMgr != null)
            {
                swPropMgr.GetAll(ref propNames, ref propTypes, ref propValues); // Obtém todas as propriedades personalizadas.
            }

            string[] todasPropriedades = propNames as string[];
            // Cria uma lista normalizada (minúsculas, sem espaços/pontos) dos campos a serem excluídos,
            // para uma comparação eficiente.
            List<string> camposExcluirNormalizados = camposExcluir.Select(f => f.Trim().ToLower()).ToList();

            if (todasPropriedades != null)
            {
                foreach (string nomePropriedade in todasPropriedades)
                {
                    // Ignora as propriedades que já foram tratadas nos campos "fixos".
                    if (camposExcluirNormalizados.Contains(nomePropriedade.Trim().ToLower()))
                        continue;

                    string valorRes;
                    string valorBruto = ObterValorPropriedade(nomePropriedade, out valorRes) ?? "";

                    bool isVar = ContemVariavel(valorBruto) || (!string.IsNullOrEmpty(valorBruto) && valorBruto != valorRes);

                    if (isVar) // Se a propriedade é uma variável.
                    {
                        AdicionarComboBoxVariavel(tabAdicionaisPanel, nomePropriedade, valorRes);
                    }
                    else if (nomePropriedade.Trim().ToLower() == "linha") // Tratamento específico para "Linha".
                    {
                        AdicionarComboBoxComOpcoes(tabAdicionaisPanel, nomePropriedade, null, valorRes);
                    }
                    else // Outras propriedades como TextBox.
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
        /// Gera um manipulador de eventos `TextChanged` para um `TextBox`.
        /// Este manipulador é responsável por atualizar a propriedade personalizada
        /// correspondente no SolidWorks sempre que o texto na `TextBox` é modificado.
        /// </summary>
        private TextChangedEventHandler CriarAtualizadorPropriedade(string nomePropriedade)
        {
            return delegate (object sender, TextChangedEventArgs e)
            {
                TextBox caixa = sender as TextBox;
                if (caixa != null && swPropMgr != null)
                {
                    // Usa `Add3` para adicionar ou substituir o valor da propriedade.
                    swPropMgr.Add3(nomePropriedade, (int)swCustomInfoType_e.swCustomInfoText, caixa.Text, (int)swCustomPropertyAddOption_e.swCustomPropertyReplaceValue);
                }
            };
        }

        /// <summary>
        /// Gera um manipulador de eventos `SelectionChanged` para um `ComboBox`.
        /// Quando a seleção em um `ComboBox` muda, este manipulador atualiza
        /// a propriedade personalizada correspondente no SolidWorks.
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
                        // Atualiza a propriedade com o item selecionado do ComboBox.
                        swPropMgr.Add3(nomePropriedade, (int)swCustomInfoType_e.swCustomInfoText, combo.SelectedItem.ToString(), (int)swCustomPropertyAddOption_e.swCustomPropertyReplaceValue);
                    }
                    else
                    {
                        // Se nada for selecionado, a propriedade é limpa.
                        swPropMgr.Add3(nomePropriedade, (int)swCustomInfoType_e.swCustomInfoText, "", (int)swCustomPropertyAddOption_e.swCustomPropertyReplaceValue);
                    }
                }
            };
        }

        /// <summary>
        /// Obtém o valor bruto (como aparece na caixa de diálogo de propriedades)
        /// e o valor resolvido (o valor avaliado pelo SolidWorks) de uma propriedade personalizada.
        /// </summary>
        private string ObterValorPropriedade(string nomePropriedade, out string valorResolvido)
        {
            string valOut = "", valRes = "";
            if (swModel != null)
            {
                swPropMgr = swModel.Extension.CustomPropertyManager[""];
                // Usa `Get4` para obter o valor bruto (`valOut`) e o valor resolvido (`valRes`).
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
        /// Verifica se uma string de valor de propriedade representa uma variável do SolidWorks.
        /// Variáveis são expressões que o SolidWorks resolve dinamicamente, como dimensões
        /// de um esboço (ex: "${D1@Sketch1}").
        /// </summary>
        private bool ContemVariavel(string valor)
        {
            if (string.IsNullOrEmpty(valor)) return false;
            return valor.StartsWith("${") && valor.EndsWith("}");
        }

        /// <summary>
        /// Adiciona um `ComboBox` a um `StackPanel` que é usado para exibir
        /// o valor resolvido de uma propriedade que é uma variável do SolidWorks.
        /// Este `ComboBox` é somente leitura, pois o valor é derivado de uma expressão.
        /// </summary>
        private void AdicionarComboBoxVariavel(StackPanel destino, string nomePropriedade, string valorRes)
        {
            TextBlock label = new TextBlock { Text = nomePropriedade };
            ComboBox combo = new ComboBox
            {
                Margin = new Thickness(0, 0, 0, 13),
                Height = 20,
                FontSize = 13,
                IsReadOnly = true, // Torna o ComboBox somente leitura
                IsEditable = false,
                ItemsSource = new List<string> { valorRes }, // Exibe apenas o valor resolvido
                SelectedIndex = 0,
                HorizontalAlignment = HorizontalAlignment.Stretch
            };
            destino.Children.Add(label);
            destino.Children.Add(combo);
        }

        /// <summary>
        /// Adiciona um `ComboBox` a um `StackPanel` com uma lista predefinida de opções.
        /// Isso é usado para campos como "Linha", onde o usuário deve selecionar
        /// um valor de uma lista específica em vez de digitar um texto arbitrário.
        /// </summary>
        private void AdicionarComboBoxComOpcoes(StackPanel destino, string nomePropriedade, List<string> opcoes, string valorAtual)
        {
            TextBlock label = new TextBlock { Text = nomePropriedade };
            // A lista de opções (`opcoes`) deve ser preenchida aqui ou passada.
            // Para "Linha", as opções viriam de uma fonte de dados (e.g., banco de dados, arquivo).
            // Atualmente, está `null`, o que significa que o ComboBox aparecerá vazio até ser populado.
            ComboBox combo = new ComboBox
            {
                Margin = new Thickness(0, 0, 0, 13),
                Height = 20,
                FontSize = 13,
                ItemsSource = opcoes,
                HorizontalAlignment = HorizontalAlignment.Stretch
            };

            // Tenta pré-selecionar o valor atual da propriedade na ComboBox.
            if (!string.IsNullOrEmpty(valorAtual))
            {
                // Busca o item na lista de opções que corresponde ao valor atual (ignorando maiúsculas/minúsculas).
                combo.SelectedItem = opcoes.FirstOrDefault(o => o.Equals(valorAtual, StringComparison.OrdinalIgnoreCase));
            }

            combo.SelectionChanged += CriarAtualizadorPropriedadeComboBox(nomePropriedade);
            destino.Children.Add(label);
            destino.Children.Add(combo);
        }

        /// <summary>
        /// Obtém o nome do material atribuído ao modelo SolidWorks ativo.
        /// Retorna uma string vazia se nenhum material estiver atribuído ou se não houver modelo ativo.
        /// </summary>
        private string ObterNomeMaterialAtivo()
        {
            if (swModel != null)
            {
                return swModel.MaterialIdName ?? string.Empty; // Retorna o nome do material ou uma string vazia.
            }
            return string.Empty;
        }
    }
}