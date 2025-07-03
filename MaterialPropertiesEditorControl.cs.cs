using FormsAndWpfControls;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;
using System.IO; // Para Path.GetFileName
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input; // Para Command
using System.Xml.Linq; // Para XElement

namespace FormsAndWpfControls // Use o namespace correto do seu projeto
{
    /// <summary>
    /// Controle WPF para exibir e editar dinamicamente as propriedades customizadas de um material SolidWorks.
    /// </summary>
    public partial class MaterialPropertiesEditorControl : UserControl
    {
        private SldWorks swApp;
        private IModelDoc2 swModel;
        private StackPanel propertiesPanel; // Onde os TextBoxes/ComboBoxes dinâmicos serão adicionados
        private string currentMaterialName;
        private string currentSldmatFilePath; // O caminho do arquivo .sldmat onde o material foi encontrado
        private List<Tuple<TextBox, TextBox>> dynamicTextBoxes; // Para rastrear os campos de propriedade e valor

        // Lista de caminhos para os arquivos .sldmat (pode vir de FunçõesExternas.Acao4 ou ser configurável)
        private IEnumerable<string> sldmatFilePaths;

        public MaterialPropertiesEditorControl()
        {
            InitializeComponent(); // Chama o método gerado pelo WPF para carregar o XAML (se você usar XAML).

            // Se você não usa XAML, crie os elementos de UI manualmente aqui.
            // Para simplicidade e compatibilidade com sua estrutura atual, vou assumir uma construção manual.
            InitializeCustomComponents();

            // Defina os caminhos dos arquivos .sldmat.
            // Idealmente, isso deveria vir de uma configuração ou ser injetado.
            // Por enquanto, vamos replicar os caminhos da Acao4 para teste.
            sldmatFilePaths = new List<string>
            {
                @"C:\Program Files\SOLIDWORKS Corp\SOLIDWORKS\lang\portuguese-brazilian\sldmaterials\solidworks materials.sldmat",
                @"C:\Program Files\SOLIDWORKS Corp\SOLIDWORKS\lang\portuguese-brazilian\sldmaterials\sustainability extras.sldmat",
                @"C:\Users\Usuario\Desktop\meus materiais teste\EF material.sldmat",
                @"C:\Program Files\SOLIDWORKS Corp\SOLIDWORKS\lang\portuguese-brazilian\sldmaterials\SolidWorks DIN Materials.sldmat",
                @"\\toronto\AIZI\TEMPLATES AIZI\MATERIAIS CADASTRADOS AIZ.sldmat", // Verifique se este caminho existe e é acessível
                @"\\toronto\AIZ IMPLEMENTOS\TEMPLATES AIZI\Bancos de Dados de Material\Materiais personalizados.sldmat"
// Verifique
            };

            dynamicTextBoxes = new List<Tuple<TextBox, TextBox>>();

            // Evento para atualizar a UI quando o documento SolidWorks ativo muda ou é salvo.
            // Você precisará de um mecanismo no seu add-in principal para notificar este controle.
            // Por enquanto, teremos um botão "Atualizar".
        }

        private void InitializeCustomComponents()
        {
            this.Content = new ScrollViewer
            {
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                Content = new StackPanel
                {
                    Margin = new Thickness(2),
                    Width=300,
                    Children =
                    {
                        new TextBlock { Text = "Propriedades do Material", FontWeight = FontWeights.Bold, Margin = new Thickness(0,0,0,10) },
                        new TextBlock { Name = "MaterialNameTextBlock", Text = "Material: N/A", Margin = new Thickness(0,0,0,5) },
                        new TextBlock { Name = "MaterialFilePathTextBlock", Text = "Arquivo SLDMAT: N/A", Margin = new Thickness(0,0,0,15) },
                        new StackPanel { Name = "PropertiesDisplayPanel" }, // Este painel será preenchido dinamicamente
                        new StackPanel
                        {
                            Orientation = Orientation.Horizontal,
                            HorizontalAlignment = HorizontalAlignment.Center,
                            Width = 300,
                            Children =
                            {
                                new Button { Content = "Atualizar", Margin = new Thickness(5), Padding = new Thickness(10, 5, 10, 5) , Tag = "Refresh" },
                                new Button { Content = "Adicionar", Margin = new Thickness(5), Padding = new Thickness(10, 5, 10, 5), Tag = "Add" },
                                new Button { Content = "Salvar", Margin = new Thickness(5), Padding = new Thickness(10, 5, 10, 5), Tag = "Save" }
                            }
                        }
                    }
                }
            };

            propertiesPanel = (this.Content as ScrollViewer).Content as StackPanel;
            var buttonsPanel = propertiesPanel.Children.OfType<StackPanel>().Last();

            // Adiciona manipuladores de eventos para os botões
            (buttonsPanel.Children[0] as Button).Click += RefreshMaterialProperties_Click;
            (buttonsPanel.Children[1] as Button).Click += AddNewPropertyField_Click;
            (buttonsPanel.Children[2] as Button).Click += SaveChanges_Click;

            // Atribui os TextBlocks a campos para acesso posterior
            MaterialNameTextBlock = propertiesPanel.Children.OfType<TextBlock>().First(tb => tb.Name == "MaterialNameTextBlock");
            MaterialFilePathTextBlock = propertiesPanel.Children.OfType<TextBlock>().First(tb => tb.Name == "MaterialFilePathTextBlock");
            propertiesPanel = propertiesPanel.Children.OfType<StackPanel>().First(sp => sp.Name == "PropertiesDisplayPanel"); // Atualiza a referência
        }

        // Referências aos TextBlocks para atualização
        private TextBlock MaterialNameTextBlock;
        private TextBlock MaterialFilePathTextBlock;

        /// <summary>
        /// Método para ser chamado pelo SolidWorks Add-in principal quando o documento ativo muda.
        /// </summary>
        public void OnSolidWorksDocumentChanged()
        {
            // Tenta obter a instância ativa do SolidWorks.
            try
            {
                swApp = (SldWorks)System.Runtime.InteropServices.Marshal.GetActiveObject("SldWorks.Application");
                swModel = swApp?.IActiveDoc2 as IModelDoc2;

                if (swModel == null || swModel.GetType() != (int)swDocumentTypes_e.swDocPART)
                {
                    MaterialNameTextBlock.Text = "Material: N/A (Peça não aberta ou ativa)";
                    MaterialFilePathTextBlock.Text = "Arquivo SLDMAT: N/A";
                    propertiesPanel.Children.Clear(); // Limpa as propriedades
                    currentMaterialName = null;
                    currentSldmatFilePath = null;
                    dynamicTextBoxes.Clear();
                    return;
                }

                UpdateMaterialDisplay();
            }
            catch (Exception ex)
            {
                MaterialNameTextBlock.Text = "Material: Erro ao conectar ao SolidWorks";
                MaterialFilePathTextBlock.Text = "Arquivo SLDMAT: N/A";
                propertiesPanel.Children.Clear();
                currentMaterialName = null;
                currentSldmatFilePath = null;
                dynamicTextBoxes.Clear();
                Console.WriteLine($"Erro ao obter SolidWorks App: {ex.Message}");
            }
        }

        private void RefreshMaterialProperties_Click(object sender, RoutedEventArgs e)
        {
            OnSolidWorksDocumentChanged(); // Simplesmente re-carrega as propriedades
        }

        private void UpdateMaterialDisplay()
        {
            string fullMaterialName = swModel.MaterialIdName;
            if (string.IsNullOrEmpty(fullMaterialName))
            {
                MaterialNameTextBlock.Text = "Material: Nenhum material aplicado";
                MaterialFilePathTextBlock.Text = "Arquivo SLDMAT: N/A";
                propertiesPanel.Children.Clear();
                currentMaterialName = null;
                currentSldmatFilePath = null;
                dynamicTextBoxes.Clear();
                return;
            }

            currentMaterialName = fullMaterialName.Split('|')[1];
            MaterialNameTextBlock.Text = $"Material: {currentMaterialName}";

            // Tenta encontrar o material em um dos arquivos .sldmat
            XElement foundMaterialElement = null;
            XElement foundCategoryElement = null;
            string foundSldmatFilePath = null;

            foreach (string path in sldmatFilePaths)
            {
                if (File.Exists(path))
                {
                    try
                    {
                        XDocument doc = XDocument.Load(path);
                        var searchResult = FuncoesExternas.BuscarMaterialRecursivo(doc.Root, currentMaterialName);
                        if (searchResult.material != null)
                        {
                            foundMaterialElement = searchResult.material;
                            foundCategoryElement = searchResult.categoria;
                            foundSldmatFilePath = path;
                            break;
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Erro ao carregar ou analisar '{Path.GetFileName(path)}': {ex.Message}");
                    }
                }
            }

            if (foundMaterialElement == null)
            {
                MaterialFilePathTextBlock.Text = "Arquivo SLDMAT: Material não encontrado nos arquivos configurados.";
                propertiesPanel.Children.Clear();
                currentSldmatFilePath = null;
                dynamicTextBoxes.Clear();
                MessageBox.Show($"O material '{currentMaterialName}' não foi encontrado em nenhum dos arquivos .sldmat especificados. Você pode adicionar novas propriedades, mas elas não serão salvas em um arquivo existente.", "Aviso", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            currentSldmatFilePath = foundSldmatFilePath;
            MaterialFilePathTextBlock.Text = $"Arquivo SLDMAT: {Path.GetFileName(foundSldmatFilePath)}";

            // Limpa o painel e os campos antigos
            propertiesPanel.Children.Clear();
            dynamicTextBoxes.Clear();

            // Carrega as propriedades do XML (usando MaterialPropertyUpdater)
            var materialProps = MaterialPropertyUpdater.GetMaterialCustomProperties(currentSldmatFilePath, currentMaterialName);

            if (materialProps != null && materialProps.Any())
            {
                foreach (var propEntry in materialProps)
                {
                    string propName = propEntry.Key;
                    string propValue = propEntry.Value.ContainsKey("description") ? propEntry.Value["description"] : string.Empty;
                    AddPropertyField(propName, propValue);
                }
            }
            else
            {
                propertiesPanel.Children.Add(new TextBlock { Text = "Nenhuma propriedade customizada encontrada para este material.", Margin = new Thickness(0, 10, 0, 0) });
            }
        }

        private void AddPropertyField(string name = "", string value = "")
        {
            var stackPanel = new StackPanel { Orientation = Orientation.Vertical, Margin = new Thickness(0, 5, 0, 0) };

            var nameTextBox = new TextBox { Width = 120, Margin = new Thickness(0, 0, 5, 0), Text = name, ToolTip = "nome" };
            var valueTextBox = new TextBox { Width = 180, Text = value, ToolTip = "descriçao" };

            // Adiciona um botão para remover a linha
            var removeButton = new Button { Content = "X", Width = 25, Height = 25, Margin = new Thickness(5, 0, 0, 0), Tag = stackPanel, ToolTip = "X" };
            removeButton.Click += RemovePropertyField_Click;

            stackPanel.Children.Add(nameTextBox);
            stackPanel.Children.Add(valueTextBox);
            stackPanel.Children.Add(removeButton);

            propertiesPanel.Children.Add(stackPanel);
            dynamicTextBoxes.Add(Tuple.Create(nameTextBox, valueTextBox)); // Armazena a referência
        }

        private void AddNewPropertyField_Click(object sender, RoutedEventArgs e)
        {
            AddPropertyField(); // Adiciona uma nova linha vazia
        }

        private void RemovePropertyField_Click(object sender, RoutedEventArgs e)
        {
            Button removeButton = sender as Button;
            StackPanel parentPanel = removeButton.Tag as StackPanel;
            if (parentPanel != null)
            {
                // Remove do UI
                propertiesPanel.Children.Remove(parentPanel);

                // Remove da lista de controle dinâmico
                var nameTextBox = parentPanel.Children.OfType<TextBox>().FirstOrDefault();
                var valueTextBox = parentPanel.Children.OfType<TextBox>().LastOrDefault();
                if (nameTextBox != null && valueTextBox != null)
                {
                    var itemToRemove = dynamicTextBoxes.FirstOrDefault(t => t.Item1 == nameTextBox && t.Item2 == valueTextBox);
                    if (itemToRemove != null)
                    {
                        dynamicTextBoxes.Remove(itemToRemove);
                    }
                }
            }
        }

        private void SaveChanges_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(currentMaterialName) || string.IsNullOrEmpty(currentSldmatFilePath))
            {
                MessageBox.Show("Nenhum material ativo ou arquivo .sldmat para salvar. Por favor, selecione um material válido no SolidWorks e certifique-se de que ele existe em um dos arquivos .sldmat configurados.", "Erro ao Salvar", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            var propertiesToSave = new Dictionary<string, Dictionary<string, string>>();

            foreach (var tuple in dynamicTextBoxes)
            {
                string propName = tuple.Item1.Text.Trim();
                string propValue = tuple.Item2.Text.Trim();

                if (string.IsNullOrEmpty(propName))
                {
                    MessageBox.Show("Todos os nomes de propriedades devem ser preenchidos. Por favor, corrija antes de salvar.", "Erro de Validação", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                // Cria um dicionário de atributos para a propriedade
                // Aqui você pode adicionar outros atributos padrão se desejar, como "description", "units", etc.
                // Por simplicidade, estamos focando em 'name' e 'value' primariamente para a edição dinâmica.
                // Se o sldmat tiver outros atributos, GetMaterialCustomProperties já os retornou.
                // Para salvar, queremos garantir que pelo menos 'name' e 'value' estão presentes.
                var attributes = new Dictionary<string, string>
                {
                    { "name", propName },
                    { "description", propValue }
                };

                // Tenta preservar outros atributos se existirem na propriedade original
                // Isso requer buscar a propriedade original no XML para obter todos os seus atributos
                // e então sobrescrever apenas 'value' e 'name'.
                // Para este exemplo, vamos manter simples e sobrescrever.
                // Se precisar de persistência total de atributos, o MaterialPropertyUpdater.UpdateMaterialCustomProperties
                // precisaria de uma lógica mais avançada para mesclar atributos, ou passar a lista completa aqui.
                // Por agora, o MaterialPropertyUpdater já é capaz de adicionar/atualizar atributos existentes.

                propertiesToSave[propName] = attributes;
            }

            // Chama a função de atualização do MaterialPropertyUpdater
            bool success = MaterialPropertyUpdater.UpdateMaterialCustomProperties(currentSldmatFilePath, currentMaterialName, propertiesToSave);


            if (success)
            {
                // Após salvar, atualiza a exibição para refletir as alterações (se houver alguma otimização no arquivo)
                UpdateMaterialDisplay();
            }
        }

        /// <summary>
        /// Método de preenchimento do InitializeComponent para quem não usa XAML
        /// (apenas para compatibilidade com o formato do seu código)
        /// </summary>
        private void InitializeComponent()
        {
            // Este método normalmente é gerado automaticamente quando se usa XAML.
            // Como você está construindo a UI em código, InitializeCustomComponents()
            // faz o papel de configurar os controles.
        }
    }
}