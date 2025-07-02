// Importa namespaces necessários para integração com o SolidWorks e para UI
using System;
using System.Runtime.InteropServices;
using Xarial.XCad.SolidWorks;
using Xarial.XCad.SolidWorks.UI.PropertyPage;
using Xarial.XCad.UI.Commands;
using Xarial.XCad.UI.Commands.Attributes;
using Xarial.XCad.UI.Commands.Enums;
using Xarial.XCad.UI.PropertyPage.Attributes;
using FormsAndWpfControls;
using Xarial.XCad.Base.Attributes;

// Torna a classe visível para COM e atribui um GUID único para identificação
[ComVisible(true)]
[Guid("1A185498-0FD9-4190-A326-E19940BA0595")]
public class ControlsCsAddIn : SwAddInEx // Herda de SwAddInEx, base para Add-Ins no SOLIDWORKS
{
    // Enumera os comandos que o Add-In irá suportar, cada item representa uma ação no UI
    enum ControlCommands_e
    {
        // Os atributos CommandItemInfo definem onde cada comando estará disponível (Part, Assembly, AllDocuments)
        [CommandItemInfo(WorkspaceTypes_e.Part | WorkspaceTypes_e.Assembly)]
        CreateWinFormModelViewTab,
        [CommandItemInfo(WorkspaceTypes_e.Part | WorkspaceTypes_e.Assembly)]
        CreateWpfModelViewTab,
        [CommandItemInfo(WorkspaceTypes_e.AllDocuments)]
        CreateWinFormFeatMgrTab,
        [CommandItemInfo(WorkspaceTypes_e.AllDocuments)]
        CreateWpfFeatMgrTab,
        CreateWinFormTaskPane,
        SW_cod,
        [CommandItemInfo(WorkspaceTypes_e.AllDocuments)]
        CreateWinFormPmPage,
        [CommandItemInfo(WorkspaceTypes_e.AllDocuments)]
        CreateWpfPmPage
    }

    // Classe para gerenciar uma página do PropertyManager com um controle WinForms
    [ComVisible(true)]
    public class WinFormsPMPage : SwPropertyManagerPageHandler
    {
        // Adiciona o controle WinFormsUserControl como um CustomControl na página
        [CustomControl(typeof(WinFormsUserControl))]
        public object WinFormCtrl { get; set; }
    }

    // Classe para gerenciar uma página do PropertyManager com um controle WPF
    [ComVisible(true)]
    public class WpfPMPage : SwPropertyManagerPageHandler
    {
        // Adiciona o controle WpfUserControl como um CustomControl na página
        [CustomControl(typeof(WpfUserControl))]
        public object WpfCtrl { get; set; }
    }

    // Instâncias das páginas do PropertyManager para WinForms e WPF
    private ISwPropertyManagerPage<WinFormsPMPage> m_WinFormsPMPage;
    private ISwPropertyManagerPage<WpfPMPage> m_WpfPMPage;

    // Método chamado quando o Add-In é carregado no SOLIDWORKS
    public override void OnConnect()
    {
        // Cria um grupo de comandos baseado no enum, e associa o evento de clique de botão
        CommandManager.AddCommandGroup<ControlCommands_e>().CommandClick += OnButtonClick;

        // Cria instâncias das páginas do PropertyManager para uso posterior
        m_WinFormsPMPage = CreatePage<WinFormsPMPage>();
        m_WpfPMPage = CreatePage<WpfPMPage>();

        this.CreateTaskPaneWpf<DynamicWpfTaskPaneControl>();
    }

    // Evento disparado quando um comando é clicado na UI
    
    private void OnButtonClick(ControlCommands_e cmd)
    {
        // Obtém o documento ativo no SOLIDWORKS
        var activeDoc = Application.Documents.Active;

        // Executa a ação correspondente ao comando selecionado
        switch (cmd)
        {
            case ControlCommands_e.CreateWinFormModelViewTab:
                {
                    // Cria uma nova aba no ModelView usando um controle WinForms
                    this.CreateDocumentTabWinForm<WinFormsUserControl>(activeDoc);
                    break;
                }

            case ControlCommands_e.CreateWpfModelViewTab:
                {
                    // Cria uma nova aba no ModelView usando um controle WPF
                    this.CreateDocumentTabWpf<WpfUserControl>(activeDoc);
                    break;
                }

            case ControlCommands_e.CreateWinFormFeatMgrTab:
                {
                    // Cria uma nova aba no FeatureManager usando um controle WinForms
                    this.CreateFeatureManagerTabWinForm<WinFormsUserControl>(activeDoc);
                    break;
                }

            case ControlCommands_e.CreateWpfFeatMgrTab:
                {
                    // Cria uma nova aba no FeatureManager usando um controle WPF
                    this.CreateFeatureManagerTabWpf<WpfUserControl>(activeDoc);
                    break;
                }

            case ControlCommands_e.CreateWinFormTaskPane:
                {
                    // Cria um novo painel de tarefas (TaskPane) usando um controle WinForms
                    this.CreateTaskPaneWinForm<WinFormsUserControl>();
                    break;
                }
                //#################################
            case ControlCommands_e.SW_cod:
                {
                    // Cria um novo painel de tarefas (TaskPane) usando um controle WPF
                    this.CreateTaskPaneWpf<DynamicWpfTaskPaneControl>();
                    
                    break;
                }
                //#################################
            case ControlCommands_e.CreateWinFormPmPage:
                {
                    // Exibe a página do PropertyManager com o controle WinForms
                    m_WinFormsPMPage.Show(new WinFormsPMPage());
                    break;
                }

            case ControlCommands_e.CreateWpfPmPage:
                {
                    // Exibe a página do PropertyManager com o controle WPF
                    m_WpfPMPage.Show(new WpfPMPage());
                    break;
                }


        }
        
    }
}
