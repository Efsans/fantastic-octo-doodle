// GreenTaskPaneControl.cs
// UserControl simples para exibir um painel verde no Task Pane do SolidWorks

using System.Windows.Forms;
using System.Drawing;

public class GreenTaskPaneControl : UserControl
{
    public GreenTaskPaneControl()
    {
        this.BackColor = Color.Green; // Cor de fundo verde
        this.Dock = DockStyle.Fill;   // Preenche todo o painel do Task Pane
    }
}