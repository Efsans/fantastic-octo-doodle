using System.Windows;
using System.Windows.Controls;

namespace FormsAndWpfControls
{
    public class AdicionarCampoWindow : Window
    {
        public string NomeCampo { get; private set; }
        public bool EhVariavel { get; private set; }

        private TextBox txtNome;
        private RadioButton rbTexto;
        private RadioButton rbVariavel;

        public AdicionarCampoWindow()
        {
            Title = "Adicionar Campo";
            Width = 300;
            Height = 180;
            WindowStartupLocation = WindowStartupLocation.CenterOwner;
            ResizeMode = ResizeMode.NoResize;

            var panel = new StackPanel { Margin = new Thickness(16) };

            panel.Children.Add(new TextBlock { Text = "Nome do campo:" });
            txtNome = new TextBox { Margin = new Thickness(0, 0, 0, 10) };
            panel.Children.Add(txtNome);

            rbTexto = new RadioButton { Content = "Campo de texto", IsChecked = true, Margin = new Thickness(0, 0, 0, 4) };
            rbVariavel = new RadioButton { Content = "Variável (expressão SolidWorks)" };
            panel.Children.Add(rbTexto);
            panel.Children.Add(rbVariavel);

            var btnPanel = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right, Margin = new Thickness(0, 12, 0, 0) };
            var btnOk = new Button { Content = "OK", Width = 70, Margin = new Thickness(0, 0, 8, 0) };
            var btnCancel = new Button { Content = "Cancelar", Width = 70 };
            btnOk.Click += BtnOk_Click;
            btnCancel.Click += (s, e) => DialogResult = false;
            btnPanel.Children.Add(btnOk);
            btnPanel.Children.Add(btnCancel);

            panel.Children.Add(btnPanel);

            Content = panel;
        }

        private void BtnOk_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtNome.Text))
            {
                MessageBox.Show("Informe o nome do campo.", "Atenção", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            NomeCampo = txtNome.Text.Trim();
            EhVariavel = rbVariavel.IsChecked == true;
            DialogResult = true;
        }
    }
}