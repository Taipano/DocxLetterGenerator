using System.Windows;

namespace DocxLetterGenerator
{
    public partial class AttachmentWindow : Window
    {
        public Attachment Attachment { get; private set; }
        private bool isEditMode;

        public AttachmentWindow(Attachment attachment = null)
        {
            InitializeComponent();

            if (attachment != null)
            {

                isEditMode = true;
                Attachment = attachment;
                txtTitle.Text = Attachment.Title;
                txtContent.Text = Attachment.Content;
            }
            else
            {
  
                isEditMode = false;
                Attachment = new Attachment();
            }
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {

            if (string.IsNullOrWhiteSpace(txtTitle.Text))
            {
                MessageBox.Show("Введите заголовок приложения!", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (string.IsNullOrWhiteSpace(txtContent.Text))
            {
                if (MessageBox.Show("Текст приложения пуст. Продолжить?", "Предупреждение",
                    MessageBoxButton.YesNo, MessageBoxImage.Question) != MessageBoxResult.Yes)
                {
                    return;
                }
            }


            Attachment.Title = txtTitle.Text.Trim();
            Attachment.Content = txtContent.Text;

            DialogResult = true;
            Close();
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}