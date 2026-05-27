using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Windows;

namespace DocxLetterGenerator
{
    public partial class MainWindow : Window
    {
        private string templatePath;
        private ObservableCollection<Attachment> attachments;

        public MainWindow()
        {
            InitializeComponent();
            templatePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Templates", "LetterTemplate.docx");
            attachments = new ObservableCollection<Attachment>();
            lvAttachments.ItemsSource = attachments;
        }

        
        private bool ValidateFields()
        {
            var errors = new List<string>();

           
            if (string.IsNullOrWhiteSpace(txtBankName.Text))
                errors.Add("• Название банка");

            if (string.IsNullOrWhiteSpace(txtSenderPosition.Text))
                errors.Add("• Должность отправителя");

            if (string.IsNullOrWhiteSpace(txtSenderName.Text))
                errors.Add("• ФИО отправителя");

           
            if (string.IsNullOrWhiteSpace(txtRecipientPost.Text))
                errors.Add("• Должность получателя");

            if (string.IsNullOrWhiteSpace(txtRecipientOrg.Text))
                errors.Add("• Организация получателя");

            if (string.IsNullOrWhiteSpace(txtRecipientName.Text))
                errors.Add("• ФИО получателя");

            if (string.IsNullOrWhiteSpace(txtLetterNumber.Text))
                errors.Add("• Номер письма");

            if (string.IsNullOrWhiteSpace(txtSubject.Text))
                errors.Add("• Тема письма");

            if (string.IsNullOrWhiteSpace(txtBody.Text))
                errors.Add("• Текст письма");

            if (errors.Any())
            {
                string message = "Пожалуйста, заполните следующие обязательные поля:\n\n" +
                                 string.Join("\n", errors) +
                                 "\n\nПоля, отмеченные (*), обязательны для заполнения.";

                MessageBox.Show(message, "Валидация данных",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return false;
            }

            return true;
        }

        private void GenerateLetter_Click(object sender, RoutedEventArgs e)
        {
            if (!ValidateFields())
                return;

            if (!File.Exists(templatePath))
            {
                MessageBox.Show("Шаблон LetterTemplate.docx не найден!\n\nПроверьте наличие файла в папке Templates.",
                    "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            try
            {
                string outputPath = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                    $"Письмо_{DateTime.Now:yyyyMMdd_HHmmss}.docx");

                File.Copy(templatePath, outputPath, true);
                ReplacePlaceholders(outputPath);

                MessageBox.Show($"✅ Письмо успешно создано!\n\n📁 Путь: {outputPath}",
                    "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"❌ Ошибка при создании письма:\n\n{ex.Message}",
                    "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ReplacePlaceholders(string filePath)
        {
            using (WordprocessingDocument doc = WordprocessingDocument.Open(filePath, true))
            {
                string dateValue = string.IsNullOrWhiteSpace(txtLetterDate.Text)
                    ? DateTime.Now.ToString("dd.MM.yyyy")
                    : txtLetterDate.Text.Trim();

                var replacements = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                {
                    { "{HEADER_BANK_NAME}", txtBankName.Text ?? "" },
                    { "{HEADER_BANK_FULL_NAME}", txtBankFullName.Text ?? "" },
                    { "{HEADER_ADDRESS}", txtBankAddress.Text ?? "" },
                    { "{HEADER_PHONE}", txtBankContacts.Text ?? "" },
                    { "{RECIPIENT_POST}", txtRecipientPost.Text ?? "" },
                    { "{RECIPIENT_ORGANIZATION}", txtRecipientOrg.Text ?? "" },
                    { "{RECIPIENT_NAME}", txtRecipientName.Text ?? "" },
                    { "{RECIPIENT_FIRSTNAME}", GetFirstName(txtRecipientName.Text) },
                    { "{LETTER_DATE}", dateValue },
                    { "{LETTER_NUMBER}", txtLetterNumber.Text ?? "" },
                    { "{LETTER_SUBJECT}", txtSubject.Text ?? "" },
                    { "{SENDER_POSITION}", txtSenderPosition.Text ?? "" },
                    { "{SENDER_NAME}", txtSenderName.Text ?? "" }
                };

             
                foreach (var textElement in doc.MainDocumentPart.Document.Descendants<Text>().ToList())
                {
                    if (string.IsNullOrEmpty(textElement.Text)) continue;

                    string original = textElement.Text;
                    string modified = original;

                    foreach (var item in replacements)
                    {
                        if (modified.Contains(item.Key, StringComparison.OrdinalIgnoreCase))
                        {
                            modified = modified.Replace(item.Key, item.Value, StringComparison.OrdinalIgnoreCase);
                        }
                    }

                    if (modified != original)
                        textElement.Text = modified;
                }

                ReplaceLetterBody(doc);

              
                InsertAttachments(doc);

                doc.MainDocumentPart.Document.Save();
            }
        }

        private void ReplaceLetterBody(WordprocessingDocument doc)
        {
            string bodyText = txtBody.Text?.Trim() ?? "";

            if (string.IsNullOrEmpty(bodyText))
                return;

            foreach (var textElement in doc.MainDocumentPart.Document.Descendants<Text>().ToList())
            {
                if (textElement.Text.Contains("{LETTER_BODY}"))
                {
                    textElement.Text = textElement.Text.Replace("{LETTER_BODY}", bodyText);
                    break;
                }
            }
        }

        private string GetFirstName(string fullName)
        {
            if (string.IsNullOrWhiteSpace(fullName)) return "";
            var parts = fullName.Trim().Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            return parts.Length > 0 ? parts[parts.Length - 1] : fullName;
        }

        private void InsertAttachments(WordprocessingDocument doc)
        {
            var body = doc.MainDocumentPart.Document.Body;

          
            if (attachments.Count > 0)
            {
              
                bool placeholderFound = false;

                foreach (var paragraph in body.Descendants<Paragraph>().ToList())
                {
                    if (paragraph.InnerText.Contains("{ATTACHMENTS_LIST}"))
                    {
                        placeholderFound = true;
                        paragraph.Remove(); 

                      
                        string headerText = attachments.Count == 1 ? "Приложение:" : "Приложения:";
                        var headerRun = new Run(new Text(headerText));
                        var headerPara = new Paragraph(headerRun);
                        headerPara.ParagraphProperties = new ParagraphProperties(
                            new SpacingBetweenLines() { After = "100" });
                        body.InsertBefore(headerPara, body.FirstChild);

                        var currentPara = headerPara;

                 
                        foreach (var att in attachments)
                        {
                         
                            var listRun = new Run(new Text(att.GetListString()));
                            var listPara = new Paragraph(listRun);
                            listPara.ParagraphProperties = new ParagraphProperties(
                                new Indentation() { Left = "720" },
                                new SpacingBetweenLines() { After = "50" });
                            body.InsertBefore(listPara, currentPara);
                            currentPara = listPara;
                        }

                        
                        var emptyPara = new Paragraph(new Run(new Text("")));
                        emptyPara.ParagraphProperties = new ParagraphProperties(
                            new SpacingBetweenLines() { After = "200" });
                        body.InsertBefore(emptyPara, currentPara);

                        break;
                    }
                }

             
                if (!placeholderFound)
                {
                 
                    var lastParagraph = body.Descendants<Paragraph>().LastOrDefault();

                    if (lastParagraph != null)
                    {
                       
                        string headerText = attachments.Count == 1 ? "Приложение:" : "Приложения:";
                        var headerPara = new Paragraph(new Run(new Text(headerText)));
                        headerPara.ParagraphProperties = new ParagraphProperties(
                            new SpacingBetweenLines() { Before = "200", After = "100" });
                        body.InsertAfter(headerPara, lastParagraph);

                        var currentPara = headerPara;

                        
                        foreach (var att in attachments)
                        {
                            var listRun = new Run(new Text(att.GetListString()));
                            var listPara = new Paragraph(listRun);
                            listPara.ParagraphProperties = new ParagraphProperties(
                                new Indentation() { Left = "720" });
                            body.InsertAfter(listPara, currentPara);
                            currentPara = listPara;
                        }
                    }
                }
            }

           
            if (attachments.Count > 0)
            {
               
                var pageBreak = new Paragraph(new Run(new Break() { Type = BreakValues.Page }));
                body.AppendChild(pageBreak);

              
                var appendixHeaderRun = new Run(new Text("ПРИЛОЖЕНИЯ"));
                appendixHeaderRun.RunProperties = new RunProperties(new Bold(), new FontSize() { Val = "28" });
                var appendixHeaderPara = new Paragraph(appendixHeaderRun);
                appendixHeaderPara.ParagraphProperties = new ParagraphProperties(
                    new Justification() { Val = JustificationValues.Center },
                    new SpacingBetweenLines() { After = "400", Before = "200" });
                body.AppendChild(appendixHeaderPara);

               
                for (int i = 0; i < attachments.Count; i++)
                {
                    var att = attachments[i];

                  
                    if (i > 0)
                    {
                        var breakPara = new Paragraph(new Run(new Break() { Type = BreakValues.Page }));
                        body.AppendChild(breakPara);
                    }

                    
                    var numberRun = new Run(new Text(att.GetFormattedHeader()));
                    numberRun.RunProperties = new RunProperties(new Bold());
                    var numberPara = new Paragraph(numberRun);
                    numberPara.ParagraphProperties = new ParagraphProperties(
                        new Justification() { Val = JustificationValues.Right },
                        new SpacingBetweenLines() { After = "200", Before = "200" });
                    body.AppendChild(numberPara);

                 
                    var titleRun = new Run(new Text(att.Title.ToUpper()));
                    titleRun.RunProperties = new RunProperties(new Bold(), new Underline() { Val = UnderlineValues.Single });
                    var titlePara = new Paragraph(titleRun);
                    titlePara.ParagraphProperties = new ParagraphProperties(
                        new Justification() { Val = JustificationValues.Center },
                        new SpacingBetweenLines() { After = "300" });
                    body.AppendChild(titlePara);

                    
                    var contentLines = att.Content.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                    foreach (var line in contentLines)
                    {
                        if (!string.IsNullOrWhiteSpace(line))
                        {
                            var contentRun = new Run(new Text(line.Trim()));
                            var contentPara = new Paragraph(contentRun);
                            contentPara.ParagraphProperties = new ParagraphProperties(
                                new Indentation() { FirstLine = "720" }, 
                                new SpacingBetweenLines() { After = "100" });
                            body.AppendChild(contentPara);
                        }
                    }

                    var emptyPara = new Paragraph(new Run(new Text("")));
                    emptyPara.ParagraphProperties = new ParagraphProperties(
                        new SpacingBetweenLines() { After = "200" });
                    body.AppendChild(emptyPara);
                }
            }
        }


        private void AddAttachment_Click(object sender, RoutedEventArgs e)
        {
            var window = new AttachmentWindow();
            if (window.ShowDialog() == true)
            {
                window.Attachment.Number = attachments.Count + 1;
                attachments.Add(window.Attachment);
                RefreshAttachmentsList();
            }
        }

        private void EditAttachment_Click(object sender, RoutedEventArgs e)
        {
            if (lvAttachments.SelectedItem is Attachment selected)
            {
                var window = new AttachmentWindow(selected);
                if (window.ShowDialog() == true)
                {
                    RefreshAttachmentsList();
                }
            }
            else
            {
                MessageBox.Show("Выберите приложение для редактирования!",
                    "Предупреждение", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void DeleteAttachment_Click(object sender, RoutedEventArgs e)
        {
            if (lvAttachments.SelectedItem is Attachment selected)
            {
                if (MessageBox.Show($"Удалить приложение \"{selected.Title}\"?",
                    "Подтверждение", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
                {
                    attachments.Remove(selected);
                    RefreshAttachmentsList();
                }
            }
            else
            {
                MessageBox.Show("Выберите приложение для удаления!",
                    "Предупреждение", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void RefreshAttachmentsList()
        {
            for (int i = 0; i < attachments.Count; i++)
            {
                attachments[i].Number = i + 1;
            }
            lvAttachments.Items.Refresh();
        }
    }
}