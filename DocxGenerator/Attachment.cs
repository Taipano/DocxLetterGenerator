using System;

namespace DocxLetterGenerator
{
    public class Attachment
    {
        public int Number { get; set; }
        public string Title { get; set; }
        public string Content { get; set; }

        public Attachment()
        {
            Title = string.Empty;
            Content = string.Empty;
        }

        public Attachment(int number, string title, string content)
        {
            Number = number;
            Title = title;
            Content = content;
        }

        public string GetFormattedHeader()
        {
            return Number == 1 ? "Приложение" : $"Приложение {Number}";
        }

        public string GetListString()
        {
            return $"{Number}. {Title}";
        }
    }
}