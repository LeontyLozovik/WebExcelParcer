namespace WebParcer.Models
{
    public class ErrorViewModel     //Модель ошибки
    {
        public string? RequestId { get; set; }

        public bool ShowRequestId => !string.IsNullOrEmpty(RequestId);
    }
}