namespace DocumentHandler.Models
{
    public class Document
    {
        public long Id { get; set; }

        public string Name { get; set; }

        public string DocumentId { get; set; }

        public byte[] FileBytes { get; set; }
    }
}
