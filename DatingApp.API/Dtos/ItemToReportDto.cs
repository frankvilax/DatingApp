namespace DatingApp.API.Dtos
{
    public class ItemToReportDto
    {
        public string Key { get; set; }

        public string Label { get; set; }

        public string Type { get; set; }

        public PropertiesDto Properties { get; set; }

        public string Status { get; set; }
    }
}