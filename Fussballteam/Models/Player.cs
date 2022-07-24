namespace Fussballteam.Models
{
    public class PlayerModel
    {
        public string? Name { get; set; }
        public string? Position { get; set; }
        public string? Age { get; set; }
        public string? Salary { get; set; }

        public bool isValid()
        {
            return !string.IsNullOrEmpty(Name);
        }

        public override string ToString()
        {
            return Name;
        }
    }
}
