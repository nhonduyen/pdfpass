namespace PassPdf
{
    public class Employee
    {
        public Employee() { }
        public Employee(string vnName, string name, string pw)
        {
            VietnameseName = vnName;
            Name = name;
            Password = pw;
        }
        public string Name { get; set; }
        public string Password { get; set; }
        public string VietnameseName { get; set; }

    }
}
