namespace PaperMgr
{
    class Person
    {
        public Person(string firstName, string lastName, string surname)
        {
            this.FirstName = firstName;
            this.LastName = lastName;
            this.Surname = surname;
        }
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string Surname { get; set; }

        public string GetFullName(bool surnameFirst = true)
        {
            if (surnameFirst)
                return Surname + " " + FirstName + " " + LastName;
            else
                return FirstName + " " + LastName + " " + Surname;
        }

        public string GetInitials()
        {
            return FirstName.Substring(0, 1) + ". " + LastName.Substring(0, 1) + ".";
        }

        public string GetName(bool surnameFirst = true)
        {
            if (surnameFirst)
                return Surname + " " + GetInitials();
            else
                return GetInitials() + " " + Surname;
        }
    }
}
