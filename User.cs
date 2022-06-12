namespace DataStructures {

    class User {

        public string Name { get; set; }
        public int Age { get; set; }

        public User(){}
        public User(string Name, int Age) {
            Console.WriteLine($"Create user name: {Name} and age: {Age}");
            this.Name = Name;
            this.Age = Age;
        }
    }

}