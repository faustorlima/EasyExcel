using System;
using System.Collections.Generic;
using System.Text;

namespace EasyExcel.Tests.TestModels
{
    public enum Gender
    {
        Male,
        Female
    }

    public class Employee
    {
        public string Name { get; set; }
        public Gender Gender { get; set; }
        public DateTime DateOfBirth { get; set; }
        public decimal Height { get; set; }
        public decimal Weight { get; set; }

        public Employee() { }

        public Employee(
            string name,
            Gender gender,
            DateTime dateOfBirth,
            decimal height,
            decimal weight)
        {
            Name = name;
            Gender = gender;
            DateOfBirth = dateOfBirth;
            Height = height;
            Weight = weight;
        }
    }
}
