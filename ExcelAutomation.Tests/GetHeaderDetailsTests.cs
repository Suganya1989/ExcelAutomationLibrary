using ExcelAutomationLibrary;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Should;

namespace ExcelAutomation.Tests
{
    public class GetHeaderDetailsTests
    {
        public void ShouldGetAllHeaderDetails()
        {
            List<Employee> lstEmployee = new List<Employee>();
            lstEmployee = CreateRequestObject();
            var headers = new GetHeaders<Employee>().Execute(lstEmployee.First());
            headers.Headers.ShouldNotBeNull();

        }

        public void ShouldBeAbleToWriteToExcel()
        {
            List<Employee> lstEmployee = new List<Employee>();
            lstEmployee = CreateRequestObject();
            var result = new ExcelWriter<Employee>();
            result.Write(lstEmployee);
          
        }


        private List<Employee> CreateRequestObject()
        {
            List<Employee> lstEmployee = new List<Employee>();
            lstEmployee.Add(new Employee()
            {
                Name = "Suganya",
                Id = "1206660",
                Experience = "5 years"
            });
            lstEmployee.Add(new Employee()
            {
                Name = "Abc",
                Id = "1222",
                Experience = "1 years"
            });
            lstEmployee.Add(new Employee()
            {
                Name = "Pqrs",
                Id = "7885",
                Experience = "8 years"
            });
            return lstEmployee;
        }
    }
}
