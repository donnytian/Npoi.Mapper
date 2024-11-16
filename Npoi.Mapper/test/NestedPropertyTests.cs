using System;
using System.IO;
using System.Linq;
using Npoi.Mapper;
using NUnit.Framework;

namespace test;

[TestFixture]
public class NestedPropertyTests : TestBase
{
    private class Customer
    {
        public string Name { get; set; }
        public int? Age { get; set; }
        public CustomerBilling Billing { get; set; }
    }

    private class CustomerBilling
    {
        public BillingAddress Address { get; set; }
        public BillingContact Contact { get; set; }
        public MyStruct Dates;
    }

    private class BillingAddress
    {
        public int AddressId { get; set; }
    }

    private class BillingContact
    {
        public string Name { get; set; }
    }

    private struct MyStruct
    {
        public DateTimeOffset Dto { get; set; }
    }

    [Test]
    public void ImportWithNestedPropertiesTest()
    {
        // Arrange
        const string customerName = "donny";
        const string contactName = "tian";
        const int customerAge = 33;
        const int addressId = 4321;
        var date = DateTime.Now;
        var workbook = GetBlankWorkbook();
        var sheet = workbook.GetSheetAt(0);
        var row0 = sheet.CreateRow(0);
        var row1 = sheet.CreateRow(1);
        var row2 = sheet.CreateRow(2);

        row0.CreateCell(0).SetCellValue("customer name");
        row0.CreateCell(1).SetCellValue("customer age");
        row0.CreateCell(2).SetCellValue("contact name");
        row0.CreateCell(3).SetCellValue("address id");
        row0.CreateCell(4).SetCellValue("birth date");

        row1.CreateCell(0).SetCellValue(customerName);
        row1.CreateCell(1).SetCellValue(customerAge);
        row1.CreateCell(2).SetCellValue(contactName);
        row1.CreateCell(3).SetCellValue(addressId);
        row1.CreateCell(4).SetCellValue(date);

        row2.CreateCell(0).SetCellValue("");
        row2.CreateCell(1).SetCellValue(customerAge.ToString());
        row2.CreateCell(2).SetCellValue((string)null);
        row2.CreateCell(3).SetCellValue("");
        row2.CreateCell(4).SetCellValue("");

        // Act
        var mapper = new Mapper(workbook);
        mapper.Map<Customer>(0, c => c.Name);
        mapper.Map<Customer>(1, c => c.Age);
        mapper.Map<Customer>(2, c => c.Billing.Contact.Name);
        mapper.Map<Customer>(3, c => c.Billing.Address.AddressId);
        mapper.Map<Customer>(4, c => c.Billing.Dates.Dto);

        var objs = mapper.Take<Customer>().ToList();

        // Assert
        Assert.That(objs, Is.Not.Null);
        Assert.That(objs[0].Value.Name, Is.EqualTo(customerName));
        Assert.That(objs[0].Value.Age, Is.EqualTo(customerAge));
        Assert.That(objs[0].Value.Billing.Contact.Name, Is.EqualTo(contactName));
        Assert.That(objs[0].Value.Billing.Address.AddressId, Is.EqualTo(addressId));
        Assert.That(objs[0].Value.Billing.Dates.Dto.Date, Is.EqualTo(date.Date));

        Assert.That(objs[1].Value.Name, Is.EqualTo(null));
        Assert.That(objs[1].Value.Age, Is.EqualTo(customerAge));
        Assert.That(objs[1].Value.Billing.Contact.Name, Is.EqualTo(null));
        Assert.That(objs[1].Value.Billing.Address, Is.EqualTo(null));
        Assert.That(objs[1].Value.Billing.Dates.Dto.Date, Is.EqualTo(DateTime.MinValue));
    }

    [Test]
    public void ExportWithNestedPropertiesTest()
    {
        // Arrange
        const string fileName = "ExportWithNestedPropertiesTest.xlsx";
        if (File.Exists(fileName)) File.Delete(fileName);
        const string customerName = "donny";
        const string contactName = "tian";
        const int customerAge = 33;
        const int addressId = 4321;
        const int addressId2 = 3333;
        var date = DateTime.Now;
        var customer1 = new Customer
        {
            Age = customerAge,
            Name = customerName,
            Billing = new CustomerBilling
            {
                Address = new BillingAddress { AddressId = addressId },
                Contact = new BillingContact { Name = contactName },
                Dates = new MyStruct { Dto = date },
            },
        };
        var customer2 = new Customer
        {
            Age = null,
            Name = null,
            Billing = new CustomerBilling
            {
                Address = new BillingAddress { AddressId = addressId2 },
            },
        };
        var entities = new[] { customer1, customer2 };

        // Act
        var mapper = new Mapper();
        mapper.Map<Customer>(0, c => c.Name);
        mapper.Map<Customer>(1, c => c.Age);
        mapper.Map<Customer>(2, c => c.Billing.Contact.Name);
        mapper.Map<Customer>(3, c => c.Billing.Address.AddressId);
        mapper.Map<Customer>(4, c => c.Billing.Dates.Dto);

        mapper.Save(fileName, entities, false);

        // Assert
        var sheet = mapper.Workbook.GetSheetAt(0);
        var row0 = sheet.GetRow(0);
        var row1 = sheet.GetRow(1);
        var row2 = sheet.GetRow(2);

        Assert.That(sheet, Is.Not.Null);
        Assert.That(row0.GetCell(0).StringCellValue, Is.EqualTo(nameof(Customer.Name)));
        Assert.That(row0.GetCell(1).StringCellValue, Is.EqualTo(nameof(Customer.Age)));
        Assert.That(row0.GetCell(2).StringCellValue, Is.EqualTo(nameof(Customer.Billing.Contact.Name)));
        Assert.That(row0.GetCell(3).StringCellValue, Is.EqualTo(nameof(Customer.Billing.Address.AddressId)));
        Assert.That(row0.GetCell(4).StringCellValue, Is.EqualTo(nameof(Customer.Billing.Dates.Dto)));

        Assert.That(row1.GetCell(0).StringCellValue, Is.EqualTo(customer1.Name));
        Assert.That(row1.GetCell(1).NumericCellValue, Is.EqualTo(customer1.Age));
        Assert.That(row1.GetCell(2).StringCellValue, Is.EqualTo(customer1.Billing.Contact.Name));
        Assert.That(row1.GetCell(3).NumericCellValue, Is.EqualTo(customer1.Billing.Address.AddressId));
        Assert.That(DateTimeOffset.Parse(row1.GetCell(4).StringCellValue).Date, Is.EqualTo(customer1.Billing.Dates.Dto.Date));

        Assert.That(row2.GetCell(0).StringCellValue, Is.EqualTo(customer2.Name ?? ""));
        Assert.That(row2.GetCell(1).NumericCellValue, Is.EqualTo(customer2.Age ?? 0.0));
        Assert.That(row2.GetCell(2).StringCellValue, Is.EqualTo(""));
        Assert.That(row2.GetCell(3).NumericCellValue, Is.EqualTo(customer2.Billing.Address.AddressId));
        Assert.That(DateTimeOffset.Parse(row2.GetCell(4).StringCellValue).Date, Is.EqualTo(customer2.Billing.Dates.Dto.Date));
    }
}
