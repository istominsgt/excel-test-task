namespace ConsoleApp.Models
{
    public class SheetColumnIndexes
    {
        public const int HeaderRow = 1;
        public class ProductColumns
        {
            public const int Code = 1;
            public const int Name = 2;
            public const int Unit = 3;
            public const int Price = 4;
        }

        public class CustomerColumns
        {
            public const int Code = 1;
            public const int OrganizationName = 2;
            public const int Address = 3;
            public const int ContactPerson = 4;
        }

        public class OrderColumns
        {
            public const int Code = 1;
            public const int ProductCode = 2;
            public const int CustomerCode = 3;
            public const int RequestNumber = 4;
            public const int Quantity = 5;
            public const int OrderDate = 6;
        }
    }
}
