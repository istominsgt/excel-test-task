using ConsoleApp.Models;
using System;
using System.IO;
using System.Linq;
using ClosedXML.Excel;

namespace ConsoleApp
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;
            Console.InputEncoding = System.Text.Encoding.UTF8;

            string filePath;
            while (true)
            {
                Console.Clear();
                Console.Write("Введите путь до файла с данными (или введите 'выход' для завершения): ");
                filePath = Console.ReadLine()?.Trim().Trim('\"') ?? string.Empty;

                if (string.Equals(filePath, "выход", StringComparison.OrdinalIgnoreCase))
                {
                    Console.WriteLine("Выход из программы.");
                    return;
                }

                if (!File.Exists(filePath))
                {
                    Console.WriteLine("Файл не найден. Попробуйте еще раз или введите 'выход' для завершения.");
                    WaitForKeyPress();
                    continue;
                }

                try
                {
                    using var workbook = new XLWorkbook(filePath);
                    Console.WriteLine("Файл найден и открыт успешно.");
                    WaitForKeyPress();
                    var productsSheet = workbook.Worksheet("Товары");
                    var customersSheet = workbook.Worksheet("Клиенты");
                    var ordersSheet = workbook.Worksheet("Заявки");

                    ProcessUserChoices(productsSheet, ordersSheet, customersSheet, filePath);
                    break;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Ошибка при открытии файла: {ex.Message}");
                    WaitForKeyPress();
                }
            }
        }

        /// <summary>
        /// Обработка выбора пользователя и вызов соответствующих функций.
        /// </summary>
        /// <param name="productsSheet">Лист с данными о товарах.</param>
        /// <param name="ordersSheet">Лист с данными о заказах.</param>
        /// <param name="customersSheet">Лист с данными о клиентах.</param>
        /// <param name="filePath">Путь к файлу с данными.</param>
        static void ProcessUserChoices(IXLWorksheet productsSheet, IXLWorksheet ordersSheet, IXLWorksheet customersSheet, string filePath)
        {
            while (true)
            {
                Console.Clear();
                Console.WriteLine("Выберите действие:");
                Console.WriteLine("1. Вывести информацию о клиентах по наименованию товара");
                Console.WriteLine("2. Добавить/Удалить/Изменить контактное лицо");
                Console.WriteLine("3. Определить золотого клиента");
                Console.WriteLine("4. Выход");

                var input = Console.ReadLine()?.Trim() ?? string.Empty;

                // Проверка на корректность ввода
                if (!int.TryParse(input, out int action) || action < 1 || action > 4)
                {
                    Console.WriteLine("Неверный выбор. Пожалуйста, введите число от 1 до 4.");
                    WaitForKeyPress();
                    continue;
                }

                switch (action)
                {
                    case 1:
                        ShowCustomerInfoByProduct(productsSheet, ordersSheet, customersSheet);
                        break;
                    case 2:
                        UpdateContactPerson(customersSheet, filePath);
                        break;
                    case 3:
                        FindTopCustomer(ordersSheet, customersSheet);
                        break;
                    case 4:
                        return;
                }
            }
        }

        /// <summary>
        /// Ожидание нажатия любой клавиши от пользователя.
        /// </summary>
        static void WaitForKeyPress()
        {
            Console.WriteLine("Нажмите любую клавишу для продолжения...");
            Console.ReadKey();
        }

        /// <summary>
        /// Вывести информацию о клиентах, которые заказали указанный товар.
        /// </summary>
        /// <param name="productsSheet">Лист с данными о товарах.</param>
        /// <param name="ordersSheet">Лист с данными о заказах.</param>
        /// <param name="customersSheet">Лист с данными о клиентах.</param>
        static void ShowCustomerInfoByProduct(IXLWorksheet productsSheet, IXLWorksheet ordersSheet, IXLWorksheet customersSheet)
        {
            Console.Clear();
            Console.Write("Введите наименование товара: ");
            var productName = Console.ReadLine()?.Trim();

            var productRow = productsSheet.RowsUsed()
                .FirstOrDefault(row => row.Cell(SheetColumnIndexes.ProductColumns.Name).GetString().ToLower() == productName);

            if (productRow == null)
            {
                Console.WriteLine("Товар не найден.");
                WaitForKeyPress();
                return;
            }

            var productCode = productRow.Cell(SheetColumnIndexes.ProductColumns.Code).GetString();
            Console.WriteLine($"Клиенты, заказавшие товар '{productName}':");

            foreach (var orderRow in ordersSheet.RowsUsed().Skip(SheetColumnIndexes.HeaderRow))
            {
                if (orderRow.Cell(SheetColumnIndexes.OrderColumns.ProductCode).GetString() == productCode)
                {
                    var customerCode = orderRow.Cell(SheetColumnIndexes.OrderColumns.CustomerCode).GetString();
                    var customerRow = customersSheet.RowsUsed()
                        .FirstOrDefault(row => row.Cell(SheetColumnIndexes.CustomerColumns.Code).GetString() == customerCode);

                    if (customerRow != null)
                    {
                        var contactPerson = customerRow.Cell(SheetColumnIndexes.CustomerColumns.ContactPerson).GetString();
                        Console.WriteLine(string.Format(
                            "Контактное лицо (ФИО): {0}, Количество: {1}, Цена: {2}, Дата заказа: {3}",
                            contactPerson,
                            orderRow.Cell(SheetColumnIndexes.OrderColumns.Quantity).GetValue<int>(),
                            productRow.Cell(SheetColumnIndexes.ProductColumns.Price).GetValue<decimal>(),
                            orderRow.Cell(SheetColumnIndexes.OrderColumns.OrderDate).GetDateTime().ToString("dd.MM.yyyy")));
                    }
                }
            }

            WaitForKeyPress();
        }

        /// <summary>
        /// Обновить контактное лицо (добавить, удалить или изменить).
        /// </summary>
        /// <param name="customersSheet">Лист с данными о клиентах.</param>
        /// <param name="filePath">Путь к файлу с данными.</param>
        static void UpdateContactPerson(IXLWorksheet customersSheet, string filePath)
        {
            while (true)
            {
                Console.Clear();
                Console.WriteLine("Выберите действие:");
                Console.WriteLine("1. Добавить новый контакт");
                Console.WriteLine("2. Удалить контакт");
                Console.WriteLine("3. Изменить контакт");
                Console.WriteLine("4. Вернуться в главное меню");

                var input = Console.ReadLine()?.Trim() ?? string.Empty;

                // Проверка на корректность ввода
                if (!int.TryParse(input, out int action) || action < 1 || action > 4)
                {
                    Console.WriteLine("Неверный выбор. Пожалуйста, введите число от 1 до 4.");
                    WaitForKeyPress();
                    continue;
                }

                switch (action)
                {
                    case 1:
                        AddContact(customersSheet);
                        break;
                    case 2:
                        RemoveContact(customersSheet);
                        break;
                    case 3:
                        ChangeContact(customersSheet);
                        break;
                    case 4:
                        return;
                }
            }
        }


        /// <summary>
        /// Добавить новый контакт в список клиентов.
        /// </summary>
        /// <param name="customersSheet">Лист с данными о клиентах.</param>
        static void AddContact(IXLWorksheet customersSheet)
        {
            ShowAllContactPersons(customersSheet);

            Console.Write("Введите ФИО нового контактного лица: ");
            var newContactPerson = Console.ReadLine()?.Trim();

            var existingRow = customersSheet.RowsUsed()
                .FirstOrDefault(row => row.Cell(SheetColumnIndexes.CustomerColumns.ContactPerson).GetString() == newContactPerson);

            if (existingRow != null)
            {
                Console.WriteLine("Контактное лицо уже существует.");
                WaitForKeyPress();
                return;
            }

            Console.Write("Введите название организации: ");
            var organizationName = Console.ReadLine()?.Trim();

            var organizationRow = customersSheet.RowsUsed()
                .FirstOrDefault(row => row.Cell(SheetColumnIndexes.CustomerColumns.OrganizationName).GetString() == organizationName);

            if (organizationRow == null)
            {
                var lastRow = customersSheet.RowsUsed().LastOrDefault();
                var newRowNumber = (lastRow?.RowNumber() ?? 0) + 1;
                var newRow = customersSheet.Row(newRowNumber);
                newRow.Cell(SheetColumnIndexes.CustomerColumns.OrganizationName).Value = organizationName;
                newRow.Cell(SheetColumnIndexes.CustomerColumns.ContactPerson).Value = newContactPerson;

                Console.WriteLine($"Контакт {newContactPerson} добавлен.");
            }
            else
                Console.WriteLine("Организация уже существует. Добавьте контактное лицо к существующей организации.");

            SaveChanges(customersSheet);
            WaitForKeyPress();
        }

        /// <summary>
        /// Удалить контактное лицо из списка клиентов.
        /// </summary>
        /// <param name="customersSheet">Лист с данными о клиентах.</param>
        static void RemoveContact(IXLWorksheet customersSheet)
        {
            ShowAllContactPersons(customersSheet);

            Console.Write("Введите ФИО контактного лица для удаления: ");
            var contactPersonToRemove = Console.ReadLine()?.Trim();

            var rowToRemove = customersSheet.RowsUsed()
                .FirstOrDefault(row => row.Cell(SheetColumnIndexes.CustomerColumns.ContactPerson).GetString() == contactPersonToRemove);

            if (rowToRemove == null)
            {
                Console.WriteLine("Контактное лицо не найдено.");
                WaitForKeyPress();
                return;
            }

            rowToRemove.Delete();
            Console.WriteLine("Контакт удален.");
            SaveChanges(customersSheet);
            WaitForKeyPress();
        }

        /// <summary>
        /// Изменить контактное лицо на новое.
        /// </summary>
        /// <param name="customersSheet">Лист с данными о клиентах.</param>
        static void ChangeContact(IXLWorksheet customersSheet)
        {
            ShowAllContactPersons(customersSheet);

            Console.Write("Введите ФИО текущего контактного лица: ");
            var currentContactPerson = Console.ReadLine()?.Trim();
            Console.Write("Введите ФИО нового контактного лица: ");
            var newContactPerson = Console.ReadLine()?.Trim();

            var customerRow = customersSheet.RowsUsed()
                .FirstOrDefault(row => row.Cell(SheetColumnIndexes.CustomerColumns.ContactPerson).GetString() == currentContactPerson);

            if (customerRow == null)
            {
                Console.WriteLine("Контактное лицо не найдено.");
                WaitForKeyPress();
                return;
            }

            customerRow.Cell(SheetColumnIndexes.CustomerColumns.ContactPerson).Value = newContactPerson;
            Console.WriteLine($"Контактное лицо {currentContactPerson} обновлено на {newContactPerson}.");

            Console.WriteLine("Список всех клиентов после изменения:");
            foreach (var row in customersSheet.RowsUsed().Skip(SheetColumnIndexes.HeaderRow))
            {
                var name = row.Cell(SheetColumnIndexes.CustomerColumns.OrganizationName).GetString();
                var contactPerson = row.Cell(SheetColumnIndexes.CustomerColumns.ContactPerson).GetString();
                Console.WriteLine($"Организация: {name}, Контактное лицо: {contactPerson}");
            }

            SaveChanges(customersSheet);
            WaitForKeyPress();
        }

        /// <summary>
        /// Сохранить изменения в файле Excel.
        /// </summary>
        /// <param name="customersSheet">Лист с данными о клиентах.</param>
        static void SaveChanges(IXLWorksheet customersSheet)
        {
            try
            {
                customersSheet.Workbook.Save();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка при сохранении файла: {ex.Message}");
            }
        }

        /// <summary>
        /// Определить золотого клиента за указанный месяц и год.
        /// </summary>
        /// <param name="ordersSheet">Лист с данными о заказах.</param>
        /// <param name="customersSheet">Лист с данными о клиентах.</param>
        static void FindTopCustomer(IXLWorksheet ordersSheet, IXLWorksheet customersSheet)
        {
            Console.Clear();

            int year;
            while (true)
            {
                Console.Write("Введите год (например, 2023) или 'выход' для отмены: ");
                var yearInput = Console.ReadLine()?.Trim() ?? string.Empty;

                if (yearInput.Equals("выход", StringComparison.OrdinalIgnoreCase))
                {
                    Console.WriteLine("Операция отменена.");
                    WaitForKeyPress();
                    return;
                }

                if (int.TryParse(yearInput, out year) && year >= 1900 && year <= DateTime.Now.Year)
                    break;

                Console.WriteLine("Некорректный год. Пожалуйста, введите год в диапазоне от 1900 до текущего года.");
            }

            int month;
            while (true)
            {
                Console.Write("Введите месяц (например, 8) или 'выход' для отмены: ");
                var monthInput = Console.ReadLine()?.Trim() ?? string.Empty;

                if (monthInput.Equals("выход", StringComparison.OrdinalIgnoreCase))
                {
                    Console.WriteLine("Операция отменена.");
                    WaitForKeyPress();
                    return;
                }

                if (int.TryParse(monthInput, out month) && month >= 1 && month <= 12)
                    break;

                Console.WriteLine("Некорректный месяц. Пожалуйста, введите номер месяца от 1 до 12.");
            }

            var customerOrders = ordersSheet.RowsUsed()
                .Skip(SheetColumnIndexes.HeaderRow)
                .Where(row => row.Cell(SheetColumnIndexes.OrderColumns.OrderDate).GetDateTime().Year == year &&
                              row.Cell(SheetColumnIndexes.OrderColumns.OrderDate).GetDateTime().Month == month)
                .GroupBy(row => row.Cell(SheetColumnIndexes.OrderColumns.CustomerCode).GetString())
                .Select(g => new { CustomerCode = g.Key, OrderCount = g.Count() })
                .OrderByDescending(c => c.OrderCount)
                .FirstOrDefault();

            if (customerOrders == null)
            {
                Console.WriteLine("Нет заказов за указанный период.");
                WaitForKeyPress();
                return;
            }

            var customerRow = customersSheet.RowsUsed()
                .FirstOrDefault(row => row.Cell(SheetColumnIndexes.CustomerColumns.Code).GetString() == customerOrders.CustomerCode);

            var customerName = customerRow?.Cell(SheetColumnIndexes.CustomerColumns.ContactPerson).GetString();
            Console.WriteLine(customerName != null
                ? $"Золотой клиент: {customerName}, Количество заказов: {customerOrders.OrderCount}"
                : "Клиент с наибольшим количеством заказов не найден.");

            WaitForKeyPress();
        }


        /// <summary>
        /// Показать список всех контактных лиц.
        /// </summary>
        /// <param name="customersSheet">Лист с данными о клиентах.</param>
        static void ShowAllContactPersons(IXLWorksheet customersSheet)
        {
            Console.Clear();
            Console.WriteLine("Список всех контактных лиц:" + Environment.NewLine);

            var rows = customersSheet.RowsUsed().Skip(SheetColumnIndexes.HeaderRow);

            if (!rows.Any())
            {
                Console.WriteLine("Нет данных для отображения.");
                WaitForKeyPress();
                return;
            }

            foreach (var row in rows)
            {
                var organizationName = row.Cell(SheetColumnIndexes.CustomerColumns.OrganizationName).GetString();
                var contactPerson = row.Cell(SheetColumnIndexes.CustomerColumns.ContactPerson).GetString();
                Console.WriteLine($"Организация: {organizationName}, Контактное лицо: {contactPerson}");
            }

            Console.WriteLine(Environment.NewLine);
            WaitForKeyPress();
        }
    }
}
