using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Text_to_Image.Models;

namespace Text_to_Image.Services
{
    public static class DateProcessor
    {
        public static ProcessingOptions ProcessDateInput(string dateInput)
        {
            var options = new ProcessingOptions();

            if (string.IsNullOrWhiteSpace(dateInput))
            {
                // If no input, use current date
                DateTime currentDate = DateTime.Now;
                options.CustomDate = currentDate.ToString("dd-MM-yyyy");
                options.SelectedDay = currentDate.Day.ToString();
                options.SelectedMonth = currentDate.Month.ToString();
                Console.WriteLine($"Using current date: {options.CustomDate}");
                return options;
            }

            try
            {
                DateTime currentDate = DateTime.Now;

                if (dateInput.Contains("-"))
                {
                    // Format: day-month (e.g., "22-5")
                    string[] parts = dateInput.Split('-');
                    if (parts.Length == 2)
                    {
                        int day = int.Parse(parts[0]);
                        int month = int.Parse(parts[1]);

                        if (day >= 1 && day <= 31 && month >= 1 && month <= 12)
                        {
                            DateTime customDateTime = new DateTime(currentDate.Year, month, day);
                            options.CustomDate = customDateTime.ToString("dd-MM-yyyy");
                            options.SelectedDay = day.ToString();
                            options.SelectedMonth = month.ToString();
                            Console.WriteLine($"Date selected: {options.CustomDate}");
                        }
                        else
                        {
                            throw new ArgumentException("Day must be 1-31 and month must be 1-12!");
                        }
                    }
                    else
                    {
                        throw new ArgumentException("Invalid format! Use day-month (e.g., 22-5)");
                    }
                }
                else
                {
                    // Format: day only (e.g., "22")
                    int day = int.Parse(dateInput);
                    if (day >= 1 && day <= 31)
                    {
                        DateTime customDateTime = new DateTime(currentDate.Year, currentDate.Month, day);
                        options.CustomDate = customDateTime.ToString("dd-MM-yyyy");
                        options.SelectedDay = day.ToString();
                        options.SelectedMonth = currentDate.Month.ToString();
                        Console.WriteLine($"Date selected: {options.CustomDate}");
                    }
                    else
                    {
                        throw new ArgumentException("Date must be between 1-31!");
                    }
                }
            }
            catch (FormatException)
            {
                throw new ArgumentException("Please enter valid numbers! (day or day-month)");
            }
            catch (ArgumentOutOfRangeException)
            {
                throw new ArgumentException("Invalid date for the specified month!");
            }

            return options;
        }
    }


}
