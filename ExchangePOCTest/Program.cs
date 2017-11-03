using System;
using System.Text;
using ExchangeEmailService;

namespace ExchangePOCTest
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Please enter the username:");
            string username = Console.ReadLine();
            Console.WriteLine("Please enter the email:");
            string email = Console.ReadLine();
            Console.WriteLine("Please enter the password");
            string password = ReadLineMasked();

            Exchange expoc = new Exchange(username, password, email);

            expoc.SaveUnreadEmail();
            Console.ReadKey();
        }

        public static string ReadLineMasked(char mask = '*')
        {
            var sb = new StringBuilder();
            ConsoleKeyInfo keyInfo;
            while ((keyInfo = Console.ReadKey(true)).Key != ConsoleKey.Enter)
            {
                if (!char.IsControl(keyInfo.KeyChar))
                {
                    sb.Append(keyInfo.KeyChar);
                    Console.Write(mask);
                }
                else if (keyInfo.Key == ConsoleKey.Backspace && sb.Length > 0)
                {
                    sb.Remove(sb.Length - 1, 1);

                    if (Console.CursorLeft == 0)
                    {
                        Console.SetCursorPosition(Console.BufferWidth - 1, Console.CursorTop - 1);
                        Console.Write(' ');
                        Console.SetCursorPosition(Console.BufferWidth - 1, Console.CursorTop - 1);
                    }
                    else Console.Write("\b \b");
                }
            }
            Console.WriteLine();
            return sb.ToString();
        }
    }
}
