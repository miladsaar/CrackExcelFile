using System;
using System.Net.Sockets;
using System.Threading;
using CrackExcelFile;
namespace TestExcelCrack
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.ForegroundColor = ConsoleColor.DarkGreen;

            //var currentNumber = false;

            //while (!currentNumber)
            //{
            //    currentNumber = true;

            //    Console.ForegroundColor = ConsoleColor.DarkGreen;
            //    Console.WriteLine("what you want do?");
            //    Console.WriteLine("1-remove password.");
            //    Console.WriteLine("2-read passwords");
            //    Console.WriteLine("3-give back password.");
            //    Console.WriteLine("4-set new password.");
            //    Console.Write("select one of the up number:");
            //    var num = Console.ReadLine();
            //    var path = string.Empty;
            //    if (num != null && (num.Equals("1") || num.Equals("2") || num.Equals("3") || num.Equals("4")))
            //    {
            //        Console.Write("Where is target file:");
            //        path = Console.ReadLine();

            //        Console.WriteLine("please wait ...");

            //    }

            //    switch (num)
            //    {
            //        case "1":
            //            RemoveExcelPass.OpenPass(path, CrackOption.RemovePassAndKeep);
            //            break;
            //        case "2":
            //            RemoveExcelPass.ReadSavedPasswords(path);
            //            break;
            //        case "3":
            //            //Todo:return password
            //            break;
            //        case "4":
            //            //Todo:set new password
            //            break;
            //        default:
            //            Console.Clear();
            //            Console.WriteLine("select curent number");
            //            currentNumber = false;
            //            break;
            //    }


            //}
            RemoveExcelPass.OpenPass(@"d:\temp\mach.xlsx", CrackOption.RemovePassAndKeep);
            RemoveExcelPass.ReadSavedPasswords(@"d:\temp\mach_new.xlp");
            Console.ReadKey();
        }
    }
}
