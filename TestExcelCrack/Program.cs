using System;
using System.Net.Sockets;
using System.Security.Cryptography;
using System.Threading;
using System.Threading.Tasks;
using CrackExcelFile;
namespace TestExcelCrack
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.ForegroundColor = ConsoleColor.DarkGreen;

            var currentNumber = false;

            while (!currentNumber)
            {
                currentNumber = true;

                Console.ForegroundColor = ConsoleColor.DarkGreen;
                Console.WriteLine("what you want do?");
                Console.WriteLine("1-remove password.");
                Console.WriteLine("2-read passwords");
                Console.WriteLine("3-give back password.");
                Console.WriteLine("4-set new password.");
                Console.Write("select one of the up number:");
                var num = Console.ReadLine();
                var path = string.Empty;
                if (num != null && (num.Equals("1") || num.Equals("2") || num.Equals("3") || num.Equals("4")))
                {
                    Console.Write("Where is target file:");
                    path = Console.ReadLine();

                    Console.WriteLine("please wait ...");

                }

                switch (num)
                {
                    case "1":
                        RemoveExcelPass.OpenPass(path, CrackOption.RemovePassAndKeep);
                        break;
                    case "2":
                        Console.WriteLine(RemoveExcelPass.ReadSavedPasswords(path));
                        break;
                    case "3":
                        RemoveExcelPass.ReturnPassword(path);
                        break;
                    case "4":
                        //Todo:set new password
                        break;
                    default:
                        Console.Clear();
                        Console.WriteLine("select curent number");
                        currentNumber = false;
                        break;
                }


            }



            Thread thr = new Thread(() =>
            {
                RemoveExcelPass.OpenPass(@"d:\temp\mach.xlsx", CrackOption.RemovePassAndKeep);
            });

            Thread thr2 = new Thread(() =>
              {
                  Console.WriteLine(RemoveExcelPass.Message);

              });
            thr2.Priority = ThreadPriority.Highest;
            thr.Priority = ThreadPriority.BelowNormal;
            thr.Start();
            thr2.Start();
            while (true)
            {
                if (thr.ThreadState == ThreadState.Stopped)
                {
                    ;
                    break;
                }
            }



            ////RemoveExcelPass.ShowMessagesAsync(() =>
            ////{
            ////    Console.WriteLine(RemoveExcelPass.Message);
            ////});
            ////RemoveExcelPass.OpenPass(@"d:\temp\mach.xlsx", CrackOption.RemovePassAndKeep);
            ////Console.WriteLine(RemoveExcelPass.ReadSavedPasswords(@"d:\temp\mach_new.xlp"));

            Console.ReadKey();
        }

       

       
        private static string SHA512(string input)
        {
          

            var bytes = System.Text.Encoding.UTF8.GetBytes(input);
            using (var hash=System.Security.Cryptography.SHA512.Create())
            {
                var hashInputByte = hash.ComputeHash(bytes);
                var hashInputStringBuilder = new System.Text.StringBuilder(128);
                foreach (var b in hashInputByte)
                {
                    hashInputStringBuilder.Append(b.ToString("X2"));

                }
              
                return hashInputStringBuilder.ToString( );
            }

          

        }
    }
}
