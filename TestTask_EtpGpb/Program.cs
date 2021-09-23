using System;
using System.Collections.Generic;
using System.Threading;
using Word = Microsoft.Office.Interop.Word;

namespace TestTask_EtpGpb
{
    class Program
    {
        static Word.Application _application;
        static string _fileName = @"C:\Temp\test.docx";

        static readonly Mutex _createDocMutex = new();
        static readonly Mutex _fillDocMutex = new();
        static readonly Mutex _saveDocMutex = new();

        static void Main(string[] args)
        {
            _application = new();

            Coordinator(args);
        }

        static void Coordinator(string[] args)
        {
            var mutexes = new List<Mutex>();

            foreach (var arg in args)
            {
                if (MutexAndActionByArg(arg, out var mutex, out var action))
                {
                    mutex.WaitOne();

                    new Thread(() => action()).Start();

                    mutexes.Add(mutex);
                }
            }

            Console.WriteLine($"Потоки запущены");
            Thread.Sleep(500);
            Console.WriteLine($"Потоки начинают работу");

            foreach (var mutex in mutexes)
            {
                mutex.ReleaseMutex();
                mutex.WaitOne();
            }

            Console.WriteLine($"Потоки завершили работу");
        }

        static void CreateDoc()
        {
            _createDocMutex.WaitOne();

            Console.WriteLine($"Поток {nameof(CreateDoc)} начал работу");

            _application.Documents.Add();

            //Имитация долгой работы
            Thread.Sleep(500);
            Console.WriteLine($"Поток {nameof(CreateDoc)} закончил работу\r\n");

            _createDocMutex.ReleaseMutex();
        }

        static void FillDoc()
        {
            _fillDocMutex.WaitOne();

            Console.WriteLine($"Поток {nameof(FillDoc)} начал работу");

            var doc = _application.ActiveDocument;

            doc.Content.SetRange(0, 0);
            doc.Content.Text = "Тестовый текст";

            //Имитация долгой работы
            Thread.Sleep(500);
            Console.WriteLine($"Поток {nameof(FillDoc)} закончил работу\r\n");

            _fillDocMutex.ReleaseMutex();
        }

        static void SaveDoc()
        {
            _saveDocMutex.WaitOne();

            Console.WriteLine($"Поток {nameof(SaveDoc)} начал работу");

            var doc = _application.ActiveDocument;

            doc.SaveAs(_fileName);
            doc.Close();

            //Имитация долгой работы
            Thread.Sleep(500);
            Console.WriteLine($"Поток {nameof(SaveDoc)} закончил работу\r\n");

            _saveDocMutex.ReleaseMutex();
        }

        static bool MutexAndActionByArg(string arg, out Mutex mutex, out Action action)
        {
            switch (arg)
            {
                case "create":
                    {
                        mutex = _createDocMutex;
                        action = CreateDoc;
                        return true;
                    }
                case "fill":
                    {
                        mutex = _fillDocMutex;
                        action = FillDoc;
                        return true;
                    }
                case "save":
                    {
                        mutex = _saveDocMutex;
                        action = SaveDoc;
                        return true;
                    }

                default:
                    {
                        mutex = null;
                        action = null;
                        return false;
                    }
            }
        }
    }
}
