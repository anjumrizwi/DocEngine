using System;
using System.Collections.Generic;
using System.IO;
using NGS.Templater;

namespace DocEngine.Processor
{
    public class TemplaterMailMerge
    {
        public static void MailMerge()
        {
            // 1. Prepare your data model
            var data = new
            {
                CustomerName = "John Smith",
                OrderDate = DateTime.Now,
                Items = new[] {
                new { Product = "Laptop", Price = 999.99 },
                new { Product = "Mouse", Price = 19.99 }
                }
            };

            // 2. File paths
            string templatePath = @"D:\Users\anjum.rizwi\source\repos\anjumrizwi\DocEngine\DocEngine.TestClient\template\Business-Plan-Word-Template.docx";
            string outputPath = @"D:\Users\anjum.rizwi\source\repos\anjumrizwi\DocEngine\DocEngine.TestClient\template\docx\templater_output.docx";

            // 3. Process the document (CORRECT 2024 SYNTAX)
            using (var input = File.OpenRead(templatePath))
            using (var output = File.Create(outputPath))
            {
                // The modern way to use Templater
                var factory = Configuration.Builder.Build();
                using (var doc = factory.Open(input, "docx", output))
                {
                    doc.Process(data);
                }
            }

            Console.WriteLine("Document generated successfully!");
        }
    }
}