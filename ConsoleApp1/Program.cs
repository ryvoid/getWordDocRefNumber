using System;
using System.Reflection;
using Microsoft.Office.Interop.Word;

namespace ConsoleApp1
{
    class Program
    {
        static void Main(string[] args)
        {
            object isReadOnly = false;
            object isVisible = true;
            object missing = Missing.Value;
            object fileName = args[0];
            string propRevNumber = "Revision Number";

            Application wordApp = new Application();
            wordApp.Visible = false;
            
            Document wordDocument = wordApp.Documents.Open(
                ref fileName,
                ref missing,
                ref isReadOnly,
                ref missing,
                ref missing,
                ref missing,
                ref missing,
                ref missing,
                ref missing,
                ref missing,
                ref missing,
                ref isVisible);

            object wordDocumentProps = wordDocument.BuiltInDocumentProperties;

            Type propertyType = wordDocumentProps.GetType();
            object property = propertyType.InvokeMember("Item", BindingFlags.Default | BindingFlags.GetProperty, null, wordDocumentProps, new object[] {propRevNumber});

            Type validatedType = property.GetType();
            string propValue = validatedType.InvokeMember("Value", BindingFlags.Default | BindingFlags.GetProperty, null, property, new object[] { }).ToString();

            Console.WriteLine(propValue);

            wordDocument.Close(false, ref missing, ref missing);
            wordApp.Quit(false, ref missing, ref missing);
        }
    }
}