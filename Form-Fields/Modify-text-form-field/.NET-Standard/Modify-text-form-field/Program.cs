using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System;
using System.IO;
using System.Drawing;

namespace Modify_text_form_field
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStream = new FileStream(Path.GetFullPath(@"C:/Docs/Template.docx"), FileMode.Open, FileAccess.ReadWrite))
            {
                //Opens the template document.
                using (WordDocument document = new WordDocument(fileStream, FormatType.Docx))
                {
                    //Iterates through section.
                    foreach (WSection section in document.Sections)
                    {
                        //Iterates through section child elements.
                        foreach (WTextBody textBody in section.ChildEntities)
                        {
                            //Iterates through form fields.
                            foreach (WFormField formField in textBody.FormFields)
                            {
                                switch (formField.FormFieldType)
                                {
                                    case FormFieldType.TextInput:
                                        WTextFormField textField = formField as WTextFormField;


                                        if (textField.Name == "txt_1")
                                        {
                                            //Modifies the text form field.
                                            textField.Type = TextFormFieldType.RegularText;
                                            textField.StringFormat = "";
                                            textField.DefaultText = "";
                                            textField.Text = "Hyundai";
                                            textField.CalculateOnExit = false;
                                        }

                                        else if (textField.Type == TextFormFieldType.DateText)
                                        {
                                            //Modifies the text form field.
                                            textField.Type = TextFormFieldType.RegularText;                                           
                                            textField.StringFormat = "";
                                            textField.DefaultText = "";
                                            textField.Text = " UMUT";
                                            textField.CalculateOnExit = false;
                                        }

                                        else if (textField.Type == TextFormFieldType.RegularText)
                                        {
                                            //Modifies the text form field.
                                            textField.Type = TextFormFieldType.RegularText;                                           
                                            textField.StringFormat = "MM/DD/YY";
                                            textField.DefaultText = "";
                                            textField.Text = "152657";
                                            textField.CalculateOnExit = false;
                                        }

                                        else
                                        {
                                            //Modifies the text form field.
                                            textField.Type = TextFormFieldType.RegularText;
                                            textField.CharacterFormat.FontName = "Calibri";
                                            textField.CharacterFormat.Bold = true;
                                            textField.StringFormat = "";
                                            textField.DefaultText = "";
                                            textField.Text = " Some Signature Here.";
                                            textField.CalculateOnExit = false;
                                        }

                                        break;
                                }                            
                            }
                        }
                    }
                    //Creates file stream.
                    using (FileStream outputStream = new FileStream(Path.GetFullPath(@"C:/Docs/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Saves the Word document to file stream.
                        document.Save(outputStream, FormatType.Docx);
                    }

                    //This version makes the old file as new one.

                    //document.Save(fileStream, FormatType.Docx);

                    //Closes the document
                    //document.Close();
                }
            }
        }
    }
}
