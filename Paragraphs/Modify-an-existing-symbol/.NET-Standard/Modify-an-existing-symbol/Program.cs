﻿using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Modify_an_existing_symbol
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"../../../Template.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Creates a new Word document.
                using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Automatic))
                {
                    //Gets the textbody content.
                    WTextBody textbody = document.Sections[0].Body;
                    //Iterates through the paragraphs.
                    foreach (WParagraph paragraph in textbody.Paragraphs)
                    {
                        //Gets the symbol from the paragraph items.
                        foreach (ParagraphItem item in paragraph.ChildEntities)
                        {
                            if (item is WSymbol)
                            {
                                WSymbol symbol = item as WSymbol;
                                if (symbol.CharacterCode == 100)
                                {
                                    //Modifies the character code.
                                    symbol.CharacterCode = 40;
                                    symbol.FontName = "Wingdings";
                                }
                            }
                        }
                    }
                    //Creates file stream.
                    using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"../../../Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Saves the Word document to file stream.
                        document.Save(outputFileStream, FormatType.Docx);
                    }
                }
            }
        }
    }
}
