using DocxEncrypter;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xceed.Words.NET;

namespace DocxEncrypterConsole {
    class Program {

        enum ContainerElement {
            Any,
            Pos,
            Neg
        }

        static void Main(string[ ] args) {
            DocX sourceDoc = DocX.Load(@"E:/temp.docx");
            string text = sourceDoc.Text;
            Console.WriteLine("Document size is {0}", text.Length);

            string encryptedText = "Hello world";
            byte[ ] encryptedBytes = Encoding.ASCII.GetBytes(encryptedText);
            bool[ ] encryptedBits = Utils.ConvertToBits(encryptedBytes);
            if (encryptedBits.Length > text.Length) {
                throw new Exception("Encrypted text is larger than container");
            }

            ContainerElement[ ] container = new ContainerElement[text.Length];
            var shuffle = Utils.GetShuffled(container.Length);
            for (var i = 0; i < encryptedBits.Length; i++) {
                container[shuffle[i]] = encryptedBits[i] ? ContainerElement.Pos : ContainerElement.Neg;
            }
            Console.WriteLine("Start encrypting");
            //DocX resultDoc = DocX.Create(@"E:/tempRes.docx");
            DocX resultDoc = sourceDoc.Copy( );
            foreach (var a in  resultDoc.Paragraphs) {
                resultDoc.RemoveParagraph(a);
            }
            var lastContainerIndex = 0;
            foreach (var sourceParagraph in sourceDoc.Paragraphs) {
                Paragraph newParagraph = resultDoc.InsertParagraph( );
                foreach (var paragraphPart in sourceParagraph.MagicText) {
                    Paragraph append(string textPart) {
                        if (paragraphPart.formatting == null) {
                            return newParagraph.Append(textPart);
                        }
                        return newParagraph.Append(textPart, paragraphPart.formatting);
                    }
                    ContainerElement lastUsedElement = ContainerElement.Any;
                    int firstCharIndex = 0;
                    for (var j = 0; j < paragraphPart.text.Length; j++) {
                        var containerIndex = lastContainerIndex + j;
                        if (container[containerIndex] != ContainerElement.Any) {
                            if (lastUsedElement != ContainerElement.Any) {
                                if (lastUsedElement != container[containerIndex]) {
                                    append(paragraphPart.text.Substring(firstCharIndex, j - firstCharIndex - 1)).Spacing(GenerateSpacing(paragraphPart, container[containerIndex] == ContainerElement.Neg));
                                    firstCharIndex = j;
                                }
                            }
                            lastUsedElement = container[containerIndex];
                        }
                    }
                    if (firstCharIndex < paragraphPart.text.Length) {
                        append(paragraphPart.text.Substring(firstCharIndex, paragraphPart.text.Length - firstCharIndex)).Spacing(GenerateSpacing(paragraphPart, lastUsedElement == ContainerElement.Neg));
                    }
                    lastContainerIndex += paragraphPart.text.Length;
                }
                newParagraph.StyleName = sourceParagraph.StyleName;
            }
            Console.WriteLine("Saving");
            resultDoc.SaveAs(@"E:/tempRes.docx");
            Console.WriteLine("Press any key to exit...");
#if DEBUG
            Console.ReadKey( );
#endif
        }

        public static float GenerateSpacing(FormattedText text, bool isLong) {
            if (text.formatting == null) {
                return isLong ? 0.5f : 0f;
            }
            return (float)(Math.Round(text.formatting.Spacing == null ? 0d : (double)text.formatting.Spacing) + (isLong ? 0.5f : 0f));
        }
    }
}
