using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

 
    public static class OpenxmlClass
    {
        public static WordprocessingDocument InsertText(this WordprocessingDocument doc, string contentControlTag, string text)
        {
            SdtElement element = doc.MainDocumentPart.Document.Body.Descendants<SdtElement>()
              .FirstOrDefault(sdt => sdt.SdtProperties.GetFirstChild<Tag>()?.Val == contentControlTag);

            if (element == null)
                throw new ArgumentException($"ContentControlTag \"{contentControlTag}\" doesn't exist.");

            element.Descendants<Text>().First().Text = text;
            element.Descendants<Text>().Skip(1).ToList().ForEach(t => t.Remove()); 
            return doc;
        }


    internal static WordprocessingDocument RemoveSdtBlocks(this WordprocessingDocument doc, IEnumerable<string> contentBlocks)
    {
        List<SdtElement> SdtBlocks = doc.MainDocumentPart.Document.Descendants<SdtElement>().ToList();

        if (contentBlocks == null)
            return doc;

        foreach (var s in contentBlocks)
        {
            SdtElement currentElement = SdtBlocks.FirstOrDefault(sdt => sdt.SdtProperties.GetFirstChild<Tag>()?.Val == s);
            if (currentElement == null)
                continue;
            IEnumerable<OpenXmlElement> elements = null;

            if (currentElement is SdtBlock)
                elements = (currentElement as SdtBlock).SdtContentBlock.Elements();
            else if (currentElement is SdtCell)
                elements = (currentElement as SdtCell).SdtContentCell.Elements();
            else if (currentElement is SdtRun)
                elements = (currentElement as SdtRun).SdtContentRun.Elements();

            foreach (var el in elements)
                currentElement.InsertBeforeSelf(el.CloneNode(true));
            currentElement.Remove();
        }
        return doc;
    }




}
