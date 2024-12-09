using iText.Kernel.Pdf.Canvas.Parser.Data;
using iText.Kernel.Pdf.Canvas.Parser.Listener;
using iText.Kernel.Pdf.Canvas.Parser;
using iText.Kernel.Pdf;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using iText.Kernel.Geom;
using System.Text;
using System.Threading.Tasks;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Shared
{
    public class PDFHelper : IDisposable
    {
        private PdfDocument pdfDocument;

        public PDFHelper(string filePath)
        {
            pdfDocument = new PdfDocument(new PdfReader(filePath));
        }

        public List<PDFImage> GetImages(int pageNumber)
        {
            var page = pdfDocument.GetPage(pageNumber);
            var imageListener = new MyImageRenderListener();
            var processor = new PdfCanvasProcessor(imageListener);
            processor.ProcessPageContent(page);
            return imageListener.Images;
        }

        public List<PDFText> GetTexts(int pageNumber)
        {
            var page = pdfDocument.GetPage(pageNumber);
            var textListener = new MyTextRenderListener();
            var processor = new PdfCanvasProcessor(textListener);
            processor.ProcessPageContent(page);
            return textListener.Texts;
        }

        public double CalculateDistance(PDFImage image, PDFText text)
        {
            return Math.Abs(image.Position.Y - text.Position.Y);
        }

        public void Dispose()
        {
            pdfDocument.Close();
        }

        public PDFText GetTextByContent(string content, int pageNumber)
        {
            var texts = GetTexts(pageNumber);
            return texts.FirstOrDefault(t => t.Content.ToUpper().Contains(content));
        }

        private class MyImageRenderListener : IEventListener
        {
            public List<PDFImage> Images { get; } = new List<PDFImage>();

            public void EventOccurred(IEventData data, EventType type)
            {
                if (type == EventType.RENDER_IMAGE)
                {
                    var renderInfo = (ImageRenderInfo)data;
                    var ctm = renderInfo.GetImageCtm();
                    var x = ctm.Get(Matrix.I31);
                    var y = ctm.Get(Matrix.I32);
                    var width = ctm.Get(Matrix.I11);
                    var height = ctm.Get(Matrix.I22);

                    Images.Add(new PDFImage
                    {
                        Position = new PointF((float)x, (float)y),
                        Width = (float)width,
                        Height = (float)height
                    });
                }
            }

            public ICollection<EventType> GetSupportedEvents()
            {
                return new HashSet<EventType> { EventType.RENDER_IMAGE };
            }
        }

        private class MyTextRenderListener : ITextExtractionStrategy
        {
            public List<PDFText> Texts { get; } = new List<PDFText>();

            public void RenderText(TextRenderInfo renderInfo)
            {
                var baseline = renderInfo.GetBaseline();
                var startPoint = baseline.GetStartPoint();
                var x = (float)startPoint.Get(Vector.I1);
                var y = (float)startPoint.Get(Vector.I2);
                var text = renderInfo.GetText();

                Texts.Add(new PDFText
                {
                    Position = new PointF(x, y),
                    Content = text
                });
            }

            public void BeginTextBlock() { }
            public void EndTextBlock() { }
            public void RenderImage(ImageRenderInfo renderInfo) { }

            public string GetResultantText()
            {
                return string.Join(" ", Texts.ConvertAll(t => t.Content));
            }

            public void EventOccurred(IEventData data, EventType type)
            {
                if (type == EventType.RENDER_TEXT)
                {
                    RenderText((TextRenderInfo)data);
                }
            }

            public ICollection<EventType> GetSupportedEvents()
            {
                return new HashSet<EventType> { EventType.RENDER_TEXT };
            }
        }

        public class PDFImage
        {
            public PointF Position { get; set; }
            public float Width { get; set; }
            public float Height { get; set; }
        }

        public class PDFText
        {
            public PointF Position { get; set; }
            public string Content { get; set; }
        }
    }
}
