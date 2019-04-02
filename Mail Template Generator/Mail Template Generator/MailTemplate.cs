using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;

namespace Mail_Template_Generator
{
    class MailTemplate
    {
        public const int MaxItems = 6;
        public const int width = 1700, height = 2200;
        public List<Bitmap[]> Pages = new List<Bitmap[]>();
        public int itemCount = 0;
        public int currentTemplate = 0;

        public int currentPage
        {
            get
            {
                return (int)Math.Floor((decimal)(itemCount / MaxItems));
            }
        }

        public int currentSlot
        {
            get
            {
                return (6 + (itemCount - (Pages.Count * 6)));
            }
        }

        public MailTemplate()
        {
        }

        public void DeleteLast()
        {
            // get the current amount of pages
            var recentPage = Pages[Pages.Count - 1];

            var currentItem = itemCount % 6;

            if (currentItem == 0 && Pages.Count > itemCount / 6)
                Pages.Remove(Pages[Pages.Count - 1]);

            Pages[Pages.Count - 1][itemCount / 6] = null;
            itemCount--;

           // if the page has 0 items delete the page
        }

        public void Add(Bitmap b)
        {


            if (currentPage == Pages.Count)
            {
                Bitmap[] items = new Bitmap[MaxItems];
                Pages.Add(items);
            }

            
            currentTemplate = currentPage;

            Pages[currentPage][currentSlot] = (Bitmap)b.Clone();
            itemCount++;
        }

        public Bitmap Template(int number)
        {
            var offsetLRM = 5; // Left , Right, Middle
            var offsetTB = 7; // top bottom
            var itemWidth = (width + offsetLRM) / 2;
            var itemHeight = (height + offsetTB) / 3;

            Bitmap template = new Bitmap(width, height);

            using (Graphics g = Graphics.FromImage(template))
            {
                var currentPage = (int)Math.Floor((decimal)(itemCount / MaxItems));
                int max = number < currentPage ? 6 : currentSlot;

                for (var u = 1; u < max + 1; u++)
                {
                    int row = (int)Math.Floor((decimal)(u / 2.5));
                    int col = u % 2 == 0 ? 1 : 0;
                    var boundary = new Rectangle(offsetLRM + (col * itemWidth), offsetTB + (row * itemHeight), itemWidth - (offsetLRM / 2), itemHeight - (offsetTB / 2));
                    //g.DrawRectangle(new Pen(Brushes.Black, 3), boundary);
                    g.DrawImage(Pages[number][u - 1],boundary);
                    // draw the items on the page after every odd item change rows
                    // once it reaches 6 return the template
                    // I want the row to change every even number

                }
            }

            return template;

        }

        public Bitmap[] AllTemplates
        {
            get
            {
                Bitmap[] templates = new Bitmap[Pages.Count];

                for (var j = 0; j < Pages.Count; j++)
                {
                    templates[j] = Template(j);
                }

                return templates;
            }
        }

        public static Bitmap createTemplate(string name, string address)
        {

            Bitmap b = (Bitmap)Properties.Resources.Hunter_Mail_Template_Box__Single_.Clone();

            var addressSplit = address.Split(',');

            using (Graphics g = Graphics.FromImage(b))
            {

                g.DrawString(name, new Font("Arial", 30), Brushes.Black, new Point(115, 300));
                g.DrawString(addressSplit[0], new Font("Arial", 20), Brushes.Black, new Point(115, 400));
                g.DrawString(addressSplit[1], new Font("Arial", 20), Brushes.Black, new Point(115, 450));

            }

            return b;

        }

    }
}
