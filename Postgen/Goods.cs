using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Postgen
{
    class Goods
    {
        private string name = string.Empty;
        private string code = string.Empty;
        private string colorID = string.Empty;
        private int gCount = 0;
        private int mass = 0;

        public Goods()
        {

        }
        public Goods(string st)
        {
            if (!string.IsNullOrEmpty(st))
            {
                st = st.Trim();
                st = st.ToUpper();
                st = st.Replace(" ", "");

                int posCount = st.LastIndexOf("-");

                gCount = Int32.Parse(st.Substring(posCount + 1));

                st = st.Substring(0, st.Length - 2);

                string tmp = st.Substring(st.Length - 2);
                if (tmp == "-B" || tmp == "-K" || tmp == "-P" || tmp == "-Z")
                {
                    colorID = st.Substring(st.Length - 5, 5);
                    code = st.Replace(colorID, "");
                }
                else
                {
                    tmp = st.Substring(st.Length - 1);
                    if (tmp == "B" || tmp == "K" || tmp == "P" || tmp == "Z")
                    {
                        colorID = st.Substring(st.Length - 4, 3) + "-" + tmp;
                        code = st.Replace(st.Substring(st.Length - 4, 4), "");
                    }
                    else
                    {
                        colorID = st.Substring(st.Length - 3, 3);
                        code = st.Replace(colorID, "");
                    }
                }
                code = code.Replace("-", "");                                                      
            }       
        }

        public string Name
        {
            get { return name; }
            set { name = value; }
        }

        public string Code
        {
            get { return code; }
            set { code = value; }
        }

        public string ColorID
        {
            get { return colorID; }
            set { colorID = value; }
        }

        public int Mass
        {
            get { return mass; }  
            set { mass = value;}
        }

        public int Count
        {
            get { return gCount; }
            set { gCount = value; }
        }
    }
}
