using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net.Sockets;
using System.IO;
using System.Net;

namespace CST.ZebraUtils
{
    public class ZplLabel
    {
        public const char FrmtPrefix = '^';
        public const char CtrlPrefix = '~';

        public List<string> _instrs = new List<string>();

        public void AddCtrl(string instr, params object[] parms)
        {
            string parmStr = String.Join(",", parms);
            _instrs.Add(CtrlPrefix + instr + parmStr);
        }
        public void AddFrmt(string instr, params object[] parms)
        {
            string parmStr = String.Join(",", parms);
            _instrs.Add(FrmtPrefix + instr + parmStr);
        }
        public void AddXA()
        {
            AddFrmt("XA");
        }
        public void AddXZ()
        {
            AddFrmt("XZ");
        }
        public void AddFT(int x, int y)
        {
            AddFrmt("FT", x, y);
        }
        public void AddFS()
        {
            AddFrmt("FS");
        }
        public void AddA0N(int size)
        {
            AddFrmt("A0", "N", size);
        }
        public void AddA0B(int size)
        {
            AddFrmt("A0", "B", size);
        }
        public void AddFD(string text)
        {
            AddFrmt("FD", text);
        }
        public void AddField(string text)
        {
            AddFD(text);
            AddFS();
        }
        public void AddText(int x, int y, int size, string text)
        {
            if (!String.IsNullOrWhiteSpace(text))
            {
                AddFT(x, y);
                AddA0N(size);
                AddField(text);
            }
        }
        public void Add270Text(int x, int y, int size, string text)
        {
            if (!String.IsNullOrWhiteSpace(text))
            {
                AddFT(x, y);
                AddA0B(size);
                AddField(text);
            }
        }
        public void AddLblText(int lblX, int y, int size, string lbl, int txtX, string text)
        {
            if (!String.IsNullOrWhiteSpace(text))
            {
                AddText(lblX, y, size, lbl);
                AddText(txtX, y, size, text);
            }
        }
        public void AddTextBlock(int origX, int origY, int size, int rightX, int height, string text)
        {
            AddFrmt("CF", 0, size);
            AddFrmt("FO", origX, origY);               //AddFrmt("FB", PriorityLeftX-LeftX, 2, null, null);
            AddFrmt("TB", "N", rightX - origX, height);
            AddField(text);
        }

        protected virtual List<string> GetInstructions()
        {
            return _instrs;
        }

        public List<string> Instructions
        {
            get
            {
                return GetInstructions();
            }
        }
    }

    public class ZplBracketed : ZplLabel
    {
        protected override List<string> GetInstructions()
        {
            List<string> instructions = new List<string>();
            instructions.Add(FrmtPrefix + "XA");
            instructions.AddRange(_instrs);
            instructions.Add(FrmtPrefix + "XZ");
            return instructions;
        }
    }

    public class ZplSetup : ZplBracketed
    {
        public ZplSetup()
        {
            AddFrmt("MN", "Y");       //Media Tracking
            AddFrmt("PW", 1015);      //Print Width
            AddFrmt("LH", 0, 0);      //Home Position
            //AddFrmt("MF", "N", "N");  //Media Feed
            AddFrmt("PR", 5, 5);      //Print Rate
            AddFrmt("LT", 0);         //Label Top
            AddCtrl("TA", "000");     //Tear Off Adjust Position
            AddFrmt("MM", "T");       //Print Mode
            AddFrmt("MT", "T");       //Thermal Media
            AddCtrl("JS", "N");       //Sensor Select
            AddFrmt("PO", "N");       //Print Orientation
            AddFrmt("PM", "N");       //Print Mirror Image of Label
            AddFrmt("JM", "A");       //Set Dots/Millimeter
            AddFrmt("LR", "N");       //Label Reverse Print
            AddFrmt("CI", 0);         //Change International Font
        }
    }

    public class ZplClear : ZplBracketed
    {
        public ZplClear()
        {
            AddFrmt("MC", "Y");     //Map Clear
            AddFrmt("LR", "N");     //Label Reverse
        }
    }

    public class ZplMultiParam
    {
        public string SerialNo { get; set; }
        public string ProdCD { get; set; }
        public string Descr { get; set; }
        public string GenNote { get; set; }
        public string Priority { get; set; }
        public string Pattern { get; set; }
        public string Col1 { get; set; }
        public string ColDescr1 { get; set; }
        public string Col2 { get; set; }
        public string ColDescr2 { get; set; }
        public string Col3 { get; set; }
        public string ColDescr3 { get; set; }
        public string Emb { get; set; }
        public string EmbThrds { get; set; }
        public string EmbNote { get; set; }
        public string HS { get; set; }
        public string HSNote { get; set; }
        public string PI { get; set; }
        public string PINote { get; set; }
        public string PrdnOrder { get; set; }
        public string ShipCD { get; set; }
    }

    public class ZplCstMulti : ZplBracketed
    {
        public ZplCstMulti() { }

        public ZplCstMulti(ZplMultiParam parm) : this()
        {
            SerialNo = parm.SerialNo;
            ProdCD = parm.ProdCD;
            ProdDescr = parm.Descr;

            int colLength = BarCodeX - LeftX - 2;
            if (!String.IsNullOrWhiteSpace(parm.Col2) && !String.IsNullOrWhiteSpace(parm.Col3)) // 3tone
            {
                colLength = colLength / 3;
                ColorDescr = "3T: " + String.Join("/", 
                    new string[] { parm.ColDescr1.SafeSub(0, colLength), parm.ColDescr2.SafeSub(0, colLength), parm.ColDescr3.SafeSub(0, colLength) });
                ColorCD = String.Join("/", new string[] { parm.Col1, parm.Col2, parm.Col3 });
            }
            else if (!String.IsNullOrWhiteSpace(parm.Col2) && String.IsNullOrWhiteSpace(parm.Col3)) //2tone
            {
                colLength = colLength / 2;
                ColorDescr = "2T: " + String.Join("/", new string[] { parm.ColDescr1.SafeSub(0, colLength), parm.ColDescr2.SafeSub(0, colLength) });
                ColorCD = String.Join("/", new string[] { parm.Col1, parm.Col2 });
            }
            else // 1tone
            {
                ColorDescr = parm.ColDescr1.SafeSub(0, colLength);
                ColorCD = parm.Col1;
            }
            
            Note = parm.GenNote;

            Priority = parm.Priority;

            Pattern = parm.Pattern;

            DecoStr = "";
            string embDtls = String.Join(" ", new string[] { parm.Emb, parm.EmbThrds, parm.EmbNote});
            if (!String.IsNullOrWhiteSpace(embDtls))
            {
                EmbStr = "E-" + embDtls;
                DecoStr = DecoStr + "E";
            }
            string hsDtls = String.Join(" ", new string[] { parm.HS, parm.HSNote });
            if (!String.IsNullOrWhiteSpace(hsDtls))
            {
                HSStr = "H-" + hsDtls;
                DecoStr = DecoStr + "H";
            }
            string perfDtls = String.Join(" ", new string[] { parm.PI, parm.PINote });
            if (!String.IsNullOrWhiteSpace(perfDtls))
            {
                PerfStr = "P-" + perfDtls;
                DecoStr = DecoStr + "P";
            }

            PrdnOrder = parm.PrdnOrder;
            ShipCD = parm.ShipCD;
        }

        public string SerialNo { get; set; }
        public string ProdCD { get; set; }
        public string ProdDescr { get; set; }
        public string ColorCD { get; set; }
        public string ColorDescr { get; set; }
        public string DecoStr { get; set; }
        public string Priority { get; set; }
        public string Note { get; set; }
        public string Pattern { get; set; }
        public string EmbStr { get; set; }
        public string HSStr { get; set; }
        public string PerfStr { get; set; }
        public string PrdnOrder { get; set; }
        public string ShipCD { get; set; }

        const int LeftX = 38;
        const int PriorityLeftX = 675;
        const int BarCodeX = 512;
        const int sectSize = 34;
        const int col2Off = 55;
        const int line2Off = 40;
        public void AddSection(int x, int y, string cNo, string pNo, string color, string deco)
        {
            string serial = cNo;
            if (!String.IsNullOrWhiteSpace(deco))
            {
                serial = serial + " " + deco;
            }
            AddText(x, y, sectSize, "C#:");
            AddText(x + col2Off, y, sectSize, serial);

            string part = pNo;
            if (!String.IsNullOrWhiteSpace(color))
            {
                part = part + " " + color;
            } 
            AddText(x, y + line2Off, sectSize, "P#:");
            AddText(x + col2Off, y + line2Off, sectSize, part);
        }
        const int sectXOff = 340;
        const int sectYOff = 103;
        public void AddSections(int x, int y, string cNo, string pNo, string color, string deco)
        {
            int baseX = x;
            int baseY = y;

            for (int iy = 0; iy < 4; iy++)
            {
                for (int ix = 0; ix < 3; ix++)
                {
                    AddSection(baseX + (ix * sectXOff), baseY + (iy * sectYOff), cNo, pNo, color, deco);
                }

            }
        }
        protected override List<string> GetInstructions()
        {
            _instrs = new List<string>();
            
            AddFrmt("MC", "Y");
            AddFrmt("LR", "N");
            //??lbl.AddFrmt("LL", "0799");  // label length
            AddFrmt("LS", "0");         // Label Shift

            AddFT(BarCodeX, 340);
            AddFrmt("BY", 2, 3, 160);
            AddFrmt("B3", "N", "N", null/*144*/, "Y", "N");
            AddField(SerialNo);

            AddLblText(LeftX, 50, 45, "C#:", 105, SerialNo);
            AddLblText(350, 50, 45, "Part#:", 470, ProdCD);
            AddLblText(650, 50, 45, "Color:", 770, ColorCD);

            //AddFrmt("CF", 0, 34);
            //AddFrmt("FO", LeftX, 62);            
            //AddFrmt("TB", "N", 1000 - LeftX, 67);
            //AddField(ProdDescr);
            AddTextBlock(LeftX, 64, 34, 1000, 67, ProdDescr);

            //AddFrmt("CF", 0, 34);
            //AddFrmt("FO", LeftX, 140);
            //AddFrmt("TB", "N", PriorityLeftX - LeftX, 33);
            //AddField(Note);
            AddTextBlock(LeftX, 140, 34, PriorityLeftX, 33, Note);

            AddText(PriorityLeftX, 165, 45, Priority);

            Add270Text(970, 250, 56, ShipCD);

            AddTextBlock(LeftX, 180, 28, BarCodeX, 25, Pattern);
            AddTextBlock(LeftX, 215, 28, BarCodeX, 25, ColorDescr);
            AddTextBlock(LeftX, 250, 28, BarCodeX, 25, EmbStr);
            AddTextBlock(LeftX, 285, 28, BarCodeX, 25, HSStr);
            AddTextBlock(LeftX, 320, 28, BarCodeX, 25, PerfStr);
            AddTextBlock(LeftX, 355, 28, BarCodeX, 25, PrdnOrder);

            AddSections(LeftX-10, 430, SerialNo, ProdCD, ColorCD, DecoStr);

            return base.GetInstructions();
        }
    }

    public static class ZplPrinterHelper
    {
        public static string GetLabelMultiZpl(params ZplMultiParam[] parms)
        {
            if ((parms == null) || (parms.Length < 1))
            {
                return null;
            }

            ZplSetup setup = new ZplSetup(); // setup clear/default

            ZplClear clear = new ZplClear(); // clear after labels print

            StringBuilder sb = new StringBuilder();

            foreach (string instr in setup.Instructions)
            {
                sb.Append(instr);
            }

            foreach (var parm in parms)
            {
                ZplCstMulti lbl = new ZplCstMulti(parm);

                sb.Append(Environment.NewLine);
                foreach (string instr in lbl.Instructions)
                {
                    sb.Append(instr);
                }
            }

            sb.Append(Environment.NewLine);
            foreach (string instr in clear.Instructions)
            {
                sb.Append(instr);
            }

            return sb.ToString();
        }

        public static void NetworkMultiZpl(string hostName, int port, params ZplMultiParam[] parms)
        {
            string zpl = GetLabelMultiZpl(parms);

            if (zpl == null)
            {
                return;
            }

            NetworkStream ns = null;
            Socket socket = null;
            try
            {
                IPEndPoint printerIP = new IPEndPoint(IPAddress.Parse(hostName), port);

                socket = new Socket(AddressFamily.InterNetwork,
                    SocketType.Stream,
                    ProtocolType.Tcp);
                socket.Connect(printerIP);

                ns = new NetworkStream(socket);
                
                byte[] toSend = Encoding.ASCII.GetBytes(zpl);

                ns.Write(toSend, 0, toSend.Length);
            }
            finally
            {
                if (ns != null)
                    ns.Close();

                if (socket != null && socket.Connected)
                    socket.Close();
            }
        }


    }
}
