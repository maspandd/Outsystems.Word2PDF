using Org.BouncyCastle.Crypto.Generators;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Outsystems.Word2PDF
{
    class RemoveWatermark
    {

        public void removeWatermark(byte[] bytes, out byte[] bitBytes, out string Text)
        {
            bitBytes = new byte[] { };
            Text = "";

            string result = Encoding.UTF8.GetString(bytes);
            var test = result.Replace("Created with a trial version of Syncfusion Word library", "");
            byte[] bitTest = Encoding.UTF8.GetBytes(test);

            string result2 = Encoding.UTF8.GetString(bitTest);
            var test2 = result2.Replace("or registered the wrong key in your", "");
            byte[] bitTest2 = Encoding.UTF8.GetBytes(test2);

            string result3 = Encoding.UTF8.GetString(bitTest2);
            var test3 = result3.Replace("application.", "").Replace(" Click ", "").Replace("here", "").Replace(" to obtain the valid key.", "");
            byte[] bitTest3 = Encoding.UTF8.GetBytes(test3);

            bitBytes = bitTest;
            Text = test3;

        }

    }
}
