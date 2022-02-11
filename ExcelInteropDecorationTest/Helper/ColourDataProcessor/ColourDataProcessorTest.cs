using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelInteropDecoration;
using ExcelInteropDecoration.Helper.ColourDataProcessor;
using ExcelInteropDecorationTest._testBase;
using NUnit.Framework;

namespace ExcelInteropDecorationTest.Helper.ColourDataProcessor
{
    [TestFixture]
    public class ColourDataProcessorTest : InteropDTestBase
    {
        [Test]
        [Category("Unit")]
        [Category("Quick")]
        public void TestGivenHexColourWhenConvertedToRgbThenValueIsCorrect()
        {
            IColourDataProcessor processor = InteropDApi.NewColourDataProcessor();

            string hex = "AA1911";
            Assert.AreEqual(11147537, processor.HexToRgbColour(hex));
            Assert.AreEqual(hex, processor.RgbColourToHex(processor.HexToRgbColour(hex)));

            hex = "A05181";
            Assert.AreEqual(hex, processor.RgbColourToHex(processor.BgrColourToRgb(processor.RgbColourToBgr(processor.HexToRgbColour(hex)))));
        }

        [Test]
        [Category("Unit")]
        [Category("Quick")]
        public void TestGivenRgbColourWhenConvertedToHexThenValueIsCorrect()
        {
            IColourDataProcessor processor = InteropDApi.NewColourDataProcessor();

            int rgb = 125242;
            Assert.AreEqual("1E93A", processor.RgbColourToHex(rgb));
            Assert.AreEqual(rgb, processor.HexToRgbColour(processor.RgbColourToHex(rgb)));

        }

        [Test]
        [Category("Unit")]
        [Category("Quick")]
        public void TestGivenBgrColourWhenConvertedToRgbThenValueIsCorrect()
        {
            IColourDataProcessor processor = InteropDApi.NewColourDataProcessor();

            int bgr = 11171393;
            Assert.AreEqual(4290218, processor.BgrColourToRgb(bgr));
            Assert.AreEqual(bgr, processor.BgrColourToRgb(processor.BgrColourToRgb(bgr)));
        }

        [Test]
        [Category("Unit")]
        [Category("Quick")]
        public void TestGivenHexColourWhenConvertedToRgbBgrRgbAndHexThenOutputMatchesInput()
        {
            IColourDataProcessor processor = InteropDApi.NewColourDataProcessor();

            string hex = "A05181";
            Assert.AreEqual(hex, processor.RgbColourToHex(processor.BgrColourToRgb(processor.RgbColourToBgr(processor.HexToRgbColour(hex)))));
        }
    }
}
