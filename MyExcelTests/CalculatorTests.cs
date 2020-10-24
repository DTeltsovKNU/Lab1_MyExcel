using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;

namespace MyExcel.Tests
{
    [TestClass()]
    public class CalculatorTests
    {
        [TestMethod()]
        public void GeneralTest()
        {
            //arrange
            string s = "34+35";
            double expected = 69;
            double tolerance = 0.001;

            //act
            double actual = Calculator.Evaluate(s);

            //assert
            Assert.AreEqual(expected, actual, tolerance);
        }

        [TestMethod()]
        [ExpectedException(typeof(DivideByZeroException))]
        public void DivideByZeroTest()
        {
            //arrange
            string s = "1/0";
            Exception expected = new DivideByZeroException();

            //act
            double actual = Calculator.Evaluate(s);

            //assert
            Assert.AreEqual(expected, actual);
        }

        [TestMethod()]
        public void NegativeNumberTest()
        {
            //arrange
            int A1 = 23;
            string s = "-" + A1.ToString();
            double expected = -23;
            double tolerance = 0.001;

            //act
            double actual = Calculator.Evaluate(s);

            //assert
            Assert.AreEqual(expected, actual, tolerance);
        }
    }
}