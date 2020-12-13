using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using HawaiiWeatherApp;

namespace HawaiiWeatherApp_locationUnitTest
{
    public class locationTextBoxTestFixture
    {
        [
            Test,
            TestCase("123Ohao123", false),
            TestCase("ohao", false),
            TestCase("123", false),
            TestCase("ogjeo3242o!+_f", false),
            TestCase("Ohao", true),
            TestCase("Lanai City", true),
            TestCase("Bradshaw Army Air Field", true),
            TestCase("Kailua-Kona", true)
        ]
        public void TestValidateLocation(string location, bool expectedResult)
        {
            var locationTextBox = new locationTextBox();
            var actualResult = locationTextBox.validateLocation(location);
            Assert.AreEqual(expectedResult, actualResult);
        }
    }
}
