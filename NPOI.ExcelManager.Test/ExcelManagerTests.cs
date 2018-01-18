using NUnit.Framework;
using ExcelManager;
using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;

namespace NPOI.ExcelManager.Test {

    [TestFixture]
    public class ExcelManagerTests {

        private string docfilePath = @"C: \Users\HP\Desktop\SOUNDSYNQ Project.doc";
        private string invalidFilePath = @"C: \Users\HP\Desktop\SOUNDSYNQ Project.doca";
        private string xlsFilePath = @"C:\Users\HP\Desktop\customer.xls";
        private string xlsxFilePath = @"C:\Users\HP\Desktop\customer.xlsx";

        [Test]
        [TestCase("")]
        [TestCase(" ")]
        [TestCase(null)]
        public void EmptyOrNullPath_ThrowsException(string path) {

            Assert.Throws<ArgumentNullException>(() => new ExcelReader<BaseDto>(path));
        }


        [Test]
        public void NullStream_ThrowsException() {

            Stream nullStream = null;
            Assert.Throws<ArgumentNullException>(() => new ExcelReader<BaseDto>(nullStream));
        }


        [Test]
        public void ValidStream_DoesNotThrowsError() {

            using (Stream validStream = new FileStream(docfilePath, FileMode.Open)) {
                Assert.DoesNotThrow(() => new ExcelReader<BaseDto>(validStream));
            }
        }

        [Test]
        public void ValidFilePath_DoesNotThrowsError() {
            Assert.DoesNotThrow(() => new ExcelReader<BaseDto>(docfilePath));
        }

        [Test]
        public void ValidFilePathInvalidFile_ThrowsError() {
            Assert.Throws<FileNotFoundException>(() => new ExcelReader<BaseDto>(invalidFilePath));
        }

        [Test]
        public void PropertyMissExcelCelReader_ThrowsException() {
            var reader = new ExcelReader<BaseDto>(docfilePath);
            Assert.Throws<ExcelCellReaderAttributeException>(() => reader.Read());
        }

        [Test]
        public void ValidWordFile_ThrowsError() {

            var reader = new ExcelReader<CustomerDto>(docfilePath);
            Assert.Throws<ArgumentException>(() => reader.Read());
        }

        [Test]
        public void ValidFIlePath_ReadsAllRecord() {

            var reader = new ExcelReader<CustomerDto>(xlsFilePath);

            var customers = reader.Read();

            Assert.GreaterOrEqual(5, customers.Count());
        }


        [Test]
        public void ValidStream_ReadsAllRecord() {

            using (Stream validStream = new FileStream(xlsFilePath, FileMode.Open)) {
                var reader = new ExcelReader<CustomerDto>(validStream);

                var customers = reader.Read();

                Assert.GreaterOrEqual(5, customers.Count());
            }
        }

        [Test]
        public void ValidXls_ReadValidPropertyValue() {
            var reader = new ExcelReader<CustomerDto>(xlsFilePath);
            var customers = reader.Read();

            Assert.NotNull(customers);

            var firstCustomer = customers.FirstOrDefault();

            Assert.That(firstCustomer.id == 1);
            Assert.That(() => !string.IsNullOrEmpty(firstCustomer.Name));
        }


        [Test]
        public void XlxsFile_ReadWithoutException() {
            var reader = new ExcelReader<CustomerDto>(xlsxFilePath);
            IEnumerable<CustomerDto> customers = null;

            Assert.DoesNotThrow(() => customers = reader.Read());

            Assert.IsNotNull(customers);
        }
    }
}