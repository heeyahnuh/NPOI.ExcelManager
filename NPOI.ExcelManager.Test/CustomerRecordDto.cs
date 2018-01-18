using ExcelManager;

namespace NPOI.ExcelManager.Test {

    public class BaseDto {
    }

    public class CustomerDto {
        [ExcelReaderCell(0)]
        public int Id { get; set; }

        [ExcelReaderCell(1)]
        public string Name { get; set; }

        [ExcelReaderCell(Name = "ID")]
        public int id { get; set; }
    }

    public class IndexOutOfRangeDto {
        [ExcelReaderCell(10)]
        public int Id { get; set; }
    }
}