// Thay đổi tên file ở đây
const fileName = "B2 2K DD";
// Thay đổi tên cột cần điểm danh: A tương đương với 1
const columnNumber = 21;
// Thay đổi tên hàng bắt đầu của học sinh 1
const defaultRow = 5;
// Thay đổi số học sinh
const studentsNumber = 44;
// Thay đổi tên cột ghi First Name
const firstNameRowNumber = 2;
// Thay đổi tên cột ghi Last Name
const lastNameRowNumber = 3;
// In report ở hàng bao nhiêu
const printReportRow = "A49";

// Tạo giao diện
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("Tính năng riêng chỉ cho Nghé Basi")
    .addItem("Tạo tổng kết buổi", "onGenerateReportForOneSection")
    .addToUi();
}

function onGenerateReportForOneSection() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(fileName);
  var content = "";
  var countThieuBaiTap = 0;
  var contentThieuBaiTap = "";
  var countMuonHoc = 0;
  var contentMuonHoc = "";
  var countNghiHoc = 0;
  var contentNghiHoc = "";

  var DDrange = sheet.getRange(defaultRow, columnNumber, studentsNumber, 1);
  var DDdate = sheet.getRange(3, columnNumber);
  content =
    content + "_________________________________________________________\n";
  content =
    content +
    "TỔNG KẾT LỚP - Buổi " +
    sheet.getRange(2, columnNumber).getValue() +
    " (" +
    Utilities.formatDate(DDdate.getValue(), "GMT", "dd-MM-yyyy") +
    ")\n";
  content = content + "Nội dung buổi học\n";
  for (var i = 1; i < studentsNumber; i++) {
    var DDcell = DDrange.getCell(i, 1);
    var DDlastName = sheet.getRange(i + defaultRow - 1, lastNameRowNumber);

    if (DDlastName.isBlank() == false) {
      var DDfirstName = sheet.getRange(i + defaultRow - 1, firstNameRowNumber);

      if (DDcell.getValue() != 0) {
        //Check lỗi thiếu bài tập:
        if (
          DDcell.getValue().indexOf("QBT") > -1 ||
          DDcell.getValue().indexOf("ThieuBaiTap") > -1
        ) {
          if (countThieuBaiTap == 0) {
            contentThieuBaiTap = contentThieuBaiTap + "➡Thiếu bài tập:\n";
          }
          countThieuBaiTap = countThieuBaiTap + 1;
          contentThieuBaiTap =
            contentThieuBaiTap +
            countThieuBaiTap +
            ") " +
            DDfirstName.getValue() +
            " " +
            DDlastName.getValue() +
            "\n";
        }

        //Check lỗi muộn học:
        if (DDcell.getValue().indexOf("MuonHoc") > -1) {
          if (countMuonHoc == 0) {
            contentMuonHoc = contentMuonHoc + "➡Đi học muộn:\n";
          }
          countMuonHoc = countMuonHoc + 1;
          contentMuonHoc =
            contentMuonHoc +
            countMuonHoc +
            ") " +
            DDfirstName.getValue() +
            " " +
            DDlastName.getValue() +
            "\n";
        }

        //Check lỗi nghỉ học:
        if (DDcell.getValue().indexOf("CP") > -1) {
          if (countNghiHoc == 0) {
            contentNghiHoc = contentNghiHoc + "➡Nghỉ học:\n";
          }
          countNghiHoc = countNghiHoc + 1;
          contentNghiHoc =
            contentNghiHoc +
            countNghiHoc +
            ") " +
            DDfirstName.getValue() +
            " " +
            DDlastName.getValue() +
            " (CP)" +
            "\n";
        }

        if (DDcell.getValue().indexOf("KP") > -1) {
          if (countNghiHoc == 0) {
            contentNghiHoc = contentNghiHoc + "➡Nghỉ học:\n";
          }
          countNghiHoc = countNghiHoc + 1;
          contentNghiHoc =
            contentNghiHoc +
            countNghiHoc +
            ") " +
            DDfirstName.getValue() +
            " " +
            DDlastName.getValue() +
            " (KP)" +
            "\n";
        }
      }
    }
  }

  content = content + contentThieuBaiTap + contentMuonHoc + contentNghiHoc;

  sheet.getRange(printReportRow).setValue(content);
  sheet.setActiveSelection(printReportRow);
}
