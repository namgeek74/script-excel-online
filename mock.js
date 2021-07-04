// Thay đổi số buổi ở đây
const soBuoi = 8;
// Thay đổi tên file ở đây
const fileName = "B2 2K DD";
// Thay đổi tên Đợt nếu cần

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

  var DDrange = sheet.getRange(5, 17, 76, soBuoi);
  var j = 21;
  //   for (var j = 1; j < soBuoi + 1; j++) {
  var DDdate = sheet.getRange(3, j + 16);
  var DDblankDate = sheet.getRange(3, j + 17);
  if (
    (DDblankDate.isBlank() == true || DDblankDate.getValue() == "Trường") &&
    DDdate.isBlank() == false
  ) {
    content =
      content + "_________________________________________________________\n";
    content =
      content +
      "🔰🔰🔰TỔNG KẾT LỚP B2 2K - Buổi " +
      sheet.getRange(2, j + 16).getValue() +
      " Đợt 8 " +
      " (" +
      Utilities.formatDate(DDdate.getValue(), "GMT", "dd-MM-yyyy") +
      ")🔰🔰🔰\n";
    content = content + "Nội dung buổi học";
    content = content + "Test";
    for (var i = 1; i < 77; i++) {
      //77 = 76 + 1; 76 là số hàng học sinh có sẵn.
      var DDcell = DDrange.getCell(i, j);
      var DDlastName = sheet.getRange(i + 4, 3);

      if (DDlastName.isBlank() == false) {
        var DDfirstName = sheet.getRange(i + 4, 2);

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

    //   break;
    // }
  }

  content = content + contentThieuBaiTap + contentMuonHoc + contentNghiHoc;

  sheet.getRange("A84").setValue(content);
  sheet.setActiveSelection("A84");
}
