//all of these are for 'B2 2K'.
function SO_BUOI() {
  return SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName("B2 2K TQ")
    .getRange(1, 9)
    .getValue();
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu("Trợ giúp từ Alex")
    .addItem("Hiện các cột điểm (cột E đến cột L)", "myFuncshowGradeColumns")
    .addItem("Ẩn các cột điểm (cột E đến cột L)", "myFunchideGradeColumns")
    .addItem("Reset Format", "myFuncResetFormat")
    .addItem("Reset Formulas", "myFuncResetFormulas")
    .addSeparator()
    .addSubMenu(
      ui
        .createMenu("Tổng kết cuối đợt")
        .addItem("Tự động thực hiện tất cả các bước tổng kết", "myFuncDoItAll")
        .addSeparator()
        .addItem("Bước 1: Nhân bản tab Tổng kết", "myFuncDuplicate")
        .addItem(
          "Bước 2: Xóa mối liên hệ với Tab Điểm Danh (để tiện chỉnh sửa)",
          "myFuncCopyPaste"
        )
        .addItem(
          'Bước 3: Xóa và chỉ giữ lại những học sinh có trạng thái "Bình thường"',
          "myFuncDeleteStudents"
        )
        .addItem("Bước 4: Điền lại STT", "myFuncRefillTheSeries")
        .addItem(
          "Bước 5: Xóa những cột điểm thừa và xóa cột TRẠNG THÁI",
          "myFuncDeleteBlankGradeColumnsandStatusColumn"
        )
        .addItem(
          "Bước 6: Điền công thức tính ĐIỂM TRUNG BÌNH",
          "myFuncFillTheAverageFormulas"
        )
        .addItem(
          "Bước 7: Điền công thức tính XẾP HẠNG và sắp xếp theo xếp hạng",
          "myFuncFillTheRankFormulasandSortByRank"
        )
    )
    .addSeparator()
    .addItem("Tạo câu hỏi Inbox", "myFuncGenerateInboxQuestions")
    .addItem("Tạo tổng kết buổi", "myFuncGenerateReport")
    .addToUi();
}

//------------------------------------Các hàm phục vụ cho tổng kết cuối đợt-----------------------------------------------------

function myFuncDoItAll() {
  myFuncDuplicate();
  myFuncCopyPaste();
  myFuncDeleteStudents();
  myFuncRefillTheSeries();
  myFuncDeleteBlankGradeColumnsandStatusColumn();
  myFuncFillTheAverageFormulas();
  myFuncFillTheRankFormulasandSortByRank();
}

function myFuncDuplicate() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("B2 2K TK").copyTo(ss); //nhớ kiểm tra xem có bị lỡ tay bấm 2 lần không :))
  sheet.setName("B2 2K TK (FINAL)");
  ss.setActiveSheet(sheet);
}

function myFuncCopyPaste() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("B2 2K TK (FINAL)");
  var range = sheet.getRange("A2:W79");
  ss.setActiveSheet(sheet);

  range.activate();
  range.copyTo(sheet.getRange("A2"), { contentsOnly: true });
}

function myFuncDeleteStudents() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("B2 2K TK (FINAL)");
  var range = sheet.getRange("D4:D79");
  ss.setActiveSheet(sheet);

  range.activate();

  for (var i = 1; i < 77; i++) {
    //77 = 76 + 1; 76 là số hàng học sinh có sẵn.
    var cell = range.getCell(i, 1);
    if (
      cell.getValue() != "Bình thường" &&
      sheet.getRange(i + 3, 1).getValue() != ""
    ) {
      sheet.deleteRow(i + 3);
      i--;
      /*for (var j = i; j < 77; j++) {
          sheet.getRange(j + 4, 2, 1, 23).copyTo(sheet.getRange(j + 3, 2), {contentsOnly:true});
          }*/
    }
  }
}

function myFuncRefillTheSeries() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("B2 2K TK (FINAL)");
  ss.setActiveSheet(sheet);

  var cell = sheet.getRange("A5");
  cell.setFormula("=A4+1");

  var oneRowCopy = sheet.getRange(5, 1, 1, 1);
  var targetRows = sheet.getRange(6, 1, sheet.getLastRow() - 5, 1);
  oneRowCopy.copyTo(targetRows);
}

function myFuncDeleteBlankGradeColumnsandStatusColumn() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("B2 2K TK (FINAL)");
  var range = sheet.getRange("F3:L3");
  ss.setActiveSheet(sheet);

  range.activate();

  for (var i = 1; i < 8; i++) {
    //8 = 7 + 1; 7 là số hàng điểm có sẵn.
    var cell = range.getCell(1, i);
    if (
      cell.getValue() == "" &&
      (sheet.getRange(2, i + 5).getValue() == "1" ||
        sheet.getRange(2, i + 5).getValue() == "2" ||
        sheet.getRange(2, i + 5).getValue() == "3" ||
        sheet.getRange(2, i + 5).getValue() == "4" ||
        sheet.getRange(2, i + 5).getValue() == "5" ||
        sheet.getRange(2, i + 5).getValue() == "6" ||
        sheet.getRange(2, i + 5).getValue() == "7")
    ) {
      sheet.deleteColumn(i + 5);
      i--;
    }
  }

  sheet.deleteColumn(4);
}

function myFuncFillTheAverageFormulas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("B2 2K TK (FINAL)");
  ss.setActiveSheet(sheet);

  for (var i = 1; i < 13; i++) {
    //13 = 7 + 6; 7 là số hàng điểm có sẵn.
    if (sheet.getRange(2, i).getValue() == "ĐIỂM TRUNG BÌNH") {
      var cell = sheet.getRange(4, i);
      var str = "E4";
      if (i == 7) {
        str = "F4";
      } else if (i == 8) {
        str = "G4";
      } else if (i == 9) {
        str = "H4";
      } else if (i == 10) {
        str = "I4";
      } else if (i == 11) {
        str = "J4";
      } else if (i == 12) {
        str = "K4";
      }
      sheet.getRange(4, i).setFormula("=ROUND(AVERAGE(E4:" + str + "); 2)");

      var oneRowCopy = sheet.getRange(4, i);
      var targetRows = sheet.getRange(5, i, sheet.getLastRow() - 4, 1);
      oneRowCopy.copyTo(targetRows);
      break;
    }
  }
}

function myFuncFillTheRankFormulasandSortByRank() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("B2 2K TK (FINAL)");
  ss.setActiveSheet(sheet);

  for (var i = 1; i < 14; i++) {
    //14 = 7 + 7; 7 là số hàng điểm có sẵn.
    if (sheet.getRange(2, i).getValue() == "XẾP HẠNG") {
      var cell = sheet.getRange(4, i);
      var str = "E4";
      var str2 = "$E$4";
      var str3 = "$E";
      if (i == 7) {
        str = "F4";
        str2 = "$F$4";
        str3 = "$F";
      } else if (i == 8) {
        str = "G4";
        str2 = "$G$4";
        str3 = "$G";
      } else if (i == 9) {
        str = "H4";
        str2 = "$H$4";
        str3 = "$H";
      } else if (i == 10) {
        str = "I4";
        str2 = "$I$4";
        str3 = "$I";
      } else if (i == 11) {
        str = "J4";
        str2 = "$J$4";
        str3 = "$J";
      } else if (i == 12) {
        str = "K4";
        str2 = "$K$4";
        str3 = "$K";
      } else if (i == 13) {
        str = "L4";
        str2 = "$L$4";
        str3 = "$L";
      }
      sheet
        .getRange(4, i)
        .setFormula("=RANK(" + str + ";" + str2 + ":" + str3 + ")");

      var oneRowCopy = sheet.getRange(4, i);
      var targetRows = sheet.getRange(5, i, sheet.getLastRow() - 4, 1);
      oneRowCopy.copyTo(targetRows);

      //Sắp xếp theo thứ hạng:
      var range = sheet.getRange(4, 2, sheet.getLastRow() - 3, i + 8);
      range.sort(i);

      break;
    }
  }
}

//------------------------------------Các hàm ẩn/ hiện cột điểm-----------------------------------------------------

function myFuncshowGradeColumns() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  sheet.showColumns(5, 8); // E-L, 8 columns starting from 5th
}

function myFunchideGradeColumns() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  sheet.hideColumns(5, 8); // E-L, 8 columns starting from 5th
}

//---------------------------------------------------------------------------------------------------------------------
//---------------------------------------------------------------------------------------------------------------------
//------------------------------------Gửi gmail nhắc nhở công việc-----------------------------------------------------
//---------------------------------------------------------------------------------------------------------------------
//---------------------------------------------------------------------------------------------------------------------
function myFuncRemindViaGmail() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("2K1 DD");

  var sheet2 = ss.getSheetByName("2K1 TQ");
  var nameCell = sheet2.getRange(16, 1).getValue();
  var mailCell = sheet2.getRange(16, 3).getValue();

  var emailContent =
    "Xin chào " +
    nameCell +
    " - trợ giảng của lớp 2K1,\n\nTôi là Alex, người nhắc việc cho các trợ giảng tại Sam English House.\n\nHiện tại có một số vấn đề sau cần bạn giải quyết:";

  var totalCount = 0;
  //Kiểm tra thông tin còn trống
  var TTrange = sheet.getRange(5, 17 + SO_BUOI(), 76, 3);
  var TTboolean = false;
  var TTcount = 0;
  for (var i = 1; i < 77; i++) {
    //77 = 76 + 1; 76 là số hàng học sinh có sẵn.
    var TTcell1 = TTrange.getCell(i, 2);
    var TTcell2 = TTrange.getCell(i, 3);

    if (
      TTcell1.isBlank() &&
      TTcell2.isBlank() &&
      sheet.getRange(i + 4, 3).isBlank() == false &&
      sheet.getRange(i + 4, 4).getValue() == "Bình thường"
    ) {
      TTboolean = true;
      TTcount = TTcount + 2;
      TTcell1.setNote(
        "From Alex: Vui lòng bổ sung thông tin." + "\n" + TTcell1.getNote()
      );
      TTcell2.setNote(
        "From Alex: Vui lòng bổ sung thông tin." + "\n" + TTcell2.getNote()
      );
    }
  }
  if (TTboolean) {
    totalCount = totalCount + 1;
    emailContent =
      emailContent +
      "\n   " +
      totalCount +
      ") Có " +
      TTcount +
      " ô trống ở phần Thông tin học viên. Vui lòng bổ sung.\n   Lưu ý: Bạn chỉ được nhập thông tin ở Tab 'CÔNG VIỆC HÀNG NGÀY'";
  }

  //Kiểm tra ô điểm danh còn trống
  var DDrange = sheet.getRange(5, 17, 76, SO_BUOI());
  var DDboolean = false;
  var DDcount = 0;
  for (var i = 1; i < 77; i++) {
    //77 = 76 + 1; 76 là số hàng học sinh có sẵn.
    for (var j = 1; j < SO_BUOI() + 1; j++) {
      var DDcell = DDrange.getCell(i, j);
      if (
        DDcell.isBlank() &&
        sheet.getRange(i + 4, 3).isBlank() == false &&
        sheet.getRange(3, j + 16).isBlank() == false
      ) {
        DDboolean = true;
        DDcount = DDcount + 1;
        DDcell.setNote(
          "From Alex: Có phải học viên này nghỉ học không? Nếu phải thì bạn vui lòng hỏi Ms Sam học viên này nghỉ học CP hay KP và bổ sung vào đây." +
            "\n" +
            DDcell.getNote()
        );
      }
    }
  }
  if (DDboolean) {
    totalCount = totalCount + 1;
    emailContent =
      emailContent +
      "\n   " +
      totalCount +
      ") Có " +
      DDcount +
      " ô trống ở phần Điểm danh. Vui lòng bổ sung.";
  }

  //Kiểm tra vào điểm còn trống
  var VDrange = sheet.getRange("F5:L80");
  var VDboolean = false;
  var VDcount = 0;
  for (var i = 1; i < 77; i++) {
    //77 = 76 + 1; 76 là số hàng học sinh có sẵn.
    for (var j = 1; j < 7; j++) {
      var VDcell = VDrange.getCell(i, j);
      if (
        VDcell.isBlank() &&
        sheet.getRange(i + 4, 3).isBlank() == false &&
        sheet.getRange(3, j + 5).isBlank() == false &&
        sheet.getRange(i + 4, 4).getValue() == "Bình thường"
      ) {
        VDboolean = true;
        VDcount = VDcount + 1;
        VDcell.setNote(
          "From Alex: Vui lòng bổ sung điểm của học viên này." +
            "\n" +
            VDcell.getNote()
        );
      }
    }
  }
  if (VDboolean) {
    totalCount = totalCount + 1;
    emailContent =
      emailContent +
      "\n   " +
      totalCount +
      ") Có " +
      VDcount +
      " ô trống ở phần Vào điểm. Vui lòng bổ sung.\n   Lưu ý: Bạn chỉ được vào điểm ở Tab 'CÔNG VIỆC HÀNG NGÀY'";
  }

  //Kiểm tra xem đã ghi BTVN chưa?
  var BTrange = sheet2.getRange(3, 7, SO_BUOI(), 1);
  var BTboolean = false;
  var BTcount = 0;
  for (var i = 1; i < SO_BUOI() + 1; i++) {
    var BTcell = BTrange.getCell(i, 1);
    if (BTcell.isBlank() && sheet2.getRange(i + 2, 2).isBlank() == false) {
      BTboolean = true;
      BTcount = BTcount + 1;
      BTcell.setNote(
        "From Alex: Vui lòng bổ sung BTVN của buổi này." +
          "\n" +
          BTcell.getNote()
      );
    }
  }
  if (BTboolean) {
    totalCount = totalCount + 1;
    emailContent =
      emailContent +
      "\n   " +
      totalCount +
      ") Có " +
      BTcount +
      " ô trống ở cột 'Bài tập về nhà' trong Tab 'TỔNG QUAN ĐỢT HỌC'. Vui lòng bổ sung.";
  }

  //Kiểm tra xem đã ghi Bài kiểm tra chưa?
  var KTrange = sheet2.getRange(3, 6, SO_BUOI(), 1);
  var KTboolean = false;
  var KTcount = 0;
  for (var i = 1; i < SO_BUOI() + 1; i++) {
    var KTcell = KTrange.getCell(i, 1);
    if (KTcell.isBlank() && sheet2.getRange(i + 2, 2).isBlank() == false) {
      KTboolean = true;
      KTcount = KTcount + 1;
      KTcell.setNote(
        "From Alex: Vui lòng bổ sung bài kiểm tra của buổi này. Nếu không có thì ghi là 'Không có'." +
          "\n" +
          KTcell.getNote()
      );
    }
  }
  if (KTboolean) {
    totalCount = totalCount + 1;
    emailContent =
      emailContent +
      "\n   " +
      totalCount +
      ") Có " +
      KTcount +
      " ô trống ở cột 'Bài kiểm tra' trong Tab 'TỔNG QUAN ĐỢT HỌC'. Vui lòng bổ sung.";
  }

  if (totalCount > 0) {
    //Gửi email
    emailContent =
      emailContent + "\n\nChúc bạn có một ngày làm việc vui vẻ,\nAlex.";
    var d = new Date();

    MailApp.sendEmail({
      to: mailCell,
      subject:
        "Nhắc nhở công việc trợ giảng (Ngày " +
        d.getDate() +
        "/" +
        (d.getMonth() + 1) +
        "/" +
        d.getFullYear() +
        ")",
      body: emailContent,
    });
  }
}

//---------------------------------------------------------------------------------------------------------------------
//---------------------------------------------------------------------------------------------------------------------
//------------------------------------Reset Format-----------------------------------------------------
//---------------------------------------------------------------------------------------------------------------------
//---------------------------------------------------------------------------------------------------------------------
function myFuncResetFormat() {
  var ss1 = SpreadsheetApp.openByUrl(
    "https://docs.google.com/spreadsheets/d/1Npz3XBsxtjJbSvIKmjS0cIWd13ycM82M2Y6PnBCLwEI/edit"
  ); //The Format Template
  var sheetsou1 = ss1.getSheetByName("FT DD");
  var sheetsou2 = ss1.getSheetByName("FT TK");
  var sheetsou3 = ss1.getSheetByName("FT HP");
  var sheetsou4 = ss1.getSheetByName("FT TQ");

  var ss2 = SpreadsheetApp.getActiveSpreadsheet();

  var sheetsou1copy = sheetsou1.copyTo(ss2);
  var sheetsou2copy = sheetsou2.copyTo(ss2);
  var sheetsou3copy = sheetsou3.copyTo(ss2);
  var sheetsou4copy = sheetsou4.copyTo(ss2);

  var sheetdes1 = ss2.getSheetByName("2K1 DD");
  var range1 = sheetsou1copy.getRange("A1:AC80");
  range1.copyTo(sheetdes1.getRange("A1"), { formatOnly: true });
  //range1.copyFormatToRange(sheetdes1, 1, 28, 1,80);
  var sheetdes2 = ss2.getSheetByName("2K1 TK");
  var range2 = sheetsou2copy.getRange("A1:W79");
  range2.copyTo(sheetdes2.getRange("A1"), { formatOnly: true });
  var sheetdes3 = ss2.getSheetByName("2K1 HP");
  var range3 = sheetsou3copy.getRange("A1:K83");
  range3.copyTo(sheetdes3.getRange("A1"), { formatOnly: true });
  var sheetdes4 = ss2.getSheetByName("2K1 TQ");
  var range4 = sheetsou4copy.getRange("A1:I15");
  range4.copyTo(sheetdes4.getRange("A1"), { formatOnly: true });

  ss2.deleteSheet(sheetsou1copy);
  ss2.deleteSheet(sheetsou2copy);
  ss2.deleteSheet(sheetsou3copy);
  ss2.deleteSheet(sheetsou4copy);
}

//---------------------------------------------------------------------------------------------------------------------
//---------------------------------------------------------------------------------------------------------------------
//------------------------------------Reset Formulas-----------------------------------------------------
//---------------------------------------------------------------------------------------------------------------------
//---------------------------------------------------------------------------------------------------------------------
function myFuncResetFormulas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet1 = ss.getSheetByName("2K1 TK");

  var oneCellCopy = sheet1.getRange(3, 6);
  oneCellCopy.setFormula("'2K1 DD'!F3");
  for (var i = 1; i < 7; i++) {
    var targetCells = sheet1.getRange(3, i + 6);
    oneCellCopy.copyTo(targetCells);
  }

  oneCellCopy = sheet1.getRange(4, 6);
  oneCellCopy.setFormula("'2K1 DD'!F5");
  for (var j = 1; j < 8; j++) {
    for (var i2 = 4; i2 < 80; i2++) {
      targetCells = sheet1.getRange(i2, j + 5);
      oneCellCopy.copyTo(targetCells);
    }
  }
}

//---------------------------------------------------------------------------------------------------------------------
//---------------------------------------------------------------------------------------------------------------------
//------------------------------------Tự tạo ra những câu hỏi (phép, BTVN, bài kiểm tra) để hỏi chị Sam cuối mỗi ca làm việc-----------------------------------------------------
//---------------------------------------------------------------------------------------------------------------------
//---------------------------------------------------------------------------------------------------------------------
function myFuncGenerateInboxQuestions() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("2K1 DD");
  var content = "";

  //Kiểm tra ô điểm danh còn trống
  var DDrange = sheet.getRange(5, 17, 76, SO_BUOI());
  for (var j = 1; j < SO_BUOI() + 1; j++) {
    var DDdate = sheet.getRange(3, j + 16);
    var DDblankDate = sheet.getRange(3, j + 17);
    if (
      (DDblankDate.isBlank() == true || DDblankDate.getValue() == "Trường") &&
      DDdate.isBlank() == false
    ) {
      var DDboolean = false;
      var DDcount = 0;
      content =
        content +
        "🏫Lớp 2K1 (ngày " +
        Utilities.formatDate(DDdate.getValue(), "GMT", "dd-MM-yyyy") +
        "):\n";
      for (var i = 1; i < 77; i++) {
        //77 = 76 + 1; 76 là số hàng học sinh có sẵn.
        var DDcell = DDrange.getCell(i, j);
        var DDlastName = sheet.getRange(i + 4, 3);

        if (DDcell.isBlank() == true && DDlastName.isBlank() == false) {
          var DDfirstName = sheet.getRange(i + 4, 2);

          if (DDboolean == false) {
            DDboolean = true;
            content =
              content +
              "A/ Chị ơi hôm nay các em sau nghỉ học có phép không ạ?\n";
          }

          DDcount = DDcount + 1;

          var sdtPhuHuynh = sheet.getRange(i + 4, 19 + SO_BUOI()).getValue();
          content =
            content +
            DDcount +
            ") " +
            DDfirstName.getValue() +
            " " +
            DDlastName.getValue() +
            " [ " +
            sdtPhuHuynh +
            " ] " +
            "\n";
        }
      }
      if (DDcount == 0) {
        content = content + "A/ Hôm nay lớp 2K1 đi học đầy đủ hết ạ (y).\n";
      }
      break;
    }
  }

  content = content + "B/ Chị ơi BTVN hôm nay của lớp 2K1 là gì ạ?\n";
  content = content + "C/ Chị ơi hôm nay lớp 2K1 có bài kiểm tra gì không ạ?";
  sheet.getRange("A82").setValue(content);
  sheet.setActiveSelection("A82");
}

//DDcell.getValue().indexOf("N") > -1
//---------------------------------------------------------------------------------------------------------------------
//---------------------------------------------------------------------------------------------------------------------
//------------------------------------Tự tạo phần tổng kết cuối buổi-----------------------------------------------------
//---------------------------------------------------------------------------------------------------------------------
//---------------------------------------------------------------------------------------------------------------------
function myFuncGenerateReport() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("B2 2K DD");
  var sheet2 = ss.getSheetByName("B2 2K TQ");
  var content = "";

  var countTBT = 0;
  var contentTBT = "";
  var countQT = 0;
  var contentQT = "";
  var countMH = 0;
  var contentMH = "";
  var countMT = 0;
  var contentMT = "";
  var countNH = 0;
  var contentNH = "";

  var DDrange = sheet.getRange(5, 17, 76, SO_BUOI());
  for (var j = 1; j < SO_BUOI() + 1; j++) {
    var DDdate = sheet.getRange(3, j + 16);
    var DDblankDate = sheet.getRange(3, j + 17);
    if (
      (DDblankDate.isBlank() == true || DDblankDate.getValue() == "Trường") &&
      DDdate.isBlank() == false
    ) {
      var DDboolean = false;
      var DDcount = 0;
      content =
        content + "_________________________________________________________\n";
      content =
        content +
        "🔰🔰🔰TỔNG KẾT LỚP B2 2K - Buổi " +
        sheet.getRange(2, j + 16).getValue() +
        " Đợt " +
        sheet2.getRange(1, 3).getValue() +
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
            //Check lỗi quên thẻ:
            if (DDcell.getValue().indexOf("QT") > -1) {
              if (countQT == 0) {
                contentQT = contentQT + "➡Quên thẻ:\n";
              }
              countQT = countQT + 1;
              contentQT =
                contentQT +
                countQT +
                ") " +
                DDfirstName.getValue() +
                " " +
                DDlastName.getValue() +
                "\n";
            }

            //Check lỗi thiếu bài tập:
            if (
              DDcell.getValue().indexOf("QBT") > -1 ||
              DDcell.getValue().indexOf("TBT") > -1
            ) {
              if (countTBT == 0) {
                contentTBT = contentTBT + "➡Thiếu bài tập:\n";
              }
              countTBT = countTBT + 1;
              contentTBT =
                contentTBT +
                countTBT +
                ") " +
                DDfirstName.getValue() +
                " " +
                DDlastName.getValue() +
                "\n";
            }

            //Check lỗi muộn học:
            if (DDcell.getValue().indexOf("MH") > -1) {
              if (countMH == 0) {
                contentMH = contentMH + "➡Đi học muộn:\n";
              }
              countMH = countMH + 1;
              contentMH =
                contentMH +
                countMH +
                ") " +
                DDfirstName.getValue() +
                " " +
                DDlastName.getValue() +
                "\n";
            }

            //Check lỗi mất thẻ:
            if (DDcell.getValue().indexOf("MT") > -1) {
              if (countMT == 0) {
                contentMT = contentMT + "➡Mất thẻ:\n";
              }
              countMT = countMT + 1;
              contentMT =
                contentMT +
                countMT +
                ") " +
                DDfirstName.getValue() +
                " " +
                DDlastName.getValue() +
                "\n";
            }

            //Check lỗi nghỉ học:
            if (DDcell.getValue().indexOf("CP") > -1) {
              if (countNH == 0) {
                contentNH = contentNH + "➡Nghỉ học:\n";
              }
              countNH = countNH + 1;
              contentNH =
                contentNH +
                countNH +
                ") " +
                DDfirstName.getValue() +
                " " +
                DDlastName.getValue() +
                " (CP)" +
                "\n";
            }

            if (DDcell.getValue().indexOf("KP") > -1) {
              if (countNH == 0) {
                contentNH = contentNH + "➡Nghỉ học:\n";
              }
              countNH = countNH + 1;
              contentNH =
                contentNH +
                countNH +
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

      break;
    }
  }

  content =
    content + contentTBT + contentQT + contentMH + contentMT + contentNH;
  content =
    content + "_________________________________________________________\n";
  content = content + "🚧 NHẮC NHỞ BUỔI SAU 🚧\n";
  content =
    content + "📌Điểm danh, thu thẻ và kiểm tra bài tập trước khi lên lớp.\n";

  var BTrange = sheet2.getRange(3, 7, SO_BUOI(), 1);
  for (var i = 1; i < SO_BUOI() + 1; i++) {
    //12 = 11 + 1; 11 là tổng số buổi trong 1 đợt.
    var BTcell = BTrange.getCell(i, 1);
    if (
      sheet2.getRange(i + 2, 2).isBlank() == false &&
      sheet2.getRange(i + 3, 2).isBlank() == true
    ) {
      content = content + "🏠BTVN: " + BTcell.getValue() + ".";
      break;
    }
  }

  sheet.getRange("A84").setValue(content);
  sheet.setActiveSelection("A84");
}
