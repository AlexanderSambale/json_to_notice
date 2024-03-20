import 'package:excel/excel.dart';

void downloadExcel() {
  var excel = Excel.createExcel();
  excel.save(fileName: 'Aushang.xlsx');
}
