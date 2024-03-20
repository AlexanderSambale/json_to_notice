import 'package:excel/excel.dart';

void downloadExcel() {
  var excel = Excel.createExcel();
  excel.rename('Sheet1', 'Basis');
  excel.save(fileName: 'Aushang.xlsx');
}
