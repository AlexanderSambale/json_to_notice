import 'package:excel/excel.dart';

const basis = 'Basis';

void downloadExcel() {
  var excel = Excel.createExcel();
  excel.rename('Sheet1', 'Basis');

  Sheet basisSheet = excel.sheets[basis]!;
  basisSheet.appendRow(createHeaderRow());

  excel.save(fileName: 'Aushang.xlsx');
}

List<CellValue?> createHeaderRow() {
  List<CellValue?> cellValues = [];
  cellValues.add(const TextCellValue(''));
  cellValues.add(const TextCellValue(''));
  for (int courtNumber = 1; courtNumber <= 4; courtNumber++) {
    cellValues.add(TextCellValue('Halle Feld $courtNumber'));
  }
  for (int courtNumber = 1; courtNumber <= 10; courtNumber++) {
    cellValues.add(TextCellValue('Feld $courtNumber'));
  }
  cellValues.add(const TextCellValue(''));
  cellValues.add(const TextCellValue(''));
  return cellValues;
}
