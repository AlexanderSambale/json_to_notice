import 'package:excel/excel.dart';

const basis = 'Basis';

void downloadExcel() {
  var excel = Excel.createExcel();
  excel.rename('Sheet1', 'Basis');

  Sheet basisSheet = excel.sheets[basis]!;
  basisSheet.appendRow(createHeaderRow());

  DAYOFTHEWEEK.values.forEach(
    (day) {
      basisSheet.appendRow(createBlankRow());
      createWeekDay(day).forEach((element) {
        basisSheet.appendRow(element);
      });
    },
  );

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

List<List<CellValue?>> createWeekDay(DAYOFTHEWEEK day) {
  List<List<CellValue?>> cells = [];
  List<CellValue?> cellValues = [];

  int startingHour, maxRows;
  if (day == DAYOFTHEWEEK.Samstag || day == DAYOFTHEWEEK.Sonntag) {
    startingHour = 10;
    maxRows = 22;
  } else {
    startingHour = 15;
    maxRows = 12;
  }
  String fullHourMinutes = ':00';
  String halfHourMinutes = ':30';
  String formatedTime = '';
  for (int rowIndex = 1; rowIndex <= maxRows; rowIndex++) {
    cellValues = [];
    if (rowIndex % 2 == 0) {
      formatedTime =
          '$startingHour$halfHourMinutes-${startingHour + 1}$fullHourMinutes';
      startingHour++;
    } else {
      formatedTime =
          '$startingHour$fullHourMinutes-$startingHour$halfHourMinutes';
    }

    cellValues.add(const TextCellValue(''));
    cellValues.add(TextCellValue(formatedTime));
    for (int cellIndex = 1; cellIndex <= 14; cellIndex++) {
      cellValues.add(const TextCellValue(''));
    }
    cellValues.add(TextCellValue(formatedTime));
    cellValues.add(const TextCellValue(''));
    cells.add(cellValues);
  }
  return cells;
}

enum DAYOFTHEWEEK {
  Montag,
  Dienstag,
  Mittwoch,
  Donnerstag,
  Freitag,
  Samstag,
  Sonntag
}

Map day_of_the_week_mapping = {
  DAYOFTHEWEEK.Montag: 'MO',
  DAYOFTHEWEEK.Dienstag: 'DI',
  DAYOFTHEWEEK.Mittwoch: 'MI',
  DAYOFTHEWEEK.Donnerstag: 'DO',
  DAYOFTHEWEEK.Freitag: 'FR',
  DAYOFTHEWEEK.Samstag: 'SA',
  DAYOFTHEWEEK.Sonntag: 'SO'
};

List<CellValue?> createBlankRow() {
  List<CellValue?> cellValues = [];
  for (int cellIndex = 1; cellIndex <= 18; cellIndex++) {
    cellValues.add(const TextCellValue(''));
  }
  return cellValues;
}
