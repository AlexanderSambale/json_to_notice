import 'package:excel/excel.dart';
import 'package:collection/collection.dart';
import 'package:json_to_notices/model.dart';

const basis = 'Basis';

void downloadExcel() {
  var excel = Excel.createExcel();
  excel.rename('Sheet1', 'Basis');

  Sheet basisSheet = excel.sheets[basis]!;
  basisSheet.appendRow(createHeaderRow());

  int start = 3;
  DAYOFTHEWEEK.values.forEachIndexed(
    (index, day) {
      basisSheet.appendRow(createBlankRow());
      createWeekDay(day).forEach((element) {
        basisSheet.appendRow(element);
      });

      merge(basisSheet, start, day, 'A');
      merge(basisSheet, start, day, 'R');
      start += dayOfTheWeekMapping[day]!.numberOfRows + 1;
    },
  );

  excel.save(fileName: 'Aushang.xlsx');
}

CellStyle cellStyleWeek = CellStyle(
  leftBorder: Border(borderStyle: BorderStyle.Thin),
  rightBorder: Border(borderStyle: BorderStyle.Thin),
  topBorder: Border(borderStyle: BorderStyle.Thin),
  bottomBorder: Border(borderStyle: BorderStyle.Thin),
  fontFamily: getFontFamily(FontFamily.Arial),
  fontSize: 18,
  bold: true,
  verticalAlign: VerticalAlign.Center,
  horizontalAlign: HorizontalAlign.Center,
);

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

  int startingHour = dayOfTheWeekMapping[day]!.startingHour;
  int numberOfRows = dayOfTheWeekMapping[day]!.numberOfRows;
  String fullHourMinutes = ':00';
  String halfHourMinutes = ':30';
  String formatedTime = '';
  for (int rowIndex = 0; rowIndex < numberOfRows; rowIndex++) {
    cellValues = [];
    if (rowIndex % 2 == 1) {
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
  montag,
  dienstag,
  mittwoch,
  donnerstag,
  freitag,
  samstag,
  sonntag
}

Map<DAYOFTHEWEEK, DayProps> dayOfTheWeekMapping = {
  DAYOFTHEWEEK.montag: const DayProps(
    startingHour: 15,
    endingHour: 21,
    name: 'MO',
  ),
  DAYOFTHEWEEK.dienstag: const DayProps(
    startingHour: 15,
    endingHour: 21,
    name: 'DI',
  ),
  DAYOFTHEWEEK.mittwoch: const DayProps(
    startingHour: 15,
    endingHour: 21,
    name: 'MI',
  ),
  DAYOFTHEWEEK.donnerstag: const DayProps(
    startingHour: 15,
    endingHour: 21,
    name: 'DO',
  ),
  DAYOFTHEWEEK.freitag: const DayProps(
    startingHour: 15,
    endingHour: 21,
    name: 'FR',
  ),
  DAYOFTHEWEEK.samstag: const DayProps(
    startingHour: 10,
    endingHour: 19,
    name: 'SA',
  ),
  DAYOFTHEWEEK.sonntag: const DayProps(
    startingHour: 10,
    endingHour: 21,
    name: 'SO',
  )
};

List<CellValue?> createBlankRow() {
  List<CellValue?> cellValues = [];
  for (int cellIndex = 1; cellIndex <= 18; cellIndex++) {
    cellValues.add(const TextCellValue(''));
  }
  return cellValues;
}

void merge(Sheet sheet, int start, DAYOFTHEWEEK day, String colName) {
  int end = start + dayOfTheWeekMapping[day]!.numberOfRows - 1;
  sheet.merge(
    CellIndex.indexByString('$colName$start'),
    CellIndex.indexByString('$colName$end'),
    customValue: TextCellValue(dayOfTheWeekMapping[day]!.name),
  );

  sheet.setMergedCellStyle(
      CellIndex.indexByString('$colName$start'), cellStyleWeek);
}
