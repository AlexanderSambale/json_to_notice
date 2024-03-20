import 'dart:math';

import 'package:excel/excel.dart';
import 'package:collection/collection.dart';
import 'package:json_to_notices/mocks.dart';
import 'package:json_to_notices/models.dart/day_props.dart';

import 'models.dart/event.dart';

const basis = 'Basis';

void downloadExcel() {
  var excel = Excel.createExcel();
  excel.rename('Sheet1', 'Basis');

  Sheet basisSheet = excel.sheets[basis]!;
  basisSheet.setDefaultColumnWidth(12.0);

  basisSheet.appendRow(createHeaderRow());

  DAYOFTHEWEEK.values.forEachIndexed(
    (index, day) {
      basisSheet.appendRow(createBlankRow());
      createWeekDay(day).forEach((element) {
        basisSheet.appendRow(element);
      });

      merge(basisSheet, day, 'A');
      merge(basisSheet, day, 'R');
      setStyleForMiddlePart(basisSheet, day);
    },
  );

  for (var element in events) {
    insertEvent(basisSheet, element);
  }

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

  int numberOfRows = dayOfTheWeekMapping[day]!.numberOfRows;
  for (int rowIndex = 0; rowIndex < numberOfRows; rowIndex++) {
    cellValues = [];

    String formatedTime = getDateRangeFormating(day, rowIndex);

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

const Map<DAYOFTHEWEEK, DayProps> dayOfTheWeekMapping = {
  DAYOFTHEWEEK.montag: DayProps(
    startingHour: 15,
    endingHour: 21,
    name: 'MO',
  ),
  DAYOFTHEWEEK.dienstag: DayProps(
    startingHour: 15,
    endingHour: 21,
    name: 'DI',
  ),
  DAYOFTHEWEEK.mittwoch: DayProps(
    startingHour: 15,
    endingHour: 21,
    name: 'MI',
  ),
  DAYOFTHEWEEK.donnerstag: DayProps(
    startingHour: 15,
    endingHour: 21,
    name: 'DO',
  ),
  DAYOFTHEWEEK.freitag: DayProps(
    startingHour: 15,
    endingHour: 21,
    name: 'FR',
  ),
  DAYOFTHEWEEK.samstag: DayProps(
    startingHour: 10,
    endingHour: 19,
    name: 'SA',
  ),
  DAYOFTHEWEEK.sonntag: DayProps(
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

void merge(Sheet sheet, DAYOFTHEWEEK day, String colName) {
  int start = startList[day.index];
  int end = start + dayOfTheWeekMapping[day]!.numberOfRows - 1;
  sheet.merge(
    CellIndex.indexByString('$colName$start'),
    CellIndex.indexByString('$colName$end'),
    customValue: TextCellValue(dayOfTheWeekMapping[day]!.name),
  );

  sheet.setMergedCellStyle(
      CellIndex.indexByString('$colName$start'), cellStyleWeek);
}

void setStyleForMiddlePart(Sheet sheet, DAYOFTHEWEEK day) {
  int start = startList[day.index];
  for (int colIndex = 1; colIndex < 17; colIndex++) {
    for (int rowIndex = 0;
        rowIndex < dayOfTheWeekMapping[day]!.numberOfRows;
        rowIndex++) {
      sheet
          .cell(
            CellIndex.indexByColumnRow(
              columnIndex: colIndex,
              rowIndex: start + rowIndex - 1,
            ),
          )
          .cellStyle = cellStyleMiddle;
    }
    sheet
        .cell(
          CellIndex.indexByColumnRow(
            columnIndex: colIndex,
            rowIndex: start - 1,
          ),
        )
        .cellStyle = cellStyleMiddle.copyWith(
      topBorderVal: Border(borderStyle: BorderStyle.Thin),
    );
    sheet
        .cell(
          CellIndex.indexByColumnRow(
            columnIndex: colIndex,
            rowIndex: start + dayOfTheWeekMapping[day]!.numberOfRows - 2,
          ),
        )
        .cellStyle = cellStyleMiddle.copyWith(
      bottomBorderVal: Border(borderStyle: BorderStyle.Thin),
    );
  }
}

CellStyle cellStyleMiddle = CellStyle(
  leftBorder: Border(borderStyle: BorderStyle.Thin),
  rightBorder: Border(borderStyle: BorderStyle.Thin),
  fontFamily: getFontFamily(FontFamily.Arial),
  fontSize: 10,
  verticalAlign: VerticalAlign.Center,
  horizontalAlign: HorizontalAlign.Center,
);

List<int> startList = startValues();

List<int> startValues() {
  int start = 3;
  List<int> startValues = [];
  DAYOFTHEWEEK.values.forEachIndexed((index, day) {
    startValues.add(start);
    start += dayOfTheWeekMapping[day]!.numberOfRows + 1;
  });
  return startValues;
}

CellStyle cellStyleBooking = cellStyleMiddle.copyWith(
  topBorderVal: Border(borderStyle: BorderStyle.Thin),
  bottomBorderVal: Border(borderStyle: BorderStyle.Thin),
);

insertEvent(Sheet sheet, Event event) {
  Point<int> rowIndex = getRowRangeFromEvent(event);
  int rowIndexStart = rowIndex.x;
  int rowIndexEnd = rowIndex.y;

  sheet.merge(
    CellIndex.indexByColumnRow(
        columnIndex: event.courts.start, rowIndex: rowIndexStart),
    CellIndex.indexByColumnRow(
        columnIndex: event.courts.end, rowIndex: rowIndexEnd),
    customValue: TextCellValue(event.name),
  );

  sheet.setMergedCellStyle(
      CellIndex.indexByColumnRow(
          columnIndex: event.courts.start, rowIndex: rowIndexStart),
      cellStyleBooking.copyWith(
          backgroundColorHexVal:
              ExcelColor.fromInt(eventColorMapping[event.type]!.value)));
}

String getDateRangeFormating(DAYOFTHEWEEK day, int rowIndex) {
  int startingHour = dayOfTheWeekMapping[day]!.startingHour + rowIndex ~/ 2;
  String fullHourMinutes = ':00';
  String halfHourMinutes = ':30';
  String formatedTime = '';
  if (rowIndex % 2 == 1) {
    formatedTime =
        '$startingHour$halfHourMinutes-${startingHour + 1}$fullHourMinutes';
    startingHour++;
  } else {
    formatedTime =
        '$startingHour$fullHourMinutes-$startingHour$halfHourMinutes';
  }
  return formatedTime;
}

Point<int> getRowRangeFromEvent(Event event) {
  int startHour = int.parse(event.start.substring(0, 2));
  int startHalfHour = int.parse(event.start.substring(3, 5));

  int endHour = int.parse(event.end.substring(0, 2));
  int endHalfHour = int.parse(event.end.substring(3, 5));

  DayProps dayProps = dayOfTheWeekMapping[event.day]!;

  int rowIndexStart, rowIndexEnd;
  rowIndexStart = startList[event.day.index] +
      (startHour - dayProps.startingHour) * 2 +
      (startHalfHour ~/ 30) -
      1;
  rowIndexEnd = startList[event.day.index] +
      (endHour - dayProps.startingHour) * 2 +
      (endHalfHour ~/ 30) -
      2;
  Point<int> rowRange = Point(
    rowIndexStart,
    rowIndexEnd,
  );
  return rowRange;
}
