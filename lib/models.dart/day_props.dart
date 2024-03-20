class DayProps {
  const DayProps(
      {required this.startingHour,
      required this.endingHour,
      required this.name});
  final String name;
  final int startingHour;
  final int endingHour;
  int get numberOfRows => (endingHour - startingHour) * 2;
}
