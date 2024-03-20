import 'package:json_to_notices/utils.dart';

import 'models.dart/event.dart';

List<Event> events = [
  const Event(
    start: '16:00',
    end: '17:30',
    day: DAYOFTHEWEEK.donnerstag,
    type: EventType.sg,
    name: 'SG1',
    courts: (start: 3, end: 3),
  ),
  const Event(
    start: '16:00',
    end: '21:00',
    day: DAYOFTHEWEEK.mittwoch,
    type: EventType.minitreff,
    name: 'Minitreff 1',
    courts: (start: 5, end: 7),
  ),
  const Event(
    start: '12:00',
    end: '19:30',
    day: DAYOFTHEWEEK.sonntag,
    type: EventType.turnier,
    name: 'Turnier 1',
    courts: (start: 6, end: 13),
  ),
  const Event(
    start: '15:00',
    end: '17:30',
    day: DAYOFTHEWEEK.dienstag,
    type: EventType.training,
    name: 'Training 1',
    courts: (start: 4, end: 5),
  ),
  const Event(
    start: '11:00',
    end: '17:30',
    day: DAYOFTHEWEEK.samstag,
    type: EventType.unisport,
    name: 'Unisport 1',
    courts: (start: 12, end: 15),
  ),
];
