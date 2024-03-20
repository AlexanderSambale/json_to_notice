import 'package:flutter/material.dart';

class Event {
  const Event({
    required this.start,
    required this.end,
    required this.type,
    required this.name,
  });
  final String start;
  final String end;
  final EventType type;
  final String name;
}

enum EventType {
  sg,
  training,
  minitreff,
  turnier,
  unisport,
}

const Map<EventType, Color> eventColorMapping = {
  EventType.sg: Colors.green,
  EventType.training: Colors.blue,
  EventType.minitreff: Color.fromARGB(255, 234, 209, 220),
  EventType.turnier: Colors.orange,
  EventType.unisport: Color.fromARGB(255, 217, 210, 233),
};