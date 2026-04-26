import 'dart:convert';
import 'dart:typed_data';

import 'package:excel/excel.dart' hide Border;
import 'package:file_picker/file_picker.dart';
import 'package:flutter/material.dart';

void main() => runApp(
  MaterialApp(
    home: const LogAnalyzer(),
    debugShowCheckedModeBanner: false,
    theme: ThemeData(fontFamily: 'NotoSansKR', useMaterial3: false),
  ),
);

enum EntryType { button, bypass, restore, info }

enum LogViewMode { summary, detail }

enum OperationMode { auto, driverless, manual, emergency, yard, unknown }

enum AnalysisBlockType {
  stationStop,
  running,
  interstationStop,
  overrunSuspectedStop,
}

const Duration kDefaultSampleDuration = Duration(milliseconds: 500);
const double kStopSpeedThresholdKmh = 0.1;
const double kDepartureSpeedThresholdKmh = 5.0;
const double kStationNearDistanceThreshold = 5.0;
const double kInterstationStopDistanceThreshold = 100.0;
const Duration kStopConfirmDuration = Duration(seconds: 2);
const Duration kContextWindow = Duration(seconds: 3);
const Duration kFsbrNCodeContextDuration = Duration(seconds: 2);
const Duration kNCodeWarnDelay = Duration(seconds: 4);
const Duration kNCodeSummaryInterval = Duration(seconds: 30);
const Duration kLongPersistenceDuration = Duration(seconds: 120);
const Duration kDoorCloseDelayWarnDuration = Duration(seconds: 5);
const String kLatestBuildDateLabel = '최신 반영: 2026-04-25';

const List<int> kEffectiveDoorCars = [0, 1, 2, 7];
const List<String> kActiveDoors = ['DOOR0', 'DOOR1', 'DOOR2', 'DOOR7'];

const Map<String, String> kSignalDisplayNames = {
  'ADC': 'ALL DOOR CLOSE',
  'OPEN-L': '좌측 문열림 버튼 취급',
  'OPEN-R': '우측 문열림 버튼 취급',
  'CLOSE': '문닫힘 버튼 취급',
  'REOPEN': '재개폐 버튼 취급',
  'S_OPEN-L': '사이드 좌측 문열림 버튼 취급',
  'S_OPEN-R': '사이드 우측 문열림 버튼 취급',
  'S_CLOSE': '사이드 문닫힘 버튼 취급',
  'ADBS': '전차 출입문 바이패스',
  'FSBR': '전상용제동(FSB) 체결',
  'DPT-PM': '출발 허가',
  'EMPB': '비상제동 버튼 취급',
  'EBCOS': '비상제동 차단 버튼',
  'AUTO': '자동운전',
  'DRIVL': '무인운전',
  'MANUAL': '수동운전',
  'EMERG': '비상운전',
  'EMEGR': '비상모드(ATC 제어 해제)',
  'YARD': '기지운전',
  'NCODE': '무코드 수신',
  'PAN UP': '판토상승 버튼 취급',
  'PAN DN': '판토하강 버튼 취급',
};

const Map<int, String> kStationNames = {
  100: '녹동역',
  101: '소태역',
  102: '학동·증심사입구역',
  103: '남광주역',
  104: '문화전당역',
  105: '금남로4가역',
  106: '금남로5가역',
  107: '양동시장역',
  108: '돌고개역',
  109: '농성역',
  110: '화정역',
  111: '쌍촌역',
  112: '운천역',
  113: '상무역',
  114: '김대중컨벤션센터역',
  115: '공항역',
  116: '송정공원역',
  117: '광주송정역',
  118: '도산역',
  119: '평동역',
};

class LogEntry {
  final String time;
  final String message;
  final EntryType type;
  final bool isSummary;

  const LogEntry({
    required this.time,
    required this.message,
    required this.type,
    this.isSummary = false,
  });
}

class AnalysisBlock {
  final String title;
  final AnalysisBlockType type;
  final List<LogEntry> entries;

  const AnalysisBlock({
    required this.title,
    required this.type,
    required this.entries,
  });
}

class _StopBlockSeed {
  final AnalysisBlockType type;
  final int startLogOffset;
  final int stopLogEndOffsetExclusive;
  final int departureLogOffset;
  final String? previousStationName;
  final String? approachStationName;
  final String? stationName;
  final String? nextStationName;
  final String? note;

  const _StopBlockSeed({
    required this.type,
    required this.startLogOffset,
    required this.stopLogEndOffsetExclusive,
    required this.departureLogOffset,
    required this.previousStationName,
    required this.approachStationName,
    required this.stationName,
    required this.nextStationName,
    this.note,
  });
}

class _ActiveStopContext {
  final int startIndex;
  final String startTime;
  final int startLogOffset;
  final String? previousStationName;
  final int? approachStationCode;
  final String? approachStationName;
  final double? approachDist;
  final int? startNextStaCode;
  final String? startNextStationName;
  final double? startDist;
  bool hadDoorActivity;
  bool hadAdcRelease;
  bool hadDoorCycle;
  bool hadNextStaChange = false;

  _ActiveStopContext({
    required this.startIndex,
    required this.startTime,
    required this.startLogOffset,
    required this.previousStationName,
    required this.approachStationCode,
    required this.approachStationName,
    required this.approachDist,
    required this.startNextStaCode,
    required this.startNextStationName,
    required this.startDist,
    this.hadDoorActivity = false,
    this.hadAdcRelease = false,
    this.hadDoorCycle = false,
  });
}

class LogAnalyzer extends StatefulWidget {
  const LogAnalyzer({super.key});

  @override
  State<LogAnalyzer> createState() => _LogAnalyzerState();
}

class _LogAnalyzerState extends State<LogAnalyzer> {
  List<LogEntry> logs = [];
  List<AnalysisBlock> blocks = [];
  String statusText = '운행기록 파일을 선택해 주세요. (.xlsx / .csv)';
  String summaryText = '';
  bool isLoading = false;
  LogViewMode viewMode = LogViewMode.detail;
  List<List<dynamic>>? _sourceRows;
  String _sourceFileSummary = '';
  final TextEditingController _startTimeController = TextEditingController();
  final TextEditingController _endTimeController = TextEditingController();

  @override
  void dispose() {
    _startTimeController.dispose();
    _endTimeController.dispose();
    super.dispose();
  }

  bool _isXlsFile(String fileName, Uint8List bytes) {
    if (fileName.toLowerCase().endsWith('.xls')) return true;
    return bytes.length > 4 &&
        bytes[0] == 0xD0 &&
        bytes[1] == 0xCF &&
        bytes[2] == 0x11 &&
        bytes[3] == 0xE0;
  }

  String _cellToStr(dynamic cell) {
    if (cell == null) return '';
    final cv = (cell is Data) ? cell.value : cell;
    if (cv == null) return '';

    if (cv is TextCellValue) return (cv.value.text ?? '').trim();
    if (cv is IntCellValue) return cv.value.toString().trim();
    if (cv is DoubleCellValue) return cv.value.toString().trim();
    if (cv is BoolCellValue) return cv.value ? '1' : '0';
    return cv.toString().trim();
  }

  String _getVal(List<dynamic> row, int idx) {
    if (idx == -1 || idx >= row.length || row[idx] == null) return '0';
    final cv = (row[idx] is Data) ? (row[idx] as Data).value : row[idx];
    if (cv == null) return '0';

    final s = (cv is TextCellValue)
        ? (cv.value.text ?? '').trim()
        : cv.toString().trim();
    return (int.tryParse(s) ?? 0) != 0 ? '1' : '0';
  }

  double _getDouble(List<dynamic> row, int idx) {
    if (idx == -1 || idx >= row.length) return 0;
    return double.tryParse(_cellToStr(row[idx])) ?? 0;
  }

  bool _isRisingEdge(List<dynamic> prev, List<dynamic> curr, int idx) {
    if (idx == -1) return false;
    return _getVal(prev, idx) == '0' && _getVal(curr, idx) == '1';
  }

  List<OperationMode> _activeModes(
    List<dynamic> row,
    Map<OperationMode, int> modeIndices,
  ) {
    return modeIndices.entries
        .where((entry) => entry.value != -1 && _getVal(row, entry.value) == '1')
        .map((entry) => entry.key)
        .toList();
  }

  OperationMode _determineOperationMode(
    List<dynamic> row,
    Map<OperationMode, int> modeIndices,
  ) {
    final activeModes = _activeModes(row, modeIndices);
    if (activeModes.length == 1) {
      return activeModes.first;
    }
    return OperationMode.unknown;
  }

  String _modeLabel(OperationMode mode) {
    switch (mode) {
      case OperationMode.auto:
        return '자동운전';
      case OperationMode.driverless:
        return '무인운전';
      case OperationMode.manual:
        return '수동운전';
      case OperationMode.emergency:
        return '비상운전';
      case OperationMode.yard:
        return '기지운전';
      case OperationMode.unknown:
        return '미상';
    }
  }

  String _modeTransitionMessage(OperationMode? prev, OperationMode curr) {
    if (prev == null) {
      return '${_modeLabel(curr)}으로 전환된 정황이 확인됩니다.';
    }
    return '${_modeLabel(prev)}에서 ${_modeLabel(curr)}으로 전환된 정황이 확인됩니다.';
  }

  Duration? _parseLogTime(String raw) {
    final value = raw.trim();
    if (value.isEmpty) return null;

    final hhmmss = RegExp(
      r'^(\d{1,2}):(\d{2}):(\d{2})(?:\.(\d{1,3}))?$',
    ).firstMatch(value);
    if (hhmmss != null) {
      final milliText = (hhmmss.group(4) ?? '').padRight(3, '0');
      return Duration(
        hours: int.parse(hhmmss.group(1)!),
        minutes: int.parse(hhmmss.group(2)!),
        seconds: int.parse(hhmmss.group(3)!),
        milliseconds: milliText.isEmpty ? 0 : int.parse(milliText),
      );
    }

    final compact = RegExp(
      r'^(\d{2})(\d{2})(\d{2})(\d{0,3})$',
    ).firstMatch(value);
    if (compact != null) {
      final milliText = (compact.group(4) ?? '').padRight(3, '0');
      return Duration(
        hours: int.parse(compact.group(1)!),
        minutes: int.parse(compact.group(2)!),
        seconds: int.parse(compact.group(3)!),
        milliseconds: milliText.isEmpty ? 0 : int.parse(milliText),
      );
    }

    return null;
  }

  String? _validateFilterTime(String raw) {
    final value = raw.trim();
    if (value.isEmpty) return null;
    return _parseLogTime(value) == null ? '시간 형식은 HH:mm:ss로 입력해 주세요.' : null;
  }

  String _timeFilterLabel(Duration? start, Duration? end) {
    String format(Duration value) {
      String two(int number) => number.toString().padLeft(2, '0');
      final hours = two(value.inHours);
      final minutes = two(value.inMinutes.remainder(60));
      final seconds = two(value.inSeconds.remainder(60));
      return '$hours:$minutes:$seconds';
    }

    if (start == null && end == null) return '전체 시간';
    if (start != null && end != null) return '${format(start)}~${format(end)}';
    if (start != null) return '${format(start)} 이후';
    return '${format(end!)} 이전';
  }

  Future<List<List<dynamic>>> _parseLogFile(PlatformFile file) async {
    final bytes = file.bytes;
    if (bytes == null) {
      throw '${file.name} 파일을 읽을 수 없습니다.';
    }
    if (_isXlsFile(file.name, bytes)) {
      throw '.xls 형식은 지원하지 않습니다. .xlsx 또는 .csv로 변환해 주세요.';
    }

    final rows = <List<dynamic>>[];
    if (bytes.length > 2 && bytes[0] == 0x50 && bytes[1] == 0x4B) {
      final excel = Excel.decodeBytes(bytes);
      if (excel.tables.isEmpty) return rows;
      final table = excel.tables[excel.tables.keys.first]!;
      for (final row in table.rows) {
        rows.add(row.toList());
      }
    } else {
      final lines = utf8
          .decode(bytes, allowMalformed: true)
          .split(RegExp(r'\r?\n'))
          .where((line) => line.trim().isNotEmpty);
      for (final line in lines) {
        rows.add(line.split(','));
      }
    }
    return rows;
  }

  List<String> _normalizedHeader(List<dynamic> row) {
    return row.map((cell) => _cellToStr(cell).toUpperCase()).toList();
  }

  List<List<dynamic>> _mergeRows(List<List<List<dynamic>>> files) {
    if (files.isEmpty || files.any((rows) => rows.isEmpty)) {
      throw '데이터가 비어 있습니다.';
    }

    final baseHeader = _normalizedHeader(files.first.first);
    const requiredColumns = ['TIME', 'NEXTSTA', 'DIST', 'VEL', 'ADC'];
    for (final column in requiredColumns) {
      if (!baseHeader.contains(column)) {
        throw '필수 컬럼($column)이 없어 분석할 수 없습니다.';
      }
    }

    final mergedDataRows = <List<dynamic>>[];
    for (final rows in files) {
      final header = _normalizedHeader(rows.first);
      if (header.toSet().length != baseHeader.toSet().length ||
          !header.toSet().containsAll(baseHeader)) {
        throw '두 파일의 운행기록 형식이 달라 병합할 수 없습니다.';
      }

      final order = baseHeader.map((name) => header.indexOf(name)).toList();
      for (final row in rows.skip(1)) {
        mergedDataRows.add([
          for (final idx in order) idx >= 0 && idx < row.length ? row[idx] : '',
        ]);
      }
    }

    final timeIndex = baseHeader.indexOf('TIME');
    mergedDataRows.sort((a, b) {
      final aTime = _parseLogTime(_cellToStr(a[timeIndex])) ?? Duration.zero;
      final bTime = _parseLogTime(_cellToStr(b[timeIndex])) ?? Duration.zero;
      return aTime.compareTo(bTime);
    });

    return [files.first.first, ...mergedDataRows];
  }

  List<List<dynamic>> _filterRowsByTime(
    List<List<dynamic>> rows,
    Duration? start,
    Duration? end,
  ) {
    if (start == null && end == null) return rows;
    if (rows.isEmpty) return rows;

    final header = _normalizedHeader(rows.first);
    final timeIndex = header.indexOf('TIME');
    if (timeIndex == -1) {
      throw 'TIME 컬럼이 없어 시간대 필터를 적용할 수 없습니다.';
    }

    final filtered = <List<dynamic>>[rows.first];
    for (final row in rows.skip(1)) {
      if (timeIndex >= row.length) continue;
      final time = _parseLogTime(_cellToStr(row[timeIndex]));
      if (time == null) continue;
      if (start != null && time < start) continue;
      if (end != null && time > end) continue;
      filtered.add(row);
    }
    return filtered;
  }

  Duration _resolveSampleDelta(String prevTime, String currTime) {
    final prev = _parseLogTime(prevTime);
    final curr = _parseLogTime(currTime);
    if (prev != null && curr != null && curr >= prev) {
      final diff = curr - prev;
      if (diff > Duration.zero && diff <= kLongPersistenceDuration) {
        return diff;
      }
    }
    return kDefaultSampleDuration;
  }

  int? _normalizeNextStaCode(String raw) {
    final value = int.tryParse(raw.trim());
    if (value == null) return null;
    if (value >= 1 && value <= 19) return value + 100;
    switch (value) {
      case 40:
        return 140;
      case 41:
        return 141;
      case 42:
        return 142;
      case 43:
        return 143;
      case 50:
        return 150;
      case 65:
        return 165;
      case 66:
        return 166;
      default:
        return value;
    }
  }

  String? _currentStationNameFromNext(int? nextSta, int direction) {
    if (nextSta == null || direction == 0) return null;
    return kStationNames[nextSta - direction];
  }

  String? _nextStationDisplay(int? nextSta) {
    if (nextSta == null) return null;
    return kStationNames[nextSta] ?? '다음역 코드 $nextSta';
  }

  String? _formatStopLocationText({
    required int? nextSta,
    required int direction,
    required String? lastBerthedStation,
  }) {
    final nextStation = _nextStationDisplay(nextSta);
    final currentStation =
        lastBerthedStation ?? _currentStationNameFromNext(nextSta, direction);

    if (currentStation != null && nextStation != null) {
      return '$currentStation 정차 중 / 다음역 $nextStation';
    }
    if (currentStation != null) {
      return '$currentStation 정차 중으로 해석됩니다.';
    }
    if (nextStation != null) {
      return '정차 중이며 다음역 $nextStation 정보가 확인됩니다.';
    }
    return null;
  }

  String? _formatDepartureLocationText({
    required String? lastBerthedStation,
    required int? nextSta,
  }) {
    final nextStation = _nextStationDisplay(nextSta);
    if (lastBerthedStation != null && nextStation != null) {
      return '$lastBerthedStation 발차 후 $nextStation 방면으로 해석됩니다.';
    }
    if (nextStation != null) {
      return '$nextStation 방면으로 운행 중인 정황이 확인됩니다.';
    }
    return null;
  }

  String _formatDoorName(String doorSignal) {
    final carNo = doorSignal.replaceFirst('DOOR', '');
    return '$carNo호차 출입문';
  }

  bool _isAnyRisingEdge(
    List<dynamic> prev,
    List<dynamic> curr,
    List<int> indices,
  ) {
    for (final idx in indices) {
      if (_isRisingEdge(prev, curr, idx)) {
        return true;
      }
    }
    return false;
  }

  List<String> _uncClosedDoorCars(List<dynamic> row, Map<String, int> doorIdx) {
    final cars = <String>[];
    for (final entry in doorIdx.entries) {
      if (entry.value != -1 && _getVal(row, entry.value) == '0') {
        cars.add('${entry.key.replaceFirst('DOOR', '')}호차');
      }
    }
    return cars;
  }

  String _doorBypassTargetLabel(
    List<dynamic> prev,
    List<dynamic> curr,
    Map<String, int> doorIdx,
  ) {
    final prevCars = _uncClosedDoorCars(prev, doorIdx);
    final cars = prevCars.isNotEmpty ? prevCars : _uncClosedDoorCars(curr, doorIdx);
    return cars.isEmpty ? '대상 불명' : cars.join(', ');
  }

  String _doorBypassReportMessage(
    String target,
    String adcValue, {
    String? suffix,
  }) {
    return '기관사 출입문 바이패스 취급($target), ${_allDoorCloseText(adcValue)}${suffix ?? ''}';
  }

  String? _doorLocationPrefix({
    required bool isStopped,
    required String? lastBerthedStation,
    required int? nextSta,
    required int direction,
  }) {
    if (!isStopped) return null;
    final currentStation =
        lastBerthedStation ?? _currentStationNameFromNext(nextSta, direction);
    if (currentStation != null) {
      return '$currentStation 정차 중';
    }
    return '정차 중';
  }

  String? _stationNameFromLocation(String? location) {
    if (location == null || location.trim().isEmpty) return null;
    final value = location.trim();
    if (value == '정차 중') return null;
    final stopIndex = value.indexOf(' 정차 중');
    if (stopIndex != -1) {
      return value.substring(0, stopIndex);
    }
    return value;
  }

  String _currentStationSuffix(String? location) {
    final station = _stationNameFromLocation(location);
    return station == null ? ' (현재 역: 미확인)' : ' (현재 역: $station)';
  }

  String _allDoorCloseText(String adcValue) => 'ALL DOOR CLOSE "$adcValue"';

  String _doorButtonActionLabel(String signal) {
    switch (signal) {
      case 'OPEN-L':
        return '기관사 좌측 출입문 열림 버튼 취급';
      case 'OPEN-R':
        return '기관사 우측 출입문 열림 버튼 취급';
      case 'CLOSE':
        return '기관사 닫힘버튼 취급';
      case 'REOPEN':
        return '기관사 재개폐 버튼 취급';
      case 'S_OPEN-L':
        return '사이드 좌측 출입문 열림 버튼 취급';
      case 'S_OPEN-R':
        return '사이드 우측 출입문 열림 버튼 취급';
      case 'S_CLOSE':
        return '사이드 출입문 닫힘 버튼 취급';
      default:
        return _signalName(signal);
    }
  }

  String _doorModeLabel(List<dynamic> row, Map<String, int> doorModeIndices) {
    final labels = <String>[];
    if (_getVal(row, doorModeIndices['AO/AC'] ?? -1) == '1') {
      labels.add('자/자');
    }
    if (_getVal(row, doorModeIndices['AO/MC'] ?? -1) == '1') {
      labels.add('자/수');
    }
    if (_getVal(row, doorModeIndices['MO/MC'] ?? -1) == '1') {
      labels.add('수/수');
    }
    if (labels.length > 1) {
      return '출입문 모드 복수 신호';
    }
    if (labels.isEmpty) {
      return '출입문 모드 미확인';
    }
    return labels.first;
  }

  String _signalName(String key) => kSignalDisplayNames[key] ?? key;

  bool _isBlockRelevantLog(LogEntry entry) {
    return entry.isSummary ||
        entry.type == EntryType.button ||
        entry.type == EntryType.bypass;
  }

  String _trimBlockMessage(String message, AnalysisBlockType type) {
    if (type == AnalysisBlockType.stationStop ||
        type == AnalysisBlockType.overrunSuspectedStop) {
      const marker = '정차 중 ';
      final index = message.indexOf(marker);
      if (index != -1) {
        return message.substring(index + marker.length);
      }
    }
    return message;
  }

  String _runningBlockTitle(String? fromStation, String? toStation) {
    if (fromStation != null && toStation != null && fromStation != toStation) {
      return '[$fromStation → $toStation 주행 구간]';
    }
    if (toStation != null) {
      return '[$toStation 방면 주행 구간]';
    }
    if (fromStation != null) {
      return '[$fromStation 출발 후 주행 구간]';
    }
    return '[역간 주행 구간]';
  }

  String _stationStopBlockTitle(String? stationName) {
    return '[${stationName ?? '역'} 정차 구간]';
  }

  String _interstationStopBlockTitle(
    String? previousStation,
    String? approachStation,
    String? nextStation,
  ) {
    if (approachStation != null) {
      return '[$approachStation 진입 전 역간 정차 구간]';
    }
    if (previousStation != null && nextStation != null) {
      return '[$previousStation → $nextStation 역간 정차 구간]';
    }
    return '[역간 정차 구간]';
  }

  String _overrunStopBlockTitle(String? stationName) {
    return '[${stationName ?? '접근'} 접근 후 정차 구간]';
  }

  String? _resolveStopStationName({
    required String? approachStationName,
    required String? currentNextStationName,
    required String? lastBerthedStation,
    required double? stopDist,
  }) {
    if (approachStationName != null) {
      return approachStationName;
    }
    if (stopDist != null && stopDist <= kStationNearDistanceThreshold) {
      return currentNextStationName;
    }
    return lastBerthedStation;
  }

  List<AnalysisBlock> _buildBlocks(
    List<LogEntry> sourceLogs,
    List<_StopBlockSeed> stopSeeds, {
    String? tailApproachStationName,
  }) {
    final filteredLogs = sourceLogs.where(_isBlockRelevantLog).toList();
    if (filteredLogs.isEmpty) {
      return const [];
    }

    final blocks = <AnalysisBlock>[];
    int cursor = 0;
    String? lastStationName;

    for (final stop in stopSeeds) {
      final safeStart = stop.startLogOffset.clamp(0, sourceLogs.length);
      final safeStopEnd = stop.stopLogEndOffsetExclusive.clamp(
        0,
        sourceLogs.length,
      );
      final safeDeparture = stop.departureLogOffset.clamp(0, sourceLogs.length);

      final runningEntries = sourceLogs
          .sublist(cursor, safeStart)
          .where(_isBlockRelevantLog)
          .toList();
      final runningTitle = _runningBlockTitle(
        stop.previousStationName ?? lastStationName,
        stop.approachStationName ?? stop.stationName ?? stop.nextStationName,
      );
      if (runningEntries.isNotEmpty) {
        blocks.add(
          AnalysisBlock(
            title: runningTitle,
            type: AnalysisBlockType.running,
            entries: runningEntries,
          ),
        );
      }

      final stopEntries = sourceLogs
          .sublist(safeStart, safeStopEnd)
          .where(_isBlockRelevantLog)
          .map(
            (entry) => LogEntry(
              time: entry.time,
              message: _trimBlockMessage(entry.message, stop.type),
              type: entry.type,
              isSummary: entry.isSummary,
            ),
          )
          .toList();
      if (stop.note != null) {
        stopEntries.insert(
          0,
          LogEntry(
            time: '-',
            message: stop.note!,
            type: EntryType.info,
            isSummary: true,
          ),
        );
      }

      if (stopEntries.isNotEmpty) {
        final title = switch (stop.type) {
          AnalysisBlockType.stationStop => _stationStopBlockTitle(
            stop.stationName,
          ),
          AnalysisBlockType.running => _runningBlockTitle(
            stop.previousStationName,
            stop.stationName ?? stop.nextStationName,
          ),
          AnalysisBlockType.interstationStop => _interstationStopBlockTitle(
            stop.previousStationName,
            stop.approachStationName,
            stop.nextStationName,
          ),
          AnalysisBlockType.overrunSuspectedStop => _overrunStopBlockTitle(
            stop.stationName ?? stop.approachStationName,
          ),
        };
        blocks.add(
          AnalysisBlock(title: title, type: stop.type, entries: stopEntries),
        );
      }

      cursor = safeDeparture;
      if (stop.type != AnalysisBlockType.interstationStop &&
          stop.stationName != null) {
        lastStationName = stop.stationName;
      }
    }

    final tailEntries = sourceLogs
        .sublist(cursor)
        .where(_isBlockRelevantLog)
        .toList();
    if (tailEntries.isNotEmpty) {
      blocks.add(
        AnalysisBlock(
          title: _runningBlockTitle(lastStationName, tailApproachStationName),
          type: AnalysisBlockType.running,
          entries: tailEntries,
        ),
      );
    }

    if (blocks.isEmpty) {
      return [
        AnalysisBlock(
          title: '[블록 요약]',
          type: AnalysisBlockType.running,
          entries: filteredLogs,
        ),
      ];
    }
    return blocks;
  }

  Color _blockColor(AnalysisBlockType type) {
    switch (type) {
      case AnalysisBlockType.stationStop:
        return Colors.teal;
      case AnalysisBlockType.running:
        return Colors.indigo;
      case AnalysisBlockType.interstationStop:
        return Colors.orange;
      case AnalysisBlockType.overrunSuspectedStop:
        return Colors.deepOrange;
    }
  }

  IconData _blockIcon(AnalysisBlockType type) {
    switch (type) {
      case AnalysisBlockType.stationStop:
        return Icons.train_outlined;
      case AnalysisBlockType.running:
        return Icons.timeline;
      case AnalysisBlockType.interstationStop:
        return Icons.pause_circle_outline;
      case AnalysisBlockType.overrunSuspectedStop:
        return Icons.location_searching_outlined;
    }
  }

  List<String> _validTimes(Iterable<LogEntry> entries) {
    return entries
        .map((entry) => entry.time.trim())
        .where((time) => time.isNotEmpty && time != '-')
        .toList();
  }

  String _entryTimeRange(Iterable<LogEntry> entries) {
    final times = _validTimes(entries);
    if (times.isEmpty) return '';
    final first = times.first;
    final last = times.last;
    return first == last ? first : '$first ~ $last';
  }

  int _countMessages(List<LogEntry> entries, String keyword) {
    return entries.where((entry) => entry.message.contains(keyword)).length;
  }

  String _summaryTimeSuffix(List<LogEntry> entries) {
    final range = _entryTimeRange(entries);
    return range.isEmpty ? '' : ' 주요 시간대: $range.';
  }

  bool _hasAdcDoorMismatch(
    List<dynamic> row,
    int adcIdx,
    Map<String, int> doorIdx,
  ) {
    if (adcIdx == -1 || _getVal(row, adcIdx) != '1') return false;
    return doorIdx.values.any((idx) => idx != -1 && _getVal(row, idx) == '0');
  }

  bool _isRecent(Duration now, Duration? target, Duration window) {
    if (target == null) return false;
    return now >= target && now - target <= window;
  }

  String _buildSummary(List<LogEntry> entries) {
    if (entries.isEmpty) {
      return '감지된 주요 이벤트가 없습니다.';
    }

    final messages = entries.map((entry) => entry.message).join('\n');
    final hasDoor = messages.contains('ALL DOOR CLOSE') || messages.contains('출입문');
    final hasDoorCloseCycle =
        messages.contains('ALL DOOR CLOSE') ||
        messages.contains('기관사 닫힘버튼 취급');
    final hasMode = messages.contains('전환된 정황');
    final hasFsbr = messages.contains('전상용제동(FSB) 체결');
    final hasNCode = messages.contains('무코드');
    final hasStop =
        messages.contains('정차 중') || messages.contains('정차가 감지되었습니다.');
    final hasDeparture =
        messages.contains('발차') ||
        messages.contains('방면으로 해석됩니다.') ||
        messages.contains('출발 허가');
    final nCodeCount = _countMessages(entries, '무코드');
    final fsbrCount = _countMessages(entries, '전상용제동(FSB) 체결');
    final modeCount = _countMessages(entries, '전환된 정황');
    final doorCloseCycleCount = _countMessages(entries, 'ALL DOOR CLOSE');
    final timeSuffix = _summaryTimeSuffix(entries);

    if (hasNCode && hasFsbr && hasMode) {
      return '무코드 수신 $nCodeCount건, 전상용제동(FSB) 체결 $fsbrCount건, 운전모드 전환 $modeCount건이 확인됩니다.$timeSuffix';
    }
    if (hasDoorCloseCycle && hasDeparture) {
      return '출입문 닫힘 사이클 $doorCloseCycleCount건과 출발 흐름이 확인됩니다.$timeSuffix';
    }
    if (hasDoorCloseCycle) {
      return '출입문 닫힘 버튼 취급 이후 ALL DOOR CLOSE 형성 흐름이 $doorCloseCycleCount건 확인됩니다.$timeSuffix';
    }
    if (hasNCode && hasFsbr) {
      return '무코드 수신 $nCodeCount건 이후 전상용제동(FSB) 체결 흐름이 확인됩니다.$timeSuffix';
    }
    if (hasDoor && hasDeparture) {
      return '출입문 닫힘 상태 변화와 발차 흐름이 함께 확인됩니다.$timeSuffix';
    }
    if (hasDoor) {
      return '출입문 닫힘 상태 변화 정황이 확인됩니다.$timeSuffix';
    }
    if (hasMode && hasFsbr) {
      return '운전모드 전환 $modeCount건 이후 전상용제동(FSB) 체결 흐름이 확인됩니다.$timeSuffix';
    }
    if (hasMode) {
      return '운전모드 전환 흐름이 $modeCount건 확인됩니다.$timeSuffix';
    }
    if (hasFsbr) {
      return '전상용제동(FSB) 체결 정황이 $fsbrCount건 확인됩니다.$timeSuffix';
    }
    if (hasStop && hasDeparture) {
      return '정차와 발차 흐름이 확인됩니다.$timeSuffix';
    }
    return '주요 신호 변화 흐름이 확인됩니다.$timeSuffix';
  }

  Future<void> _pickAndAnalyze() async {
    final result = await FilePicker.pickFiles(
      type: FileType.custom,
      allowedExtensions: ['xlsx', 'csv', 'xls'],
      allowMultiple: true,
      withData: true,
    );
    if (result == null) return;
    if (result.files.length > 2) {
      setState(() {
        statusText = '운행기록 파일은 최대 2개까지 선택할 수 있습니다.';
      });
      return;
    }

    _startTimeController.clear();
    _endTimeController.clear();
    setState(() {
      isLoading = true;
      logs = [];
      blocks = [];
      summaryText = '';
      _sourceRows = null;
      _sourceFileSummary = '';
      statusText = result.files.length == 2
          ? '파일 2개 병합 중...'
          : '전동차 운행로그를 분석 중입니다...';
    });

    try {
      final parsedFiles = <List<List<dynamic>>>[];
      final rowCounts = <String>[];
      for (final file in result.files) {
        final rows = await _parseLogFile(file);
        parsedFiles.add(rows);
        rowCounts.add('${file.name} ${rows.isNotEmpty ? rows.length - 1 : 0}행');
      }
      final mergedRows = _mergeRows(parsedFiles);
      _sourceRows = mergedRows;
      _sourceFileSummary = result.files.length == 2
          ? '${rowCounts.join(' / ')} | 병합 후 ${mergedRows.length - 1}행'
            : rowCounts.first;
      await _analyzeRows(mergedRows, filterLabel: '전체 시간');
    } catch (e) {
      setState(() {
        isLoading = false;
        statusText = e.toString().contains('.xls')
            ? '오류: .xls 형식은 지원하지 않습니다. .xlsx 또는 .csv로 변환해 주세요.'
            : '오류: $e';
      });
      if (e.toString().contains('.xls')) {
        _showXlsWarning();
      }
    }
  }

  Future<void> _applyTimeFilter() async {
    final sourceRows = _sourceRows;
    if (sourceRows == null) return;

    final startError = _validateFilterTime(_startTimeController.text);
    final endError = _validateFilterTime(_endTimeController.text);
    if (startError != null || endError != null) {
      setState(() {
        statusText = startError ?? endError!;
      });
      return;
    }

    final start = _parseLogTime(_startTimeController.text);
    final end = _parseLogTime(_endTimeController.text);
    if (start != null && end != null && start > end) {
      setState(() {
        statusText = '시작 시간이 종료 시간보다 늦습니다.';
      });
      return;
    }

    setState(() {
      isLoading = true;
      logs = [];
      blocks = [];
      summaryText = '';
      statusText = '선택한 시간대를 다시 분석 중입니다...';
    });

    try {
      final rows = _filterRowsByTime(sourceRows, start, end);
      if (rows.length <= 1) {
        setState(() {
          isLoading = false;
          logs = [];
          blocks = [];
          summaryText = '';
          statusText = '선택한 시간대에 해당하는 운행기록이 없습니다.';
        });
        return;
      }
      await _analyzeRows(rows, filterLabel: _timeFilterLabel(start, end));
    } catch (e) {
      setState(() {
        isLoading = false;
        statusText = '오류: $e';
      });
    }
  }

  Future<void> _analyzeRows(
    List<List<dynamic>> rows, {
    required String filterLabel,
  }) async {
    if (rows.isEmpty) {
      throw '데이터가 비어 있습니다.';
    }

      final header = rows.first
          .map((cell) => _cellToStr(cell).toUpperCase())
          .toList();

      int findIdx(String name) {
        final target = name.toUpperCase();
        final exact = header.indexOf(target);
        return exact != -1
            ? exact
            : header.indexWhere((h) => h.contains(target));
      }

      List<int> findIndices(List<String> names) {
        final seen = <int>{};
        final indices = <int>[];
        for (final name in names) {
          final idx = findIdx(name);
          if (idx != -1 && seen.add(idx)) {
            indices.add(idx);
          }
        }
        return indices;
      }

      final idxTime = findIdx('TIME');
      final idxVel = findIdx('VEL');
      final idxAdc = findIdx('ADC');
      final idxAdbs = findIdx('ADBS');
      final idxDist = findIdx('DIST');
      final idxNextSta = findIdx('NEXTSTA');
      final idxNCode = findIdx('NCODE');
      final idxFsbr = findIdx('FSBR');
      final idxFsb = idxFsbr != -1 ? idxFsbr : findIdx('FSB');
      final dptPmIndices = findIndices(['ATC1DPT-PM', 'ATC2DPT-PM', 'DPT-PM']);
      final spIndices = findIndices(['ATC1S&P', 'ATC2S&P', 'S&P']);
      final idxEmegr = findIdx('EMEGR');
      final modeIndices = <OperationMode, int>{
        OperationMode.auto: findIdx('AUTO'),
        OperationMode.driverless: findIdx('DRIVL'),
        OperationMode.manual: findIdx('MANUAL'),
        OperationMode.emergency: findIdx('EMERG'),
        OperationMode.yard: findIdx('YARD'),
      };
      final doorModeIndices = <String, int>{
        'AO/AC': findIdx('AO/AC'),
        'AO/MC': findIdx('AO/MC'),
        'MO/MC': findIdx('MO/MC'),
      };

      final doorButtons = <String, int>{
        'OPEN-L': findIdx('OPEN-L'),
        'OPEN-R': findIdx('OPEN-R'),
        'CLOSE': findIdx('CLOSE'),
        'REOPEN': findIdx('REOPEN'),
        'S_OPEN-L': findIdx('S_OPEN-L'),
        'S_OPEN-R': findIdx('S_OPEN-R'),
        'S_CLOSE': findIdx('S_CLOSE'),
      };
      final autoDoorOpenCommands = <String, List<int>>{
        '좌측': findIndices(['ATC1/2OPEN-L', 'ATC1OPEN-L', 'ATC2OPEN-L']),
        '우측': findIndices(['ATC1/2OPEN-R', 'ATC1OPEN-R', 'ATC2OPEN-R']),
      };

      final eventButtons = <String, int>{
        'PAN UP': findIdx('PAN UP'),
        'PAN DN': findIdx('PAN DN'),
        'EMPB': findIdx('EMPB'),
        'EBCOS': findIdx('EBCOS'),
      };

      final doorIdx = {for (final door in kActiveDoors) door: findIdx(door)};
      final tempLogs = <LogEntry>[];

      final firstDataRow = rows.length > 1 ? rows[1] : const <dynamic>[];
      bool adcOpen = idxAdc != -1 && _getVal(firstDataRow, idxAdc) == '0';
      bool adcDoorMismatchActive = _hasAdcDoorMismatch(
        firstDataRow,
        idxAdc,
        doorIdx,
      );
      bool emegrActive =
          idxEmegr != -1 && _getVal(firstDataRow, idxEmegr) == '1';
      bool isStopped = false;
      bool nCodeActive =
          idxNCode != -1 && _getVal(firstDataRow, idxNCode) == '1';
      bool nCodeWarnIssued = false;
      int nextNCodeSummarySecond = kNCodeSummaryInterval.inSeconds;
      Duration nCodeDuration = nCodeActive
          ? kDefaultSampleDuration
          : Duration.zero;
      Duration zeroDuration = Duration.zero;
      Duration movingDuration = Duration.zero;
      Duration rowClock = Duration.zero;
      Duration? lastDoorContextAt;
      Duration? lastModeChangeAt;
      String? departureStartTime;
      String? lastBerthedStation;
      String? lastApproachStationName;
      int? lastApproachStationCode;
      double? lastApproachDist;
      String? lastConfirmedStationName;
      _ActiveStopContext? activeStop;
      final stopSeeds = <_StopBlockSeed>[];
      int? lastNextStaCode = idxNextSta != -1 && firstDataRow.isNotEmpty
          ? _normalizeNextStaCode(_cellToStr(firstDataRow[idxNextSta]))
          : null;
      int travelDirection = 0;

      final initialMode = _determineOperationMode(firstDataRow, modeIndices);
      OperationMode? lastKnownMode = initialMode == OperationMode.unknown
          ? null
          : initialMode;
      bool modeConflictActive =
          _activeModes(firstDataRow, modeIndices).length > 1;
      int? lastDoorCloseIndex;
      String? lastDoorCloseMode;
      List<String> lastDoorCloseUnclosedCars = const [];
      Duration? lastDoorCloseAt;
      Duration? lastDoorCycleClosedAt;
      int? lastDptLoggedCloseIndex;
      Duration? lastModeChangeDptConsumedAt;
      bool doorClosePending = false;
      bool doorCloseDelayLogged = false;
      bool doorCloseReopenLogged = false;
      bool doorCloseBypassLogged = false;
      bool lastDoorCycleHadDelay = false;
      String lastDoorBypassTarget = '대상 불명';

      for (var i = 1; i < rows.length; i++) {
        final curr = rows[i];
        if (curr.length < 2) continue;

        final prev = i > 1 ? rows[i - 1] : curr;
        final rowLogStartIndex = tempLogs.length;
        final tIdx = idxTime != -1 ? idxTime : 1;
        final time = tIdx < curr.length ? _cellToStr(curr[tIdx]) : '-';
        if (time.isEmpty) continue;

        final prevTime = i > 1 && tIdx < prev.length
            ? _cellToStr(prev[tIdx])
            : time;
        final sampleDelta = i > 1
            ? _resolveSampleDelta(prevTime, time)
            : Duration.zero;
        rowClock += sampleDelta;

        final speed = _getDouble(curr, idxVel);
        final currentDist = idxDist != -1 && idxDist < curr.length
            ? _getDouble(curr, idxDist)
            : null;
        final nextStaCode = idxNextSta != -1 && idxNextSta < curr.length
            ? _normalizeNextStaCode(_cellToStr(curr[idxNextSta]))
            : null;

        if (nextStaCode != null &&
            lastNextStaCode != null &&
            nextStaCode != lastNextStaCode) {
          travelDirection = nextStaCode > lastNextStaCode ? 1 : -1;
        }
        lastNextStaCode = nextStaCode ?? lastNextStaCode;

        if (speed > kStopSpeedThresholdKmh && nextStaCode != null) {
          final shouldUpdateApproach =
              lastApproachStationCode != nextStaCode ||
              currentDist == null ||
              lastApproachDist == null ||
              currentDist <= lastApproachDist;
          if (shouldUpdateApproach) {
            lastApproachStationCode = nextStaCode;
            lastApproachStationName = _nextStationDisplay(nextStaCode);
            lastApproachDist = currentDist;
          }
        }

        final prevAdcValue = idxAdc != -1
            ? _getVal(prev, idxAdc)
            : (adcOpen ? '0' : '1');
        final currAdcValue = idxAdc != -1
            ? _getVal(curr, idxAdc)
            : prevAdcValue;
        bool rowHadDoorActivity = prevAdcValue == '1' && currAdcValue == '0';
        bool rowHadDoorCycle = false;
        final doorLocation = _doorLocationPrefix(
          isStopped: isStopped,
          lastBerthedStation: lastBerthedStation,
          nextSta: nextStaCode,
          direction: travelDirection,
        );
        final currentDoorMode = _doorModeLabel(curr, doorModeIndices);

        final activeModes = _activeModes(curr, modeIndices);
        final hasModeConflict = activeModes.length > 1;
        if (hasModeConflict && !modeConflictActive) {
          tempLogs.add(
            LogEntry(
              time: time,
              message:
                  '복수 운전모드 신호가 동시에 감지되었습니다: ${activeModes.map(_modeLabel).join(', ')}',
              type: EntryType.bypass,
              isSummary: true,
            ),
          );
        } else if (!hasModeConflict && modeConflictActive) {
          tempLogs.add(
            LogEntry(
              time: time,
              message: '복수 운전모드 신호 상태가 해소된 정황이 확인됩니다.',
              type: EntryType.restore,
              isSummary: true,
            ),
          );
        }
        modeConflictActive = hasModeConflict;

        final currentMode = _determineOperationMode(curr, modeIndices);
        if (currentMode != OperationMode.unknown) {
          if (lastKnownMode != null && lastKnownMode != currentMode) {
            tempLogs.add(
              LogEntry(
                time: time,
                message: _modeTransitionMessage(lastKnownMode, currentMode),
                type: EntryType.info,
                isSummary: true,
              ),
            );
            lastModeChangeAt = rowClock;
          }
          lastKnownMode = currentMode;
        }

        if (idxEmegr != -1) {
          final currentEmegr = _getVal(curr, idxEmegr) == '1';
          if (currentEmegr != emegrActive) {
            tempLogs.add(
              LogEntry(
                time: time,
                message: currentEmegr
                    ? 'ATC 제어가 해제된 비상모드 상태가 감지되었습니다.'
                    : '비상모드 상태가 해제된 정황이 확인됩니다.',
                type: currentEmegr ? EntryType.bypass : EntryType.restore,
                isSummary: true,
              ),
            );
          }
          emegrActive = currentEmegr;
        }

        for (final entry in eventButtons.entries) {
          if (_isRisingEdge(prev, curr, entry.value)) {
            tempLogs.add(
              LogEntry(
                time: time,
                message: '${_signalName(entry.key)}이 감지되었습니다.',
                type: EntryType.button,
              ),
            );
          }
        }

        if (prevAdcValue == '0' || currAdcValue == '0') {
          for (final entry in doorButtons.entries) {
            if (_isRisingEdge(prev, curr, entry.value)) {
              rowHadDoorActivity = true;
              rowHadDoorCycle = true;
              tempLogs.add(
                LogEntry(
                  time: time,
                  message:
                      '${_doorButtonActionLabel(entry.key)}${_currentStationSuffix(doorLocation)}',
                  type: EntryType.button,
                  isSummary:
                      entry.key == 'CLOSE' ||
                      entry.key == 'S_CLOSE' ||
                      entry.key == 'REOPEN',
                ),
              );

              if (entry.key == 'CLOSE' || entry.key == 'S_CLOSE') {
                lastDoorCloseIndex = i;
                lastDoorCloseMode = currentDoorMode;
                lastDoorCloseUnclosedCars = _uncClosedDoorCars(curr, doorIdx);
                lastDoorCloseAt = rowClock;
                lastDoorCycleClosedAt = null;
                doorClosePending = currAdcValue == '0';
                doorCloseDelayLogged = false;
                doorCloseReopenLogged = false;
                doorCloseBypassLogged = false;
                lastDoorCycleHadDelay = false;
                lastDoorBypassTarget = '대상 불명';
              } else if (entry.key == 'REOPEN' &&
                  doorClosePending &&
                  !doorCloseReopenLogged) {
                tempLogs.add(
                  LogEntry(
                    time: time,
                    message:
                        '기관사 재개폐 버튼 취급, ${_allDoorCloseText(currAdcValue)} (출입문 모드: ${lastDoorCloseMode ?? currentDoorMode})',
                    type: EntryType.button,
                    isSummary: true,
                  ),
                );
                doorCloseReopenLogged = true;
                lastDoorCycleHadDelay = true;
              }
            }
          }
        }

        final autoModeIdx = modeIndices[OperationMode.auto] ?? -1;
        final isAutoMode =
            autoModeIdx != -1 && _getVal(curr, autoModeIdx) == '1';
        if (isAutoMode) {
          for (final entry in autoDoorOpenCommands.entries) {
            if (_isAnyRisingEdge(prev, curr, entry.value)) {
              rowHadDoorActivity = true;
              rowHadDoorCycle = true;
              if (activeStop != null && isStopped) {
                activeStop.hadDoorActivity = true;
                activeStop.hadDoorCycle = true;
              }
              tempLogs.add(
                LogEntry(
                  time: time,
                  message:
                      '자동운전 ${entry.key} 출입문 개방 명령, ${_allDoorCloseText(currAdcValue)}',
                  type: EntryType.button,
                  isSummary: true,
                ),
              );
              lastDoorContextAt = rowClock;
            }
          }
        }

        if (_isRisingEdge(prev, curr, idxAdbs) &&
            doorClosePending &&
            !doorCloseBypassLogged) {
          rowHadDoorActivity = true;
          rowHadDoorCycle = true;
          lastDoorBypassTarget = _doorBypassTargetLabel(prev, curr, doorIdx);
          tempLogs.add(
            LogEntry(
              time: time,
              message: _doorBypassReportMessage(
                lastDoorBypassTarget,
                currAdcValue,
                suffix: ' (출입문 모드: ${lastDoorCloseMode ?? currentDoorMode})',
              ),
              type: EntryType.bypass,
              isSummary: true,
            ),
          );
          if (currAdcValue == '0') {
            tempLogs.add(
              LogEntry(
                time: time,
                message: '바이패스 취급 후에도 ${_allDoorCloseText('0')} 유지',
                type: EntryType.bypass,
                isSummary: true,
              ),
            );
          }
          doorCloseBypassLogged = true;
          lastDoorCycleHadDelay = true;
        }

        if (idxAdc != -1) {
          if (prevAdcValue != currAdcValue) {
            if (currAdcValue == '0') {
              tempLogs.add(
                LogEntry(
                  time: time,
                  message: _allDoorCloseText('0'),
                  type: EntryType.bypass,
                  isSummary: true,
                ),
              );
            } else if (doorClosePending && lastDoorCloseAt != null) {
              tempLogs.add(
                LogEntry(
                  time: time,
                  message: doorCloseBypassLogged
                      ? _doorBypassReportMessage(lastDoorBypassTarget, '1')
                      : '${_allDoorCloseText('1')} (출입문 모드: ${lastDoorCloseMode ?? currentDoorMode})',
                  type: EntryType.restore,
                  isSummary: true,
                ),
              );
              lastDoorCycleClosedAt = rowClock;
              doorClosePending = false;
            } else {
              tempLogs.add(
                LogEntry(
                  time: time,
                  message: _allDoorCloseText('1'),
                  type: EntryType.restore,
                  isSummary: true,
                ),
              );
            }
            lastDoorContextAt = rowClock;
          }
          adcOpen = currAdcValue == '0';
        }

        final hasAdcDoorMismatch = _hasAdcDoorMismatch(curr, idxAdc, doorIdx);
        if (hasAdcDoorMismatch != adcDoorMismatchActive) {
          tempLogs.add(
            LogEntry(
              time: time,
              message: hasAdcDoorMismatch
                  ? '${_doorBypassReportMessage(_doorBypassTargetLabel(prev, curr, doorIdx), '1')} (바이패스 대상 불일치 가능성)'
                  : _allDoorCloseText(currAdcValue),
              type: hasAdcDoorMismatch ? EntryType.bypass : EntryType.restore,
              isSummary: true,
            ),
          );
          lastDoorContextAt = rowClock;
        }
        adcDoorMismatchActive = hasAdcDoorMismatch;

        for (final entry in doorIdx.entries) {
          final idx = entry.value;
          if (idx == -1) continue;

          final prevValue = _getVal(prev, idx);
          final currValue = _getVal(curr, idx);
          if (prevValue == currValue) continue;
          if (currValue == '0') {
            rowHadDoorActivity = true;
          }

          tempLogs.add(
            LogEntry(
              time: time,
              message: currValue == '0'
                  ? '${_formatDoorName(entry.key)} 닫힘 상태가 해제된 정황이 감지되었습니다.'
                  : '${_formatDoorName(entry.key)} 닫힘 상태가 복구된 정황이 확인됩니다.',
              type: currValue == '0' ? EntryType.bypass : EntryType.restore,
            ),
          );
          lastDoorContextAt = rowClock;
        }

        if (doorClosePending) {
          lastDoorCloseUnclosedCars = _uncClosedDoorCars(curr, doorIdx);
          if (lastDoorCloseAt != null) {
            final elapsed = rowClock - lastDoorCloseAt;
            if (!doorCloseDelayLogged &&
                currAdcValue == '0' &&
                elapsed >= kDoorCloseDelayWarnDuration) {
              final carDetail = lastDoorCloseUnclosedCars.length == 1
                  ? ' 출입문 닫힘 과정에서 ${lastDoorCloseUnclosedCars.first} 출입문 닫힘 신호가 지연되었습니다.'
                  : (lastDoorCloseUnclosedCars.isNotEmpty
                        ? ' 닫힘 미형성 호차: ${lastDoorCloseUnclosedCars.join(', ')}.'
                        : '');
              tempLogs.add(
                LogEntry(
                  time: time,
                  message:
                      '기관사 닫힘버튼 취급, ${_allDoorCloseText('0')} (출입문 모드: ${lastDoorCloseMode ?? currentDoorMode})${carDetail.isEmpty ? '' : ' $carDetail'}',
                  type: EntryType.bypass,
                  isSummary: true,
                ),
              );
              doorCloseDelayLogged = true;
              lastDoorCycleHadDelay = true;
            }
          }
        }

        if (_isAnyRisingEdge(prev, curr, spIndices)) {
          tempLogs.add(
            LogEntry(
              time: time,
              message: '정지 후 진행(S&P) 신호가 감지되었습니다.',
              type: EntryType.info,
              isSummary: true,
            ),
          );
        }

        if (_isAnyRisingEdge(prev, curr, dptPmIndices)) {
          final isFirstDptAfterDoorCycle =
              lastDoorCloseIndex != null &&
              lastDptLoggedCloseIndex != lastDoorCloseIndex &&
              currAdcValue == '1' &&
              (lastDoorCycleClosedAt != null || lastDoorCycleHadDelay);
          final isModeChangeDpt =
              _isRecent(rowClock, lastModeChangeAt, kContextWindow) &&
              lastModeChangeAt != lastModeChangeDptConsumedAt;

          if (isFirstDptAfterDoorCycle || isModeChangeDpt) {
            // 출발 허가(DPT-PM)는 내부 맥락 판단에만 사용하고 사용자 로그에는 출력하지 않는다.
            if (isFirstDptAfterDoorCycle) {
              lastDptLoggedCloseIndex = lastDoorCloseIndex;
            }
            if (isModeChangeDpt) {
              lastModeChangeDptConsumedAt = lastModeChangeAt;
            }
          }
        }

        if (idxNCode != -1) {
          final currNCode = _getVal(curr, idxNCode) == '1';
          if (currNCode) {
            nCodeDuration = nCodeActive
                ? nCodeDuration + sampleDelta
                : sampleDelta;

            if (!nCodeWarnIssued && nCodeDuration >= kNCodeWarnDelay) {
              tempLogs.add(
                LogEntry(
                  time: time,
                  message: '무코드 수신이 4초 이상 지속되고 있습니다.',
                  type: EntryType.info,
                  isSummary: true,
                ),
              );
              nCodeWarnIssued = true;
            }

            if (nCodeWarnIssued &&
                nCodeDuration.inSeconds >= nextNCodeSummarySecond &&
                nextNCodeSummarySecond >= kNCodeSummaryInterval.inSeconds) {
              tempLogs.add(
                LogEntry(
                  time: time,
                  message:
                      '무코드 수신이 ${nCodeDuration.inSeconds}초 이상 지속되는 정황이 확인됩니다.',
                  type: EntryType.info,
                  isSummary: true,
                ),
              );
              nextNCodeSummarySecond += kNCodeSummaryInterval.inSeconds;
            }
          } else {
            nCodeDuration = Duration.zero;
            nCodeWarnIssued = false;
            nextNCodeSummarySecond = kNCodeSummaryInterval.inSeconds;
          }
          nCodeActive = currNCode;
        }

        if (_isRisingEdge(prev, curr, idxFsb)) {
          tempLogs.add(
            LogEntry(
              time: time,
              message: '${_signalName('FSBR')} 정황이 확인됩니다.',
              type: EntryType.info,
              isSummary: true,
            ),
          );

          if (nCodeDuration >= kFsbrNCodeContextDuration) {
            tempLogs.add(
              LogEntry(
                time: time,
                message: '무코드 수신이 이어진 뒤 전상용제동(FSB) 체결 흐름이 확인됩니다.',
                type: EntryType.info,
                isSummary: true,
              ),
            );
          } else if (_isRecent(rowClock, lastModeChangeAt, kContextWindow)) {
            tempLogs.add(
              LogEntry(
                time: time,
                message: '운전모드 전환 이후 전상용제동(FSB) 체결 흐름이 확인됩니다.',
                type: EntryType.info,
                isSummary: true,
              ),
            );
          } else if (_isRecent(rowClock, lastDoorContextAt, kContextWindow)) {
            tempLogs.add(
              LogEntry(
                time: time,
                message: '출입문 닫힘 신호 변화 이후 전상용제동(FSB) 체결 정황이 확인됩니다.',
                type: EntryType.info,
                isSummary: true,
              ),
            );
          }
        }

        if (idxVel != -1) {
          if (speed <= kStopSpeedThresholdKmh) {
            zeroDuration += sampleDelta;
          } else {
            zeroDuration = Duration.zero;
          }

          if (!isStopped && zeroDuration >= kStopConfirmDuration) {
            isStopped = true;
            movingDuration = Duration.zero;
            departureStartTime = null;
            activeStop = _ActiveStopContext(
              startIndex: i,
              startTime: time,
              startLogOffset: rowLogStartIndex,
              previousStationName: lastConfirmedStationName,
              approachStationCode: lastApproachStationCode,
              approachStationName: lastApproachStationName,
              approachDist: lastApproachDist,
              startNextStaCode: nextStaCode,
              startNextStationName: _nextStationDisplay(nextStaCode),
              startDist: currentDist,
              hadDoorActivity: rowHadDoorActivity,
              hadAdcRelease: prevAdcValue == '1' && currAdcValue == '0',
              hadDoorCycle: rowHadDoorCycle,
            );

            final inferredStation =
                activeStop.approachStationName ??
                lastConfirmedStationName ??
                _currentStationNameFromNext(nextStaCode, travelDirection);
            if (inferredStation != null) {
              lastBerthedStation = inferredStation;
            }

            tempLogs.add(
              LogEntry(
                time: time,
                message: '정차가 감지되었습니다.',
                type: EntryType.info,
                isSummary: true,
              ),
            );

            final locationText = _formatStopLocationText(
              nextSta: nextStaCode,
              direction: travelDirection,
              lastBerthedStation: lastBerthedStation,
            );
            if (locationText != null) {
              tempLogs.add(
                LogEntry(
                  time: time,
                  message: locationText,
                  type: EntryType.info,
                  isSummary: true,
                ),
              );
            }
          }

          if (activeStop != null && isStopped) {
            final stop = activeStop;
            stop.hadDoorActivity = stop.hadDoorActivity || rowHadDoorActivity;
            stop.hadAdcRelease =
                stop.hadAdcRelease ||
                (prevAdcValue == '1' && currAdcValue == '0');
            stop.hadDoorCycle = stop.hadDoorCycle || rowHadDoorCycle;
            if (nextStaCode != stop.startNextStaCode) {
              stop.hadNextStaChange = true;
            }
          }

          if (isStopped) {
            if (speed > kStopSpeedThresholdKmh && departureStartTime == null) {
              departureStartTime = time;
            }

            if (speed >= kDepartureSpeedThresholdKmh) {
              movingDuration += sampleDelta;
            } else if (speed <= kStopSpeedThresholdKmh) {
              movingDuration = Duration.zero;
              departureStartTime = null;
            }

            if (movingDuration >= kStopConfirmDuration) {
              final departureLogOffset = tempLogs.length;
              final stop = activeStop;
              if (stop != null) {
                final stopStationName = _resolveStopStationName(
                  approachStationName: stop.approachStationName,
                  currentNextStationName: stop.startNextStationName,
                  lastBerthedStation: lastBerthedStation,
                  stopDist: stop.startDist,
                );
                final hasDoorEvidence =
                    stop.hadDoorActivity ||
                    stop.hadAdcRelease ||
                    stop.hadDoorCycle;
                final hasLargeStopDist =
                    stop.startDist != null &&
                    stop.startDist! > kInterstationStopDistanceThreshold;
                final blockType = hasDoorEvidence
                    ? ((stop.hadNextStaChange || hasLargeStopDist)
                          ? AnalysisBlockType.overrunSuspectedStop
                          : AnalysisBlockType.stationStop)
                    : AnalysisBlockType.interstationStop;
                final blockNote =
                    blockType == AnalysisBlockType.overrunSuspectedStop
                    ? (stop.hadNextStaChange
                          ? '정차 시점에 다음역 정보가 갱신된 정황이 확인됩니다.'
                          : '정차 위치와 다음역 정보 간 불일치 정황이 확인됩니다.')
                    : null;
                stopSeeds.add(
                  _StopBlockSeed(
                    type: blockType,
                    startLogOffset: stop.startLogOffset,
                    stopLogEndOffsetExclusive: departureLogOffset,
                    departureLogOffset: departureLogOffset,
                    previousStationName: stop.previousStationName,
                    approachStationName: stop.approachStationName,
                    stationName: stopStationName,
                    nextStationName: stop.startNextStationName,
                    note: blockNote,
                  ),
                );
                if (blockType != AnalysisBlockType.interstationStop &&
                    stopStationName != null) {
                  lastConfirmedStationName = stopStationName;
                }
                activeStop = null;
              }

              isStopped = false;
              zeroDuration = Duration.zero;
              movingDuration = Duration.zero;

              tempLogs.add(
                LogEntry(
                  time: departureStartTime ?? time,
                  message: '발차가 감지되었습니다.',
                  type: EntryType.info,
                  isSummary: true,
                ),
              );

              final locationText = _formatDepartureLocationText(
                lastBerthedStation: lastBerthedStation,
                nextSta: nextStaCode,
              );
              if (locationText != null) {
                tempLogs.add(
                  LogEntry(
                    time: departureStartTime ?? time,
                    message: locationText,
                    type: EntryType.info,
                    isSummary: true,
                  ),
                );
              }

              departureStartTime = null;
            }
          }
        }
      }

      final stop = activeStop;
      if (stop != null) {
        final stopStationName = _resolveStopStationName(
          approachStationName: stop.approachStationName,
          currentNextStationName: stop.startNextStationName,
          lastBerthedStation: lastBerthedStation,
          stopDist: stop.startDist,
        );
        final hasDoorEvidence =
            stop.hadDoorActivity || stop.hadAdcRelease || stop.hadDoorCycle;
        final hasLargeStopDist =
            stop.startDist != null &&
            stop.startDist! > kInterstationStopDistanceThreshold;
        final blockType = hasDoorEvidence
            ? ((stop.hadNextStaChange || hasLargeStopDist)
                  ? AnalysisBlockType.overrunSuspectedStop
                  : AnalysisBlockType.stationStop)
            : AnalysisBlockType.interstationStop;
        final blockNote = blockType == AnalysisBlockType.overrunSuspectedStop
            ? (stop.hadNextStaChange
                  ? '정차 시점에 다음역 정보가 갱신된 정황이 확인됩니다.'
                  : '정차 위치와 다음역 정보 간 불일치 정황이 확인됩니다.')
            : null;
        stopSeeds.add(
          _StopBlockSeed(
            type: blockType,
            startLogOffset: stop.startLogOffset,
            stopLogEndOffsetExclusive: tempLogs.length,
            departureLogOffset: tempLogs.length,
            previousStationName: stop.previousStationName,
            approachStationName: stop.approachStationName,
            stationName: stopStationName,
            nextStationName: stop.startNextStationName,
            note: blockNote,
          ),
        );
        if (blockType != AnalysisBlockType.interstationStop &&
            stopStationName != null) {
          lastConfirmedStationName = stopStationName;
        }
      }

      final resolvedLogs = tempLogs.isEmpty
          ? const [
              LogEntry(
                time: '-',
                message: '감지된 주요 이벤트가 없습니다.',
                type: EntryType.info,
                isSummary: true,
              ),
            ]
          : tempLogs;
      final resolvedBlocks = _buildBlocks(
        resolvedLogs,
        stopSeeds,
        tailApproachStationName: lastApproachStationName,
      );

      setState(() {
        statusText =
            '분석 완료 | $filterLabel 구간 ${rows.length - 1}개 행 처리${_sourceFileSummary.isEmpty ? '' : ' | $_sourceFileSummary'}';
        summaryText = _buildSummary(resolvedLogs);
        logs = resolvedLogs;
        blocks = resolvedBlocks;
        isLoading = false;
      });
  }

  @override
  Widget build(BuildContext context) {
    final bypassCount = logs
        .where((log) => log.type == EntryType.bypass)
        .length;
    final restoreCount = logs
        .where((log) => log.type == EntryType.restore)
        .length;
    final buttonCount = logs
        .where((log) => log.type == EntryType.button)
        .length;
    final infoCount = logs.where((log) => log.type == EntryType.info).length;

    return Scaffold(
      appBar: AppBar(
        title: const Text(
          '광주도시철도 운행로그 분석기',
          style: TextStyle(fontWeight: FontWeight.bold),
        ),
        backgroundColor: const Color(0xFF2C3E50),
        foregroundColor: Colors.white,
      ),
      body: Column(
        children: [
          Container(
            width: double.infinity,
            padding: const EdgeInsets.all(12),
            color: Colors.grey[200],
            child: Column(
              crossAxisAlignment: CrossAxisAlignment.start,
              children: [
                Text(
                  statusText,
                  style: const TextStyle(fontSize: 12, color: Colors.blueGrey),
                ),
                const SizedBox(height: 4),
                const Text(
                  kLatestBuildDateLabel,
                  style: TextStyle(
                    fontSize: 11,
                    color: Colors.blueGrey,
                    fontWeight: FontWeight.w600,
                  ),
                ),
              ],
            ),
          ),
          Padding(
            padding: const EdgeInsets.all(16),
            child: ElevatedButton.icon(
              onPressed: isLoading ? null : _pickAndAnalyze,
              icon: const Icon(Icons.analytics_outlined),
              label: Text(isLoading ? '분석 중...' : '운행기록 파일(.xlsx / .csv) 불러오기'),
              style: ElevatedButton.styleFrom(
                minimumSize: const Size(double.infinity, 55),
                backgroundColor: const Color(0xFF34495E),
                foregroundColor: Colors.white,
                shape: RoundedRectangleBorder(
                  borderRadius: BorderRadius.circular(8),
                ),
              ),
            ),
          ),
          if (_sourceRows != null)
            Padding(
              padding: const EdgeInsets.fromLTRB(16, 0, 16, 12),
              child: Row(
                children: [
                  Expanded(
                    child: TextField(
                      controller: _startTimeController,
                      enabled: !isLoading,
                      decoration: const InputDecoration(
                        labelText: '시작 시간',
                        hintText: 'HH:mm:ss',
                        isDense: true,
                        border: OutlineInputBorder(),
                      ),
                    ),
                  ),
                  const SizedBox(width: 8),
                  Expanded(
                    child: TextField(
                      controller: _endTimeController,
                      enabled: !isLoading,
                      decoration: const InputDecoration(
                        labelText: '종료 시간',
                        hintText: 'HH:mm:ss',
                        isDense: true,
                        border: OutlineInputBorder(),
                      ),
                    ),
                  ),
                  const SizedBox(width: 8),
                  SizedBox(
                    height: 48,
                    child: ElevatedButton(
                      onPressed: isLoading ? null : _applyTimeFilter,
                      style: ElevatedButton.styleFrom(
                        backgroundColor: const Color(0xFF2C3E50),
                        foregroundColor: Colors.white,
                      ),
                      child: const Text('적용'),
                    ),
                  ),
                ],
              ),
            ),
          if (logs.isEmpty)
            Padding(
              padding: const EdgeInsets.fromLTRB(16, 0, 16, 16),
              child: Column(
                crossAxisAlignment: CrossAxisAlignment.stretch,
                children: [
                  Text(
                    '운행기록 분석을 보조하기 위한 시스템으로,\n최종 판단은 사용자의 확인이 필요합니다.',
                    textAlign: TextAlign.center,
                    style: TextStyle(
                      fontSize: 12,
                      color: Colors.black.withValues(alpha: 0.78),
                      fontWeight: FontWeight.w500,
                      height: 1.4,
                    ),
                  ),
                  const SizedBox(height: 20),
                  Align(
                    alignment: Alignment.centerRight,
                    child: Text(
                      '제작자: 강경현',
                      style: TextStyle(
                        fontSize: 11,
                        color: Colors.black.withValues(alpha: 0.72),
                        fontWeight: FontWeight.w500,
                      ),
                    ),
                  ),
                ],
              ),
            ),
          if (summaryText.isNotEmpty)
            Padding(
              padding: const EdgeInsets.fromLTRB(16, 0, 16, 12),
              child: Container(
                width: double.infinity,
                padding: const EdgeInsets.all(12),
                decoration: BoxDecoration(
                  color: const Color(0xFF34495E).withValues(alpha: 0.08),
                  borderRadius: BorderRadius.circular(10),
                  border: Border.all(
                    color: const Color(0xFF34495E).withValues(alpha: 0.2),
                  ),
                ),
                child: SelectableText(
                  summaryText,
                  style: const TextStyle(
                    fontSize: 13,
                    fontWeight: FontWeight.w600,
                  ),
                ),
              ),
            ),
          if (logs.isNotEmpty)
            Padding(
              padding: const EdgeInsets.fromLTRB(16, 0, 16, 12),
              child: Row(
                children: [
                  Expanded(
                    child: Wrap(
                      spacing: 8,
                      runSpacing: 6,
                      children: [
                        _Badge(
                          label: '이상',
                          count: bypassCount,
                          color: Colors.red,
                        ),
                        _Badge(
                          label: '복구',
                          count: restoreCount,
                          color: Colors.green,
                        ),
                        _Badge(
                          label: '취급',
                          count: buttonCount,
                          color: Colors.orange,
                        ),
                        _Badge(
                          label: '정보',
                          count: infoCount,
                          color: Colors.blue,
                        ),
                      ],
                    ),
                  ),
                  const SizedBox(width: 8),
                  ToggleButtons(
                    isSelected: [
                      viewMode == LogViewMode.summary,
                      viewMode == LogViewMode.detail,
                    ],
                    onPressed: (index) {
                      setState(() {
                        viewMode = index == 0
                            ? LogViewMode.summary
                            : LogViewMode.detail;
                      });
                    },
                    borderRadius: BorderRadius.circular(8),
                    constraints: const BoxConstraints(
                      minHeight: 36,
                      minWidth: 58,
                    ),
                    children: const [
                      Padding(
                        padding: EdgeInsets.symmetric(horizontal: 10),
                        child: Text('블록'),
                      ),
                      Padding(
                        padding: EdgeInsets.symmetric(horizontal: 10),
                        child: Text('상세'),
                      ),
                    ],
                  ),
                ],
              ),
            ),
          Expanded(
            child: viewMode == LogViewMode.summary
                ? ListView.builder(
                    itemCount: blocks.length,
                    itemBuilder: (context, index) {
                      final block = blocks[index];
                      final color = _blockColor(block.type);
                      final icon = _blockIcon(block.type);
                      final timeRange = _entryTimeRange(block.entries);

                      return Card(
                        margin: const EdgeInsets.symmetric(
                          horizontal: 16,
                          vertical: 5,
                        ),
                        elevation: 0,
                        shape: RoundedRectangleBorder(
                          side: BorderSide(
                            color: color.withValues(alpha: 0.25),
                          ),
                          borderRadius: BorderRadius.circular(10),
                        ),
                        color: color.withValues(alpha: 0.05),
                        child: Padding(
                          padding: const EdgeInsets.all(12),
                          child: Column(
                            crossAxisAlignment: CrossAxisAlignment.start,
                            children: [
                              Row(
                                children: [
                                  Icon(icon, color: color, size: 20),
                                  const SizedBox(width: 8),
                                  Expanded(
                                    child: Text(
                                      block.title,
                                      style: TextStyle(
                                        color: color,
                                        fontSize: 14,
                                        fontWeight: FontWeight.w700,
                                      ),
                                    ),
                                  ),
                                  if (timeRange.isNotEmpty)
                                    Container(
                                      padding: const EdgeInsets.symmetric(
                                        horizontal: 8,
                                        vertical: 3,
                                      ),
                                      decoration: BoxDecoration(
                                        color: color.withValues(alpha: 0.10),
                                        borderRadius: BorderRadius.circular(12),
                                      ),
                                      child: Text(
                                        timeRange,
                                        style: TextStyle(
                                          color: color,
                                          fontSize: 11,
                                          fontWeight: FontWeight.w600,
                                        ),
                                      ),
                                    ),
                                ],
                              ),
                              if (block.entries.isNotEmpty)
                                const SizedBox(height: 8),
                              ...block.entries.map(
                                (entry) => Padding(
                                  padding: const EdgeInsets.only(top: 4),
                                  child: SelectableText(
                                    '• ${entry.message}',
                                    style: const TextStyle(
                                      fontSize: 13,
                                      color: Colors.black87,
                                    ),
                                  ),
                                ),
                              ),
                            ],
                          ),
                        ),
                      );
                    },
                  )
                : ListView.builder(
                    itemCount: logs.length,
                    itemBuilder: (context, index) {
                      final entry = logs[index];
                      final color = switch (entry.type) {
                        EntryType.bypass => Colors.red,
                        EntryType.restore => Colors.green,
                        EntryType.button => Colors.orange,
                        EntryType.info => Colors.blue,
                      };
                      final icon = switch (entry.type) {
                        EntryType.bypass => Icons.warning_amber_rounded,
                        EntryType.restore => Icons.check_circle_outline,
                        EntryType.button => Icons.touch_app_outlined,
                        EntryType.info => Icons.directions_subway,
                      };

                      return Card(
                        margin: const EdgeInsets.symmetric(
                          horizontal: 16,
                          vertical: 3,
                        ),
                        elevation: 0,
                        shape: RoundedRectangleBorder(
                          side: BorderSide(color: color.withValues(alpha: 0.3)),
                          borderRadius: BorderRadius.circular(8),
                        ),
                        color: color.withValues(alpha: 0.05),
                        child: ListTile(
                          dense: true,
                          leading: Icon(icon, color: color, size: 22),
                          title: SelectableText(
                            '[${entry.time}] ${entry.message}',
                            style: TextStyle(
                              color: Colors.black87,
                              fontSize: 13,
                              fontWeight: entry.type == EntryType.bypass
                                  ? FontWeight.bold
                                  : FontWeight.normal,
                            ),
                          ),
                        ),
                      );
                    },
                  ),
          ),
        ],
      ),
    );
  }

  void _showXlsWarning() {
    showDialog(
      context: context,
      builder: (ctx) => AlertDialog(
        title: const Row(
          children: [
            Icon(Icons.warning_amber_rounded, color: Colors.orange),
            SizedBox(width: 8),
            Text('구형 파일 감지'),
          ],
        ),
        content: const Text(
          'Excel 97-2003(.xls) 형식은 지원하지 않습니다. .xlsx 또는 .csv로 변환한 뒤 다시 업로드해 주세요.',
        ),
        actions: [
          TextButton(
            onPressed: () => Navigator.pop(ctx),
            child: const Text('확인'),
          ),
        ],
      ),
    );
  }
}

class _Badge extends StatelessWidget {
  final String label;
  final int count;
  final Color color;

  const _Badge({required this.label, required this.count, required this.color});

  @override
  Widget build(BuildContext context) {
    return Container(
      padding: const EdgeInsets.symmetric(horizontal: 12, vertical: 6),
      decoration: BoxDecoration(
        color: color.withValues(alpha: 0.1),
        borderRadius: BorderRadius.circular(20),
        border: Border.all(color: color.withValues(alpha: 0.4)),
      ),
      child: Text(
        '$label: $count건',
        style: TextStyle(
          fontSize: 12,
          color: color,
          fontWeight: FontWeight.bold,
        ),
      ),
    );
  }
}
