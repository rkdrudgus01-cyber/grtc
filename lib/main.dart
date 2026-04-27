import 'dart:convert';
import 'dart:js_interop';

import 'package:excel/excel.dart' hide Border;
import 'package:file_picker/file_picker.dart';
import 'package:flutter/material.dart';
import 'package:flutter/services.dart';

void main() => runApp(
  MaterialApp(
    home: const LogAnalyzer(),
    debugShowCheckedModeBanner: false,
    theme: ThemeData(fontFamily: 'NotoSansKR', useMaterial3: false),
  ),
);

@JS('grtcCanReadLegacyXls')
external bool _grtcCanReadLegacyXls();

@JS('grtcReadXlsRowsFromBase64')
external JSPromise<JSAny?> _grtcReadXlsRowsFromBase64(JSString base64);

enum EntryType { button, bypass, restore, info }

enum LogViewMode { summary, detail, report }

enum OperationMode { auto, driverless, manual, emergency, yard, unknown }

enum IncidentType {
  doorCloseFailure,
  doorOpenFailure,
  doorBypass,
  ncodeEmergency,
  brake,
  modeChange,
  doorFlow,
  general,
}

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
const double kAtcOverspeedMarginKmh = 2.0;
const Duration kStopConfirmDuration = Duration(seconds: 2);
const Duration kContextWindow = Duration(seconds: 3);
const Duration kFsbrNCodeContextDuration = Duration(seconds: 2);
const Duration kNCodeWarnDelay = Duration(seconds: 4);
const Duration kNCodeSummaryInterval = Duration(seconds: 30);
const Duration kLongPersistenceDuration = Duration(seconds: 120);
const Duration kDoorCloseDelayWarnDuration = Duration(seconds: 8);
const Duration kDoorOpenDelayWarnDuration = Duration(seconds: 5);
const Duration kDoorBypassHoldWarnDuration = Duration(seconds: 3);
const Duration kDepartureMovementConfirmDuration = Duration(seconds: 3);
const Duration kModeConflictConfirmDuration = Duration(seconds: 3);
const Duration kSignalSpikeIgnoreDuration = Duration(seconds: 3);
const int kUiYieldInterval = 250;

const List<int> kEffectiveDoorCars = [0, 1, 2, 7];
const List<String> kActiveDoors = ['DOOR0', 'DOOR1', 'DOOR2', 'DOOR7'];
const List<String> kKoreanFontFallback = [
  'Noto Sans KR',
  'Noto Sans CJK KR',
  'Malgun Gothic',
  'Apple SD Gothic Neo',
  'Nanum Gothic',
];

const Map<String, String> kSignalDisplayNames = {
  'ADC': '전체 출입문 닫힘',
  'OPEN-L': '좌측 문열림 버튼 취급',
  'OPEN-R': '우측 문열림 버튼 취급',
  'CLOSE': '문닫힘 버튼 취급',
  'REOPEN': '재개폐 버튼 취급',
  'S_OPEN-L': '사이드 좌측 문열림 버튼 취급',
  'S_OPEN-R': '사이드 우측 문열림 버튼 취급',
  'S_CLOSE': '사이드 문닫힘 버튼 취급',
  'S_REOPEN': '사이드 재개폐 버튼 취급',
  'ADBS': '전차 출입문 바이패스',
  'DOOR0': '0호차 출입문 바이패스 신호',
  'DOOR1': '1호차 출입문 바이패스 신호',
  'DOOR2': '2호차 출입문 바이패스 신호',
  'DOOR7': '7호차 출입문 바이패스 신호',
  'ATC1/2OPEN-L': '자동운전 좌측 출입문 개방 명령',
  'ATC1/2OPEN-R': '자동운전 우측 출입문 개방 명령',
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
  'MC N': '주간제어기 Not N Position',
  'FOR WAR': '전방예고',
  'PBR': '주차제동 완해',
  'SBS': '보안제동 취급',
  'CRPB': '강제완해 푸쉬버튼',
  'PBPS': '주차제동 압력스위치',
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
  double? minDist;
  double? endDist;
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
    double? minDist,
    double? endDist,
    this.hadDoorActivity = false,
    this.hadAdcRelease = false,
    this.hadDoorCycle = false,
  }) : minDist = minDist ?? startDist,
       endDist = endDist ?? startDist;
}

class _ParsedLogFile {
  final String name;
  final List<List<dynamic>> rows;

  const _ParsedLogFile({required this.name, required this.rows});
}

class _PreparedLogRows {
  final List<List<dynamic>> rows;
  final String scopeLabel;

  const _PreparedLogRows({required this.rows, required this.scopeLabel});
}

class LogAnalyzer extends StatefulWidget {
  const LogAnalyzer({super.key});

  @override
  State<LogAnalyzer> createState() => _LogAnalyzerState();
}

class _LogAnalyzerState extends State<LogAnalyzer> {
  List<LogEntry> logs = [];
  List<AnalysisBlock> blocks = [];
  String statusText = '운행기록 파일을 선택해 주세요. (.xls / .xlsx / .csv)';
  String summaryText = '';
  String reportText = '';
  bool isLoading = false;
  LogViewMode viewMode = LogViewMode.detail;
  List<_ParsedLogFile> _lastParsedFiles = const [];
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

  Future<List<List<dynamic>>> _readLegacyXlsRows(
    String fileName,
    Uint8List bytes,
  ) async {
    try {
      if (!_grtcCanReadLegacyXls()) {
        throw '구형 Excel 파서가 준비되지 않았습니다.';
      }

      final result = await _grtcReadXlsRowsFromBase64(
        base64Encode(bytes).toJS,
      ).toDart;
      final dartRows = result.dartify();
      if (dartRows is! List) {
        throw '구형 Excel 데이터 구조를 읽을 수 없습니다.';
      }
      return [
        for (final row in dartRows)
          if (row is List)
            row.map((cell) => cell?.toString().trim() ?? '').toList()
          else
            [row?.toString().trim() ?? ''],
      ];
    } catch (error) {
      throw '$fileName 구형 Excel(.xls) 자동 변환에 실패했습니다. Excel에서 “다른 이름으로 저장 → Excel 통합 문서(.xlsx)”로 변환 후 업로드해 주세요. ($error)';
    }
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

  bool _isSustainedValueAt(
    List<List<dynamic>> rows,
    int startIndex,
    int idx,
    int timeIdx,
    String targetValue,
    Duration minDuration,
  ) {
    if (idx == -1 ||
        startIndex < 0 ||
        startIndex >= rows.length ||
        _getVal(rows[startIndex], idx) != targetValue) {
      return false;
    }

    var duration = Duration.zero;
    for (var j = startIndex + 1; j < rows.length; j++) {
      final prevRow = rows[j - 1];
      final currRow = rows[j];
      if (_getVal(prevRow, idx) != targetValue) break;

      final prevTime = timeIdx != -1 && timeIdx < prevRow.length
          ? _cellToStr(prevRow[timeIdx])
          : '';
      final currTime = timeIdx != -1 && timeIdx < currRow.length
          ? _cellToStr(currRow[timeIdx])
          : '';
      duration += _resolveSampleDelta(prevTime, currTime);
      if (duration >= minDuration) return true;
      if (_getVal(currRow, idx) != targetValue) break;
    }

    return false;
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

  String? _directionLabel(int direction) {
    if (direction > 0) return '하선';
    if (direction < 0) return '상선';
    return null;
  }

  String _stationWithDirection(String station, int direction) {
    final directionLabel = _directionLabel(direction);
    return directionLabel == null ? station : '$station $directionLabel';
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
      return '${_stationWithDirection(currentStation, direction)} 정차 중 / 다음역 $nextStation';
    }
    if (currentStation != null) {
      return '${_stationWithDirection(currentStation, direction)} 정차 중으로 해석됩니다.';
    }
    if (nextStation != null) {
      return '정차 중이며 다음역 $nextStation 정보가 확인됩니다.';
    }
    return null;
  }

  String? _formatDepartureLocationText({
    required String? lastBerthedStation,
    required int? nextSta,
    required int direction,
  }) {
    final nextStation = _nextStationDisplay(nextSta);
    if (lastBerthedStation != null && nextStation != null) {
      return '${_stationWithDirection(lastBerthedStation, direction)} 발차 후 $nextStation 방면으로 해석됩니다.';
    }
    if (nextStation != null) {
      return '$nextStation 방면으로 운행 중인 정황이 확인됩니다.';
    }
    return null;
  }

  String _formatDoorName(String doorSignal) {
    final carNo = doorSignal.replaceFirst('DOOR', '');
    return '$carNo호차 출입문 바이패스 신호';
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

  String _formatElapsed(Duration duration) {
    if (duration.inMilliseconds % 1000 == 0) {
      return '${duration.inSeconds}초';
    }
    return '${(duration.inMilliseconds / 1000).toStringAsFixed(1)}초';
  }

  List<String> _activeDoorBypassCars(
    List<dynamic> row,
    Map<String, int> doorIdx,
  ) {
    final cars = <String>[];
    for (final entry in doorIdx.entries) {
      if (entry.value != -1 && _getVal(row, entry.value) == '0') {
        cars.add('${entry.key.replaceFirst('DOOR', '')}호차');
      }
    }
    return cars;
  }

  List<String> _activeDoorBypassCarsFromValues(Map<String, String> doorValues) {
    return doorValues.entries
        .where((entry) => entry.value == '0')
        .map((entry) => '${entry.key.replaceFirst('DOOR', '')}호차')
        .toList();
  }

  List<String> _doorBypassTargetCars(
    List<dynamic> prev,
    List<dynamic> curr,
    Map<String, int> doorIdx,
  ) {
    final prevTargets = _activeDoorBypassCars(prev, doorIdx);
    return prevTargets.isNotEmpty
        ? prevTargets
        : _activeDoorBypassCars(curr, doorIdx);
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
      return '${_stationWithDirection(currentStation, direction)} 정차 중';
    }
    return '정차 중';
  }

  String _currentStationSuffix(String? location) {
    if (location == null || location.trim().isEmpty) return '';
    final station = location.split(' 정차 중').first.trim();
    if (station.isEmpty || station == location) return '';
    if (station.endsWith(' 상선')) {
      return ' (현재 역: ${station.substring(0, station.length - 3)}, 방향: 상선)';
    }
    if (station.endsWith(' 하선')) {
      return ' (현재 역: ${station.substring(0, station.length - 3)}, 방향: 하선)';
    }
    return ' (현재 역: $station)';
  }

  String _adcReportValue(String value) {
    return 'ALL DOOR CLOSE "$value"';
  }

  String _doorModeSuffix(String doorMode) {
    return ' (출입문 모드: $doorMode)';
  }

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
        return '사이드 닫힘버튼 취급';
      case 'S_REOPEN':
        return '사이드 재개폐 버튼 취급';
      default:
        return _signalName(signal);
    }
  }

  String _autoDoorOpenCommandLabel(String signal) {
    switch (signal) {
      case 'ATC1/2OPEN-L':
        return '자동운전 좌측 출입문 개방 명령';
      case 'ATC1/2OPEN-R':
        return '자동운전 우측 출입문 개방 명령';
      default:
        return '자동운전 출입문 개방 명령';
    }
  }

  String _doorButtonReportMessage({
    required String signal,
    required String adcValue,
    required String doorMode,
    required String? location,
  }) {
    final label = _doorButtonActionLabel(signal);
    final locationSuffix = _currentStationSuffix(location);
    if (signal == 'CLOSE' || signal == 'S_CLOSE') {
      return '$label$locationSuffix';
    }
    if (signal == 'REOPEN' || signal == 'S_REOPEN') {
      return '$label, ${_adcReportValue(adcValue)}${_doorModeSuffix(doorMode)}$locationSuffix';
    }
    return '$label$locationSuffix';
  }

  String _adcRestoreReportMessage(String doorMode) {
    return '${_adcReportValue('1')}${_doorModeSuffix(doorMode)}';
  }

  String _doorBypassTargetLabel(List<String> targetCars) {
    return targetCars.isEmpty ? '대상 불명' : targetCars.join(', ');
  }

  String _doorBypassReportMessage(
    String adcValue, [
    List<String> targetCars = const [],
  ]) {
    return '기관사 출입문 바이패스 취급(${_doorBypassTargetLabel(targetCars)}), ${_adcReportValue(adcValue)}';
  }

  int? _activeAtcSpeedCode(
    List<dynamic> row,
    Map<int, List<int>> speedCodeIndices,
  ) {
    int? activeCode;
    for (final entry in speedCodeIndices.entries) {
      final hasActiveInput = entry.value.any((idx) => _getVal(row, idx) == '1');
      if (!hasActiveInput) continue;
      if (activeCode == null || entry.key < activeCode) {
        activeCode = entry.key;
      }
    }
    return activeCode;
  }

  String _formatSpeedKmh(double speed) {
    if (speed % 1 == 0) {
      return speed.toInt().toString();
    }
    return speed.toStringAsFixed(1);
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
      final marker = '정차 중 ';
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

  bool _hasLargeStopPositionMismatch(_ActiveStopContext stop) {
    final stopDist = stop.minDist ?? stop.startDist;
    if (stopDist == null || stopDist <= kInterstationStopDistanceThreshold) {
      return false;
    }
    final approachDist = stop.approachDist;
    return approachDist == null ||
        approachDist > kInterstationStopDistanceThreshold;
  }

  AnalysisBlockType _resolveStopBlockType({
    required _ActiveStopContext stop,
    required bool hasDoorEvidence,
  }) {
    if (!hasDoorEvidence) {
      return AnalysisBlockType.interstationStop;
    }

    // NEXTSTA often changes during a normal station stop. Treat it as overrun
    // context only when the approach and stop distances both remain large.
    if (_hasLargeStopPositionMismatch(stop)) {
      return AnalysisBlockType.overrunSuspectedStop;
    }
    return AnalysisBlockType.stationStop;
  }

  String? _overrunBlockNote(_ActiveStopContext stop) {
    if (!_hasLargeStopPositionMismatch(stop)) {
      return null;
    }
    if (stop.hadNextStaChange) {
      return '정차 위치와 다음역 정보 간 불일치 정황이 확인됩니다. TWC/TB 정위치 조건 미형성 가능성이 있습니다.';
    }
    return '정차 시점에 잔여거리가 큰 상태에서 출입문 취급 정황이 확인됩니다. TWC/TB 정위치 조건 미형성 가능성이 있습니다.';
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

  String _firstMatchingTime(List<LogEntry> entries, List<String> keywords) {
    for (final entry in entries) {
      final time = entry.time.trim();
      if (time.isEmpty || time == '-') continue;
      if (keywords.any((keyword) => entry.message.contains(keyword))) {
        return time;
      }
    }
    return '';
  }

  String _summaryLead(String time, String label) {
    return time.isEmpty ? '' : '최초 $time $label, ';
  }

  String _summaryTimeSuffix(List<LogEntry> entries) {
    final range = _entryTimeRange(entries);
    return range.isEmpty ? '' : ' 주요 시간대: $range.';
  }

  String _incidentTypeLabel(IncidentType type) {
    switch (type) {
      case IncidentType.doorCloseFailure:
        return '출입문 닫히지 않음 발생 정황';
      case IncidentType.doorOpenFailure:
        return '출입문 열림불량 발생 정황';
      case IncidentType.doorBypass:
        return '출입문 바이패스 현시/취급 정황';
      case IncidentType.ncodeEmergency:
        return '속도코드 무코드 발생 및 비상운전 흐름';
      case IncidentType.brake:
        return '전상용제동(FSB) 체결 흐름';
      case IncidentType.modeChange:
        return '운전모드 전환 흐름';
      case IncidentType.doorFlow:
        return '출입문 취급 흐름';
      case IncidentType.general:
        return '주요 신호 변화 흐름';
    }
  }

  IncidentType _classifyIncidentType(List<LogEntry> entries) {
    final messages = entries.map((entry) => entry.message).join('\n');
    final hasDoorBypass =
        messages.contains('출입문 바이패스') || messages.contains('바이패스 취급');
    final hasDoorOpenFailure =
        messages.contains('출입문 열림 지연') ||
        messages.contains('열림 버튼 취급, ALL DOOR CLOSE "1"') ||
        messages.contains('열림 버튼 취급 후에도 ALL DOOR CLOSE "1" 유지');
    final hasDoorCloseFailure =
        messages.contains('닫힘버튼 취급, ALL DOOR CLOSE "0"') ||
        messages.contains('닫힘 신호 지연') ||
        messages.contains('전체 출입문 닫힘 신호 미형성');
    final hasDoorFlow =
        messages.contains('출입문') || messages.contains('ALL DOOR CLOSE');
    final hasNCode = messages.contains('무코드');
    final hasEmergency = messages.contains('비상운전') || messages.contains('비상모드');
    final hasFsbr = messages.contains('전상용제동(FSB) 체결');
    final hasMode = messages.contains('전환된 정황');

    if (hasDoorCloseFailure) return IncidentType.doorCloseFailure;
    if (hasDoorOpenFailure) return IncidentType.doorOpenFailure;
    if (hasDoorBypass) return IncidentType.doorBypass;
    if (hasNCode && (hasFsbr || hasEmergency)) {
      return IncidentType.ncodeEmergency;
    }
    if (hasFsbr) return IncidentType.brake;
    if (hasMode) return IncidentType.modeChange;
    if (hasDoorFlow) return IncidentType.doorFlow;
    return IncidentType.general;
  }

  String _firstLine(
    List<LogEntry> entries,
    List<String> keywords,
    String label,
  ) {
    final time = _firstMatchingTime(entries, keywords);
    if (time.isEmpty) return '';
    return '- 최초 $label: $time';
  }

  int _countAnyMessages(List<LogEntry> entries, List<String> keywords) {
    return entries
        .where(
          (entry) => keywords.any((keyword) => entry.message.contains(keyword)),
        )
        .length;
  }

  String _reportDigestLine({
    required List<LogEntry> entries,
    required List<String> keywords,
    required String label,
  }) {
    final count = _countAnyMessages(entries, keywords);
    if (count == 0) return '';
    final firstTime = _firstMatchingTime(entries, keywords);
    final firstText = firstTime.isEmpty ? '' : ', 최초 $firstTime';
    return '- $label: $count건$firstText';
  }

  List<String> _buildReportDigest(List<LogEntry> entries) {
    return [
      _reportDigestLine(
        entries: entries,
        keywords: ['출입문 바이패스', '바이패스 취급'],
        label: '출입문 바이패스 현시/취급',
      ),
      _reportDigestLine(
        entries: entries,
        keywords: [
          '전체 출입문 닫힘 신호 미형성',
          '전체 출입문 닫힘 신호가 형성되지',
          '닫힘버튼 취급, ALL DOOR CLOSE "0"',
        ],
        label: '출입문 닫히지 않음 정황',
      ),
      _reportDigestLine(
        entries: entries,
        keywords: ['출입문 열림 지연', '열림 버튼 취급 후에도 ALL DOOR CLOSE "1" 유지'],
        label: '출입문 열림불량 정황',
      ),
      _reportDigestLine(
        entries: entries,
        keywords: ['무코드'],
        label: '속도코드 무코드 발생',
      ),
      _reportDigestLine(
        entries: entries,
        keywords: ['전상용제동(FSB) 체결'],
        label: '전상용제동(FSB) 체결',
      ),
      _reportDigestLine(
        entries: entries,
        keywords: ['속도코드', '과속'],
        label: 'ATC-과속검지 정황',
      ),
      _reportDigestLine(
        entries: entries,
        keywords: ['이동 중 ALL DOOR CLOSE'],
        label: '이동 중 출입문 닫힘 신호 해제',
      ),
      _reportDigestLine(
        entries: entries,
        keywords: ['TWC/TB'],
        label: 'TWC/TB 정위치 조건 의심',
      ),
      _reportDigestLine(
        entries: entries,
        keywords: ['전환된 정황'],
        label: '운전모드 전환',
      ),
      _reportDigestLine(
        entries: entries,
        keywords: ['비상모드', '비상운전'],
        label: '비상모드/비상운전',
      ),
    ].where((line) => line.isNotEmpty).toList();
  }

  bool _isGenericReportEvent(LogEntry entry) {
    final message = entry.message;
    return message == '정차가 감지되었습니다.' ||
        message == '발차가 감지되었습니다.' ||
        message.contains('정차 중 / 다음역') ||
        message.contains('정차 중으로 해석됩니다.') ||
        message.contains('방면으로 운행 중인 정황') ||
        message.contains('발차 후') ||
        message.contains('정차 중이며 다음역');
  }

  int _reportEventPriority(LogEntry entry) {
    final message = entry.message;
    if (_isGenericReportEvent(entry)) return 100;
    if (message.contains('바이패스') ||
        message.contains('전체 출입문 닫힘 신호 미형성') ||
        message.contains('전체 출입문 닫힘 신호가 형성되지') ||
        message.contains('불일치') ||
        message.contains('운전모드 신호 불일치') ||
        message.contains('이동 중 ALL DOOR CLOSE')) {
      return 0;
    }
    if (message.contains('무코드') ||
        message.contains('전상용제동(FSB) 체결') ||
        message.contains('전환된 정황') ||
        message.contains('비상모드') ||
        message.contains('비상운전') ||
        message.contains('정차 위치') ||
        message.contains('잔여거리') ||
        message.contains('속도코드') ||
        message.contains('TWC/TB')) {
      return 1;
    }
    if (message.contains('열림 버튼 취급') ||
        message.contains('닫힘버튼 취급') ||
        message.contains('재개폐 버튼 취급') ||
        message.contains('출입문 열림 지연') ||
        message.contains('ALL DOOR CLOSE "0"')) {
      return 2;
    }
    if (message.contains('ALL DOOR CLOSE "1"') ||
        message.contains('정지 후 진행(S&P)')) {
      return 3;
    }
    if (entry.type == EntryType.bypass) return 2;
    if (entry.isSummary) return 4;
    return 100;
  }

  String _reportEventSignature(String message) {
    if (message.contains('바이패스')) return 'door_bypass';
    if (message.contains('이동 중 ALL DOOR CLOSE')) return 'door_motion_open';
    if (message.contains('전체 출입문 닫힘 신호 미형성') || message.contains('불일치')) {
      return 'door_mismatch';
    }
    if (message.contains('열림')) return 'door_open';
    if (message.contains('닫힘') || message.contains('ALL DOOR CLOSE')) {
      return 'door_close';
    }
    if (message.contains('무코드')) return 'ncode';
    if (message.contains('전상용제동(FSB) 체결')) return 'fsb';
    if (message.contains('속도코드')) return 'speed_code';
    if (message.contains('전환된 정황')) return 'mode_change';
    if (message.contains('비상모드') || message.contains('비상운전')) {
      return 'emergency';
    }
    if (message.contains('정차 위치') || message.contains('잔여거리')) {
      return 'position';
    }
    if (message.contains('정지 후 진행(S&P)')) return 'sp';
    return message;
  }

  int _reportEventRepeatLimit(int priority) {
    if (priority == 0) return 8;
    if (priority == 1) return 6;
    if (priority == 2) return 4;
    return 2;
  }

  List<LogEntry> _selectReportEvents(List<LogEntry> entries) {
    final ranked = <({int index, LogEntry entry, int priority})>[];
    for (var i = 0; i < entries.length; i++) {
      final priority = _reportEventPriority(entries[i]);
      if (priority < 100) {
        ranked.add((index: i, entry: entries[i], priority: priority));
      }
    }
    ranked.sort((a, b) {
      final priorityCompare = a.priority.compareTo(b.priority);
      if (priorityCompare != 0) return priorityCompare;
      return a.index.compareTo(b.index);
    });

    final hasCoreEvents = ranked.any((candidate) => candidate.priority <= 1);
    final selectedIndices = <int>{};
    final signatureCounts = <String, int>{};
    for (final candidate in ranked) {
      if (hasCoreEvents && candidate.priority > 1) {
        continue;
      }
      final signature = _reportEventSignature(candidate.entry.message);
      final count = signatureCounts[signature] ?? 0;
      if (count >= _reportEventRepeatLimit(candidate.priority)) {
        continue;
      }
      selectedIndices.add(candidate.index);
      signatureCounts[signature] = count + 1;
      if (selectedIndices.length >= 16) break;
    }

    if (selectedIndices.isEmpty) {
      for (final candidate in ranked) {
        final signature = _reportEventSignature(candidate.entry.message);
        final count = signatureCounts[signature] ?? 0;
        if (count >= _reportEventRepeatLimit(candidate.priority)) {
          continue;
        }
        selectedIndices.add(candidate.index);
        signatureCounts[signature] = count + 1;
        if (selectedIndices.length >= 10) break;
      }
    }

    final orderedIndices = selectedIndices.toList()..sort();
    return orderedIndices.map((index) => entries[index]).toList();
  }

  String _buildReportDraft(List<LogEntry> entries, List<AnalysisBlock> blocks) {
    if (entries.isEmpty) {
      return '운행기록 분석 초안\n\n감지된 주요 이벤트가 없습니다.';
    }

    final incidentType = _incidentTypeLabel(_classifyIncidentType(entries));
    final timeRange = _entryTimeRange(entries);
    final summary = _buildSummary(entries);
    final digestLines = _buildReportDigest(entries);
    final firstLines = [
      _firstLine(
        entries,
        ['출입문 열림 지연', '열림 버튼 취급 후에도 ALL DOOR CLOSE "1" 유지'],
        '출입문 열림불량 관련 신호',
      ),
      _firstLine(
        entries,
        ['닫힘버튼 취급, ALL DOOR CLOSE "0"', '전체 출입문 닫힘 신호가 형성되지'],
        '출입문 닫히지 않음 관련 신호',
      ),
      _firstLine(entries, ['출입문 바이패스', '바이패스 취급'], '출입문 바이패스'),
      _firstLine(entries, ['무코드'], '속도코드 무코드 발생'),
      _firstLine(entries, ['전상용제동(FSB) 체결'], '전상용제동(FSB) 체결'),
      _firstLine(entries, ['전환된 정황'], '운전모드 전환'),
    ].where((line) => line.isNotEmpty).toList();

    final summaryEntries = _selectReportEvents(entries)
        .map(
          (entry) =>
              '- ${entry.time == '-' ? '' : '${entry.time} '}${entry.message}',
        )
        .toList();
    final blockLines = blocks.take(6).map((block) {
      final range = _entryTimeRange(block.entries);
      return '- ${block.title}${range.isEmpty ? '' : ' ($range)'}';
    }).toList();

    return [
      '운행기록 분석 초안',
      '',
      '1. 고장개요',
      '- 분류: $incidentType',
      if (timeRange.isNotEmpty) '- 주요 시간대: $timeRange',
      '- 분석 범위: 업로드된 운행기록 신호 기준',
      '',
      '2. 주요 흐름',
      summary,
      '',
      if (digestLines.isNotEmpty) '3. 핵심 이벤트 요약',
      if (digestLines.isNotEmpty) ...digestLines,
      if (digestLines.isNotEmpty) '',
      if (firstLines.isNotEmpty) '4. 최초 확인 신호',
      if (firstLines.isNotEmpty) ...firstLines,
      if (firstLines.isNotEmpty) '',
      if (summaryEntries.isNotEmpty) '5. 주요 이벤트',
      if (summaryEntries.isNotEmpty) ...summaryEntries,
      if (summaryEntries.isNotEmpty) '',
      if (blockLines.isNotEmpty) '6. 블록 흐름',
      if (blockLines.isNotEmpty) ...blockLines,
      if (blockLines.isNotEmpty) '',
      '7. 유의사항',
      '- 본 초안은 운행기록 기반의 사건 흐름 정리이며, 원인 확정 또는 조치 지시가 아닙니다.',
      '- 필요 시 현장 확인 결과와 정비 기록을 함께 대조해야 합니다.',
    ].join('\n');
  }

  Future<void> _allowUiRefresh([String? nextStatus]) async {
    if (!mounted) return;
    if (nextStatus != null) {
      setState(() {
        statusText = nextStatus;
      });
    }
    await Future<void>.delayed(Duration.zero);
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
    final incidentType = _incidentTypeLabel(_classifyIncidentType(entries));
    final hasDoor =
        messages.contains('ALL DOOR CLOSE') || messages.contains('출입문');
    final hasDoorCloseCycle =
        messages.contains('닫힘버튼 취급') || messages.contains('ALL DOOR CLOSE "1"');
    final hasDoorOpenCycle =
        messages.contains('열림 버튼 취급') ||
        messages.contains('ALL DOOR CLOSE "0"');
    final hasDoorBypass =
        messages.contains('출입문 바이패스') || messages.contains('바이패스 취급');
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
    final doorCloseCycleCount =
        _countMessages(entries, '닫힘버튼 취급') +
        _countMessages(entries, 'ALL DOOR CLOSE "1"');
    final doorOpenCycleCount =
        _countMessages(entries, '열림 버튼 취급') +
        _countMessages(entries, 'ALL DOOR CLOSE "0"');
    final doorBypassCount =
        _countMessages(entries, '출입문 바이패스') +
        _countMessages(entries, '바이패스 취급');
    final firstNCodeTime = _firstMatchingTime(entries, ['무코드']);
    final firstFsbrTime = _firstMatchingTime(entries, ['전상용제동(FSB) 체결']);
    final firstModeTime = _firstMatchingTime(entries, ['전환된 정황']);
    final firstDoorCloseTime = _firstMatchingTime(entries, [
      '닫힘버튼 취급',
      'ALL DOOR CLOSE "1"',
    ]);
    final firstDoorOpenTime = _firstMatchingTime(entries, [
      '열림 버튼 취급',
      'ALL DOOR CLOSE "0"',
    ]);
    final firstDoorBypassTime = _firstMatchingTime(entries, [
      '출입문 바이패스',
      '바이패스 취급',
    ]);
    final firstDepartureTime = _firstMatchingTime(entries, ['출발 허가', '발차']);
    final timeSuffix = _summaryTimeSuffix(entries);

    String withType(String body) => '분류: $incidentType. $body';

    if (hasDoorBypass) {
      return withType(
        '${_summaryLead(firstDoorBypassTime, '출입문 바이패스')}출입문 바이패스 취급 정황이 $doorBypassCount건 확인됩니다.$timeSuffix',
      );
    }
    if (hasDoorOpenCycle &&
        _classifyIncidentType(entries) == IncidentType.doorOpenFailure) {
      return withType(
        '${_summaryLead(firstDoorOpenTime, '출입문 열림 관련 신호')}출입문 열림 불량 정황이 확인됩니다.$timeSuffix',
      );
    }
    if (_classifyIncidentType(entries) == IncidentType.doorCloseFailure) {
      return withType(
        '${_summaryLead(firstDoorCloseTime, '출입문 닫힘 관련 신호')}출입문 닫힘 불량 정황이 확인됩니다.$timeSuffix',
      );
    }
    if (hasNCode && hasFsbr && hasMode) {
      return withType(
        '${_summaryLead(firstNCodeTime, '무코드 수신')}무코드 수신 $nCodeCount건, 전상용제동(FSB) 체결 $fsbrCount건, 운전모드 전환 $modeCount건이 확인됩니다.$timeSuffix',
      );
    }
    if (hasDoorCloseCycle && hasDeparture) {
      return withType(
        '${_summaryLead(firstDoorCloseTime, '출입문 닫힘 사이클')}출입문 닫힘 사이클 $doorCloseCycleCount건과 출발 흐름이 확인됩니다.$timeSuffix',
      );
    }
    if (hasDoorCloseCycle) {
      return withType(
        '${_summaryLead(firstDoorCloseTime, '출입문 닫힘 관련 신호')}전체 출입문 닫힘 형성 흐름이 $doorCloseCycleCount건 확인됩니다.$timeSuffix',
      );
    }
    if (hasDoorOpenCycle) {
      return withType(
        '${_summaryLead(firstDoorOpenTime, '출입문 열림 관련 신호')}출입문 열림 흐름이 $doorOpenCycleCount건 확인됩니다.$timeSuffix',
      );
    }
    if (hasNCode && hasFsbr) {
      return withType(
        '${_summaryLead(firstNCodeTime, '무코드 수신')}무코드 수신 $nCodeCount건 이후 전상용제동(FSB) 체결 흐름이 $fsbrCount건 확인됩니다.$timeSuffix',
      );
    }
    if (hasDoor && hasDeparture) {
      return withType(
        '${_summaryLead(firstDepartureTime, '출발 흐름')}출입문 상태 변화와 발차 흐름이 함께 확인됩니다.$timeSuffix',
      );
    }
    if (hasDoor) {
      return withType('출입문 상태 변화 정황이 확인됩니다.$timeSuffix');
    }
    if (hasMode && hasFsbr) {
      return withType(
        '${_summaryLead(firstModeTime, '운전모드 전환')}전상용제동(FSB) 체결 흐름이 $fsbrCount건 확인됩니다.$timeSuffix',
      );
    }
    if (hasMode) {
      return withType(
        '${_summaryLead(firstModeTime, '운전모드 전환')}운전모드 전환 흐름이 $modeCount건 확인됩니다.$timeSuffix',
      );
    }
    if (hasFsbr) {
      return withType(
        '${_summaryLead(firstFsbrTime, '전상용제동(FSB) 체결')}전상용제동(FSB) 체결 정황이 $fsbrCount건 확인됩니다.$timeSuffix',
      );
    }
    if (hasStop && hasDeparture) {
      return withType(
        '${_summaryLead(firstDepartureTime, '발차')}정차와 발차 흐름이 확인됩니다.$timeSuffix',
      );
    }
    return withType('주요 신호 변화 흐름이 확인됩니다.$timeSuffix');
  }

  List<String> _headerTexts(List<dynamic> row) {
    return row.map((cell) => _cellToStr(cell).toUpperCase()).toList();
  }

  int _findHeaderIndex(List<String> header, String name) {
    final target = name.toUpperCase();
    final exact = header.indexOf(target);
    return exact != -1 ? exact : header.indexWhere((h) => h.contains(target));
  }

  void _validateRequiredHeader(List<String> header) {
    final requiredColumns = ['TIME', 'NEXTSTA', 'DIST', 'VEL', 'ADC'];
    final missing = requiredColumns
        .where((name) => _findHeaderIndex(header, name) == -1)
        .toList();
    if (missing.isNotEmpty) {
      throw '운행기록 필수 컬럼이 없습니다: ${missing.join(', ')}';
    }
  }

  Future<_ParsedLogFile> _readLogFile(PlatformFile file) async {
    final bytes = file.bytes;
    if (bytes == null) {
      throw '${file.name} 파일 데이터를 읽을 수 없습니다.';
    }

    final rows = <List<dynamic>>[];
    if (_isXlsFile(file.name, bytes)) {
      await _allowUiRefresh('${file.name} 구형 Excel(.xls) 파일을 변환 중입니다...');
      rows.addAll(await _readLegacyXlsRows(file.name, bytes));
    } else if (bytes.length > 2 && bytes[0] == 0x50 && bytes[1] == 0x4B) {
      await _allowUiRefresh('${file.name} 엑셀 파일을 읽는 중입니다...');
      final excel = Excel.decodeBytes(bytes);
      if (excel.tables.isEmpty) {
        throw '${file.name} 엑셀 시트를 찾을 수 없습니다.';
      }
      var table = excel.tables.values.first;
      for (final sheet in excel.tables.values) {
        if (sheet.maxRows > 0) {
          table = sheet;
          break;
        }
      }
      var rowCount = 0;
      for (final row in table.rows) {
        rows.add(row.toList());
        rowCount++;
        if (rowCount % kUiYieldInterval == 0) {
          await _allowUiRefresh('${file.name} 엑셀 데이터를 읽는 중입니다... ($rowCount행)');
        }
      }
    } else {
      final lines = utf8
          .decode(bytes, allowMalformed: true)
          .split(RegExp(r'\r?\n'))
          .where((line) => line.trim().isNotEmpty)
          .toList();
      for (var lineIndex = 0; lineIndex < lines.length; lineIndex++) {
        rows.add(lines[lineIndex].split(','));
        if ((lineIndex + 1) % kUiYieldInterval == 0) {
          await _allowUiRefresh(
            '${file.name} CSV 데이터를 읽는 중입니다... (${lineIndex + 1}행)',
          );
        }
      }
    }

    if (rows.isEmpty) {
      throw '${file.name} 데이터가 비어 있습니다.';
    }
    _validateRequiredHeader(_headerTexts(rows.first));
    return _ParsedLogFile(name: file.name, rows: rows);
  }

  Duration? _parseTimeFilterText(String value) {
    final text = value.trim();
    if (text.isEmpty) return null;
    if (!RegExp(r'^\d{1,2}:\d{2}:\d{2}(?:\.\d{1,3})?$').hasMatch(text)) {
      throw '시간 형식은 HH:mm:ss로 입력해 주세요.';
    }
    final parsed = _parseLogTime(text);
    if (parsed == null) {
      throw '시간 형식은 HH:mm:ss로 입력해 주세요.';
    }
    return parsed;
  }

  String _timeFilterScopeLabel(Duration? start, Duration? end) {
    String format(Duration value) {
      final hours = value.inHours.toString().padLeft(2, '0');
      final minutes = (value.inMinutes % 60).toString().padLeft(2, '0');
      final seconds = (value.inSeconds % 60).toString().padLeft(2, '0');
      return '$hours:$minutes:$seconds';
    }

    if (start != null && end != null) {
      return '${format(start)}~${format(end)} 구간';
    }
    if (start != null) {
      return '${format(start)} 이후 구간';
    }
    if (end != null) {
      return '${format(end)} 이전 구간';
    }
    return '';
  }

  _PreparedLogRows _prepareLogRows(List<_ParsedLogFile> parsedFiles) {
    if (parsedFiles.isEmpty) {
      throw '선택된 운행기록 파일이 없습니다.';
    }

    final canonicalHeader = parsedFiles.first.rows.first;
    final canonicalHeaderText = _headerTexts(canonicalHeader);
    _validateRequiredHeader(canonicalHeaderText);
    final canonicalTimeIdx = _findHeaderIndex(canonicalHeaderText, 'TIME');
    final dataRows =
        <({List<dynamic> row, int order, Duration? time, String fileName})>[];
    var order = 0;

    for (final parsedFile in parsedFiles) {
      final header = _headerTexts(parsedFile.rows.first);
      _validateRequiredHeader(header);
      final comparableColumns = canonicalHeaderText
          .where((name) => name.trim().isNotEmpty)
          .toList();
      final matchedColumns = comparableColumns
          .where((name) => _findHeaderIndex(header, name) != -1)
          .length;
      final minimumMatches = comparableColumns.isEmpty
          ? 0
          : (comparableColumns.length * 0.7).ceil();
      if (parsedFiles.length > 1 && matchedColumns < minimumMatches) {
        throw '두 파일의 운행기록 형식이 달라 병합할 수 없습니다. 동일한 운행기록 양식의 파일을 선택해 주세요.';
      }

      final mapping = canonicalHeaderText
          .map((name) => _findHeaderIndex(header, name))
          .toList();
      for (final row in parsedFile.rows.skip(1)) {
        final aligned = List<dynamic>.generate(canonicalHeader.length, (index) {
          final sourceIndex = mapping[index];
          return sourceIndex != -1 && sourceIndex < row.length
              ? row[sourceIndex]
              : '';
        });
        final time = canonicalTimeIdx != -1 && canonicalTimeIdx < aligned.length
            ? _parseLogTime(_cellToStr(aligned[canonicalTimeIdx]))
            : null;
        dataRows.add((
          row: aligned,
          order: order++,
          time: time,
          fileName: parsedFile.name,
        ));
      }
    }

    dataRows.sort((a, b) {
      final aTime = a.time;
      final bTime = b.time;
      if (aTime != null && bTime != null && aTime != bTime) {
        return aTime.compareTo(bTime);
      }
      return a.order.compareTo(b.order);
    });

    final start = _parseTimeFilterText(_startTimeController.text);
    final end = _parseTimeFilterText(_endTimeController.text);
    if (start != null && end != null && start > end) {
      throw '종료 시간은 시작 시간 이후로 입력해 주세요.';
    }

    Iterable<({String fileName, int order, List<dynamic> row, Duration? time})>
    selectedRows = dataRows;
    final scopeLabel = _timeFilterScopeLabel(start, end);
    if (start != null || end != null) {
      selectedRows = dataRows.where((entry) {
        final time = entry.time;
        if (time == null) return false;
        if (start != null && time < start) return false;
        if (end != null && time > end) return false;
        return true;
      });
      if (selectedRows.isEmpty) {
        throw '선택한 시간대에 해당하는 운행기록이 없습니다.';
      }
    }

    final rows = <List<dynamic>>[
      canonicalHeader,
      ...selectedRows.map((entry) => entry.row),
    ];
    return _PreparedLogRows(rows: rows, scopeLabel: scopeLabel);
  }

  Future<void> _pickAndAnalyze() async {
    final result = await FilePicker.platform.pickFiles(
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

    setState(() {
      isLoading = true;
      logs = [];
      blocks = [];
      summaryText = '';
      reportText = '';
      _lastParsedFiles = const [];
      statusText = '전동차 운행로그를 분석 중입니다...';
    });
    await Future<void>.delayed(const Duration(milliseconds: 16));

    try {
      _startTimeController.clear();
      _endTimeController.clear();
      final parsedFiles = <_ParsedLogFile>[];
      for (var fileIndex = 0; fileIndex < result.files.length; fileIndex++) {
        final file = result.files[fileIndex];
        if (result.files.length == 2) {
          await _allowUiRefresh('파일 ${fileIndex + 1} / 2 병합 준비 중...');
        }
        parsedFiles.add(await _readLogFile(file));
      }
      if (parsedFiles.length == 2) {
        await _allowUiRefresh(
          '파일 2개 병합 중... 1번 파일 ${parsedFiles[0].rows.length - 1}행 / 2번 파일 ${parsedFiles[1].rows.length - 1}행',
        );
      }

      _lastParsedFiles = List.unmodifiable(parsedFiles);
      await _analyzeParsedFiles(parsedFiles);
    } catch (e) {
      setState(() {
        isLoading = false;
        statusText = '오류: $e';
      });
    }
  }

  Future<void> _reAnalyzeSelectedTime() async {
    if (_lastParsedFiles.isEmpty) return;
    try {
      final start = _parseTimeFilterText(_startTimeController.text);
      final end = _parseTimeFilterText(_endTimeController.text);
      if (start != null && end != null && start > end) {
        throw '종료 시간은 시작 시간 이후로 입력해 주세요.';
      }
    } catch (e) {
      setState(() {
        statusText = '오류: $e';
      });
      return;
    }
    setState(() {
      isLoading = true;
      logs = [];
      blocks = [];
      summaryText = '';
      reportText = '';
      statusText = '선택 시간대를 기준으로 재분석 중입니다...';
    });
    await Future<void>.delayed(const Duration(milliseconds: 16));
    await _analyzeParsedFiles(_lastParsedFiles);
  }

  Future<void> _analyzeParsedFiles(List<_ParsedLogFile> parsedFiles) async {
    try {
      final preparedRows = _prepareLogRows(parsedFiles);
      final rows = preparedRows.rows;

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
      final speedCodeIndices = <int, List<int>>{
        0: findIndices(['ATC1/2 0', 'ATC1 0', 'ATC2 0']),
        15: findIndices(['ATC1/2 15', 'ATC1 15', 'ATC2 15']),
        25: findIndices(['ATC1/2 25', 'ATC1 25', 'ATC2 25']),
        30: findIndices(['ATC1/2 30', 'ATC1 30', 'ATC2 30']),
        35: findIndices(['ATC1/2 35', 'ATC1 35', 'ATC2 35']),
        40: findIndices(['ATC1/2 40', 'ATC1 40', 'ATC2 40']),
        45: findIndices(['ATC1/2 45', 'ATC1 45', 'ATC2 45']),
        50: findIndices(['ATC1/2 50', 'ATC1 50', 'ATC2 50']),
        55: findIndices(['ATC1/2 55', 'ATC1 55', 'ATC2 55']),
        60: findIndices(['ATC1/2 60', 'ATC1 60', 'ATC2 60']),
        65: findIndices(['ATC1/2 65', 'ATC1 65', 'ATC2 65']),
        70: findIndices(['ATC1/2 70', 'ATC1 70', 'ATC2 70']),
        75: findIndices(['ATC1/2 75', 'ATC1 75', 'ATC2 75']),
        80: findIndices(['ATC1/2 80', 'ATC1 80', 'ATC2 80']),
      }..removeWhere((_, indices) => indices.isEmpty);
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
        'S_REOPEN': findIdx('S_REOPEN'),
      };

      final autoDoorOpenCommands = <String, List<int>>{
        'ATC1/2OPEN-L': findIndices([
          'ATC1/2OPEN-L',
          'ATC1OPEN-L',
          'ATC2OPEN-L',
        ]),
        'ATC1/2OPEN-R': findIndices([
          'ATC1/2OPEN-R',
          'ATC1OPEN-R',
          'ATC2OPEN-R',
        ]),
      }..removeWhere((_, indices) => indices.isEmpty);

      final eventButtons = <String, int>{
        'PAN UP': findIdx('PAN UP'),
        'PAN DN': findIdx('PAN DN'),
        'EMPB': findIdx('EMPB'),
        'EBCOS': findIdx('EBCOS'),
      };

      final doorIdx = {for (final door in kActiveDoors) door: findIdx(door)};
      final tempLogs = <LogEntry>[];

      final firstDataRow = rows.length > 1 ? rows[1] : const <dynamic>[];
      var stableAdcValue = idxAdc != -1 ? _getVal(firstDataRow, idxAdc) : '1';
      var stableEmegrValue = idxEmegr != -1
          ? _getVal(firstDataRow, idxEmegr)
          : '0';
      bool emegrActive = stableEmegrValue == '1';
      var stableAdbsValue = idxAdbs != -1
          ? _getVal(firstDataRow, idxAdbs)
          : '0';
      final stableDoorValues = <String, String>{
        for (final entry in doorIdx.entries)
          entry.key: entry.value == -1
              ? '1'
              : _getVal(firstDataRow, entry.value),
      };
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
      Duration highSpeedDuration = Duration.zero;
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

      final initialActiveModes = _activeModes(firstDataRow, modeIndices);
      final initialMode = initialActiveModes.length == 1
          ? initialActiveModes.first
          : OperationMode.unknown;
      OperationMode? lastKnownMode = initialMode == OperationMode.unknown
          ? null
          : initialMode;
      bool modeConflictActive = initialActiveModes.length > 1;
      Duration? modeConflictStartedAt = modeConflictActive
          ? Duration.zero
          : null;
      bool modeConflictLogged = false;
      List<OperationMode> modeConflictModes = initialActiveModes;
      int? lastDoorCloseIndex;
      String? lastDoorCloseLocation;
      String? lastDoorCloseMode;
      List<String> lastDoorCloseBypassCars = const [];
      Duration? lastDoorCloseAt;
      Duration? lastDoorCycleClosedAt;
      int? lastDptLoggedCloseIndex;
      Duration? lastModeChangeDptConsumedAt;
      bool doorClosePending = false;
      bool doorCloseDelayLogged = false;
      bool doorCloseReopenLogged = false;
      bool doorCloseBypassLogged = false;
      bool lastDoorCycleHadDelay = false;
      String? lastDoorOpenLocation;
      String? lastDoorOpenMode;
      String? lastDoorOpenLabel;
      Duration? lastDoorOpenAt;
      bool doorOpenPending = false;
      bool doorOpenDelayLogged = false;
      String? lastDoorBypassLocation;
      List<String> lastDoorBypassCars = const [];
      Duration? lastDoorBypassAt;
      bool doorBypassPending = false;
      bool doorBypassHoldLogged = false;
      final totalRecords = rows.length > 1 ? rows.length - 1 : 0;

      await _allowUiRefresh('운행기록을 분석 중입니다... (0 / $totalRecords)');

      for (var i = 1; i < rows.length; i++) {
        if (i % kUiYieldInterval == 0) {
          await _allowUiRefresh('운행기록을 분석 중입니다... ($i / $totalRecords)');
        }

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

        final prevAdcValue = stableAdcValue;
        final rawCurrAdcValue = idxAdc != -1
            ? _getVal(curr, idxAdc)
            : stableAdcValue;
        if (rawCurrAdcValue != stableAdcValue &&
            _isSustainedValueAt(
              rows,
              i,
              idxAdc,
              tIdx,
              rawCurrAdcValue,
              kSignalSpikeIgnoreDuration,
            )) {
          stableAdcValue = rawCurrAdcValue;
        }
        final currAdcValue = stableAdcValue;
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
        if (hasModeConflict) {
          if (!modeConflictActive) {
            modeConflictActive = true;
            modeConflictStartedAt = rowClock;
            modeConflictLogged = false;
          }
          modeConflictModes = activeModes;

          final conflictDuration =
              rowClock - (modeConflictStartedAt ?? rowClock);
          if (!modeConflictLogged &&
              conflictDuration >= kModeConflictConfirmDuration) {
            tempLogs.add(
              LogEntry(
                time: time,
                message:
                    '운전모드 신호 불일치가 ${_formatElapsed(conflictDuration)} 이상 지속되었습니다: ${modeConflictModes.map(_modeLabel).join(', ')}',
                type: EntryType.bypass,
                isSummary: true,
              ),
            );
            modeConflictLogged = true;
          }
        } else if (modeConflictActive) {
          if (modeConflictLogged) {
            tempLogs.add(
              LogEntry(
                time: time,
                message: '운전모드 신호 불일치 상태가 해소된 정황이 확인됩니다.',
                type: EntryType.restore,
                isSummary: true,
              ),
            );
          }
          modeConflictActive = false;
          modeConflictStartedAt = null;
          modeConflictLogged = false;
          modeConflictModes = const <OperationMode>[];
        }

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
          final rawEmegrValue = _getVal(curr, idxEmegr);
          if (rawEmegrValue != stableEmegrValue &&
              _isSustainedValueAt(
                rows,
                i,
                idxEmegr,
                tIdx,
                rawEmegrValue,
                kSignalSpikeIgnoreDuration,
              )) {
            stableEmegrValue = rawEmegrValue;
          }
          final currentEmegr = stableEmegrValue == '1';
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

        for (final entry in doorButtons.entries) {
          if (_isRisingEdge(prev, curr, entry.value)) {
            final isOpenButton =
                entry.key == 'OPEN-L' ||
                entry.key == 'OPEN-R' ||
                entry.key == 'S_OPEN-L' ||
                entry.key == 'S_OPEN-R';
            final isCloseButton =
                entry.key == 'CLOSE' || entry.key == 'S_CLOSE';
            final isReopenButton =
                entry.key == 'REOPEN' || entry.key == 'S_REOPEN';
            rowHadDoorActivity = true;
            rowHadDoorCycle = true;

            final message = isReopenButton && doorClosePending
                ? '${_doorButtonActionLabel(entry.key)}, ${_adcReportValue(currAdcValue)}${_doorModeSuffix(lastDoorCloseMode ?? currentDoorMode)}${_currentStationSuffix(lastDoorCloseLocation ?? doorLocation)}'
                : _doorButtonReportMessage(
                    signal: entry.key,
                    adcValue: currAdcValue,
                    doorMode: currentDoorMode,
                    location: doorLocation,
                  );
            tempLogs.add(
              LogEntry(
                time: time,
                message: message,
                type: EntryType.button,
                isSummary: isCloseButton || isOpenButton || isReopenButton,
              ),
            );

            if (isOpenButton) {
              lastDoorOpenLocation = doorLocation;
              lastDoorOpenMode = currentDoorMode;
              lastDoorOpenLabel = _doorButtonActionLabel(entry.key);
              lastDoorOpenAt = rowClock;
              doorOpenPending = currAdcValue == '1';
              doorOpenDelayLogged = false;
            } else if (isCloseButton) {
              lastDoorCloseIndex = i;
              lastDoorCloseLocation = doorLocation;
              lastDoorCloseMode = currentDoorMode;
              lastDoorCloseBypassCars = _activeDoorBypassCars(curr, doorIdx);
              lastDoorCloseAt = rowClock;
              lastDoorCycleClosedAt = null;
              doorClosePending = currAdcValue == '0';
              doorCloseDelayLogged = false;
              doorCloseReopenLogged = false;
              doorCloseBypassLogged = false;
              lastDoorCycleHadDelay = false;
              doorOpenPending = false;
            } else if (isReopenButton &&
                doorClosePending &&
                !doorCloseReopenLogged) {
              doorCloseReopenLogged = true;
              lastDoorCycleHadDelay = true;
            }
          }
        }

        final isAutoMode =
            _getVal(curr, modeIndices[OperationMode.auto] ?? -1) == '1';
        if (isAutoMode) {
          for (final entry in autoDoorOpenCommands.entries) {
            if (_isAnyRisingEdge(prev, curr, entry.value)) {
              final label = _autoDoorOpenCommandLabel(entry.key);
              rowHadDoorActivity = true;
              rowHadDoorCycle = true;
              final stop = activeStop;
              if (stop != null && isStopped) {
                stop.hadDoorActivity = true;
                stop.hadDoorCycle = true;
              }
              tempLogs.add(
                LogEntry(
                  time: time,
                  message:
                      '$label, ${_adcReportValue(currAdcValue)}${_currentStationSuffix(doorLocation)}',
                  type: EntryType.info,
                  isSummary: true,
                ),
              );
              lastDoorOpenLocation = doorLocation;
              lastDoorOpenMode = currentDoorMode;
              lastDoorOpenLabel = label;
              lastDoorOpenAt = rowClock;
              doorOpenPending = currAdcValue == '1';
              doorOpenDelayLogged = false;
            }
          }
        }

        final prevAdbsValue = stableAdbsValue;
        final rawCurrAdbsValue = idxAdbs != -1
            ? _getVal(curr, idxAdbs)
            : stableAdbsValue;
        if (rawCurrAdbsValue != stableAdbsValue &&
            _isSustainedValueAt(
              rows,
              i,
              idxAdbs,
              tIdx,
              rawCurrAdbsValue,
              kSignalSpikeIgnoreDuration,
            )) {
          stableAdbsValue = rawCurrAdbsValue;
        }
        if (prevAdbsValue == '0' && stableAdbsValue == '1') {
          final bypassTargetCars = _doorBypassTargetCars(prev, curr, doorIdx);
          rowHadDoorActivity = true;
          rowHadDoorCycle = true;
          tempLogs.add(
            LogEntry(
              time: time,
              message:
                  '${_doorBypassReportMessage(currAdcValue, bypassTargetCars)}${_currentStationSuffix(doorLocation)}',
              type: EntryType.bypass,
              isSummary: true,
            ),
          );
          lastDoorBypassLocation = doorLocation;
          lastDoorBypassCars = bypassTargetCars;
          lastDoorBypassAt = rowClock;
          doorBypassPending = currAdcValue == '0';
          doorBypassHoldLogged = false;
          if (doorClosePending && !doorCloseBypassLogged) {
            doorCloseBypassLogged = true;
            lastDoorCycleHadDelay = true;
          }
        }

        if (idxAdc != -1) {
          if (prevAdcValue != currAdcValue) {
            if (currAdcValue == '0') {
              if (!isStopped && speed > kStopSpeedThresholdKmh) {
                tempLogs.add(
                  LogEntry(
                    time: time,
                    message:
                        '이동 중 ${_adcReportValue('0')} 해제 정황이 확인됩니다. 주행 중 출입문 개방 관련 FSB 체결 가능성이 있습니다.',
                    type: EntryType.bypass,
                    isSummary: true,
                  ),
                );
              }
              if (doorOpenPending && lastDoorOpenAt != null) {
                final elapsed = rowClock - lastDoorOpenAt;
                tempLogs.add(
                  LogEntry(
                    time: time,
                    message:
                        '${lastDoorOpenLabel ?? '출입문 열림 버튼 취급'} 후 ${_adcReportValue('0')} 형성 (소요시간: ${_formatElapsed(elapsed)})${_doorModeSuffix(lastDoorOpenMode ?? currentDoorMode)}${_currentStationSuffix(lastDoorOpenLocation ?? doorLocation)}',
                    type: EntryType.restore,
                    isSummary: true,
                  ),
                );
                doorOpenPending = false;
              } else {
                tempLogs.add(
                  LogEntry(
                    time: time,
                    message: _adcReportValue('0'),
                    type: EntryType.bypass,
                    isSummary: true,
                  ),
                );
              }
            } else if (doorClosePending && lastDoorCloseAt != null) {
              final elapsed = rowClock - lastDoorCloseAt;
              tempLogs.add(
                LogEntry(
                  time: time,
                  message: doorCloseBypassLogged
                      ? '${_doorBypassReportMessage('1', lastDoorBypassCars.isEmpty ? lastDoorCloseBypassCars : lastDoorBypassCars)} (소요시간: ${_formatElapsed(elapsed)})${_currentStationSuffix(lastDoorCloseLocation)}'
                      : '${_adcRestoreReportMessage(lastDoorCloseMode ?? currentDoorMode)} (소요시간: ${_formatElapsed(elapsed)})${_currentStationSuffix(lastDoorCloseLocation)}',
                  type: EntryType.restore,
                  isSummary: true,
                ),
              );
              lastDoorCycleClosedAt = rowClock;
              doorClosePending = false;
              doorBypassPending = false;
            } else if (doorBypassPending && lastDoorBypassAt != null) {
              final elapsed = rowClock - lastDoorBypassAt;
              tempLogs.add(
                LogEntry(
                  time: time,
                  message:
                      '${_doorBypassReportMessage('1', lastDoorBypassCars)} (소요시간: ${_formatElapsed(elapsed)})${_currentStationSuffix(lastDoorBypassLocation)}',
                  type: EntryType.restore,
                  isSummary: true,
                ),
              );
              doorBypassPending = false;
            } else {
              tempLogs.add(
                LogEntry(
                  time: time,
                  message: _adcRestoreReportMessage(currentDoorMode),
                  type: EntryType.restore,
                  isSummary: true,
                ),
              );
            }
            lastDoorContextAt = rowClock;
          }
        }

        for (final entry in doorIdx.entries) {
          final idx = entry.value;
          if (idx == -1) continue;

          final prevValue = stableDoorValues[entry.key] ?? _getVal(prev, idx);
          final rawCurrValue = _getVal(curr, idx);
          if (rawCurrValue != prevValue &&
              _isSustainedValueAt(
                rows,
                i,
                idx,
                tIdx,
                rawCurrValue,
                kSignalSpikeIgnoreDuration,
              )) {
            stableDoorValues[entry.key] = rawCurrValue;
          }
          final currValue = stableDoorValues[entry.key] ?? prevValue;
          if (prevValue == currValue) continue;
          if (currValue == '0') {
            final bypassTargetCars = _doorBypassTargetCars(prev, curr, doorIdx);
            rowHadDoorActivity = true;
            lastDoorBypassLocation = doorLocation;
            lastDoorBypassCars = bypassTargetCars;
            lastDoorBypassAt = rowClock;
            doorBypassPending = currAdcValue == '0';
            doorBypassHoldLogged = false;
            if (doorClosePending && !doorCloseBypassLogged) {
              doorCloseBypassLogged = true;
              lastDoorCycleHadDelay = true;
            }
          } else if (_activeDoorBypassCarsFromValues(
            stableDoorValues,
          ).isEmpty) {
            doorBypassPending = false;
          }

          tempLogs.add(
            LogEntry(
              time: time,
              message: currValue == '0'
                  ? '${_formatDoorName(entry.key)}, ${_adcReportValue(currAdcValue)}${_currentStationSuffix(doorLocation)}'
                  : '${_formatDoorName(entry.key)} 해제${_currentStationSuffix(doorLocation)}',
              type: currValue == '0' ? EntryType.bypass : EntryType.restore,
              isSummary: true,
            ),
          );
          lastDoorContextAt = rowClock;
        }

        if (doorClosePending) {
          lastDoorCloseBypassCars = _activeDoorBypassCarsFromValues(
            stableDoorValues,
          );
          if (lastDoorCloseAt != null) {
            final elapsed = rowClock - lastDoorCloseAt;
            if (!doorCloseDelayLogged &&
                currAdcValue == '0' &&
                elapsed >= kDoorCloseDelayWarnDuration) {
              final carDetail = lastDoorCloseBypassCars.length == 1
                  ? ' (바이패스 신호 호차: ${lastDoorCloseBypassCars.first})'
                  : (lastDoorCloseBypassCars.isNotEmpty
                        ? ' 바이패스 신호 호차: ${lastDoorCloseBypassCars.join(', ')}.'
                        : '');
              tempLogs.add(
                LogEntry(
                  time: time,
                  message: doorCloseBypassLogged
                      ? '${_doorBypassReportMessage('0', lastDoorBypassCars.isEmpty ? lastDoorCloseBypassCars : lastDoorBypassCars)} 유지. 전체 출입문 닫힘 신호가 형성되지 않았습니다.$carDetail'
                      : '기관사 닫힘버튼 취급, ${_adcReportValue('0')}${_doorModeSuffix(lastDoorCloseMode ?? currentDoorMode)}${_currentStationSuffix(lastDoorCloseLocation)}$carDetail',
                  type: EntryType.bypass,
                  isSummary: true,
                ),
              );
              doorCloseDelayLogged = true;
              lastDoorCycleHadDelay = true;
            }
          }
        }

        if (doorOpenPending && lastDoorOpenAt != null) {
          final elapsed = rowClock - lastDoorOpenAt;
          if (!doorOpenDelayLogged &&
              currAdcValue == '1' &&
              elapsed >= kDoorOpenDelayWarnDuration) {
            tempLogs.add(
              LogEntry(
                time: time,
                message:
                    '${lastDoorOpenLabel ?? '출입문 열림 버튼 취급'} 후에도 ${_adcReportValue('1')} 유지 (출입문 열림 지연 정황)${_doorModeSuffix(lastDoorOpenMode ?? currentDoorMode)}${_currentStationSuffix(lastDoorOpenLocation)}',
                type: EntryType.bypass,
                isSummary: true,
              ),
            );
            doorOpenDelayLogged = true;
          }
        }

        if (doorBypassPending && lastDoorBypassAt != null) {
          final elapsed = rowClock - lastDoorBypassAt;
          if (!doorBypassHoldLogged &&
              currAdcValue == '0' &&
              elapsed >= kDoorBypassHoldWarnDuration) {
            tempLogs.add(
              LogEntry(
                time: time,
                message:
                    '${_doorBypassReportMessage('0', lastDoorBypassCars)} 유지. 전체 출입문 닫힘 신호가 형성되지 않았습니다.${_currentStationSuffix(lastDoorBypassLocation)}',
                type: EntryType.bypass,
                isSummary: true,
              ),
            );
            doorBypassHoldLogged = true;
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

        if (_isRisingEdge(prev, curr, idxFsb) &&
            _isSustainedValueAt(
              rows,
              i,
              idxFsb,
              tIdx,
              '1',
              kSignalSpikeIgnoreDuration,
            )) {
          final activeSpeedCode = _activeAtcSpeedCode(curr, speedCodeIndices);
          tempLogs.add(
            LogEntry(
              time: time,
              message: '${_signalName('FSBR')} 정황이 확인됩니다.',
              type: EntryType.info,
              isSummary: true,
            ),
          );

          if (activeSpeedCode != null &&
              speed > activeSpeedCode + kAtcOverspeedMarginKmh) {
            tempLogs.add(
              LogEntry(
                time: time,
                message:
                    '속도코드 ${activeSpeedCode}km/h 대비 열차속도 ${_formatSpeedKmh(speed)}km/h 초과 상태에서 전상용제동(FSB) 체결 흐름이 확인됩니다.',
                type: EntryType.info,
                isSummary: true,
              ),
            );
          } else if (nCodeDuration >= kFsbrNCodeContextDuration) {
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
            highSpeedDuration = Duration.zero;
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
            if (currentDist != null) {
              stop.endDist = currentDist;
              stop.minDist = stop.minDist == null
                  ? currentDist
                  : (currentDist < stop.minDist! ? currentDist : stop.minDist);
            }
            if (nextStaCode != stop.startNextStaCode) {
              stop.hadNextStaChange = true;
            }
          }

          if (isStopped) {
            if (speed > kStopSpeedThresholdKmh) {
              departureStartTime ??= time;
              movingDuration += sampleDelta;
            } else {
              movingDuration = Duration.zero;
              highSpeedDuration = Duration.zero;
              departureStartTime = null;
            }

            if (speed >= kDepartureSpeedThresholdKmh) {
              highSpeedDuration += sampleDelta;
            } else {
              highSpeedDuration = Duration.zero;
            }

            final stopForDeparture = activeStop;
            final hasDoorEvidenceForDeparture =
                stopForDeparture != null &&
                (stopForDeparture.hadDoorActivity ||
                    stopForDeparture.hadAdcRelease ||
                    stopForDeparture.hadDoorCycle);
            final hasDepartureDoorReady =
                !hasDoorEvidenceForDeparture ||
                idxAdc == -1 ||
                currAdcValue == '1';
            final hasMotionDeparture =
                movingDuration >= kDepartureMovementConfirmDuration ||
                highSpeedDuration >= kStopConfirmDuration;
            final hasConfirmedDeparture =
                hasMotionDeparture && hasDepartureDoorReady;

            if (hasConfirmedDeparture) {
              final departureLogOffset = tempLogs.length;
              final stop = activeStop;
              if (stop != null) {
                final stopStationName = _resolveStopStationName(
                  approachStationName: stop.approachStationName,
                  currentNextStationName: stop.startNextStationName,
                  lastBerthedStation: lastBerthedStation,
                  stopDist: stop.minDist ?? stop.startDist,
                );
                final hasDoorEvidence =
                    stop.hadDoorActivity ||
                    stop.hadAdcRelease ||
                    stop.hadDoorCycle;
                final blockType = _resolveStopBlockType(
                  stop: stop,
                  hasDoorEvidence: hasDoorEvidence,
                );
                final blockNote =
                    blockType == AnalysisBlockType.overrunSuspectedStop
                    ? _overrunBlockNote(stop)
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
              highSpeedDuration = Duration.zero;

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
                direction: travelDirection,
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
          stopDist: stop.minDist ?? stop.startDist,
        );
        final hasDoorEvidence =
            stop.hadDoorActivity || stop.hadAdcRelease || stop.hadDoorCycle;
        final blockType = _resolveStopBlockType(
          stop: stop,
          hasDoorEvidence: hasDoorEvidence,
        );
        final blockNote = blockType == AnalysisBlockType.overrunSuspectedStop
            ? _overrunBlockNote(stop)
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
        final mergeText = parsedFiles.length == 2 ? ' | 파일 2개 병합' : '';
        final scopeText = preparedRows.scopeLabel.isEmpty
            ? ''
            : ' | ${preparedRows.scopeLabel}';
        statusText = '분석 완료$mergeText$scopeText | ${rows.length - 1}개 행 처리';
        summaryText = _buildSummary(resolvedLogs);
        reportText = _buildReportDraft(resolvedLogs, resolvedBlocks);
        logs = resolvedLogs;
        blocks = resolvedBlocks;
        isLoading = false;
      });
    } catch (e) {
      setState(() {
        isLoading = false;
        statusText = '오류: $e';
      });
    }
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
          '광주교통공사 운행로그 분석기',
          style: TextStyle(
            fontFamily: 'NotoSansKR',
            fontWeight: FontWeight.bold,
            fontFamilyFallback: kKoreanFontFallback,
          ),
        ),
        backgroundColor: const Color(0xFF2C3E50),
        foregroundColor: Colors.white,
      ),
      body: DefaultTextStyle.merge(
        style: const TextStyle(
          fontFamily: 'NotoSansKR',
          fontFamilyFallback: kKoreanFontFallback,
        ),
        child: Stack(
          children: [
            Column(
              children: [
                Container(
                  width: double.infinity,
                  padding: const EdgeInsets.all(12),
                  color: Colors.grey[200],
                  child: Text(
                    statusText,
                    style: const TextStyle(
                      fontSize: 12,
                      color: Colors.blueGrey,
                    ),
                  ),
                ),
                Padding(
                  padding: const EdgeInsets.all(16),
                  child: ElevatedButton.icon(
                    onPressed: isLoading ? null : _pickAndAnalyze,
                    icon: const Icon(Icons.analytics_outlined),
                    label: Text(
                      isLoading
                          ? '분석 중...'
                          : '운행기록 파일(.xls / .xlsx / .csv) 불러오기',
                    ),
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
                if (logs.isNotEmpty)
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
                          child: FilledButton.icon(
                            onPressed: isLoading
                                ? null
                                : _reAnalyzeSelectedTime,
                            icon: const Icon(Icons.manage_search, size: 18),
                            label: const Text('시간 적용'),
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
                          'Excel 97-2003(.xls) 운행기록도 직접 업로드할 수 있습니다.\n분석이 실패하거나 오래 걸리면 Excel 통합 문서(.xlsx)로 변환 후 업로드해 주세요.',
                          textAlign: TextAlign.center,
                          style: TextStyle(
                            fontSize: 12,
                            color: Colors.black.withValues(alpha: 0.82),
                            fontWeight: FontWeight.w600,
                            height: 1.45,
                          ),
                        ),
                        const SizedBox(height: 10),
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
                          child: Column(
                            crossAxisAlignment: CrossAxisAlignment.end,
                            children: [
                              Text(
                                '제작자: 강경현',
                                style: const TextStyle(
                                  fontSize: 13,
                                  color: Colors.black,
                                  fontWeight: FontWeight.w700,
                                ),
                              ),
                              Text(
                                '검토: 김정주',
                                style: const TextStyle(
                                  fontSize: 13,
                                  color: Colors.black,
                                  fontWeight: FontWeight.w700,
                                ),
                              ),
                            ],
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
                            viewMode == LogViewMode.report,
                          ],
                          onPressed: (index) {
                            setState(() {
                              viewMode = switch (index) {
                                0 => LogViewMode.summary,
                                1 => LogViewMode.detail,
                                _ => LogViewMode.report,
                              };
                            });
                          },
                          borderRadius: BorderRadius.circular(8),
                          constraints: const BoxConstraints(
                            minHeight: 36,
                            minWidth: 64,
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
                            Padding(
                              padding: EdgeInsets.symmetric(horizontal: 10),
                              child: Text('보고'),
                            ),
                          ],
                        ),
                      ],
                    ),
                  ),
                Expanded(
                  child: viewMode == LogViewMode.report
                      ? ListView(
                          padding: const EdgeInsets.fromLTRB(16, 0, 16, 16),
                          children: [
                            Align(
                              alignment: Alignment.centerRight,
                              child: OutlinedButton.icon(
                                onPressed: reportText.isEmpty
                                    ? null
                                    : () async {
                                        await Clipboard.setData(
                                          ClipboardData(text: reportText),
                                        );
                                        if (!context.mounted) return;
                                        ScaffoldMessenger.of(
                                          context,
                                        ).showSnackBar(
                                          const SnackBar(
                                            content: Text('보고서 초안이 복사되었습니다.'),
                                            duration: Duration(seconds: 2),
                                          ),
                                        );
                                      },
                                icon: const Icon(Icons.copy_outlined, size: 18),
                                label: const Text('보고서 초안 복사'),
                              ),
                            ),
                            const SizedBox(height: 8),
                            Card(
                              elevation: 0,
                              shape: RoundedRectangleBorder(
                                side: BorderSide(
                                  color: const Color(
                                    0xFF34495E,
                                  ).withValues(alpha: 0.25),
                                ),
                                borderRadius: BorderRadius.circular(10),
                              ),
                              color: const Color(
                                0xFF34495E,
                              ).withValues(alpha: 0.05),
                              child: Padding(
                                padding: const EdgeInsets.all(14),
                                child: SelectableText(
                                  reportText.isEmpty
                                      ? '보고서 초안이 없습니다.'
                                      : reportText,
                                  style: const TextStyle(
                                    fontSize: 13,
                                    height: 1.45,
                                    color: Colors.black87,
                                  ),
                                ),
                              ),
                            ),
                          ],
                        )
                      : viewMode == LogViewMode.summary
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
                                          child: SelectableText(
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
                                              color: color.withValues(
                                                alpha: 0.10,
                                              ),
                                              borderRadius:
                                                  BorderRadius.circular(12),
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
                                side: BorderSide(
                                  color: color.withValues(alpha: 0.3),
                                ),
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
            if (isLoading)
              Positioned.fill(
                child: AbsorbPointer(
                  child: Container(
                    color: Colors.black.withValues(alpha: 0.10),
                    child: Center(
                      child: Container(
                        padding: const EdgeInsets.symmetric(
                          horizontal: 24,
                          vertical: 20,
                        ),
                        decoration: BoxDecoration(
                          color: Colors.white,
                          borderRadius: BorderRadius.circular(14),
                          boxShadow: const [
                            BoxShadow(
                              color: Color(0x22000000),
                              blurRadius: 18,
                              offset: Offset(0, 6),
                            ),
                          ],
                        ),
                        child: const Column(
                          mainAxisSize: MainAxisSize.min,
                          children: [
                            SizedBox(
                              width: 28,
                              height: 28,
                              child: CircularProgressIndicator(strokeWidth: 3),
                            ),
                            SizedBox(height: 14),
                            Text(
                              'Loading...',
                              style: TextStyle(
                                fontSize: 15,
                                fontWeight: FontWeight.w700,
                              ),
                            ),
                            SizedBox(height: 6),
                            Text(
                              '파일 크기에 따라 분석에 시간이 걸릴 수 있습니다.',
                              textAlign: TextAlign.center,
                              style: TextStyle(
                                fontSize: 12,
                                color: Colors.blueGrey,
                              ),
                            ),
                          ],
                        ),
                      ),
                    ),
                  ),
                ),
              ),
          ],
        ),
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
