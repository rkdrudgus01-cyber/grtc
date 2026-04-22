import 'dart:async';
import 'dart:convert';
import 'dart:html' as html;
import 'dart:typed_data';

import 'package:excel/excel.dart' as ex;
import 'package:flutter/material.dart';

const bool kDeveloperBuild =
    bool.fromEnvironment('DEVELOPER_BUILD', defaultValue: false);
const double kStopSpeedThreshold = 0.5;
const double kMoveSpeedThreshold = 5.0;
const double kPbActiveThreshold = 0.1;
const int kDoorCloseDelaySec = 3;
const int kControlPendingWarnSec = 3;
const int kDepartureConfirmSec = 5;
const int kSignalDebounceSec = 2;
const int kFsbrSustainSec = 2;
const int kNoCodeFsbrMinSec = 2;
const int kNoCodeNoFsbrWarnSec = 4;
const int kNoCodeNoFsbrSummaryEverySec = 30;
const int kNoCodeCriticalSec = 120;
const int kInhibitWarnSec = 10;
const int kRecentModeChangeSec = 10;
const int kRecentFsbrExplainedSec = 8;
const int kD15CooldownSec = 15;
const int kModeUnsetWarnSec = 5;
const int kDepartureFailureWarnSec = 5;

void main() {
  runApp(const LogAnalyzerApp());
}

class LogAnalyzerApp extends StatelessWidget {
  const LogAnalyzerApp({super.key});

  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      title: '광주도시철도 전동차 운행로그 분석기',
      debugShowCheckedModeBanner: false,
      theme: ThemeData(
        useMaterial3: true,
        colorScheme: ColorScheme.fromSeed(seedColor: const Color(0xFF0B6E4F)),
      ),
      home: const AnalyzerHomePage(),
    );
  }
}

class AnalyzerHomePage extends StatefulWidget {
  const AnalyzerHomePage({super.key});

  @override
  State<AnalyzerHomePage> createState() => _AnalyzerHomePageState();
}

class _AnalyzerHomePageState extends State<AnalyzerHomePage> {
  final AnalyzerEngine _engine = AnalyzerEngine(
    signalDefinitions: kSignalDefinitions,
  );

  List<LogRecord> _records = const <LogRecord>[];
  List<DiagnosticFinding> _findings = const <DiagnosticFinding>[];
  String _status = '운행기록 파일(.xlsx/.csv)을 업로드해 주세요.';
  String _fileName = '';
  int _rawRows = 0;
  bool _isAnalyzing = false;
  EntryType? _filterType;

  List<DiagnosticFinding> get _visibleFindings {
    if (_filterType == null) {
      return _findings;
    }
    return _findings
        .where((f) => f.type == _filterType)
        .toList(growable: false);
  }

  Future<void> _pickAndAnalyzeFile() async {
    final input = html.FileUploadInputElement()..accept = '.xlsx,.csv,.xls';
    input.click();
    await input.onChange.first;
    if (input.files == null || input.files!.isEmpty) {
      return;
    }
    final file = input.files!.first;
    await _handleFile(file);
  }

  Future<void> _handleFile(html.File file) async {
    final ext = _fileExtension(file.name);
    if (ext == '.xls') {
      setState(() {
        _status = '.xls 형식은 지원하지 않습니다. .xlsx 또는 .csv를 사용해 주세요.';
        _fileName = file.name;
      });
      return;
    }

    if (ext != '.xlsx' && ext != '.csv') {
      setState(() {
        _status = '지원하지 않는 파일 형식입니다: ${file.name}';
        _fileName = file.name;
      });
      return;
    }

    setState(() {
      _isAnalyzing = true;
      _status = '파일을 읽고 분석을 준비하고 있습니다...';
      _fileName = file.name;
      _records = const <LogRecord>[];
      _findings = const <DiagnosticFinding>[];
      _rawRows = 0;
    });

    try {
      final bytes = await _readFileAsBytes(file);
      if (mounted) {
        setState(() {
          _status = '데이터 파싱 중...';
        });
      }
      await Future<void>.delayed(Duration.zero);

      final rawRows =
          ext == '.csv' ? parseCsvRows(bytes) : parseXlsxRows(bytes);

      if (mounted) {
        setState(() {
          _status = '레코드 변환 중...';
        });
      }
      await Future<void>.delayed(Duration.zero);
      final parsed = _engine.parseRowsToRecords(rawRows);

      if (mounted) {
        setState(() {
          _status = '신호 정규화 중...';
        });
      }
      await Future<void>.delayed(Duration.zero);
      final normalized = _engine.normalizeSignals(parsed);

      if (mounted) {
        setState(() {
          _status = '이벤트/진단 분석 중...';
        });
      }
      await Future<void>.delayed(Duration.zero);
      final findings = _engine.analyze(normalized);

      if (!mounted) {
        return;
      }
      setState(() {
        _records = normalized;
        _findings = findings;
        _rawRows = rawRows.length;
        _status =
            '분석 완료: 레코드 ${normalized.length}건, 진단 ${findings.length}건 (${file.name})';
      });
    } catch (e) {
      if (!mounted) {
        return;
      }
      setState(() {
        _status = '분석 실패: $e';
      });
    } finally {
      if (mounted) {
        setState(() {
          _isAnalyzing = false;
        });
      }
    }
  }

  void _clearResult() {
    setState(() {
      _records = const <LogRecord>[];
      _findings = const <DiagnosticFinding>[];
      _status = '운행기록 파일(.xlsx/.csv)을 업로드해 주세요.';
      _fileName = '';
      _rawRows = 0;
      _filterType = null;
    });
  }

  @override
  Widget build(BuildContext context) {
    final scheme = Theme.of(context).colorScheme;
    final counts = _countByType(_findings);

    return Scaffold(
      appBar: AppBar(
        elevation: 0,
        backgroundColor: const Color(0xFF0E4F46),
        foregroundColor: Colors.white,
        title: const Text('광주도시철도 전동차 운행로그 분석기 v5.0'),
        actions: const <Widget>[
          Padding(
            padding: EdgeInsets.only(right: 16),
            child: Center(
              child: Text(
                '제작자: 강경현',
                style: TextStyle(fontSize: 13, fontWeight: FontWeight.w600),
              ),
            ),
          ),
        ],
      ),
      body: Container(
        decoration: const BoxDecoration(
          gradient: LinearGradient(
            begin: Alignment.topCenter,
            end: Alignment.bottomCenter,
            colors: <Color>[Color(0xFFF4FAF7), Color(0xFFEAF3EF)],
          ),
        ),
        child: Padding(
          padding: const EdgeInsets.all(16),
          child: Column(
            crossAxisAlignment: CrossAxisAlignment.stretch,
            children: <Widget>[
              Card(
                elevation: 0,
                shape: RoundedRectangleBorder(
                  borderRadius: BorderRadius.circular(14),
                  side: BorderSide(color: scheme.primary.withOpacity(0.18)),
                ),
                child: Padding(
                  padding: const EdgeInsets.all(14),
                  child: Column(
                    crossAxisAlignment: CrossAxisAlignment.start,
                    children: <Widget>[
                      Text(
                        _status,
                        style: Theme.of(context).textTheme.titleMedium?.copyWith(
                              fontWeight: FontWeight.w700,
                            ),
                      ),
                      const SizedBox(height: 8),
                      Wrap(
                        spacing: 10,
                        runSpacing: 8,
                        children: <Widget>[
                          _infoChip(
                            icon: Icons.description_outlined,
                            text: '파일: ${_fileName.isEmpty ? "-" : _fileName}',
                          ),
                          _infoChip(
                            icon: Icons.table_rows_outlined,
                            text: '원시행: $_rawRows',
                          ),
                          _infoChip(
                            icon: Icons.dataset_outlined,
                            text: '레코드: ${_records.length}',
                          ),
                        ],
                      ),
                    ],
                  ),
                ),
              ),
              const SizedBox(height: 10),
              Wrap(
                spacing: 8,
                runSpacing: 8,
                children: <Widget>[
                  FilledButton.icon(
                    onPressed: _isAnalyzing ? null : _pickAndAnalyzeFile,
                    icon: const Icon(Icons.upload_file),
                    label: const Text('파일 업로드'),
                  ),
                  OutlinedButton.icon(
                    onPressed: _isAnalyzing ? null : _clearResult,
                    icon: const Icon(Icons.delete_outline),
                    label: const Text('초기화'),
                  ),
                  _buildTypeFilterDropdown(),
                ],
              ),
              const SizedBox(height: 10),
              _buildBadges(counts, scheme),
              const SizedBox(height: 10),
              Expanded(
                child: Card(
                  elevation: 0,
                  shape: RoundedRectangleBorder(
                    borderRadius: BorderRadius.circular(14),
                    side: BorderSide(color: Colors.black.withOpacity(0.08)),
                  ),
                  child: Padding(
                    padding: const EdgeInsets.all(10),
                    child: _isAnalyzing
                        ? const Center(child: CircularProgressIndicator())
                        : _visibleFindings.isEmpty
                            ? const Center(
                                child: Text(
                                  '표시할 진단 로그가 없습니다. 파일을 업로드해 주세요.',
                                  style: TextStyle(fontSize: 15),
                                ),
                              )
                            : ListView.builder(
                                itemCount: _visibleFindings.length,
                                itemBuilder: (BuildContext context, int index) {
                                  final finding = _visibleFindings[index];
                                  final color =
                                      entryTypeColor(finding.type, scheme);
                                  final bg = color.withOpacity(0.10);

                                  return Container(
                                    margin: const EdgeInsets.only(bottom: 8),
                                    decoration: BoxDecoration(
                                      color: bg,
                                      borderRadius: BorderRadius.circular(10),
                                      border: Border(
                                        left: BorderSide(color: color, width: 4),
                                      ),
                                    ),
                                    child: ExpansionTile(
                                      tilePadding: const EdgeInsets.symmetric(
                                          horizontal: 12, vertical: 2),
                                      childrenPadding:
                                          const EdgeInsets.fromLTRB(
                                              16, 4, 16, 12),
                                      title: Text(
                                        '[${finding.time}] ${findingEventLabel(finding.code)}',
                                        style: TextStyle(
                                          color: color,
                                          fontWeight: FontWeight.w700,
                                        ),
                                      ),
                                      subtitle: Padding(
                                        padding: const EdgeInsets.only(top: 4),
                                        child: Text(finding.message),
                                      ),
                                      trailing: _typePill(finding.type, scheme),
                                      children: <Widget>[
                                        if (kDeveloperBuild)
                                          _detailField(
                                              '상태', trainStateLabel(finding.state)),
                                        if (kDeveloperBuild)
                                          _detailField('진단 코드', finding.code),
                                        _detailField('근거 신호', finding.evidence),
                                        if ((finding.checkPoint ?? '').isNotEmpty)
                                          _detailField(
                                            kDeveloperBuild
                                                ? '확인 포인트'
                                                : '해석 참고',
                                            finding.checkPoint!,
                                          ),
                                      ],
                                    ),
                                  );
                                },
                              ),
                  ),
                ),
              ),
            ],
          ),
        ),
      ),
    );
  }

  Widget _infoChip({required IconData icon, required String text}) {
    return Container(
      padding: const EdgeInsets.symmetric(horizontal: 10, vertical: 6),
      decoration: BoxDecoration(
        color: const Color(0xFFEFF7F3),
        borderRadius: BorderRadius.circular(999),
        border: Border.all(color: const Color(0xFFB9D8CB)),
      ),
      child: Row(
        mainAxisSize: MainAxisSize.min,
        children: <Widget>[
          Icon(icon, size: 14, color: const Color(0xFF0E4F46)),
          const SizedBox(width: 6),
          Text(
            text,
            style: const TextStyle(
              color: Color(0xFF0E4F46),
              fontSize: 12,
              fontWeight: FontWeight.w600,
            ),
          ),
        ],
      ),
    );
  }

  Widget _buildTypeFilterDropdown() {
    return DecoratedBox(
      decoration: BoxDecoration(
        border: Border.all(color: Colors.black26),
        borderRadius: BorderRadius.circular(8),
      ),
      child: Padding(
        padding: const EdgeInsets.symmetric(horizontal: 8),
        child: DropdownButton<EntryType?>(
          value: _filterType,
          underline: const SizedBox.shrink(),
          hint: const Text('유형 필터'),
          items: <DropdownMenuItem<EntryType?>>[
            const DropdownMenuItem<EntryType?>(
              value: null,
              child: Text('전체'),
            ),
            ...EntryType.values.map(
              (type) => DropdownMenuItem<EntryType?>(
                value: type,
                child: Text(entryTypeLabel(type)),
              ),
            ),
          ],
          onChanged: _isAnalyzing
              ? null
              : (EntryType? value) {
                  setState(() {
                    _filterType = value;
                  });
                },
        ),
      ),
    );
  }

  Widget _buildBadges(Map<EntryType, int> counts, ColorScheme scheme) {
    final ordered = <EntryType>[
      EntryType.critical,
      EntryType.warning,
      EntryType.info,
      EntryType.door,
      EntryType.atc,
      EntryType.ato,
      EntryType.tcms,
      EntryType.brake,
      EntryType.mode,
    ];

    return Wrap(
      spacing: 8,
      runSpacing: 8,
      children: ordered
          .where((type) => (counts[type] ?? 0) > 0)
          .map((type) => _badge(type, counts[type] ?? 0, scheme))
          .toList(growable: false),
    );
  }

  Widget _badge(EntryType type, int count, ColorScheme scheme) {
    final color = entryTypeColor(type, scheme);
    return Container(
      padding: const EdgeInsets.symmetric(horizontal: 10, vertical: 6),
      decoration: BoxDecoration(
        color: color.withOpacity(0.14),
        borderRadius: BorderRadius.circular(999),
        border: Border.all(color: color.withOpacity(0.45)),
      ),
      child: Text(
        '${entryTypeLabel(type)} $count',
        style: TextStyle(
          color: color,
          fontWeight: FontWeight.w700,
        ),
      ),
    );
  }

  Widget _typePill(EntryType type, ColorScheme scheme) {
    final color = entryTypeColor(type, scheme);
    return Container(
      padding: const EdgeInsets.symmetric(horizontal: 8, vertical: 4),
      decoration: BoxDecoration(
        color: color.withOpacity(0.18),
        borderRadius: BorderRadius.circular(999),
      ),
      child: Text(
        entryTypeLabel(type),
        style: TextStyle(
          color: color,
          fontSize: 12,
          fontWeight: FontWeight.w700,
        ),
      ),
    );
  }

  Widget _detailField(String label, String value) {
    return Padding(
      padding: const EdgeInsets.only(top: 4),
      child: RichText(
        text: TextSpan(
          style: DefaultTextStyle.of(context).style,
          children: <TextSpan>[
            TextSpan(
              text: '$label: ',
              style: const TextStyle(fontWeight: FontWeight.w700),
            ),
            TextSpan(text: value),
          ],
        ),
      ),
    );
  }

  Map<EntryType, int> _countByType(List<DiagnosticFinding> findings) {
    final counts = <EntryType, int>{
      for (final type in EntryType.values) type: 0,
    };
    for (final finding in findings) {
      counts[finding.type] = (counts[finding.type] ?? 0) + 1;
    }
    return counts;
  }
}

enum TrainState {
  inactive,
  controlPending,
  berthed,
  doorOpen,
  doorClosing,
  readyToDepart,
  departureCommand,
  running,
  braking,
  forcedBraking,
  inhibited,
  emergency,
}

enum OperationMode {
  unknown,
  manual,
  auto,
  yard,
  emergency,
  emergencyRescue,
}

enum EntryType {
  info,
  door,
  atc,
  ato,
  tcms,
  brake,
  mode,
  warning,
  critical,
}

class SignalDefinition {
  final String key;
  final String label;
  final String description;
  final String type;
  final String category;
  final String confidence;

  const SignalDefinition({
    required this.key,
    required this.label,
    required this.description,
    required this.type,
    required this.category,
    this.confidence = '확정',
  });
}

class LogRecord {
  final int index;
  final String time;
  final Map<String, dynamic> values;

  const LogRecord({
    required this.index,
    required this.time,
    required this.values,
  });

  String stringValue(String key) {
    final value = values[key];
    if (value == null) {
      return '';
    }
    if (value is String) {
      return value.trim();
    }
    return value.toString().trim();
  }

  double numberValue(String key) {
    final value = values[key];
    if (value == null) {
      return 0;
    }
    if (value is num) {
      return value.toDouble();
    }
    final text = value.toString().trim();
    if (text.isEmpty) {
      return 0;
    }
    final normalized = text.replaceAll(',', '');
    return double.tryParse(normalized) ?? 0;
  }

  bool boolValue(String key) {
    final value = values[key];
    if (value == null) {
      return false;
    }
    if (value is bool) {
      return value;
    }
    if (value is num) {
      return value != 0;
    }
    final text = value.toString().trim().toLowerCase();
    if (text.isEmpty) {
      return false;
    }
    if (text == '1' ||
        text == 'true' ||
        text == 'y' ||
        text == 'yes' ||
        text == 'on') {
      return true;
    }
    if (text == '0' ||
        text == 'false' ||
        text == 'n' ||
        text == 'no' ||
        text == 'off') {
      return false;
    }
    final numeric = double.tryParse(text.replaceAll(',', ''));
    if (numeric != null) {
      return numeric != 0;
    }
    return false;
  }
}

class DiagnosticFinding {
  final String code;
  final String time;
  final TrainState state;
  final EntryType type;
  final String message;
  final List<String> relatedSignals;
  final String evidence;
  final String? checkPoint;
  final int recordIndex;

  const DiagnosticFinding({
    required this.code,
    required this.time,
    required this.state,
    required this.type,
    required this.message,
    required this.relatedSignals,
    required this.evidence,
    this.checkPoint,
    required this.recordIndex,
  });
}

typedef RuleCondition = bool Function(
  LogRecord curr,
  LogRecord? prev,
  TrainState state,
  AnalyzerContext context,
);

typedef RuleMessageBuilder = String Function(
  LogRecord curr,
  LogRecord? prev,
  TrainState state,
  AnalyzerContext context,
);

class DiagnosticRule {
  final String code;
  final String title;
  final EntryType severity;
  final RuleCondition condition;
  final RuleMessageBuilder message;

  const DiagnosticRule({
    required this.code,
    required this.title,
    required this.severity,
    required this.condition,
    required this.message,
  });
}

class AnalyzerEngine {
  final List<SignalDefinition> signalDefinitions;

  AnalyzerEngine({required this.signalDefinitions});

  List<LogRecord> parseRowsToRecords(List<Map<String, dynamic>> rows) {
    final records = <LogRecord>[];

    for (var i = 0; i < rows.length; i++) {
      final row = rows[i];
      final normalizedMap = <String, dynamic>{};

      for (final entry in row.entries) {
        final key = normalizeHeader(entry.key);
        if (key.isEmpty) {
          continue;
        }
        normalizedMap[key] = entry.value;
      }

      if (_isFullyEmptyRow(normalizedMap)) {
        continue;
      }

      final rowIndex = _resolveRowIndex(normalizedMap, i);
      final time = _normalizeTimeValue(normalizedMap['TIME'], rowIndex);

      records.add(
        LogRecord(
          index: rowIndex,
          time: time,
          values: normalizedMap,
        ),
      );
    }
    return records;
  }

  List<LogRecord> normalizeSignals(List<LogRecord> records) {
    final typeByKey = <String, String>{
      for (final def in signalDefinitions) def.key: def.type,
    };

    return records.map((record) {
      final normalized = <String, dynamic>{};

      for (final key in kAllHeaders) {
        final raw = record.values[key];
        final expectedType = typeByKey[key] ?? 'bool';
        switch (expectedType) {
          case 'number':
          case 'percent':
          case 'code':
            normalized[key] = _parseNumber(raw);
            break;
          case 'bool':
          default:
            normalized[key] = _parseBool(raw) ? 1 : 0;
            break;
        }
      }

      final normalizedNextSta = normalizeStationCode(normalized['NEXTSTA']);
      normalized['TIME'] = record.time;
      normalized['NUM'] = record.index;
      normalized['NEXTSTA_NORM'] = normalizedNextSta ?? 0;

      return LogRecord(
        index: record.index,
        time: record.time,
        values: normalized,
      );
    }).toList(growable: false);
  }

  List<DiagnosticFinding> analyze(List<LogRecord> records) {
    final findings = <DiagnosticFinding>[];
    final context = AnalyzerContext();

    for (var i = 0; i < records.length; i++) {
      final curr = records[i];
      final prev = i > 0 ? records[i - 1] : null;
      final state = determineState(curr);

      findings.addAll(
        detectSignalTransitions(
          curr: curr,
          prev: prev,
          state: state,
          context: context,
        ),
      );

      findings.addAll(
        applyDiagnosticRules(
          curr: curr,
          prev: prev,
          state: state,
          context: context,
        ),
      );
    }

    findings.sort((a, b) => a.recordIndex.compareTo(b.recordIndex));
    return _compactFindings(findings);
  }

  TrainState determineState(LogRecord r) {
    final vel = r.numberValue('VEL').abs();

    if (r.boolValue('TCMS-EMTRIP') ||
        !r.boolValue('EB loop') ||
        !r.boolValue('EBR')) {
      return TrainState.emergency;
    }
    if (r.boolValue('ATC1/2FSBR') || r.boolValue('ATC EB')) {
      return TrainState.forcedBraking;
    }
    if (r.boolValue('TCMS-INBITD') || r.boolValue('ATO INBITD')) {
      return TrainState.inhibited;
    }
    if ((r.boolValue('FOR') || r.boolValue('REV')) &&
        (!r.boolValue('HCR') || !r.boolValue('TCMS-CAB ACT'))) {
      return TrainState.controlPending;
    }
    if (vel <= kStopSpeedThreshold &&
        (r.boolValue('CLOSE') || r.boolValue('S_CLOSE')) &&
        !r.boolValue('ADC')) {
      return TrainState.doorClosing;
    }
    if (vel <= kStopSpeedThreshold && !r.boolValue('ADC')) {
      return TrainState.doorOpen;
    }
    if (vel <= kStopSpeedThreshold &&
        r.boolValue('ADC') &&
        r.boolValue('ATC1/2DPT-PM') &&
        !r.boolValue('TCMS-INBITD') &&
        !r.boolValue('ATO INBITD')) {
      return TrainState.readyToDepart;
    }
    if ((r.boolValue('ATC1/2DptBP') ||
            r.numberValue('ATO P/B COM') > kPbActiveThreshold ||
            r.numberValue('P/B') > kPbActiveThreshold) &&
        vel <= kMoveSpeedThreshold &&
        r.boolValue('ADC')) {
      return TrainState.departureCommand;
    }
    if (vel > kStopSpeedThreshold &&
        r.numberValue('P/B') < -kPbActiveThreshold &&
        !r.boolValue('ATC1/2FSBR')) {
      return TrainState.braking;
    }
    if (vel > kMoveSpeedThreshold && !r.boolValue('ATC1/2FSBR')) {
      return TrainState.running;
    }
    if (vel <= kStopSpeedThreshold && r.boolValue('ADC')) {
      return TrainState.berthed;
    }
    return TrainState.inactive;
  }

  OperationMode determineOperationMode(LogRecord r) {
    if (r.boolValue('EMEGR')) {
      return OperationMode.emergencyRescue;
    }
    if (r.boolValue('EMERG')) {
      return OperationMode.emergency;
    }
    if (r.boolValue('YARD')) {
      return OperationMode.yard;
    }
    if (r.boolValue('AUTO')) {
      return OperationMode.auto;
    }
    if (r.boolValue('MANUAL')) {
      return OperationMode.manual;
    }
    return OperationMode.unknown;
  }

  String modeTransitionMessage(OperationMode prevMode, OperationMode currMode) {
    if (prevMode == OperationMode.auto && currMode == OperationMode.manual) {
      return '자동운전이 해제되고 수동운전으로 전환되었습니다.';
    }
    if (prevMode == OperationMode.manual && currMode == OperationMode.auto) {
      return '수동운전에서 자동운전으로 전환되었습니다.';
    }
    if (currMode == OperationMode.emergency) {
      return '비상운전 모드로 전환되었습니다.';
    }
    if (currMode == OperationMode.emergencyRescue) {
      return '비상구원운전 모드가 활성화되었습니다.';
    }
    if (currMode == OperationMode.yard) {
      return '기지운전 모드로 전환되었습니다.';
    }
    return '운전모드가 ${operationModeLabel(prevMode)}에서 ${operationModeLabel(currMode)}로 전환되었습니다.';
  }

  List<DiagnosticFinding> detectSignalTransitions({
    required LogRecord curr,
    required LogRecord? prev,
    required TrainState state,
    required AnalyzerContext context,
  }) {
    final out = <DiagnosticFinding>[];
    final vel = curr.numberValue('VEL').abs();
    final adcHighStreak =
        context.bump('transition_adc_high_streak', curr.boolValue('ADC'));
    final adcLowStreak =
        context.bump('transition_adc_low_streak', !curr.boolValue('ADC'));
    final fsbrEventStreak = context.bump(
      'transition_fsbr_streak',
      curr.boolValue('ATC1/2FSBR'),
    );
    final emergModeStreak = context.bump(
      'transition_emerg_mode_streak',
      curr.boolValue('EMERG') || curr.boolValue('EMEGR'),
    );
    final adbsStreak =
        context.bump('transition_adbs_streak', curr.boolValue('ADBS'));
    _updateBerthStationMemory(curr, context);

    if (prev != null) {
      final prevMode = determineOperationMode(prev);
      final currMode = determineOperationMode(curr);
      if (prevMode != OperationMode.unknown &&
          currMode != OperationMode.unknown &&
          prevMode != currMode) {
        context.memory['LAST_MODE_CHANGE_INDEX'] = '${curr.index}';
        context.memory['LAST_MODE_CHANGE_FROM'] = operationModeLabel(prevMode);
        context.memory['LAST_MODE_CHANGE_TO'] = operationModeLabel(currMode);
        out.add(
          _finding(
            code: 'EVT-MODE-CHANGE',
            type: EntryType.mode,
            state: state,
            curr: curr,
            message: _withLocation(
              curr,
              prev,
              modeTransitionMessage(prevMode, currMode),
            ),
            signals: const ['AUTO', 'MANUAL', 'EMERG', 'EMEGR', 'YARD'],
          ),
        );
      }
      if (vel > kStopSpeedThreshold && emergModeStreak == kSignalDebounceSec) {
        out.add(
          _finding(
            code: 'EVT-EMERG-RUN',
            type: EntryType.critical,
            state: state,
            curr: curr,
            message: _withLocation(curr, prev, '주행 중 비상운전 계열 모드로 전환되었습니다.'),
            signals: const ['VEL', 'EMERG', 'EMEGR'],
          ),
        );
      }
      if (adcLowStreak == kSignalDebounceSec) {
        out.add(
          _finding(
            code: 'EVT-ADC-OPEN',
            type: EntryType.door,
            state: state,
            curr: curr,
            message: 'ADC가 1→0으로 전환되어 출입문 열림 상태가 감지되었습니다.',
            signals: const ['ADC'],
          ),
        );
      }
      if (adcHighStreak == kSignalDebounceSec) {
        out.add(
          _finding(
            code: 'EVT-ADC-CLOSE',
            type: EntryType.door,
            state: state,
            curr: curr,
            message: 'ADC가 0→1로 전환되어 전차 출입문 닫힘(완폐) 상태가 형성되었습니다.',
            signals: const ['ADC'],
          ),
        );
      }
      if (wasRising(prev, curr, 'ATC1/2DPT-PM')) {
        final currentStation = context.memory['LAST_BERTHED_STATION_NAME'];
        final nextStation = stationDisplayName(normalizedNextSta(curr));
        final dptMessage = (currentStation != null && currentStation.isNotEmpty)
            ? '$currentStation 정차 중 출발 허가(DPT-PM)가 형성되었습니다.'
                '${(nextStation != null && nextStation != currentStation) ? ' 다음역은 $nextStation입니다.' : ''}'
            : _appendStation(
                curr,
                'ATC 출발 허가(DPT-PM)가 형성되었습니다.',
                prev: prev,
              );
        out.add(
          _finding(
            code: 'EVT-DPT',
            type: EntryType.atc,
            state: state,
            curr: curr,
            message: dptMessage,
            signals: const ['ATC1/2DPT-PM', 'ADC', 'VEL'],
          ),
        );
      }
      if (fsbrEventStreak == kFsbrSustainSec) {
        out.add(
          _finding(
            code: 'EVT-FSBR',
            type: EntryType.brake,
            state: state,
            curr: curr,
            message: '전상용제동(FSBR)이 체결되었습니다.',
            signals: const ['ATC1/2FSBR', 'VEL'],
          ),
        );
      }
      if (adbsStreak == kSignalDebounceSec) {
        out.add(
          _finding(
            code: 'EVT-ADBS',
            type: EntryType.warning,
            state: state,
            curr: curr,
            message: '출입문 바이패스(ADBS) 취급이 감지되었습니다. 안전 확인이 필요한 상태로 해석됩니다.',
            signals: const ['ADBS', 'ADC'],
          ),
        );
      }
      if (wasRising(prev, curr, 'FOR') && !curr.boolValue('HCR')) {
        out.add(
          _finding(
            code: 'EVT-CTRL',
            type: EntryType.warning,
            state: state,
            curr: curr,
            message: '전진 지령(FOR) 투입 후에도 HCR=0으로 제어권 형성이 되지 않았습니다.',
            signals: const ['FOR', 'HCR', 'TCMS-CAB ACT'],
          ),
        );
      }
      if (wasRising(prev, curr, 'DRIVL')) {
        out.add(
          _finding(
            code: 'EVT-DRIVL',
            type: EntryType.mode,
            state: state,
            curr: curr,
            message: _appendStation(curr, '무인운전 모드(DRIVL)가 활성화되었습니다.', prev: prev),
            signals: const ['DRIVL', 'AUTO', 'MANUAL'],
          ),
        );
      }
      if (wasRising(prev, curr, 'FOR WAR')) {
        out.add(
          _finding(
            code: 'EVT-FORWAR',
            type: EntryType.atc,
            state: state,
            curr: curr,
            message: _appendStation(curr, '전방 예고 신호(FOR WAR)가 감지되었습니다.', prev: prev),
            signals: const ['FOR WAR', 'VEL'],
          ),
        );
      }
      if (wasRising(prev, curr, 'PBR')) {
        out.add(
          _finding(
            code: 'EVT-PBR',
            type: EntryType.brake,
            state: state,
            curr: curr,
            message: _appendStation(curr, '추진/제동 요청(PBR) 신호가 감지되었습니다.', prev: prev),
            signals: const ['PBR', 'P/B', 'VEL'],
          ),
        );
      }
      if (wasRising(prev, curr, 'PAN UP')) {
        final type =
            curr.boolValue('EMEGR') ? EntryType.warning : EntryType.mode;
        final message = curr.boolValue('EMEGR')
            ? '비상구원운전 상태에서 판토 상승 버튼 취급이 감지되었습니다.'
            : '판토 상승 버튼 취급이 감지되었습니다.';
        out.add(
          _finding(
            code: 'EVT-PAN-UP',
            type: type,
            state: state,
            curr: curr,
            message: _appendStation(curr, message, prev: prev),
            signals: const ['PAN UP', 'EMEGR', 'VEL'],
          ),
        );
      }
      if (wasRising(prev, curr, 'PAN DN')) {
        final type = (curr.boolValue('EMEGR') || vel > kStopSpeedThreshold)
            ? EntryType.warning
            : EntryType.mode;
        final message = curr.boolValue('EMEGR')
            ? '비상구원운전 상태에서 판토 하강 버튼 취급이 감지되었습니다.'
            : '판토 하강 버튼 취급이 감지되었습니다.';
        out.add(
          _finding(
            code: 'EVT-PAN-DN',
            type: type,
            state: state,
            curr: curr,
            message: _appendStation(curr, message, prev: prev),
            signals: const ['PAN DN', 'EMEGR', 'VEL'],
          ),
        );
      }
      if (prev.numberValue('VEL').abs() > kStopSpeedThreshold &&
          vel <= kStopSpeedThreshold) {
        final currentStation = context.memory['LAST_BERTHED_STATION_NAME'];
        final nextStation = stationDisplayName(normalizedNextSta(curr));
        final stopMessage = (currentStation != null && currentStation.isNotEmpty)
            ? '$currentStation 정차 상태로 전환되었습니다.'
                '${(nextStation != null && nextStation != currentStation) ? ' 다음역은 $nextStation입니다.' : ''}'
            : _appendStation(curr, '열차가 정차 상태로 전환되었습니다.', prev: prev);
        out.add(
          _finding(
            code: 'EVT-STOP',
            type: EntryType.info,
            state: state,
            curr: curr,
            message: stopMessage,
            signals: const ['VEL', 'NEXTSTA'],
          ),
        );
      }
      if (prev.numberValue('VEL').abs() <= kStopSpeedThreshold &&
          vel > kStopSpeedThreshold) {
        final departedStation = context.memory['LAST_BERTHED_STATION_NAME'];
        if (departedStation != null && departedStation.isNotEmpty) {
          context.memory['LAST_DEPARTED_STATION_NAME'] = departedStation;
          context.memory['LAST_DEPARTED_STATION_CODE'] =
              context.memory['LAST_BERTHED_STATION_CODE'] ?? '';
        }
        final nextStation = stationDisplayName(normalizedNextSta(curr));
        final moveMessage = (departedStation != null && departedStation.isNotEmpty)
            ? '$departedStation 발차 후'
                '${(nextStation != null && nextStation != departedStation) ? ' $nextStation 방면으로' : ''}'
                ' 이동 상태로 전환되었습니다.'
            : _appendStation(
                curr,
                '열차가 정차 상태에서 이동 상태로 전환되었습니다.',
                prev: prev,
              );
        out.add(
          _finding(
            code: 'EVT-MOVE',
            type: EntryType.mode,
            state: state,
            curr: curr,
            message: moveMessage,
            signals: const ['VEL', 'NEXTSTA'],
          ),
        );
      }
      final currNextSta = normalizedNextSta(curr);
      final prevNextSta = normalizedNextSta(prev);
      if (currNextSta != null &&
          prevNextSta != null &&
          currNextSta != prevNextSta) {
        final prevStation = stationDisplayName(prevNextSta) ?? '미확인';
        final currStation = stationDisplayName(currNextSta) ?? '미확인';
        final dir = inferLineDirection(prev: prev, curr: curr);
        final dirText = dir == null ? '' : ' / $dir';
        final location = formatLocationText(curr, prev: prev);
        final isOperational = _isOperationalStationCode(currNextSta);
        out.add(
          _finding(
            code: 'EVT-NEXTSTA',
            type: EntryType.info,
            state: state,
            curr: curr,
            message: isOperational
                ? '$location에서 다음역 정보가 $prevStation에서 $currStation($currNextSta)으로 전환되었습니다.$dirText'
                : '$location에서 다음역 정보가 $prevStation에서 $currStation으로 갱신되었습니다.$dirText',
            signals: const ['NEXTSTA', 'DIST'],
          ),
        );
      }
    }

    for (final button in const ['OPEN-L', 'OPEN-R', 'CLOSE', 'REOPEN']) {
      if (prev != null && wasRising(prev, curr, button)) {
        out.add(
          _finding(
            code: 'EVT-$button',
            type: EntryType.door,
            state: state,
            curr: curr,
            message: '$button 버튼 입력이 감지되었습니다.',
            signals: [button, 'ADC'],
          ),
        );
      }
    }

    final stopCount =
        context.bump('transition_stop_streak', vel <= kStopSpeedThreshold);
    if (stopCount >= kDoorCloseDelaySec) {
      context.flags['departure_armed'] = true;
    }
    if (vel <= kStopSpeedThreshold) {
      context.bump('transition_speed5_streak', false);
    }

    if (context.flags['departure_armed'] == true) {
      final speed5 =
          context.bump('transition_speed5_streak', vel >= kMoveSpeedThreshold);
      if (speed5 == kDepartureConfirmSec) {
        out.add(
          _finding(
            code: 'EVT-DEPART',
            type: EntryType.mode,
            state: state,
            curr: curr,
            message: _appendStation(
              curr,
              '속도 ${kMoveSpeedThreshold.toStringAsFixed(0)}km/h 이상이 ${kDepartureConfirmSec}초 지속되어 발차가 확정되었습니다.',
              prev: prev,
            ),
            signals: const ['VEL', 'ATC1/2DptBP', 'ATO P/B COM', 'P/B'],
          ),
        );
        context.flags['departure_armed'] = false;
        context.bump('transition_speed5_streak', false);
      }
    }

    return out;
  }

  List<DiagnosticFinding> applyDiagnosticRules({
    required LogRecord curr,
    required LogRecord? prev,
    required TrainState state,
    required AnalyzerContext context,
  }) {
    final findings = <DiagnosticFinding>[];

    final vel = curr.numberValue('VEL').abs();
    final pb = curr.numberValue('P/B');
    final atoPb = curr.numberValue('ATO P/B COM');
    final adc = curr.boolValue('ADC');
    final noCode = isNoCode(curr);
    final fsbr = curr.boolValue('ATC1/2FSBR');
    final fsbrStreak = context.bump('diag_fsbr_streak', fsbr);
    final fsbrSustained = fsbrStreak >= kFsbrSustainSec;
    final fsbrSustainedRise = fsbrStreak == kFsbrSustainSec;
    final nCodeStreak = context.bump('NCodeStreak', noCode);
    final nCodeNoFsbrStreak =
        context.bump('NCodeNoFsbrStreak', noCode && !fsbrSustained);
    final activeSpeedCode = getActiveAtcSpeedCode(curr);
    final unclosedDoorCarsNow = _unclosedDoorCars(curr);
    final isDoorClosingWindow =
        vel <= kStopSpeedThreshold &&
        (curr.boolValue('CLOSE') || curr.boolValue('S_CLOSE')) &&
        !adc;

    // D1. 출발 불능
    final hasDepartureCommand =
        curr.boolValue('ATC1/2DptBP') ||
        atoPb > kPbActiveThreshold ||
        pb > kPbActiveThreshold;
    final d1Condition = (state == TrainState.readyToDepart ||
            state == TrainState.departureCommand ||
            state == TrainState.inhibited ||
            state == TrainState.controlPending) &&
        adc &&
        vel <= kStopSpeedThreshold &&
        hasDepartureCommand;
    final d1Count = context.bump('D1', d1Condition);
    if (d1Count == kDepartureFailureWarnSec) {
      String message;
      String checkPoint =
          'ATC1/2DPT-PM, TCMS-INBITD, HCR, TCMS-CAB ACT, ATO NMID';
      if (!curr.boolValue('ATC1/2DPT-PM')) {
        message = '발차 조작이 있었으나 출발 허가가 형성되지 않았습니다.';
        checkPoint = 'ATC1/2DPT-PM, ATC ON';
      } else if (curr.boolValue('TCMS-INBITD') ||
          curr.boolValue('ATO INBITD')) {
        message = '운행금지 신호로 인해 발차가 억제되고 있습니다.';
        checkPoint = 'TCMS-INBITD, ATO INBITD';
      } else if (!curr.boolValue('HCR') || !curr.boolValue('TCMS-CAB ACT')) {
        message = '발차 조건은 있으나 제어권이 형성되지 않았습니다.';
        checkPoint = 'HCR, TCMS-CAB ACT, ATC ON';
      } else if (curr.boolValue('ATO NMID')) {
        message = '동력 투입 후에도 열차 움직임이 검출되지 않았습니다.';
        checkPoint = 'ATO NMID, ATO P/B COM, VEL';
      } else {
        message = '발차 조작이 있었으나 열차가 움직이지 않았습니다.';
      }

      findings.add(
        _finding(
          code: 'D1',
          type: EntryType.warning,
          state: state,
          curr: curr,
          message: _appendStation(curr, message, prev: prev),
          signals: const [
            'ADC',
            'VEL',
            'P/B',
            'ATC1/2DptBP',
            'ATC1/2DPT-PM',
            'HCR',
            'TCMS-CAB ACT',
            'TCMS-INBITD',
            'ATO INBITD',
            'ATO NMID',
          ],
          checkPoint: checkPoint,
        ),
      );
    }

    // D2. 출입문 미완폐
    final d2Condition = isDoorClosingWindow;
    final d2Count = context.bump('D2', d2Condition);
    if (d2Count == kDoorCloseDelaySec) {
      final unclosedCars = unclosedDoorCarsNow;
      final carText =
          unclosedCars.isEmpty ? '호차 정보 미확인' : '${unclosedCars.join(", ")}호차';
      findings.add(
        _finding(
          code: 'D2',
          type: EntryType.warning,
          state: state,
          curr: curr,
          message: _withLocation(
            curr,
            prev,
            '$carText 출입문 닫힘 상태가 형성되지 않았습니다.',
          ),
          signals: const [
            'CLOSE',
            'S_CLOSE',
            'ADC',
            'DOOR0',
            'DOOR1',
            'DOOR2',
            'DOOR7'
          ],
          checkPoint: 'DOOR0, DOOR1, DOOR2, DOOR7, ADC',
        ),
      );
    }

    // D3. 제어권 형성 실패
    final d3Condition = state == TrainState.controlPending &&
        (curr.boolValue('FOR') || curr.boolValue('REV')) &&
        (!curr.boolValue('HCR') || !curr.boolValue('TCMS-CAB ACT'));
    final d3Count = context.bump('D3', d3Condition);
    if (d3Count == kControlPendingWarnSec) {
      findings.add(
        _finding(
          code: 'D3',
          type: EntryType.warning,
          state: state,
          curr: curr,
          message: _withLocation(
            curr,
            prev,
            '역전기 투입 후에도 활성 운전실이 형성되지 않았습니다.',
          ),
          signals: const ['FOR', 'REV', 'HCR', 'TCMS-CAB ACT', 'ATC ON'],
          checkPoint: 'HCR, TCMS-CAB ACT, ATC ON',
        ),
      );
    }

    // D4. 무코드 지속(2~3초 이상) + FSB
    final d4Condition =
        noCode && fsbrSustained && nCodeStreak >= kNoCodeFsbrMinSec;
    final d4Count = context.bump('D4', d4Condition);
    if (d4Count == 1) {
      findings.add(
        _finding(
          code: 'D4',
          type: EntryType.warning,
          state: state,
          curr: curr,
          message: _withLocation(
            curr,
            prev,
            '무코드가 약 ${nCodeStreak}초 지속된 뒤 전상용제동이 체결되었습니다.',
          ),
          signals: const ['ATC1/2NCode', 'ATC1/2FSBR', 'VEL'],
          checkPoint: 'ATC1/2NCode, ATC1/2FSBR',
        ),
      );
      context.memory['LAST_FSB_EXPLAINED_INDEX'] = '${curr.index}';
    }
    if (nCodeStreak == kNoCodeCriticalSec) {
      findings.add(
        _finding(
          code: 'D4',
          type: EntryType.critical,
          state: state,
          curr: curr,
          message: _withLocation(
            curr,
            prev,
            '무코드 상태가 ${kNoCodeCriticalSec}초 이상 지속되고 있습니다.',
          ),
          signals: const ['ATC1/2NCode', 'ATC1/2FSBR', 'VEL'],
          checkPoint: 'ATC1/2NCode, ATC1/2FSBR',
        ),
      );
    }

    // D5. 과속으로 인한 FSB
    final d5Condition = fsbrSustained &&
        !noCode &&
        activeSpeedCode != null &&
        vel > activeSpeedCode + 2;
    final d5Count = context.bump('D5', d5Condition);
    if (d5Count == 1) {
      findings.add(
        _finding(
          code: 'D5',
          type: EntryType.warning,
          state: state,
          curr: curr,
          message: _withLocation(
            curr,
            prev,
            '실제 속도(${vel.toStringAsFixed(1)}km/h)가 허용 속도($activeSpeedCode)를 초과하여 전상용제동이 체결되었습니다.',
          ),
          signals: const [
            'ATC1/2FSBR',
            'VEL',
            'ATC1/2NCode',
            'ATC1/2 0',
            'ATC1/2 15',
            'ATC1/2 25',
            'ATC1/2 30',
            'ATC1/2 35',
            'ATC1/2 40',
            'ATC1/2 45',
            'ATC1/2 50',
            'ATC1/2 55',
            'ATC1/2 60',
            'ATC1/2 65',
            'ATC1/2 70',
            'ATC1/2 75',
            'ATC1/2 80'
          ],
          checkPoint: 'VEL, ATC Speed Code',
        ),
      );
      context.memory['LAST_FSB_EXPLAINED_INDEX'] = '${curr.index}';
    }

    // D16. 무코드 지속 대비 FSBR 미체결(요약 알림)
    if (nCodeNoFsbrStreak == kNoCodeNoFsbrWarnSec ||
        (nCodeNoFsbrStreak >= kNoCodeNoFsbrSummaryEverySec &&
            nCodeNoFsbrStreak % kNoCodeNoFsbrSummaryEverySec == 0)) {
      final message = nCodeNoFsbrStreak == kNoCodeNoFsbrWarnSec
          ? '무코드가 ${kNoCodeNoFsbrWarnSec}초 이상 지속되었으나 전상용제동이 아직 체결되지 않았습니다.'
          : '무코드 대비 전상용제동 미체결 상태가 ${nCodeNoFsbrStreak}초 지속되고 있습니다.';
      findings.add(
        _finding(
          code: 'D16',
          type: EntryType.warning,
          state: state,
          curr: curr,
          message: _withLocation(curr, prev, message),
          signals: const ['ATC1/2NCode', 'ATC1/2FSBR', 'VEL'],
          checkPoint: 'ATC1/2NCode, ATC1/2FSBR',
        ),
      );
    }

    // D15. FSB 원인 보조 해석 (중립 기록 + 주변 신호 분기)
    final lastD15Index = int.tryParse(context.memory['LAST_D15_INDEX'] ?? '');
    final d15Cooldown = lastD15Index != null &&
        (curr.index - lastD15Index).abs() <= kD15CooldownSec;
    final lastFsbrExplainedIndex =
        int.tryParse(context.memory['LAST_FSB_EXPLAINED_INDEX'] ?? '');
    final fsbrExplainedRecently = lastFsbrExplainedIndex != null &&
        (curr.index - lastFsbrExplainedIndex).abs() <=
            kRecentFsbrExplainedSec;
    final lastModeChangeIndex =
        int.tryParse(context.memory['LAST_MODE_CHANGE_INDEX'] ?? '');
    final hasRecentModeChange = lastModeChangeIndex != null &&
        (curr.index - lastModeChangeIndex).abs() <= kRecentModeChangeSec;
    final doorInterlockRelated =
        vel > kStopSpeedThreshold &&
        !adc &&
        unclosedDoorCarsNow.isNotEmpty;
    final hasContext = hasRecentModeChange || doorInterlockRelated;

    if (fsbrSustainedRise &&
        !d4Condition &&
        !d5Condition &&
        !d15Cooldown &&
        !fsbrExplainedRecently &&
        hasContext) {
      if (doorInterlockRelated) {
        final message = _withLocation(
          curr,
          prev,
          '출입문 인터록 변화와 연계되어 전상용제동이 체결된 것으로 해석됩니다.',
        );
        const signals = <String>[
          'ATC1/2FSBR',
          'ADC',
          'DOOR0',
          'DOOR1',
          'DOOR2',
          'DOOR7',
        ];
        const checkPoint = 'ADC, 유효 호차 DOOR(0/1/2/7)';
        findings.add(
          _finding(
            code: 'D15',
            type: EntryType.warning,
            state: state,
            curr: curr,
            message: message,
            signals: signals,
            checkPoint: checkPoint,
          ),
        );
        context.memory['LAST_D15_INDEX'] = '${curr.index}';
      } else if (hasRecentModeChange) {
        final modeFrom = context.memory['LAST_MODE_CHANGE_FROM'] ?? '미확인';
        final modeTo = context.memory['LAST_MODE_CHANGE_TO'] ?? '미확인';
        final message = _withLocation(
          curr,
          prev,
          '주행 중 운전모드가 ${modeFrom}에서 ${modeTo}로 변경된 이후 전상용제동이 체결되었습니다.',
        );
        const signals = <String>[
          'ATC1/2FSBR',
          'AUTO',
          'MANUAL',
          'EMERG',
          'EMEGR'
        ];
        const checkPoint = '운전모드 전환 시점, FSBR 체결 시점';
        findings.add(
          _finding(
            code: 'D15',
            type: EntryType.warning,
            state: state,
            curr: curr,
            message: message,
            signals: signals,
            checkPoint: checkPoint,
          ),
        );
        context.memory['LAST_D15_INDEX'] = '${curr.index}';
      }
    }

    // D6. ATO 움직임 미검출
    final d6Condition = curr.boolValue('AUTO') &&
        atoPb > kPbActiveThreshold &&
        vel <= kStopSpeedThreshold &&
        curr.boolValue('ATO NMID');
    final d6Count = context.bump('D6', d6Condition);
    if (d6Count == kDepartureFailureWarnSec) {
      findings.add(
        _finding(
          code: 'D6',
          type: EntryType.warning,
          state: state,
          curr: curr,
          message: _appendStation(
            curr,
            '자동운전 명령 이후에도 열차 움직임이 검출되지 않았습니다.',
            prev: prev,
          ),
          signals: const [
            'AUTO',
            'ATO P/B COM',
            'ATO NMID',
            'VEL',
            'TCMS-INBITD',
            'ATO INBITD'
          ],
          checkPoint: 'ATO NMID, ATO P/B COM, VEL',
        ),
      );
    }

    // D7. 운행금지 활성
    final d7Condition =
        curr.boolValue('TCMS-INBITD') || curr.boolValue('ATO INBITD');
    final d7Count = context.bump('D7', d7Condition);
    if (d7Count == kSignalDebounceSec) {
      findings.add(
        _finding(
          code: 'D7',
          type: EntryType.warning,
          state: state,
          curr: curr,
          message:
              _withLocation(curr, prev, '운행금지 신호가 활성화되어 열차 출발이 억제되고 있습니다.'),
          signals: const ['TCMS-INBITD', 'ATO INBITD', 'ADC', 'ATC1/2DPT-PM'],
          checkPoint: 'TCMS-INBITD, ATO INBITD',
        ),
      );
    }
    if (d7Count == kInhibitWarnSec) {
      findings.add(
        _finding(
          code: 'D7',
          type: EntryType.warning,
          state: state,
          curr: curr,
          message: _withLocation(
            curr,
            prev,
            '운행금지 상태가 ${kInhibitWarnSec}초 이상 지속되고 있습니다.',
          ),
          signals: const ['TCMS-INBITD', 'ATO INBITD', 'ADC', 'ATC1/2DPT-PM'],
          checkPoint: 'TCMS-INBITD, ATO INBITD',
        ),
      );
    }

    // D8. 비상루프/EB 이상
    final d8Condition = !curr.boolValue('EB loop') || !curr.boolValue('EBR');
    final d8Count = context.bump('D8', d8Condition);
    if (d8Count == kSignalDebounceSec) {
      final hasCutout = curr.boolValue('EBCOS');
      final hasEmTrip = curr.boolValue('TCMS-EMTRIP');
      final extra = <String>[];
      if (hasCutout) {
        extra.add('EBCOS 차단 스위치 취급 상태');
      }
      if (hasEmTrip) {
        extra.add('TCMS 비상차단 연계');
      }
      final suffix = extra.isEmpty ? '' : ' (${extra.join(", ")})';

      findings.add(
        _finding(
          code: 'D8',
          type: EntryType.critical,
          state: state,
          curr: curr,
          message: _withLocation(
            curr,
            prev,
            '비상제동 안전루프 이상이 감지되었습니다.$suffix',
          ),
          signals: const ['EB loop', 'EBR', 'EBCOS', 'TCMS-EMTRIP'],
          checkPoint: 'EB loop, EBR, EBCOS',
        ),
      );
    }

    // D9. 바이패스 경고
    final d9Condition = curr.boolValue('ADBS') || curr.boolValue('EBCOS');
    final d9Count = context.bump('D9', d9Condition);
    if (d9Count == kSignalDebounceSec) {
      final reasons = <String>[];
      if (curr.boolValue('ADBS')) {
        reasons.add('출입문 바이패스');
      }
      if (curr.boolValue('EBCOS')) {
        reasons.add('비상제동 차단 스위치');
      }

      findings.add(
        _finding(
          code: 'D9',
          type: EntryType.warning,
          state: state,
          curr: curr,
          message: _withLocation(
            curr,
            prev,
            '안전회로 바이패스 취급이 감지되었습니다. (${reasons.join(" / ")})',
          ),
          signals: const ['ADBS', 'EBCOS', 'ADC', 'EB loop'],
          checkPoint: 'ADBS, EBCOS',
        ),
      );
    }

    // D10. 신호 불일치
    final mismatchMessages = <String>[];
    final mismatchSignals = <String>{};

    if (adc) {
      final unclosed = unclosedDoorCarsNow;
      if (unclosed.isNotEmpty) {
        mismatchMessages
            .add('ADC=1인데 유효 호차(${unclosed.join(", ")}) DOOR 비트가 0입니다');
        mismatchSignals.add('ADC');
        mismatchSignals.addAll(
          unclosed.map((c) => c == 0 ? 'DOOR0' : 'DOOR$c'),
        );
      }
    }

    if (prev != null &&
        wasRising(prev, curr, 'OPEN-L') &&
        !curr.boolValue('ATC1/2EDL')) {
      mismatchMessages.add('OPEN-L 입력이 있으나 EDL=0');
      mismatchSignals.addAll(const ['OPEN-L', 'ATC1/2EDL']);
    }

    if (prev != null &&
        wasRising(prev, curr, 'OPEN-R') &&
        !curr.boolValue('ATC1/2EDR')) {
      mismatchMessages.add('OPEN-R 입력이 있으나 EDR=0');
      mismatchSignals.addAll(const ['OPEN-R', 'ATC1/2EDR']);
    }

    final activeModes = <String>[
      if (curr.boolValue('EMEGR')) 'EMEGR',
      if (curr.boolValue('EMERG')) 'EMERG',
      if (curr.boolValue('YARD')) 'YARD',
      if (curr.boolValue('AUTO')) 'AUTO',
      if (curr.boolValue('MANUAL')) 'MANUAL',
    ];
    if (activeModes.length >= 2) {
      mismatchMessages.add('복수의 운전모드 신호가 동시에 감지되었습니다');
      mismatchSignals.addAll(activeModes);
    }

    if (prev != null) {
      final pbDiff = (curr.numberValue('P/B') - prev.numberValue('P/B')).abs();
      if (curr.boolValue('AUTO') && pbDiff >= 60) {
        mismatchMessages.add('AUTO=1 상태에서 P/B 급변($pbDiff)');
        mismatchSignals.addAll(const ['AUTO', 'P/B']);
      }
    }

    final signature = mismatchMessages.join('|');
    final prevSignature = context.memory['D10_SIGNATURE'] ?? '';
    if (signature.isNotEmpty && signature != prevSignature) {
      findings.add(
        _finding(
          code: 'D10',
          type: EntryType.warning,
          state: state,
          curr: curr,
          message: _withLocation(
            curr,
            prev,
            '제어 신호 불일치가 감지되었습니다. ${mismatchMessages.join(' / ')}',
          ),
          signals: mismatchSignals.toList(growable: false),
          checkPoint: '상황별 관련 신호 확인 필요',
        ),
      );
    }
    context.memory['D10_SIGNATURE'] = signature;

    // D11. 비상구원운전 모드 활성
    final d11Condition = curr.boolValue('EMEGR');
    final d11Count = context.bump('D11', d11Condition);
    if (d11Count == kSignalDebounceSec) {
      findings.add(
        _finding(
          code: 'D11',
          type: EntryType.critical,
          state: state,
          curr: curr,
          message: _withLocation(
            curr,
            prev,
            '비상구원운전 모드(EMEGR)가 활성화되었습니다.',
          ),
          signals: const ['EMEGR', 'AUTO', 'MANUAL', 'YARD'],
          checkPoint: 'EMEGR, 운전모드 전환 이력',
        ),
      );
    }

    // D13. 비상제동 버튼 취급
    final d13Condition = curr.boolValue('EMPB');
    final d13Count = context.bump('D13', d13Condition);
    if (d13Count == 1) {
      findings.add(
        _finding(
          code: 'D13',
          type: EntryType.warning,
          state: state,
          curr: curr,
          message: _withLocation(
            curr,
            prev,
            '운전실 비상제동 버튼(EMPB) 취급이 감지되었습니다.',
          ),
          signals: const ['EMPB', 'VEL', 'ATC1/2FSBR'],
          checkPoint: 'EMPB 취급 이력 및 제동 반응',
        ),
      );
    }

    // D14. 운전모드 미설정 지속
    final d14Condition = activeModes.isEmpty;
    final d14Count = context.bump('D14', d14Condition);
    if (d14Count == kModeUnsetWarnSec) {
      findings.add(
        _finding(
          code: 'D14',
          type: EntryType.warning,
          state: state,
          curr: curr,
          message: _withLocation(
            curr,
            prev,
            '유효한 운전모드가 설정되지 않은 상태가 지속되고 있습니다.',
          ),
          signals: const ['EMEGR', 'EMERG', 'YARD', 'AUTO', 'MANUAL'],
          checkPoint: '운전모드 비트 변화 이력',
        ),
      );
    }

    return findings;
  }

  int? getActiveAtcSpeedCode(LogRecord r) {
    if (isNoCode(r)) {
      return null;
    }
    for (final code in const <int>[
      0,
      15,
      25,
      30,
      35,
      40,
      45,
      50,
      55,
      60,
      65,
      70,
      75,
      80
    ]) {
      if (r.boolValue('ATC1/2 $code')) {
        return code;
      }
    }
    return null;
  }

  bool isNoCode(LogRecord r) {
    return r.boolValue('ATC1/2NCode');
  }

  bool _isPassengerStationCode(int code) {
    return (code >= 101 && code <= 119) || code == 140;
  }

  bool _isOperationalStationCode(int code) {
    return (code >= 141 && code <= 143) || code == 150 || code == 165 || code == 166;
  }

  String? stationDisplayName(int? nextSta) {
    final normalized = normalizeStationCode(nextSta);
    if (normalized == null) {
      return null;
    }
    final name = getStationName(normalized);
    if (name == null || name.isEmpty) {
      return null;
    }
    if (_isPassengerStationCode(normalized)) {
      return name.endsWith('역') ? name : '$name역';
    }
    if (normalized >= 141 && normalized <= 143) {
      return '시험선 구간';
    }
    if (normalized == 150) {
      return '기지입고 구간';
    }
    if (normalized == 165) {
      return '시운전 구간';
    }
    if (normalized == 166) {
      return '회송 구간';
    }
    return name;
  }

  void _updateBerthStationMemory(LogRecord curr, AnalyzerContext context) {
    final nextSta = normalizedNextSta(curr);
    if (nextSta == null) {
      return;
    }
    final vel = curr.numberValue('VEL').abs();
    final dist = curr.numberValue('DIST').round();
    final station = stationDisplayName(nextSta);
    if (station == null || station.isEmpty) {
      return;
    }
    if (vel <= kStopSpeedThreshold && dist <= 5) {
      context.memory['LAST_BERTHED_STATION_CODE'] = '$nextSta';
      context.memory['LAST_BERTHED_STATION_NAME'] = station;
    }
  }

  int? normalizeStationCode(dynamic rawValue) {
    final value = rawValue is num ? rawValue : _parseNumber(rawValue);
    final code = value.round();
    if (code <= 0) {
      return null;
    }
    if (code >= 100) {
      return code;
    }
    const specialCodes = <int>{40, 41, 42, 43, 50, 65, 66};
    if ((code >= 1 && code <= 19) || specialCodes.contains(code)) {
      return code + 100;
    }
    return code;
  }

  int? normalizedNextSta(LogRecord record) {
    final cached = record.values['NEXTSTA_NORM'];
    final cachedCode = normalizeStationCode(cached);
    if (cachedCode != null) {
      return cachedCode;
    }
    return normalizeStationCode(record.values['NEXTSTA']);
  }

  String? inferLineDirection(
      {required LogRecord prev, required LogRecord curr}) {
    final prevCode = normalizedNextSta(prev);
    final currCode = normalizedNextSta(curr);
    if (prevCode == null || currCode == null || prevCode == currCode) {
      return null;
    }
    if (!_isPassengerStationCode(prevCode) || !_isPassengerStationCode(currCode)) {
      return null;
    }
    if (currCode < prevCode) {
      return '상선';
    }
    return '하선';
  }

  String? getStationName(int? nextSta) {
    final normalized = normalizeStationCode(nextSta);
    if (normalized == null) {
      return null;
    }
    if (normalized >= 141 && normalized <= 143) {
      return '시험선';
    }
    return kStationNames[normalized];
  }

  String formatLocationText(LogRecord curr, {LogRecord? prev}) {
    final stationCode = normalizedNextSta(curr);
    final stationName = stationDisplayName(stationCode);
    final dist = curr.numberValue('DIST').round();

    final base = () {
      if (stationName == null) {
        return dist > 0 ? '위치 미확인 (${dist}m)' : '위치 미확인';
      }
      if (stationCode != null && _isOperationalStationCode(stationCode)) {
        return dist > 0 ? '$stationName (${dist}m)' : stationName;
      }
      if (dist <= 5) {
        return '$stationName 접근';
      }
      if (dist <= 80) {
        return '$stationName 직전 (${dist}m)';
      }
      if (dist <= 200) {
        return '$stationName 진입 중 (${dist}m)';
      }
      return '$stationName까지 ${dist}m';
    }();

    if (prev == null) {
      return base;
    }
    final direction = inferLineDirection(prev: prev, curr: curr);
    if (direction == null) {
      return base;
    }
    return '$base [$direction]';
  }

  bool wasRising(LogRecord prev, LogRecord curr, String key) {
    return !prev.boolValue(key) && curr.boolValue(key);
  }

  bool wasFalling(LogRecord prev, LogRecord curr, String key) {
    return prev.boolValue(key) && !curr.boolValue(key);
  }

  DiagnosticFinding _finding({
    required String code,
    required EntryType type,
    required TrainState state,
    required LogRecord curr,
    required String message,
    required List<String> signals,
    String? checkPoint,
  }) {
    return DiagnosticFinding(
      code: code,
      time: curr.time,
      state: state,
      type: type,
      message: message,
      relatedSignals: signals,
      evidence: _buildEvidence(curr, signals),
      checkPoint: checkPoint,
      recordIndex: curr.index,
    );
  }

  String _withLocation(LogRecord curr, LogRecord? prev, String baseMessage) {
    return '${formatLocationText(curr, prev: prev)}에서 $baseMessage';
  }

  String _appendStation(LogRecord r, String baseMessage, {LogRecord? prev}) {
    return _withLocation(r, prev, baseMessage);
  }

  String _buildEvidence(LogRecord curr, List<String> signals) {
    if (signals.isEmpty) {
      return '-';
    }
    final values = signals
        .map((key) => '$key=${_signalValueText(curr, key)}')
        .toList(growable: false);
    return values.join(', ');
  }

  String _signalValueText(LogRecord curr, String key) {
    if (key == 'NEXTSTA') {
      final code = normalizedNextSta(curr);
      if (code == null) {
        return '-';
      }
      final name = getStationName(code) ?? '미확인역';
      return '$code($name)';
    }
    final value = curr.values[key];
    if (value == null) {
      return '-';
    }
    if (value is num) {
      if (value == value.roundToDouble()) {
        return value.toInt().toString();
      }
      return value.toStringAsFixed(1);
    }
    return value.toString();
  }

  List<DiagnosticFinding> _compactFindings(List<DiagnosticFinding> findings) {
    if (findings.isEmpty) {
      return findings;
    }

    final filtered = _dedupeFsbrEventAtSameIndex(findings);
    if (filtered.isEmpty) {
      return filtered;
    }

    final compacted = <DiagnosticFinding>[];
    var start = filtered.first;
    var last = filtered.first;
    var duration = 1;

    bool canMerge(DiagnosticFinding a, DiagnosticFinding b) {
      return a.code == b.code &&
          a.type == b.type &&
          a.state == b.state &&
          a.message == b.message &&
          a.evidence == b.evidence &&
          a.checkPoint == b.checkPoint &&
          b.recordIndex == a.recordIndex + 1;
    }

    void flush() {
      var message = start.message;
      if (duration >= 3) {
        message = '$message (${duration}초 지속)';
      }
      compacted.add(
        DiagnosticFinding(
          code: start.code,
          time: start.time,
          state: start.state,
          type: start.type,
          message: message,
          relatedSignals: start.relatedSignals,
          evidence: start.evidence,
          checkPoint: start.checkPoint,
          recordIndex: start.recordIndex,
        ),
      );
    }

    for (var i = 1; i < filtered.length; i++) {
      final curr = filtered[i];
      if (canMerge(last, curr)) {
        duration += 1;
        last = curr;
        continue;
      }
      flush();
      start = curr;
      last = curr;
      duration = 1;
    }
    flush();
    return compacted;
  }

  List<DiagnosticFinding> _dedupeFsbrEventAtSameIndex(
    List<DiagnosticFinding> findings,
  ) {
    final explainedAt = <int>{};
    for (final finding in findings) {
      if (finding.code == 'D4' || finding.code == 'D5') {
        explainedAt.add(finding.recordIndex);
      }
    }
    if (explainedAt.isEmpty) {
      return findings;
    }
    return findings.where((finding) {
      return !(finding.code == 'EVT-FSBR' &&
          explainedAt.contains(finding.recordIndex));
    }).toList(growable: false);
  }

  List<int> _unclosedDoorCars(LogRecord r) {
    final cars = <int>[];
    for (final car in kEffectiveDoorCars) {
      final key = car == 0 ? 'DOOR0' : 'DOOR$car';
      if (!r.boolValue(key)) {
        cars.add(car);
      }
    }
    return cars;
  }
}

class AnalyzerContext {
  final Map<String, int> _streaks = <String, int>{};
  final Map<String, bool> flags = <String, bool>{};
  final Map<String, String> memory = <String, String>{};

  int bump(String key, bool active) {
    final next = active ? (_streaks[key] ?? 0) + 1 : 0;
    _streaks[key] = next;
    return next;
  }
}

Future<Uint8List> _readFileAsBytes(html.File file) {
  final completer = Completer<Uint8List>();
  final reader = html.FileReader();

  reader.onError.listen((_) {
    if (!completer.isCompleted) {
      completer.completeError(Exception('파일을 읽을 수 없습니다.'));
    }
  });

  reader.onLoadEnd.listen((_) {
    try {
      final result = reader.result;
      if (result is ByteBuffer) {
        completer.complete(Uint8List.view(result));
        return;
      }
      if (result is Uint8List) {
        completer.complete(Uint8List.fromList(result));
        return;
      }
      throw Exception('지원되지 않는 파일 데이터 형식입니다.');
    } catch (e) {
      if (!completer.isCompleted) {
        completer.completeError(e);
      }
    }
  });

  reader.readAsArrayBuffer(file);
  return completer.future;
}

List<Map<String, dynamic>> parseXlsxRows(Uint8List bytes) {
  final excel = ex.Excel.decodeBytes(bytes);
  if (excel.tables.isEmpty) {
    return const <Map<String, dynamic>>[];
  }

  final firstSheetName = excel.tables.keys.first;
  final sheet = excel.tables[firstSheetName];
  if (sheet == null || sheet.rows.isEmpty) {
    return const <Map<String, dynamic>>[];
  }

  final rows = sheet.rows;
  final headerRow = rows.first;
  final headers = <String>[];
  for (final cell in headerRow) {
    final key = normalizeHeader(_cellValueToString(cell?.value));
    headers.add(key);
  }

  final result = <Map<String, dynamic>>[];
  for (var rowIndex = 1; rowIndex < rows.length; rowIndex++) {
    final row = rows[rowIndex];
    final mapped = <String, dynamic>{};
    var hasData = false;

    for (var col = 0; col < headers.length; col++) {
      final header = headers[col];
      if (header.isEmpty) {
        continue;
      }
      final cell = col < row.length ? row[col] : null;
      final value = _cellValueToPrimitive(cell?.value);
      mapped[header] = value;
      if (!_isBlankValue(value)) {
        hasData = true;
      }
    }

    if (hasData) {
      result.add(mapped);
    }
  }

  return result;
}

List<Map<String, dynamic>> parseCsvRows(Uint8List bytes) {
  final text = _decodeText(bytes);
  final lines = const LineSplitter().convert(text);
  if (lines.isEmpty) {
    return const <Map<String, dynamic>>[];
  }

  final nonEmptyLines =
      lines.where((line) => line.trim().isNotEmpty).toList(growable: false);
  if (nonEmptyLines.isEmpty) {
    return const <Map<String, dynamic>>[];
  }

  final headerCells = _parseCsvLine(nonEmptyLines.first);
  final headers = headerCells.map(normalizeHeader).toList(growable: false);

  final result = <Map<String, dynamic>>[];
  for (var i = 1; i < nonEmptyLines.length; i++) {
    final cells = _parseCsvLine(nonEmptyLines[i]);
    final row = <String, dynamic>{};
    var hasData = false;

    for (var c = 0; c < headers.length; c++) {
      final header = headers[c];
      if (header.isEmpty) {
        continue;
      }
      final value = c < cells.length ? cells[c].trim() : '';
      row[header] = value;
      if (value.isNotEmpty) {
        hasData = true;
      }
    }

    if (hasData) {
      result.add(row);
    }
  }

  return result;
}

String _decodeText(Uint8List bytes) {
  try {
    return utf8.decode(bytes);
  } catch (_) {
    return latin1.decode(bytes);
  }
}

List<String> _parseCsvLine(String line) {
  final values = <String>[];
  var buffer = StringBuffer();
  var inQuotes = false;

  for (var i = 0; i < line.length; i++) {
    final ch = line[i];
    if (ch == '"') {
      final isEscapedQuote =
          inQuotes && i + 1 < line.length && line[i + 1] == '"';
      if (isEscapedQuote) {
        buffer.write('"');
        i++;
      } else {
        inQuotes = !inQuotes;
      }
      continue;
    }

    if (ch == ',' && !inQuotes) {
      values.add(buffer.toString());
      buffer = StringBuffer();
      continue;
    }
    buffer.write(ch);
  }
  values.add(buffer.toString());
  return values;
}

String _cellValueToString(ex.CellValue? value) {
  if (value == null) {
    return '';
  }
  if (value is ex.TextCellValue) {
    return value.value.toString();
  }
  if (value is ex.IntCellValue) {
    return value.value.toString();
  }
  if (value is ex.DoubleCellValue) {
    final numeric = value.value;
    if (numeric == numeric.roundToDouble()) {
      return numeric.toInt().toString();
    }
    return numeric.toString();
  }
  if (value is ex.BoolCellValue) {
    return value.value ? '1' : '0';
  }
  if (value is ex.TimeCellValue) {
    return value.toString();
  }
  if (value is ex.DateTimeCellValue) {
    return '${_twoDigits(value.hour)}:${_twoDigits(value.minute)}:${_twoDigits(value.second)}';
  }
  if (value is ex.DateCellValue) {
    return '${value.year}-${_twoDigits(value.month)}-${_twoDigits(value.day)}';
  }
  if (value is ex.FormulaCellValue) {
    return value.formula;
  }
  return value.toString();
}

dynamic _cellValueToPrimitive(ex.CellValue? value) {
  if (value == null) {
    return '';
  }
  if (value is ex.IntCellValue) {
    return value.value;
  }
  if (value is ex.DoubleCellValue) {
    return value.value;
  }
  if (value is ex.BoolCellValue) {
    return value.value ? 1 : 0;
  }
  if (value is ex.TimeCellValue) {
    return value.toString();
  }
  if (value is ex.DateTimeCellValue) {
    return '${_twoDigits(value.hour)}:${_twoDigits(value.minute)}:${_twoDigits(value.second)}';
  }
  if (value is ex.DateCellValue) {
    return '${value.year}-${_twoDigits(value.month)}-${_twoDigits(value.day)}';
  }
  if (value is ex.TextCellValue) {
    return value.value.toString().trim();
  }
  if (value is ex.FormulaCellValue) {
    return value.formula;
  }
  return value.toString();
}

bool _isBlankValue(dynamic value) {
  if (value == null) {
    return true;
  }
  if (value is String) {
    return value.trim().isEmpty;
  }
  if (value is num) {
    return false;
  }
  return value.toString().trim().isEmpty;
}

bool _isFullyEmptyRow(Map<String, dynamic> row) {
  for (final value in row.values) {
    if (!_isBlankValue(value)) {
      return false;
    }
  }
  return true;
}

int _resolveRowIndex(Map<String, dynamic> row, int fallbackIndex) {
  final raw = row['NUM'];
  if (raw == null) {
    return fallbackIndex + 1;
  }
  if (raw is num) {
    return raw.round();
  }
  final text = raw.toString().trim();
  final parsed = int.tryParse(text);
  return parsed ?? fallbackIndex + 1;
}

String _normalizeTimeValue(dynamic raw, int fallbackIndex) {
  if (raw == null) {
    return fallbackIndex.toString();
  }

  if (raw is ex.TimeCellValue) {
    return raw.toString();
  }

  if (raw is ex.DateTimeCellValue) {
    return '${_twoDigits(raw.hour)}:${_twoDigits(raw.minute)}:${_twoDigits(raw.second)}';
  }

  if (raw is num) {
    if (raw >= 0 && raw < 1) {
      final seconds = (raw * 24 * 60 * 60).round();
      final h = (seconds ~/ 3600) % 24;
      final m = (seconds % 3600) ~/ 60;
      final s = seconds % 60;
      return '${_twoDigits(h)}:${_twoDigits(m)}:${_twoDigits(s)}';
    }
    return raw.toString();
  }

  final text = raw.toString().trim();
  if (text.isEmpty) {
    return fallbackIndex.toString();
  }
  return text;
}

String normalizeHeader(String input) {
  return input
      .replaceAll('\uFEFF', '')
      .replaceAll('\u00A0', ' ')
      .trim()
      .replaceAll(RegExp(r'\s+'), ' ');
}

String _fileExtension(String name) {
  final dotIndex = name.lastIndexOf('.');
  if (dotIndex < 0) {
    return '';
  }
  return name.substring(dotIndex).toLowerCase();
}

double _parseNumber(dynamic value) {
  if (value == null) {
    return 0;
  }
  if (value is num) {
    return value.toDouble();
  }
  final text = value.toString().trim();
  if (text.isEmpty) {
    return 0;
  }
  return double.tryParse(text.replaceAll(',', '')) ?? 0;
}

bool _parseBool(dynamic value) {
  if (value == null) {
    return false;
  }
  if (value is bool) {
    return value;
  }
  if (value is num) {
    return value != 0;
  }
  final text = value.toString().trim().toLowerCase();
  if (text.isEmpty) {
    return false;
  }
  if (text == '1' ||
      text == 'true' ||
      text == 'y' ||
      text == 'yes' ||
      text == 'on') {
    return true;
  }
  if (text == '0' ||
      text == 'false' ||
      text == 'n' ||
      text == 'no' ||
      text == 'off') {
    return false;
  }
  final asNumber = double.tryParse(text.replaceAll(',', ''));
  return asNumber != null && asNumber != 0;
}

String _twoDigits(int value) => value.toString().padLeft(2, '0');

String trainStateLabel(TrainState state) {
  switch (state) {
    case TrainState.inactive:
      return '비활성';
    case TrainState.controlPending:
      return '제어권 대기';
    case TrainState.berthed:
      return '정차';
    case TrainState.doorOpen:
      return '문열림';
    case TrainState.doorClosing:
      return '문닫힘';
    case TrainState.readyToDepart:
      return '발차대기';
    case TrainState.departureCommand:
      return '발차명령';
    case TrainState.running:
      return '주행';
    case TrainState.braking:
      return '제동';
    case TrainState.forcedBraking:
      return '강제제동';
    case TrainState.inhibited:
      return '운행억제';
    case TrainState.emergency:
      return '비상';
  }
}

String operationModeLabel(OperationMode mode) {
  switch (mode) {
    case OperationMode.unknown:
      return '미확인';
    case OperationMode.manual:
      return '수동운전';
    case OperationMode.auto:
      return '자동운전';
    case OperationMode.yard:
      return '기지운전';
    case OperationMode.emergency:
      return '비상운전';
    case OperationMode.emergencyRescue:
      return '비상구원운전';
  }
}

String entryTypeLabel(EntryType type) {
  switch (type) {
    case EntryType.info:
      return '정보';
    case EntryType.door:
      return '도어';
    case EntryType.atc:
      return 'ATC';
    case EntryType.ato:
      return 'ATO';
    case EntryType.tcms:
      return 'TCMS';
    case EntryType.brake:
      return '제동';
    case EntryType.mode:
      return '모드';
    case EntryType.warning:
      return '경고';
    case EntryType.critical:
      return '비상';
  }
}

String findingEventLabel(String code) {
  if (code.startsWith('EVT-')) {
    switch (code) {
      case 'EVT-ADC-OPEN':
        return '출입문 열림 감지';
      case 'EVT-ADC-CLOSE':
        return '전차 완폐 확인';
      case 'EVT-DPT':
        return '출발 허가 형성';
      case 'EVT-FSBR':
        return '전상용제동 체결';
      case 'EVT-ADBS':
        return '바이패스 취급 감지';
      case 'EVT-CTRL':
        return '제어권 이상 감지';
      case 'EVT-MODE-CHANGE':
        return '운전모드 전환';
      case 'EVT-EMERG-RUN':
        return '주행 중 비상모드 전환';
      case 'EVT-DRIVL':
        return '무인운전 모드 활성';
      case 'EVT-FORWAR':
        return '전방 예고 감지';
      case 'EVT-PBR':
        return '추진/제동 요청 감지';
      case 'EVT-PAN-UP':
        return '판토 상승 버튼 취급';
      case 'EVT-PAN-DN':
        return '판토 하강 버튼 취급';
      case 'EVT-STOP':
        return '정차 전환';
      case 'EVT-MOVE':
        return '이동 전환';
      case 'EVT-NEXTSTA':
        return '다음역 갱신';
      case 'EVT-DEPART':
        return '발차 확정';
      default:
        return '이벤트';
    }
  }

  switch (code) {
    case 'D1':
      return '출발 불능';
    case 'D2':
      return '출입문 미완폐';
    case 'D3':
      return '제어권 형성 실패';
    case 'D4':
      return '무코드 + FSB';
    case 'D5':
      return '과속 + FSB';
    case 'D6':
      return 'ATO 미동작';
    case 'D7':
      return '운행금지 활성';
    case 'D8':
      return 'EB 루프 이상';
    case 'D9':
      return '바이패스 경고';
    case 'D10':
      return '신호 불일치';
    case 'D11':
      return '비상구원운전';
    case 'D12':
      return '판토 버튼 이벤트';
    case 'D13':
      return '비상제동 버튼 취급';
    case 'D14':
      return '운전모드 미설정';
    case 'D15':
      return 'FSB 보조 해석';
    case 'D16':
      return '무코드 FSBR 지연';
    default:
      return code;
  }
}

Color entryTypeColor(EntryType type, ColorScheme scheme) {
  switch (type) {
    case EntryType.info:
      return scheme.primary;
    case EntryType.door:
      return const Color(0xFF00695C);
    case EntryType.atc:
      return const Color(0xFF1E3A8A);
    case EntryType.ato:
      return const Color(0xFF1B5E20);
    case EntryType.tcms:
      return const Color(0xFF0F766E);
    case EntryType.brake:
      return const Color(0xFFB45309);
    case EntryType.mode:
      return const Color(0xFF374151);
    case EntryType.warning:
      return const Color(0xFFB45309);
    case EntryType.critical:
      return const Color(0xFFB91C1C);
  }
}

final List<SignalDefinition> kSignalDefinitions = _buildSignalDefinitions();

List<SignalDefinition> _buildSignalDefinitions() {
  final overrides = <String, SignalDefinition>{
    'NUM': const SignalDefinition(
      key: 'NUM',
      label: '순번',
      description: '로그 행 번호',
      type: 'number',
      category: 'basic',
    ),
    'TIME': const SignalDefinition(
      key: 'TIME',
      label: '시각',
      description: '1초 단위 기록 시각',
      type: 'code',
      category: 'basic',
    ),
    'NEXTSTA': const SignalDefinition(
      key: 'NEXTSTA',
      label: '다음역',
      description: '다음 역 코드',
      type: 'code',
      category: 'location',
    ),
    'DIST': const SignalDefinition(
      key: 'DIST',
      label: '잔여거리',
      description: '다음 정차점까지 거리',
      type: 'number',
      category: 'location',
    ),
    'VEL': const SignalDefinition(
      key: 'VEL',
      label: '실제속도',
      description: '열차 실제 속도(km/h)',
      type: 'number',
      category: 'basic',
    ),
    'P/B': const SignalDefinition(
      key: 'P/B',
      label: '추진/제동',
      description: '마스콘/ATO 추진·제동 명령값',
      type: 'percent',
      category: 'brake',
    ),
    'MC N': const SignalDefinition(
      key: 'MC N',
      label: 'MC N 보조신호(미확정)',
      description: '마스콘 중립 관련 보조 신호(극성 미확정, FOR/REV 직접 대응 아님)',
      type: 'bool',
      category: 'mode',
      confidence: '미확정',
    ),
    'EMERG': const SignalDefinition(
      key: 'EMERG',
      label: '비상운전 모드',
      description: 'Emergency driving mode',
      type: 'bool',
      category: 'mode',
      confidence: '확정',
    ),
    'EMEGR': const SignalDefinition(
      key: 'EMEGR',
      label: '비상구원운전 모드',
      description: 'Emergency rescue driving mode',
      type: 'bool',
      category: 'mode',
      confidence: '확정',
    ),
    'PAN UP': const SignalDefinition(
      key: 'PAN UP',
      label: '판토상승 버튼',
      description: '판토상승 버튼 입력 (0=미취급, 1=버튼 입력)',
      type: 'bool',
      category: 'mode',
      confidence: '확정',
    ),
    'PAN DN': const SignalDefinition(
      key: 'PAN DN',
      label: '판토하강 버튼',
      description: '판토하강 버튼 입력 (0=미취급, 1=버튼 입력)',
      type: 'bool',
      category: 'mode',
      confidence: '확정',
    ),
    'EMPB': const SignalDefinition(
      key: 'EMPB',
      label: '비상제동 버튼',
      description: 'Emergency brake push button',
      type: 'bool',
      category: 'brake',
      confidence: '확정',
    ),
    'FOR WAR': const SignalDefinition(
      key: 'FOR WAR',
      label: '전방 예고',
      description: 'Forward warning / pre-announcement signal',
      type: 'bool',
      category: 'atc',
      confidence: '확정',
    ),
    'DRIVL': const SignalDefinition(
      key: 'DRIVL',
      label: '무인운전 모드',
      description: 'Driverless mode status',
      type: 'bool',
      category: 'mode',
      confidence: '확정',
    ),
    'PBR': const SignalDefinition(
      key: 'PBR',
      label: '추진/제동 요청',
      description: 'Power/Brake request signal',
      type: 'bool',
      category: 'brake',
      confidence: '확정',
    ),
    'HCR': const SignalDefinition(
      key: 'HCR',
      label: '전두부 제어 계전기',
      description: '활성 운전실 형성 전달 신호',
      type: 'bool',
      category: 'mode',
    ),
    'ATC ON': const SignalDefinition(
      key: 'ATC ON',
      label: 'ATC 제어 승인',
      description: 'TCMS에서 ATC 제어권 승인',
      type: 'bool',
      category: 'atc',
    ),
    'ADC': const SignalDefinition(
      key: 'ADC',
      label: '전차 출입문 닫힘',
      description: 'All Door Closed (1=전차 완폐)',
      type: 'bool',
      category: 'door',
    ),
    'ATC1/2NCode': const SignalDefinition(
      key: 'ATC1/2NCode',
      label: '무코드',
      description: 'No Code (지상 신호 미수신)',
      type: 'bool',
      category: 'atc',
    ),
    'ATC1/2DPT-PM': const SignalDefinition(
      key: 'ATC1/2DPT-PM',
      label: '출발허가',
      description: 'Departure Permit',
      type: 'bool',
      category: 'atc',
    ),
    'ATC1/2FSBR': const SignalDefinition(
      key: 'ATC1/2FSBR',
      label: '전상용제동 요청',
      description: 'Full Service Brake Request',
      type: 'bool',
      category: 'brake',
    ),
    'DOOR0': const SignalDefinition(
      key: 'DOOR0',
      label: '0호차 도어 닫힘',
      description: '유효 호차 도어 상태 (1=닫힘)',
      type: 'bool',
      category: 'door',
    ),
    'DOOR1': const SignalDefinition(
      key: 'DOOR1',
      label: '1호차 도어 닫힘',
      description: '유효 호차 도어 상태 (1=닫힘)',
      type: 'bool',
      category: 'door',
    ),
    'DOOR2': const SignalDefinition(
      key: 'DOOR2',
      label: '2호차 도어 닫힘',
      description: '유효 호차 도어 상태 (1=닫힘)',
      type: 'bool',
      category: 'door',
    ),
    'DOOR7': const SignalDefinition(
      key: 'DOOR7',
      label: '7호차 도어 닫힘',
      description: '유효 호차 도어 상태 (1=닫힘)',
      type: 'bool',
      category: 'door',
    ),
    'DOOR3': const SignalDefinition(
      key: 'DOOR3',
      label: '3호차 도어(미사용)',
      description: '비실차/비사용 신호로 분석에서 제외',
      type: 'bool',
      category: 'ignored',
      confidence: '확정',
    ),
    'DOOR4': const SignalDefinition(
      key: 'DOOR4',
      label: '4호차 도어(미사용)',
      description: '비실차/비사용 신호로 분석에서 제외',
      type: 'bool',
      category: 'ignored',
      confidence: '확정',
    ),
    'DOOR5': const SignalDefinition(
      key: 'DOOR5',
      label: '5호차 도어(미사용)',
      description: '비실차/비사용 신호로 분석에서 제외',
      type: 'bool',
      category: 'ignored',
      confidence: '확정',
    ),
    'DOOR6': const SignalDefinition(
      key: 'DOOR6',
      label: '6호차 도어(미사용)',
      description: '비실차/비사용 신호로 분석에서 제외',
      type: 'bool',
      category: 'ignored',
      confidence: '확정',
    ),
    'ADBS': const SignalDefinition(
      key: 'ADBS',
      label: '출입문 바이패스',
      description: 'All Door Bypass',
      type: 'bool',
      category: 'door',
    ),
    'ATO NMID': const SignalDefinition(
      key: 'ATO NMID',
      label: '움직임 미검출',
      description: 'No Motion ID',
      type: 'bool',
      category: 'ato',
    ),
    'TCMS-INBITD': const SignalDefinition(
      key: 'TCMS-INBITD',
      label: 'TCMS 운행금지',
      description: 'TCMS Inhibit Driving',
      type: 'bool',
      category: 'tcms',
    ),
    'ATO INBITD': const SignalDefinition(
      key: 'ATO INBITD',
      label: 'ATO 운행금지',
      description: 'ATO Inhibit Driving',
      type: 'bool',
      category: 'ato',
    ),
    'TCMS-CAB ACT': const SignalDefinition(
      key: 'TCMS-CAB ACT',
      label: '활성 운전실',
      description: 'TCMS Cab Active',
      type: 'bool',
      category: 'tcms',
    ),
    'TCMS-ATCHBR': const SignalDefinition(
      key: 'TCMS-ATCHBR',
      label: '정차제동 유지',
      description: 'ATC Holding Brake Request',
      type: 'bool',
      category: 'brake',
    ),
    'EBCOS': const SignalDefinition(
      key: 'EBCOS',
      label: '비상제동 차단',
      description: 'Emergency Brake Cut-Out Switch',
      type: 'bool',
      category: 'safety',
    ),
    'EB loop': const SignalDefinition(
      key: 'EB loop',
      label: '비상루프',
      description: '비상제동 안전루프 도통',
      type: 'bool',
      category: 'safety',
    ),
  };

  return kAllHeaders.map((key) {
    final known = overrides[key];
    if (known != null) {
      return known;
    }
    return SignalDefinition(
      key: key,
      label: key,
      description: '$key 신호',
      type: _defaultTypeForKey(key),
      category: _defaultCategoryForKey(key),
      confidence: '입력',
    );
  }).toList(growable: false);
}

String _defaultTypeForKey(String key) {
  if (key == 'NUM' || key == 'DIST' || key == 'VEL') {
    return 'number';
  }
  if (key == 'P/B' || key == 'ATO P/B COM') {
    return 'percent';
  }
  if (key == 'NEXTSTA' || key == 'TIME') {
    return 'code';
  }
  return 'bool';
}

String _defaultCategoryForKey(String key) {
  if (const {'NUM', 'TIME', 'VEL'}.contains(key)) {
    return 'basic';
  }
  if (const {'NEXTSTA', 'DIST'}.contains(key)) {
    return 'location';
  }
  if (key.startsWith('ATC1/2') || key == 'ATC ON' || key == 'ATC EB') {
    return 'atc';
  }
  if (key.startsWith('ATO ')) {
    return 'ato';
  }
  if (key.startsWith('TCMS') || key.startsWith('TWC')) {
    return 'tcms';
  }
  if (key.startsWith('DOOR') ||
      key.contains('OPEN') ||
      key.contains('CLOSE') ||
      key.startsWith('AO/') ||
      key == 'MO/MC' ||
      key == 'ADBS' ||
      key == 'ADC') {
    return 'door';
  }
  if (const {
    'AUTO',
    'DRIVL',
    'MANUAL',
    'EMERG',
    'YARD',
    'FOR',
    'REV',
    'HCR',
    'MC N'
  }.contains(key)) {
    return 'mode';
  }
  if (key.contains('BR') || key.contains('PB') || key.contains('EB')) {
    return 'brake';
  }
  return 'safety';
}

const Map<int, String> kStationNames = <int, String>{
  101: '소태',
  102: '증심사입구',
  103: '남광주',
  104: '문화전당',
  105: '금남로4가',
  106: '금남로5가',
  107: '양동시장',
  108: '돌고개',
  109: '농성',
  110: '화정',
  111: '쌍촌',
  112: '운천',
  113: '상무',
  114: '김대중컨벤션센터',
  115: '공항',
  116: '송정공원',
  117: '광주송정',
  118: '도산',
  119: '평동',
  140: '녹동',
  150: '기지입고',
  165: '시운전',
  166: '회송',
};

const List<int> kEffectiveDoorCars = <int>[0, 1, 2, 7];

const List<String> kAllHeaders = <String>[
  'NUM',
  'TIME',
  'NEXTSTA',
  'DIST',
  'VEL',
  'FOR',
  'REV',
  'P/B',
  'ATO RDY',
  'ATC EB',
  'EB loop',
  'MC N',
  'ATC ON',
  'HCR',
  'EBR',
  'FOR WAR',
  'EMEGR',
  'PAN UP',
  'PAN DN',
  'EMPB',
  'EBCOS',
  'SBS',
  'PBR',
  'CRPB',
  'PBPS',
  'AUTO',
  'DRIVL',
  'MANUAL',
  'EMERG',
  'YARD',
  'ATC1/2NCode',
  'ATC1/2 0',
  'ATC1/2 15',
  'ATC1/2 25',
  'ATC1/2 30',
  'ATC1/2 35',
  'ATC1/2 40',
  'ATC1/2 45',
  'ATC1/2 50',
  'ATC1/2 55',
  'ATC1/2 60',
  'ATC1/2 65',
  'ATC1/2 70',
  'ATC1/2 75',
  'ATC1/2 80',
  'AO/AC',
  'AO/MC',
  'MO/MC',
  'ADBS',
  'DOOR1',
  'DOOR2',
  'DOOR3',
  'DOOR4',
  'DOOR5',
  'DOOR6',
  'DOOR7',
  'DOOR0',
  'OPEN-L',
  'OPEN-R',
  'CLOSE',
  'REOPEN',
  'S_OPEN-L',
  'S_OPEN-R',
  'S_CLOSE',
  'S_REOPEN',
  'ATC1/2DPT-PM',
  'ATC1/2OPEN-L',
  'ATC1/2OPEN-R',
  'ATC1/2S&P',
  'ATC1/2FSBR',
  'ATC1/2KEYUP',
  'ATC1/2KEYDN',
  'ATC1/2DptBP',
  'ATC1/2EDR',
  'ATC1/2EDL',
  'ATO P/B COM',
  'ATO RCVD',
  'ATO NMID',
  'ATO INBITD',
  'ATO STB',
  'ATO PSMID F1',
  'ATO PSMID F2',
  'ATO PSMID F3',
  'ATO PSMID F4',
  'ATO PSMID F5',
  'ATO PSMID F6',
  'ATO PSMID F7',
  'ATO PSMID F8',
  'TCMS-INBITD',
  'TCMS-CAB ACT',
  'TCMS-Act ATO',
  'TCMS-ATCHBR',
  'TCMS-EMTRIP',
  'TCMS D-ACKR',
  'TWC D-ACK',
  'ADC',
];

// 향후 확장 포인트
// 1) determineState() 상태 전이 기록 객체(StateMachineContext) 추가
// 2) applyDiagnosticRules()를 규칙 테이블 기반 DSL로 분리
// 3) 상/하선 라인 정보와 위치 정보(DIST/NEXTSTA) 기반 진단 강화

