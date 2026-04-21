import 'dart:html' as html;
import 'dart:typed_data';
import 'dart:convert';

import 'package:archive/archive.dart';
import 'package:flutter/material.dart';
import 'package:flutter/services.dart';
import 'package:xml/xml.dart';

void main() => runApp(const GRTCApp());

enum TrainStatus {
  maintenance,
  outOfService,
  research,
  longDia,
  mediumDia,
  shortDia,
  oneLoop,
  night7D,
  morning7D,
  afternoon7D,
  cleaning,
  okdongStay,
  okdongReserve,
  tomorrowReserve,
}

enum ViewMode {
  dial,
  doubleParking,
}

enum GenerationMode {
  quickManual,
  autoRestore,
}

enum TrainListFilter {
  all,
  needsReview,
  okdongOnly,
}

const Map<TrainStatus, String> kStatusLabel = {
  TrainStatus.maintenance: '중정비',
  TrainStatus.outOfService: '운휴',
  TrainStatus.research: '기술연구',
  TrainStatus.longDia: '장다이아 유도',
  TrainStatus.mediumDia: '중다이아 유도',
  TrainStatus.shortDia: '단다이아 유도',
  TrainStatus.oneLoop: '1회 운행',
  TrainStatus.night7D: '야간 7D',
  TrainStatus.morning7D: '오전 7D',
  TrainStatus.afternoon7D: '오후 7D',
  TrainStatus.cleaning: '대청소',
  TrainStatus.okdongStay: '금일 옥동 주박',
  TrainStatus.okdongReserve: '금일 옥동 예비',
  TrainStatus.tomorrowReserve: '명일 옥동 예비',
};

const double kChipFontSize = 17;

String statusLabel(TrainStatus status) => kStatusLabel[status] ?? '미지정';

T? firstOrNull<T>(Iterable<T> values) {
  final iterator = values.iterator;
  return iterator.moveNext() ? iterator.current : null;
}

class Train {
  final String id;
  final Set<TrainStatus> statuses;
  final double mileage;
  final int departureMins;

  const Train({
    required this.id,
    this.statuses = const {},
    this.mileage = 0,
    this.departureMins = TimeParser.unknown,
  });

  bool has(TrainStatus status) => statuses.contains(status);

  Train copyWith({
    Set<TrainStatus>? statuses,
    double? mileage,
    int? departureMins,
  }) {
    return Train(
      id: id,
      statuses: statuses ?? this.statuses,
      mileage: mileage ?? this.mileage,
      departureMins: departureMins ?? this.departureMins,
    );
  }
}

class ParsedArrivalRow {
  final String trainId;
  final int arrivalMins;
  final String lane;
  final String inspection;
  final String note;

  const ParsedArrivalRow({
    required this.trainId,
    required this.arrivalMins,
    required this.lane,
    required this.inspection,
    required this.note,
  });
}

class ParsedDepartureRow {
  final int slot;
  final String trainId;
  final int departureMins;
  final String lane;
  final String note;

  const ParsedDepartureRow({
    required this.slot,
    required this.trainId,
    required this.departureMins,
    required this.lane,
    required this.note,
  });
}

class ParsedCompanyDial {
  final Map<String, int> departures;
  final List<ParsedDepartureRow> departureRows;
  final List<ParsedArrivalRow> arrivals;
  final String? okdongReserveFromG35;

  const ParsedCompanyDial({
    required this.departures,
    required this.departureRows,
    required this.arrivals,
    required this.okdongReserveFromG35,
  });
}

class TomorrowAssignment {
  final int slot;
  final String trainId;
  final int departureMins;
  final String lane;
  final String note;
  final String reason;

  const TomorrowAssignment({
    required this.slot,
    required this.trainId,
    required this.departureMins,
    required this.lane,
    required this.note,
    this.reason = '',
  });
}

class TomorrowDialGenerationResult {
  final List<TomorrowAssignment> assignments;
  final List<String> logs;
  final List<int> unassignedSlots;

  const TomorrowDialGenerationResult({
    required this.assignments,
    required this.logs,
    required this.unassignedSlots,
  });

  int get length => assignments.length;
}

class GRTCApp extends StatelessWidget {
  const GRTCApp({super.key});

  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      title: 'GRTC',
      debugShowCheckedModeBanner: false,
      theme: ThemeData(
        colorScheme: ColorScheme.fromSeed(seedColor: Colors.indigo),
        useMaterial3: true,
        chipTheme: const ChipThemeData(
          labelStyle: TextStyle(fontSize: kChipFontSize, fontWeight: FontWeight.w700),
          padding: EdgeInsets.symmetric(horizontal: 10, vertical: 8),
        ),
      ),
      home: const MainScreen(),
    );
  }
}

class MainScreen extends StatefulWidget {
  const MainScreen({super.key});

  @override
  State<MainScreen> createState() => _MainScreenState();
}

class _MainScreenState extends State<MainScreen> {
  List<Train> trains = _initialTrains();
  Uint8List? uploadedTodayDialBytes;
  ParsedCompanyDial? parsedTodayDial;
  Uint8List? uploadedTomorrowDialBytes;
  ParsedCompanyDial? parsedTomorrowDial;
  List<TomorrowAssignment> tomorrowAssignments = const [];
  List<String> generationLogs = const [];
  List<int> unassignedTomorrowSlots = const [];
  Set<String> okdongForbidden = {};
  ViewMode viewMode = ViewMode.dial;
  GenerationMode generationMode = GenerationMode.autoRestore;
  TrainListFilter trainListFilter = TrainListFilter.all;
  String trainSearch = '';
  String message = '오늘 다이얼 미업로드';

  static List<Train> _initialTrains() {
    return List.generate(
      23,
      (i) => Train(id: '${101 + i}', mileage: 150000 + (i * 1000)),
    );
  }

  Uint8List _toBytes(Object? result) {
    if (result == null) {
      throw Exception('파일을 읽지 못했습니다.');
    }
    if (result is Uint8List) {
      return Uint8List.fromList(result);
    }
    if (result is ByteBuffer) {
      return Uint8List.view(result);
    }
    throw Exception('지원하지 않는 파일 결과 타입: ${result.runtimeType}');
  }

  void _uploadTodayDial() {
    final input = html.FileUploadInputElement()..accept = '.xlsx';

    input.onChange.listen((event) {
      if (input.files == null || input.files!.isEmpty) return;

      final file = input.files!.first;
      final reader = html.FileReader();
      reader.readAsArrayBuffer(file);

      reader.onLoadEnd.listen((event) {
        try {
          final bytes = _toBytes(reader.result);
          final parsed = CompanyDialParser.parse(bytes);

          setState(() {
            uploadedTodayDialBytes = bytes;
            parsedTodayDial = parsed;
            uploadedTomorrowDialBytes = null;
            parsedTomorrowDial = null;
            trains = TrainAutoMapper.applyTodayInfoBulk(trains, parsed);
            okdongForbidden = {
              for (final t in trains)
                if (t.has(TrainStatus.outOfService) ||
                    t.has(TrainStatus.maintenance) ||
                    t.has(TrainStatus.research))
                  t.id,
            };
            tomorrowAssignments = const [];
            generationLogs = const [];
            unassignedTomorrowSlots = const [];
            message =
                '금일 업로드 완료 / 출고 ${parsed.departures.length}건 / 입고 ${parsed.arrivals.length}건 / 자동복원 완료';
          });
        } catch (e) {
          setState(() {
            message = '금일 업로드 오류: $e';
          });
        }
      });
    });

    input.click();
  }

  void _uploadTomorrowDial() {
    final input = html.FileUploadInputElement()..accept = '.xlsx';

    input.onChange.listen((event) {
      if (input.files == null || input.files!.isEmpty) return;

      final file = input.files!.first;
      final reader = html.FileReader();
      reader.readAsArrayBuffer(file);

      reader.onLoadEnd.listen((event) {
        try {
          final bytes = _toBytes(reader.result);
          final parsed = CompanyDialParser.parse(bytes);
          setState(() {
            uploadedTomorrowDialBytes = bytes;
            parsedTomorrowDial = parsed;
            message = '명일 업로드 완료 / 출고 ${parsed.departures.length}건 / 입고 ${parsed.arrivals.length}건';
          });
        } catch (e) {
          setState(() {
            message = '명일 업로드 오류: $e';
          });
        }
      });
    });

    input.click();
  }

  void _toggleStatus(String id, TrainStatus status) {
    setState(() {
      trains = trains.map((t) {
        if (t.id != id) return t;

        final next = {...t.statuses};
        if (next.contains(status)) {
          next.remove(status);
        } else {
          next.add(status);
        }

        if (_blocksOkdong(status)) {
          if (next.contains(status)) {
            okdongForbidden.add(id);
          } else {
            final stillBlocked = next.any(_blocksOkdong);
            if (!stillBlocked) okdongForbidden.remove(id);
          }
        }

        return t.copyWith(statuses: next);
      }).toList();
      tomorrowAssignments = const [];
      generationLogs = const [];
      unassignedTomorrowSlots = const [];
    });

  }

  bool _blocksOkdong(TrainStatus status) {
    return status == TrainStatus.outOfService ||
        status == TrainStatus.maintenance ||
        status == TrainStatus.research;
  }

  void _toggleOkdongForbidden(String id) {
    setState(() {
      if (okdongForbidden.contains(id)) {
        okdongForbidden.remove(id);
      } else {
        okdongForbidden.add(id);
      }
      tomorrowAssignments = const [];
      generationLogs = const [];
      unassignedTomorrowSlots = const [];
    });
  }

  List<String> _blockingIssuesForGeneration() {
    final issues = <String>[];
    if (generationMode == GenerationMode.autoRestore && parsedTodayDial == null) {
      issues.add('자동 복원 생성은 오늘 다이얼 업로드가 필요합니다.');
    }

    final okdongSet = <String>{
      for (final t in trains)
        if (t.has(TrainStatus.okdongStay) || t.has(TrainStatus.okdongReserve)) t.id,
    };
    if (okdongSet.length != 5) {
      issues.add('옥동 주박군 5대 조건이 맞지 않습니다. 현재 ${okdongSet.length}대입니다.');
    }

    final reserveCount = trains.where((t) => t.has(TrainStatus.okdongReserve)).length;
    if (reserveCount != 1) {
      issues.add('금일 옥동 예비 차량은 1대여야 합니다. 현재 ${reserveCount}대입니다.');
    }

    return issues;
  }

  void _generateTomorrowDial() {
    final blocking = _blockingIssuesForGeneration();
    if (blocking.isNotEmpty) {
      setState(() => message = '생성 중단: ${blocking.join(' / ')}');
      return;
    }

    final effectiveDial = parsedTodayDial ?? _buildQuickModeTemplate();
    if (effectiveDial == null) {
      setState(() => message = '오늘 다이얼을 먼저 업로드하세요.');
      return;
    }

    final generated = TomorrowDialEngine.generateDetailed(
      trains: trains,
      todayDial: effectiveDial,
      okdongForbidden: okdongForbidden,
    );
    /*

    final modeText = parsedTodayDial == null ? '빠른 생성' : '자동 복원';
    final unassignedSuffix =
        generated.unassignedSlots.isEmpty ? '' : ' / 미배정 ${generated.unassignedSlots.join(', ')}';
        : ' / 미배정 ${generated.unassignedSlots.join(', ')}';
    setState(() {
      tomorrowAssignments = generated.assignments;
      generationLogs = generated.logs;
      unassignedTomorrowSlots = generated.unassignedSlots;
      message = '명일 다이얼 생성 완료 / ${generated.length}개 슬롯 배정$unassignedSuffix';
    });

    */
    final modeText = parsedTodayDial == null ? '빠른 생성' : '자동 복원 생성';
    final unassignedSuffix =
        generated.unassignedSlots.isEmpty ? '' : ' / 미배정 ${generated.unassignedSlots.join(', ')}';
    setState(() {
      tomorrowAssignments = generated.assignments;
      generationLogs = generated.logs;
      unassignedTomorrowSlots = generated.unassignedSlots;
      message = '$modeText 완료 / ${generated.length}개 슬롯 배정$unassignedSuffix';
    });

    for (final log in generated.logs) {
      debugPrint('[TomorrowDial] $log');
    }
  }

  ParsedCompanyDial? _buildQuickModeTemplate() {
    if (generationMode != GenerationMode.quickManual) return null;

    const okdongLanes = <int, String>{
      14: '옥동1',
      15: '옥동2',
      16: '옥동3',
      17: '옥동4',
      18: '옥동예비',
    };

    final rows = List<ParsedDepartureRow>.generate(18, (index) {
      final slot = index + 1;
      return ParsedDepartureRow(
        slot: slot,
        trainId: '',
        departureMins: TimeParser.unknown,
        lane: okdongLanes[slot] ?? '용산',
        note: '',
      );
    });

    return ParsedCompanyDial(
      departures: const {},
      departureRows: rows,
      arrivals: const [],
      okdongReserveFromG35: null,
    );
  }

  Future<void> _saveCompanyExcel() async {
    if (DateTime.now().millisecondsSinceEpoch < 0) {
      setState(() => message = '저장할 기준 엑셀이 없습니다. 오늘 다이얼을 먼저 업로드하세요.');
      return;
    }
    if (tomorrowAssignments.isEmpty) {
      _generateTomorrowDial();
    }
    if (tomorrowAssignments.isEmpty) {
      setState(() => message = '저장 중단: 먼저 명일 다이얼을 생성해 주세요.');
      return;
    }

    final templateBytes = await _resolveTemplateBytesForSave();
    if (templateBytes == null) {
      setState(() => message = '저장 실패: 템플릿을 찾지 못했습니다. 오늘 다이얼을 업로드해 주세요.');
      return;
    }

    try {
      final parkingResult = parsedTodayDial == null || parsedTomorrowDial == null
          ? null
          : DoubleParkingEngine.simulate(
              todayDial: parsedTodayDial!,
              tomorrowDial: parsedTomorrowDial!,
            );
      final bytes = CompanyExcelWriter.write(
        templateBytes: templateBytes,
        assignments: tomorrowAssignments,
        doubleParkingResult: parkingResult,
      );
      final blob = html.Blob([bytes], 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      final url = html.Url.createObjectUrlFromBlob(blob);
      html.AnchorElement(href: url)
        ..download = '명일입출고계획.xlsx'
        ..click();
      html.Url.revokeObjectUrl(url);

      setState(() => message = parkingResult == null ? '회사 엑셀 양식 저장 완료' : '회사 엑셀 양식 저장 완료 / 이중주차 결과 포함');
    } catch (e) {
      setState(() => message = '엑셀 저장 오류: $e');
    }
  }

  Future<Uint8List> _loadBundledTemplateBytes() async {
    final data = await rootBundle.load('assets/templates/company_blank_template.xlsx');
    return data.buffer.asUint8List(data.offsetInBytes, data.lengthInBytes);
  }

  Future<Uint8List?> _resolveTemplateBytesForSave() async {
    if (uploadedTodayDialBytes != null) {
      return uploadedTodayDialBytes!;
    }
    try {
      return await _loadBundledTemplateBytes();
    } catch (_) {
      return null;
    }
  }

  void _resetAll() {
    setState(() {
      trains = _initialTrains();
      uploadedTodayDialBytes = null;
      parsedTodayDial = null;
      uploadedTomorrowDialBytes = null;
      parsedTomorrowDial = null;
      tomorrowAssignments = const [];
      generationLogs = const [];
      unassignedTomorrowSlots = const [];
      okdongForbidden.clear();
      viewMode = ViewMode.dial;
      generationMode = GenerationMode.autoRestore;
      trainListFilter = TrainListFilter.all;
      trainSearch = '';
      message = '전체 초기화 완료';
    });
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(
        title: const Text('GRTC 명일 다이얼 자동 생성기'),
        actions: [
          FilledButton(
            onPressed: _uploadTodayDial,
            child: const Text('오늘 다이얼 업로드', style: TextStyle(fontSize: 16)),
          ),
          const SizedBox(width: 8),
          FilledButton(
            onPressed: _uploadTomorrowDial,
            child: const Text('명일 다이얼 업로드', style: TextStyle(fontSize: 16)),
          ),
          const SizedBox(width: 8),
          FilledButton.tonal(
            onPressed: _generateTomorrowDial,
            child: const Text('명일 다이얼 생성', style: TextStyle(fontSize: 16)),
          ),
          const SizedBox(width: 8),
          FilledButton.tonal(
            onPressed: _saveCompanyExcel,
            child: const Text('회사 엑셀 저장', style: TextStyle(fontSize: 16)),
          ),
          const SizedBox(width: 8),
          IconButton(
            onPressed: _resetAll,
            icon: const Icon(Icons.refresh),
            tooltip: '전체 초기화',
          ),
          const SizedBox(width: 8),
        ],
      ),
      body: Column(
        children: [
          Container(
            width: double.infinity,
            padding: const EdgeInsets.all(14),
            color: Colors.amber.shade50,
            child: Text(
              message,
              style: const TextStyle(fontSize: 19, fontWeight: FontWeight.bold),
            ),
          ),
          Container(
            width: double.infinity,
            padding: const EdgeInsets.symmetric(horizontal: 12, vertical: 8),
            child: SegmentedButton<ViewMode>(
              segments: const [
                ButtonSegment(value: ViewMode.dial, label: Text('다이얼 관리')),
                ButtonSegment(value: ViewMode.doubleParking, label: Text('이중주차 시뮬레이터')),
              ],
              selected: {viewMode},
              onSelectionChanged: (selected) {
                setState(() {
                  viewMode = selected.first;
                });
              },
            ),
          ),
          if (viewMode == ViewMode.dial)
            Container(
              width: double.infinity,
              padding: const EdgeInsets.fromLTRB(12, 0, 12, 8),
              child: SegmentedButton<GenerationMode>(
                segments: const [
                  ButtonSegment(
                    value: GenerationMode.quickManual,
                    label: Text('빠른 생성'),
                  ),
                  ButtonSegment(
                    value: GenerationMode.autoRestore,
                    label: Text('자동 복원 생성'),
                  ),
                ],
                selected: {generationMode},
                onSelectionChanged: (selected) {
                  setState(() {
                    generationMode = selected.first;
                  });
                },
              ),
            ),
          if (viewMode == ViewMode.dial) _buildWarningPanel(),
          if (viewMode == ViewMode.dial) _buildNeedsReviewPanel(),
          Expanded(
            child: viewMode == ViewMode.dial
                ? Row(
                    children: [
                      Expanded(flex: 3, child: _buildStatusPanel()),
                      Expanded(flex: 5, child: _buildTrainListV2()),
                      Expanded(flex: 4, child: _buildAssignmentPanel()),
                    ],
                  )
                : _buildDoubleParkingSimulator(),
          ),
        ],
      ),
    );
  }

  Widget _buildStatusPanel() {
    final groups = <String, List<TrainStatus>>{
      '운영 제외': [
        TrainStatus.outOfService,
        TrainStatus.maintenance,
        TrainStatus.research,
      ],
      '다이아 / 운행': [
        TrainStatus.longDia,
        TrainStatus.mediumDia,
        TrainStatus.shortDia,
        TrainStatus.oneLoop,
      ],
      '특수 작업': [
        TrainStatus.night7D,
        TrainStatus.morning7D,
        TrainStatus.afternoon7D,
        TrainStatus.cleaning,
      ],
      '옥동 상태': [
        TrainStatus.okdongStay,
        TrainStatus.okdongReserve,
        TrainStatus.tomorrowReserve,
      ],
    };

    return Container(
      color: Colors.indigo.shade50,
      child: ListView(
        padding: const EdgeInsets.all(12),
        children: groups.entries.map((entry) {
          return Card(
            child: Padding(
              padding: const EdgeInsets.all(12),
              child: Column(
                crossAxisAlignment: CrossAxisAlignment.start,
                children: [
                  Text(
                    entry.key,
                    style: const TextStyle(fontSize: 20, fontWeight: FontWeight.bold),
                  ),
                  const SizedBox(height: 10),
                  Wrap(
                    spacing: 8,
                    runSpacing: 8,
                    children: entry.value.map((s) {
                      return Chip(label: Text(statusLabel(s)));
                    }).toList(),
                  ),
                ],
              ),
            ),
          );
        }).toList(),
      ),
    );
  }

  // ignore: unused_element
  Widget _buildTrainList() {
    return ListView.builder(
      padding: const EdgeInsets.all(12),
      itemCount: trains.length,
      itemBuilder: (_, i) {
        final train = trains[i];
        final forbidden = okdongForbidden.contains(train.id);
        return Card(
          margin: const EdgeInsets.symmetric(vertical: 6),
          child: ExpansionTile(
            leading: CircleAvatar(
              radius: 24,
              backgroundColor: forbidden ? Colors.red : Colors.blue,
              child: Text(
                train.id,
                style: const TextStyle(color: Colors.white, fontSize: 16, fontWeight: FontWeight.bold),
              ),
            ),
            title: Text(
              '${train.id} / 출고 ${TimeParser.format(train.departureMins)}',
              style: const TextStyle(fontSize: 21, fontWeight: FontWeight.bold),
            ),
            subtitle: Padding(
              padding: const EdgeInsets.only(top: 8),
              child: Wrap(
                spacing: 8,
                runSpacing: 8,
                children: train.statuses.map((s) {
                  return InputChip(
                    label: Text(statusLabel(s)),
                    onDeleted: () => _toggleStatus(train.id, s),
                  );
                }).toList(),
              ),
            ),
            trailing: SizedBox(
              width: 128,
              child: forbidden
                  ? FilledButton(
                      onPressed: () => _toggleOkdongForbidden(train.id),
                      style: FilledButton.styleFrom(backgroundColor: Colors.red),
                      child: const Text('옥동금지', style: TextStyle(fontSize: 16, fontWeight: FontWeight.bold)),
                    )
                  : OutlinedButton(
                      onPressed: () => _toggleOkdongForbidden(train.id),
                      child: const Text('옥동허용', style: TextStyle(fontSize: 16, fontWeight: FontWeight.bold)),
                    ),
            ),
            children: [
              Padding(
                padding: const EdgeInsets.all(12),
                child: Wrap(
                  spacing: 8,
                  runSpacing: 8,
                  children: TrainStatus.values.map((s) {
                    final active = train.statuses.contains(s);
                    return FilterChip(
                      selected: active,
                      onSelected: (_) => _toggleStatus(train.id, s),
                      label: Text(statusLabel(s)),
                    );
                  }).toList(),
                ),
              ),
              if (tomorrowAssignments.any((a) => a.trainId == train.id))
                Padding(
                  padding: const EdgeInsets.fromLTRB(16, 0, 16, 14),
                  child: Align(
                    alignment: Alignment.centerLeft,
                    child: Text(
                      _tomorrowSummary(train.id),
                      style: const TextStyle(fontSize: 18, fontWeight: FontWeight.w700),
                    ),
                  ),
                ),
            ],
          ),
        );
      },
    );
  }

  String _tomorrowSummary(String trainId) {
    final assigned = tomorrowAssignments.where((a) => a.trainId == trainId).toList();
    return assigned
        .map((a) => '명일 ${a.slot}번 ${a.lane} ${TimeParser.format(a.departureMins)} ${a.note}'.trim())
        .join(' / ');
  }

  Widget _buildTrainListV2() {
    final visibleTrains = _visibleTrainsV2();
    return Column(
      children: [
        Padding(
          padding: const EdgeInsets.fromLTRB(12, 12, 12, 6),
          child: TextField(
            onChanged: (value) {
              setState(() {
                trainSearch = value.trim();
              });
            },
            decoration: const InputDecoration(
              isDense: true,
              border: OutlineInputBorder(),
              prefixIcon: Icon(Icons.search),
              hintText: '편성 검색',
            ),
          ),
        ),
        Padding(
          padding: const EdgeInsets.fromLTRB(12, 0, 12, 8),
          child: SegmentedButton<TrainListFilter>(
            segments: const [
              ButtonSegment(
                value: TrainListFilter.all,
                label: Text('전체'),
              ),
              ButtonSegment(
                value: TrainListFilter.needsReview,
                label: Text('확인 필요'),
              ),
              ButtonSegment(
                value: TrainListFilter.okdongOnly,
                label: Text('옥동'),
              ),
            ],
            selected: {trainListFilter},
            onSelectionChanged: (selected) {
              setState(() {
                trainListFilter = selected.first;
              });
            },
          ),
        ),
        Expanded(
          child: ListView.builder(
            padding: const EdgeInsets.all(12),
            itemCount: visibleTrains.length,
            itemBuilder: (_, i) {
              final train = visibleTrains[i];
              final forbidden = okdongForbidden.contains(train.id);
              final reviewReasons = _reviewReasons(train);

              return Card(
                margin: const EdgeInsets.symmetric(vertical: 6),
                child: ExpansionTile(
                  leading: CircleAvatar(
                    radius: 24,
                    backgroundColor: forbidden ? Colors.red : Colors.blue,
                    child: Text(
                      train.id,
                      style: const TextStyle(color: Colors.white, fontSize: 16, fontWeight: FontWeight.bold),
                    ),
                  ),
                  title: Text(
                    '${train.id} / 출고 ${TimeParser.format(train.departureMins)}',
                    style: const TextStyle(fontSize: 21, fontWeight: FontWeight.bold),
                  ),
                  subtitle: Padding(
                    padding: const EdgeInsets.only(top: 8),
                    child: Column(
                      crossAxisAlignment: CrossAxisAlignment.start,
                      children: [
                        Wrap(
                          spacing: 8,
                          runSpacing: 8,
                          children: train.statuses.map((s) {
                            return InputChip(
                              label: Text(statusLabel(s)),
                              onDeleted: () => _toggleStatus(train.id, s),
                            );
                          }).toList(),
                        ),
                        if (reviewReasons.isNotEmpty) const SizedBox(height: 8),
                        if (reviewReasons.isNotEmpty)
                          Text(
                            '확인 필요: ${reviewReasons.join(' / ')}',
                            style: const TextStyle(
                              fontSize: 15,
                              fontWeight: FontWeight.w700,
                              color: Colors.deepOrange,
                            ),
                          ),
                      ],
                    ),
                  ),
                  trailing: SizedBox(
                    width: 128,
                    child: forbidden
                        ? FilledButton(
                            onPressed: () => _toggleOkdongForbidden(train.id),
                            style: FilledButton.styleFrom(backgroundColor: Colors.red),
                            child: const Text('옥동금지', style: TextStyle(fontSize: 16, fontWeight: FontWeight.bold)),
                          )
                        : OutlinedButton(
                            onPressed: () => _toggleOkdongForbidden(train.id),
                            child: const Text('옥동허용', style: TextStyle(fontSize: 16, fontWeight: FontWeight.bold)),
                          ),
                  ),
                  children: [
                    Padding(
                      padding: const EdgeInsets.all(12),
                      child: Wrap(
                        spacing: 8,
                        runSpacing: 8,
                        children: TrainStatus.values.map((s) {
                          final active = train.statuses.contains(s);
                          return FilterChip(
                            selected: active,
                            onSelected: (_) => _toggleStatus(train.id, s),
                            label: Text(statusLabel(s)),
                          );
                        }).toList(),
                      ),
                    ),
                    if (tomorrowAssignments.any((a) => a.trainId == train.id))
                      Padding(
                        padding: const EdgeInsets.fromLTRB(16, 0, 16, 14),
                        child: Align(
                          alignment: Alignment.centerLeft,
                          child: Text(
                            _tomorrowSummaryWithReason(train.id),
                            style: const TextStyle(fontSize: 18, fontWeight: FontWeight.w700),
                          ),
                        ),
                      ),
                  ],
                ),
              );
            },
          ),
        ),
      ],
    );
  }

  List<Train> _visibleTrainsV2() {
    final query = trainSearch.trim();
    final filtered = trains.where((train) {
      if (query.isNotEmpty && !train.id.contains(query)) return false;
      switch (trainListFilter) {
        case TrainListFilter.all:
          return true;
        case TrainListFilter.needsReview:
          return _trainNeedsReview(train);
        case TrainListFilter.okdongOnly:
          return train.has(TrainStatus.okdongStay) || train.has(TrainStatus.okdongReserve);
      }
    }).toList();
    filtered.sort((a, b) => (int.tryParse(a.id) ?? 9999).compareTo(int.tryParse(b.id) ?? 9999));
    return filtered;
  }

  bool _trainNeedsReview(Train train) => _reviewReasons(train).isNotEmpty;

  List<String> _reviewReasons(Train train) {
    final reasons = <String>[];

    final hasExcluded = train.has(TrainStatus.outOfService) ||
        train.has(TrainStatus.maintenance) ||
        train.has(TrainStatus.research);
    final hasRunLike = train.has(TrainStatus.longDia) ||
        train.has(TrainStatus.mediumDia) ||
        train.has(TrainStatus.shortDia) ||
        train.has(TrainStatus.oneLoop) ||
        train.has(TrainStatus.night7D) ||
        train.has(TrainStatus.morning7D) ||
        train.has(TrainStatus.afternoon7D);
    if (hasExcluded && hasRunLike) {
      reasons.add('운영 제외와 운행 상태 동시 지정');
    }

    final sevenDCount = [
      TrainStatus.night7D,
      TrainStatus.morning7D,
      TrainStatus.afternoon7D,
    ].where((s) => train.has(s)).length;
    if (sevenDCount > 1) {
      reasons.add('7D 복수 지정');
    }

    final hasOkdongStatus =
        train.has(TrainStatus.okdongStay) || train.has(TrainStatus.okdongReserve) || train.has(TrainStatus.tomorrowReserve);
    if (okdongForbidden.contains(train.id) && hasOkdongStatus) {
      reasons.add('옥동금지와 옥동상태 충돌');
    }

    return reasons;
  }

  String _tomorrowSummaryWithReason(String trainId) {
    final assigned = tomorrowAssignments.where((a) => a.trainId == trainId).toList();
    return assigned.map((a) {
      final reason = a.reason.isEmpty ? '' : ' [${a.reason}]';
      return '명일 ${a.slot}번 ${a.lane} ${TimeParser.format(a.departureMins)} ${a.note}$reason'.trim();
    }).join(' / ');
  }

  Widget _buildNeedsReviewPanel() {
    final targets = trains.where(_trainNeedsReview).toList()
      ..sort((a, b) => (int.tryParse(a.id) ?? 9999).compareTo(int.tryParse(b.id) ?? 9999));
    if (targets.isEmpty) return const SizedBox.shrink();

    return Container(
      width: double.infinity,
      padding: const EdgeInsets.fromLTRB(12, 0, 12, 8),
      child: Card(
        color: Colors.orange.shade50,
        child: Padding(
          padding: const EdgeInsets.all(12),
          child: Wrap(
            spacing: 8,
            runSpacing: 8,
            children: targets.take(8).map((train) {
              final reason = _reviewReasons(train).join(', ');
              return Chip(
                label: Text('${train.id}: $reason'),
              );
            }).toList(),
          ),
        ),
      ),
    );
  }

  Widget _buildWarningPanel() {
    final warnings = _collectDialWarnings();
    if (warnings.isEmpty) return const SizedBox.shrink();

    return Container(
      width: double.infinity,
      padding: const EdgeInsets.fromLTRB(12, 0, 12, 8),
      child: Card(
        color: Colors.red.shade50,
        child: Padding(
          padding: const EdgeInsets.all(12),
          child: Text(
            warnings.map((w) => '경고: $w').join('\n'),
            style: const TextStyle(fontSize: 16, fontWeight: FontWeight.w700),
          ),
        ),
      ),
    );
  }

  List<String> _collectDialWarnings() {
    final warnings = <String>[];
    if (generationMode == GenerationMode.autoRestore && parsedTodayDial == null) {
      warnings.add('자동 복원 생성은 오늘 다이얼 업로드가 필요합니다.');
    }

    final okdongSet = <String>{
      for (final t in trains)
        if (t.has(TrainStatus.okdongStay) || t.has(TrainStatus.okdongReserve)) t.id,
    };
    final reserveCount = trains.where((t) => t.has(TrainStatus.okdongReserve)).length;
    if (okdongSet.length != 5) {
      warnings.add('옥동 주박군이 5대가 아닙니다. 현재 ${okdongSet.length}대입니다.');
    }
    if (reserveCount != 1) {
      warnings.add('금일 옥동 예비 차량 수가 1대가 아닙니다. 현재 ${reserveCount}대입니다.');
    }
    if (tomorrowAssignments.isNotEmpty && unassignedTomorrowSlots.isNotEmpty) {
      warnings.add('미배정 슬롯: ${unassignedTomorrowSlots.join(', ')}');
    }
    final reviewCount = trains.where(_trainNeedsReview).length;
    if (reviewCount > 0) {
      warnings.add('확인 필요 차량: ${reviewCount}대');
    }
    return warnings;
  }

  Widget _buildAssignmentPanel() {
    final sorted = [...tomorrowAssignments]..sort((a, b) => a.slot.compareTo(b.slot));
    return Container(
      color: Colors.grey.shade100,
      child: Column(
        children: [
          Padding(
            padding: const EdgeInsets.all(12),
            child: Row(
              children: [
                const Expanded(
                  child: Text(
                    '명일 배정 결과',
                    style: TextStyle(fontSize: 20, fontWeight: FontWeight.bold),
                  ),
                ),
                Text(
                  '${sorted.length}개',
                  style: const TextStyle(fontSize: 16, fontWeight: FontWeight.w700),
                ),
              ],
            ),
          ),
          Expanded(
            child: sorted.isEmpty
                ? const Center(
                    child: Text(
                      '아직 생성된 배정이 없습니다.',
                      style: TextStyle(fontSize: 16, fontWeight: FontWeight.w600),
                    ),
                  )
                : ListView.builder(
                    padding: const EdgeInsets.fromLTRB(12, 0, 12, 12),
                    itemCount: sorted.length,
                    itemBuilder: (_, i) {
                      final a = sorted[i];
                      return Card(
                        margin: const EdgeInsets.symmetric(vertical: 4),
                        child: ListTile(
                          title: Text(
                            '${a.slot}번 / ${a.trainId} / ${TimeParser.format(a.departureMins)}',
                            style: const TextStyle(fontSize: 17, fontWeight: FontWeight.w700),
                          ),
                          subtitle: Text(
                            '${a.lane} ${a.note}'.trim(),
                            style: const TextStyle(fontSize: 15),
                          ),
                          trailing: a.reason.isEmpty
                              ? null
                              : Chip(
                                  label: Text(
                                    a.reason,
                                    style: const TextStyle(fontSize: 13, fontWeight: FontWeight.w700),
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

  Widget _buildDoubleParkingSimulator() {
    if (parsedTodayDial == null) {
      return const Center(
        child: Text(
          '오늘 다이얼 업로드 후 시뮬레이션이 가능합니다.',
          style: TextStyle(fontSize: 20, fontWeight: FontWeight.w700),
        ),
      );
    }
    if (parsedTomorrowDial == null) {
      return const Center(
        child: Text(
          '명일 다이얼 업로드 후 시뮬레이션이 가능합니다.',
          style: TextStyle(fontSize: 20, fontWeight: FontWeight.w700),
        ),
      );
    }

    final result = DoubleParkingEngine.simulate(
      todayDial: parsedTodayDial!,
      tomorrowDial: parsedTomorrowDial!,
    );

    return ListView(
      padding: const EdgeInsets.all(12),
      children: [
        Card(
          color: Colors.indigo.shade50,
          child: Padding(
            padding: const EdgeInsets.all(12),
            child: Text(
              _okdongSummaryText(result.okdongSummary),
              style: const TextStyle(fontSize: 18, fontWeight: FontWeight.w700),
            ),
          ),
        ),
        const SizedBox(height: 6),
        if (result.warnings.isNotEmpty)
          Card(
            color: Colors.red.shade50,
            child: Padding(
              padding: const EdgeInsets.all(12),
              child: Text(
                result.warnings.map((w) => '경고: $w').join('\n'),
                style: const TextStyle(fontSize: 17, fontWeight: FontWeight.w700),
              ),
            ),
          ),
        if (result.warnings.isNotEmpty) const SizedBox(height: 6),
        Card(
          child: Padding(
            padding: const EdgeInsets.all(12),
            child: Text(
              '입고 순서 가이드: ${result.globalInboundGuide}',
              style: const TextStyle(fontSize: 18, fontWeight: FontWeight.w700),
            ),
          ),
        ),
        const SizedBox(height: 6),
        ...result.lines.map(_buildDoubleParkingLineCard),
      ],
    );
  }

  Widget _buildDoubleParkingLineCard(DoubleParkingLineResult line) {
    final outText = line.outbound.isEmpty
        ? '없음'
        : line.outbound
            .map((e) => '${e.trainId}(${TimeParser.format(e.timeMins)})')
            .join(', ');
    final inText = line.inbound.isEmpty
        ? '없음'
        : line.inbound
            .map((e) => '${e.trainId}(${TimeParser.format(e.timeMins)})')
            .join(', ');

    return Card(
      margin: const EdgeInsets.symmetric(vertical: 6),
      child: Padding(
        padding: const EdgeInsets.all(12),
        child: Column(
          crossAxisAlignment: CrossAxisAlignment.start,
          children: [
            Text(
              line.lineName,
              style: const TextStyle(fontSize: 22, fontWeight: FontWeight.bold),
            ),
            const SizedBox(height: 8),
            Text(
              '명일 OUT: $outText',
              style: const TextStyle(fontSize: 18, fontWeight: FontWeight.w700),
            ),
            const SizedBox(height: 6),
            Text(
              '금일 IN: $inText',
              style: const TextStyle(fontSize: 18, fontWeight: FontWeight.w700),
            ),
            const SizedBox(height: 6),
            Text(
              '입고 가이드: ${line.inboundGuide}',
              style: const TextStyle(fontSize: 17, fontWeight: FontWeight.w700),
            ),
            if (line.warnings.isNotEmpty) const SizedBox(height: 6),
            if (line.warnings.isNotEmpty)
              Text(
                line.warnings.map((w) => '경고: $w').join('\n'),
                style: const TextStyle(fontSize: 16, fontWeight: FontWeight.w700, color: Colors.red),
              ),
          ],
        ),
      ),
    );
  }

  String _okdongSummaryText(OkdongParkingSummary summary) {
    final normal = summary.normalIds.isEmpty ? '없음' : summary.normalIds.join(', ');
    final reserve = summary.reserveId ?? '미확인';
    final all = summary.allIds.isEmpty ? '없음' : summary.allIds.join(', ');
    return '옥동 주박(일반4): $normal\n옥동 예비(1): $reserve\n총 5대 기준 확인: $all';
  }
}

class TimeParser {
  static const int unknown = 9999;

  static int parse(dynamic value) {
    if (value == null) return unknown;

    try {
      if (value is Duration) return value.inMinutes;
      if (value is DateTime) return value.hour * 60 + value.minute;

      if (value is num) {
        return _parseExcelNumber(value.toDouble());
      }

      final s = value.toString().trim();
      if (s.isEmpty) return unknown;

      final match = RegExp(r'(\d{1,2}):(\d{2})(?::\d{2})?').firstMatch(s);
      if (match != null) {
        return int.parse(match.group(1)!) * 60 + int.parse(match.group(2)!);
      }

      final numeric = double.tryParse(s);
      if (numeric != null) return _parseExcelNumber(numeric);
    } catch (_) {}

    return unknown;
  }

  static int _parseExcelNumber(double value) {
    if (value < 0) return unknown;
    final fraction = value - value.floorToDouble();
    if (value < 2) {
      return (value * 24 * 60).round();
    }
    if (fraction > 0) {
      return (fraction * 24 * 60).round();
    }
    return unknown;
  }

  static double toExcelTime(int mins) {
    if (mins == unknown) return 0;
    return mins / (24 * 60);
  }

  static String format(int mins) {
    if (mins == unknown) return '-';
    final h = mins ~/ 60;
    final m = mins % 60;
    return '${h.toString().padLeft(2, '0')}:${m.toString().padLeft(2, '0')}';
  }
}

class XlsxSheetDoc {
  final Archive archive;
  final String sheetPath;
  final XmlDocument sheetDocument;
  final List<String> sharedStrings;

  XlsxSheetDoc({
    required this.archive,
    required this.sheetPath,
    required this.sheetDocument,
    required this.sharedStrings,
  });

  static XlsxSheetDoc fromBytes(Uint8List bytes, {String preferredSheetName = '평일'}) {
    final archive = ZipDecoder().decodeBytes(bytes, verify: false);
    final sheetPath = _resolveSheetPath(archive, preferredSheetName);
    final sheetXml = XmlDocument.parse(utf8.decode(_fileBytes(archive, sheetPath), allowMalformed: true));
    final sharedStrings = _parseSharedStrings(archive);

    return XlsxSheetDoc(
      archive: archive,
      sheetPath: sheetPath,
      sheetDocument: sheetXml,
      sharedStrings: sharedStrings,
    );
  }

  Iterable<XmlElement> get rows =>
      sheetDocument.descendants.whereType<XmlElement>().where((e) => e.name.local == 'row');

  int rowNumber(XmlElement row) {
    final direct = int.tryParse(row.getAttribute('r') ?? '');
    if (direct != null) return direct;

    final firstCell = firstOrNull(row.children.whereType<XmlElement>().where((e) => e.name.local == 'c'));
    if (firstCell == null) return -1;
    final ref = firstCell.getAttribute('r') ?? '';
    final digits = ref.replaceAll(RegExp(r'[^0-9]'), '');
    return int.tryParse(digits) ?? -1;
  }

  dynamic readCellValue(XmlElement row, int columnIndex) {
    final cell = _findCellByColumn(row, columnIndex);
    return _readCellValue(cell, sharedStrings);
  }

  dynamic readByRef(String cellRef) {
    final rowNum = _rowNumberFromRef(cellRef);
    final colIndex = _columnIndexFromRef(cellRef);
    if (rowNum < 1 || colIndex < 0) return null;
    final row = firstOrNull(rows.where((r) => rowNumber(r) == rowNum));
    if (row == null) return null;
    return readCellValue(row, colIndex);
  }

  void writeCell(XmlElement row, int columnIndex, Object? value, {required bool asText}) {
    final rowNum = rowNumber(row);
    if (rowNum < 1) return;
    final cell = _ensureCell(row, columnIndex, rowNum);
    _writeCellValue(cell, value, asText: asText);
  }

  void writeAt(int rowNumber1Based, int columnIndex, Object? value, {required bool asText}) {
    if (rowNumber1Based < 1) return;
    final row = _ensureRow(rowNumber1Based);
    final cell = _ensureCell(row, columnIndex, rowNumber1Based);
    _writeCellValue(cell, value, asText: asText);
  }

  Uint8List encode() {
    final updatedSheet = sheetDocument.toXmlString(pretty: false);
    final index = archive.files.indexWhere((f) => _normalizePath(f.name) == _normalizePath(sheetPath));
    if (index < 0) {
      throw Exception('시트 파일을 저장할 수 없습니다: $sheetPath');
    }
    archive.files[index] = ArchiveFile.string(sheetPath, updatedSheet);
    final encoded = ZipEncoder().encode(archive);
    if (encoded == null) {
      throw Exception('엑셀 zip 인코딩에 실패했습니다.');
    }
    return Uint8List.fromList(encoded);
  }

  static String _resolveSheetPath(Archive archive, String preferredSheetName) {
    final workbookPath = 'xl/workbook.xml';
    final workbook = XmlDocument.parse(utf8.decode(_fileBytes(archive, workbookPath), allowMalformed: true));
    final sheetElements =
        workbook.descendants.whereType<XmlElement>().where((e) => e.name.local == 'sheet').toList();
    if (sheetElements.isEmpty) {
      throw Exception('workbook.xml에서 시트를 찾지 못했습니다.');
    }

    final selected = firstOrNull(sheetElements.where((s) => s.getAttribute('name') == preferredSheetName)) ??
        sheetElements.first;
    final relId = _attributeByLocalName(selected, 'id');
    if (relId.isEmpty) {
      return _firstWorksheetPath(archive);
    }

    final relsPath = 'xl/_rels/workbook.xml.rels';
    final relsDoc = XmlDocument.parse(utf8.decode(_fileBytes(archive, relsPath), allowMalformed: true));
    final rel = firstOrNull(
      relsDoc.descendants.whereType<XmlElement>().where(
            (e) => e.name.local == 'Relationship' && e.getAttribute('Id') == relId,
          ),
    );

    final target = rel?.getAttribute('Target');
    if (target == null || target.isEmpty) {
      return _firstWorksheetPath(archive);
    }

    final normalizedTarget = _normalizeSheetTarget(target);
    final exists = archive.files.any((f) => _normalizePath(f.name) == _normalizePath(normalizedTarget));
    if (exists) return normalizedTarget;
    return _firstWorksheetPath(archive);
  }

  static String _firstWorksheetPath(Archive archive) {
    final worksheets = archive.files
        .map((f) => _normalizePath(f.name))
        .where((name) => name.startsWith('xl/worksheets/') && name.endsWith('.xml'))
        .toList()
      ..sort();
    if (worksheets.isEmpty) {
      throw Exception('엑셀 워크시트를 찾지 못했습니다.');
    }
    return worksheets.first;
  }

  static String _normalizeSheetTarget(String target) {
    var normalized = _normalizePath(target);
    while (normalized.startsWith('../')) {
      normalized = normalized.substring(3);
    }
    if (normalized.startsWith('/')) {
      normalized = normalized.substring(1);
    }
    if (normalized.startsWith('xl/')) {
      return normalized;
    }
    return 'xl/$normalized';
  }

  static List<String> _parseSharedStrings(Archive archive) {
    const sharedPath = 'xl/sharedStrings.xml';
    final exists = archive.files.any((f) => _normalizePath(f.name) == sharedPath);
    if (!exists) return const [];

    final doc = XmlDocument.parse(utf8.decode(_fileBytes(archive, sharedPath), allowMalformed: true));
    final siElements = doc.descendants.whereType<XmlElement>().where((e) => e.name.local == 'si');
    return siElements.map(_inlineTextFromSi).toList();
  }

  static String _inlineTextFromSi(XmlElement si) {
    final directT = firstOrNull(si.children.whereType<XmlElement>().where((e) => e.name.local == 't'));
    if (directT != null) return directT.innerText;
    return si
        .descendants
        .whereType<XmlElement>()
        .where((e) => e.name.local == 't')
        .map((e) => e.innerText)
        .join();
  }

  static String _attributeByLocalName(XmlElement element, String localName) {
    return firstOrNull(element.attributes.where((a) => a.name.local == localName))?.value ?? '';
  }

  static Uint8List _fileBytes(Archive archive, String path) {
    final normalizedPath = _normalizePath(path);
    final file = firstOrNull(archive.files.where((f) => _normalizePath(f.name) == normalizedPath));
    if (file == null) {
      throw Exception('엑셀 내부 파일을 찾을 수 없습니다: $path');
    }
    final content = file.content;
    if (content is Uint8List) return content;
    if (content is List<int>) return Uint8List.fromList(content);
    if (content is String) return Uint8List.fromList(utf8.encode(content));
    throw Exception('지원하지 않는 zip 파일 타입: ${content.runtimeType}');
  }

  static String _normalizePath(String path) => path.replaceAll('\\', '/');

  XmlElement _ensureRow(int rowNumber1Based) {
    final existing = firstOrNull(rows.where((r) => rowNumber(r) == rowNumber1Based));
    if (existing != null) return existing;

    final row = XmlElement(
      XmlName('row'),
      [XmlAttribute(XmlName('r'), rowNumber1Based.toString())],
      const [],
    );
    _sheetDataElement.children.add(row);
    return row;
  }

  XmlElement get _sheetDataElement {
    final element = firstOrNull(
      sheetDocument.descendants.whereType<XmlElement>().where((e) => e.name.local == 'sheetData'),
    );
    if (element == null) {
      throw Exception('sheetData 엘리먼트를 찾지 못했습니다.');
    }
    return element;
  }

  static XmlElement? _findCellByColumn(XmlElement row, int columnIndex) {
    for (final cell in row.children.whereType<XmlElement>().where((e) => e.name.local == 'c')) {
      final ref = cell.getAttribute('r') ?? '';
      final col = _columnIndexFromRef(ref);
      if (col == columnIndex) return cell;
    }
    return null;
  }

  static XmlElement _ensureCell(XmlElement row, int columnIndex, int rowIndex) {
    final existing = _findCellByColumn(row, columnIndex);
    if (existing != null) return existing;

    final ref = '${_columnName(columnIndex)}$rowIndex';
    final newCell = XmlElement(XmlName('c'), [XmlAttribute(XmlName('r'), ref)], []);
    row.children.add(newCell);
    return newCell;
  }

  static dynamic _readCellValue(XmlElement? cell, List<String> sharedStrings) {
    if (cell == null) return null;

    final type = cell.getAttribute('t') ?? '';
    if (type == 'inlineStr') {
      return _inlineTextFromCell(cell);
    }

    final vElement = firstOrNull(cell.children.whereType<XmlElement>().where((e) => e.name.local == 'v'));
    final raw = vElement?.innerText.trim() ?? '';
    if (raw.isEmpty) return null;

    if (type == 's') {
      final index = int.tryParse(raw);
      if (index == null || index < 0 || index >= sharedStrings.length) return '';
      return sharedStrings[index];
    }
    if (type == 'b') return raw == '1';

    final asNum = num.tryParse(raw);
    if (asNum != null) return asNum;
    return raw;
  }

  static String _inlineTextFromCell(XmlElement cell) {
    final isElement = firstOrNull(cell.children.whereType<XmlElement>().where((e) => e.name.local == 'is'));
    if (isElement == null) return '';
    final texts = isElement
        .descendants
        .whereType<XmlElement>()
        .where((e) => e.name.local == 't')
        .map((e) => e.innerText)
        .toList();
    return texts.join();
  }

  static void _writeCellValue(XmlElement cell, Object? value, {required bool asText}) {
    final clean = value?.toString() ?? '';
    cell.children.clear();

    if (clean.isEmpty) {
      cell.attributes.removeWhere((a) => a.name.local == 't');
      return;
    }

    if (asText) {
      cell.attributes.removeWhere((a) => a.name.local == 't');
      cell.attributes.add(XmlAttribute(XmlName('t'), 'inlineStr'));
      cell.children.add(
        XmlElement(
          XmlName('is'),
          const [],
          [XmlElement(XmlName('t'), const [], [XmlText(clean)])],
        ),
      );
      return;
    }

    cell.attributes.removeWhere((a) => a.name.local == 't');
    final numText = _normalizeNumberText(clean);
    cell.children.add(XmlElement(XmlName('v'), const [], [XmlText(numText)]));
  }

  static String _normalizeNumberText(String text) {
    final parsed = num.tryParse(text);
    if (parsed == null) return text;
    if (parsed is int) return parsed.toString();
    final formatted = parsed.toStringAsFixed(15);
    return formatted.replaceFirst(RegExp(r'0+$'), '').replaceFirst(RegExp(r'\.$'), '');
  }

  static int _columnIndexFromRef(String cellRef) {
    final letters = cellRef.toUpperCase().replaceAll(RegExp(r'[^A-Z]'), '');
    if (letters.isEmpty) return -1;
    var result = 0;
    for (final code in letters.codeUnits) {
      result = result * 26 + (code - 64);
    }
    return result - 1;
  }

  static int _rowNumberFromRef(String cellRef) {
    final digits = cellRef.replaceAll(RegExp(r'[^0-9]'), '');
    return int.tryParse(digits) ?? -1;
  }

  static String _columnName(int index) {
    var n = index + 1;
    final chars = <int>[];
    while (n > 0) {
      n--;
      chars.add(65 + (n % 26));
      n ~/= 26;
    }
    return String.fromCharCodes(chars.reversed);
  }
}

class CompanyDialParser {
  static const int headerRow = 3;
  static const int outSlotCol = 1;
  static const int outTrainCol = 3;
  static const int outTimeCol = 4;
  static const int outLaneCol = 5;
  static const int outNoteCol = 6;
  static const int inTrainCol = 9;
  static const int inTimeCol = 10;
  static const int inLaneCol = 11;
  static const int inspectCol = 12;
  static const int inNoteCol = 13;

  static ParsedCompanyDial parse(Uint8List bytes) {
    final xlsx = XlsxSheetDoc.fromBytes(bytes, preferredSheetName: '평일');
    final fixedReserveText = xlsx.readByRef('G35');
    final fixedReserveId = _normalizeId(fixedReserveText);

    final departures = <String, int>{};
    final departureRows = <ParsedDepartureRow>[];
    final arrivals = <ParsedArrivalRow>[];

    for (final row in xlsx.rows) {
      final rowNum = xlsx.rowNumber(row);
      if (rowNum <= headerRow + 1) continue;
      try {
        final outId = _normalizeId(xlsx.readCellValue(row, outTrainCol));
        if (outId.isNotEmpty) {
          final departureMins = TimeParser.parse(xlsx.readCellValue(row, outTimeCol));
          departures[outId] = departureMins;
          departureRows.add(
            ParsedDepartureRow(
              slot: _parseSlot(xlsx.readCellValue(row, outSlotCol), departureRows.length + 1),
              trainId: outId,
              departureMins: departureMins,
              lane: _stringify(xlsx.readCellValue(row, outLaneCol)),
              note: _stringify(xlsx.readCellValue(row, outNoteCol)),
            ),
          );
        }

        final inId = _normalizeId(xlsx.readCellValue(row, inTrainCol));
        if (inId.isNotEmpty) {
          arrivals.add(
            ParsedArrivalRow(
              trainId: inId,
              arrivalMins: TimeParser.parse(xlsx.readCellValue(row, inTimeCol)),
              lane: _stringify(xlsx.readCellValue(row, inLaneCol)),
              inspection: _stringify(xlsx.readCellValue(row, inspectCol)),
              note: _stringify(xlsx.readCellValue(row, inNoteCol)),
            ),
          );
        }
      } catch (_) {
        continue;
      }
    }

    return ParsedCompanyDial(
      departures: departures,
      departureRows: departureRows,
      arrivals: arrivals,
      okdongReserveFromG35: fixedReserveId.isEmpty ? null : fixedReserveId,
    );
  }

  static int _parseSlot(dynamic value, int fallback) {
    if (value is num) return value.toInt();
    final text = _stringify(value);
    final intValue = int.tryParse(text);
    if (intValue != null) return intValue;
    final doubleValue = double.tryParse(text);
    if (doubleValue != null && doubleValue == doubleValue.roundToDouble()) {
      return doubleValue.toInt();
    }
    return fallback;
  }

  static String _normalizeId(dynamic value) {
    if (value == null) return '';
    if (value is num) {
      final asInt = value.toInt();
      if (value == asInt.toDouble()) {
        if (asInt >= 101 && asInt <= 123) return asInt.toString();
      }
      return '';
    }
    final text = value.toString().trim();
    final intValue = int.tryParse(text);
    if (intValue != null) {
      if (intValue >= 101 && intValue <= 123) return intValue.toString();
      return '';
    }
    final doubleValue = double.tryParse(text);
    if (doubleValue != null && doubleValue == doubleValue.roundToDouble()) {
      final asInt = doubleValue.toInt();
      if (asInt >= 101 && asInt <= 123) return asInt.toString();
      return '';
    }
    if (!RegExp(r'^\d{3}$').hasMatch(text)) return '';
    final id = int.tryParse(text);
    if (id == null || id < 101 || id > 123) return '';
    return text;
  }

  static String _stringify(dynamic value) {
    if (value == null) return '';
    final text = value.toString().trim();
    if (text == 'null') return '';
    return text;
  }
}

class TrainAutoMapper {
  static const Set<TrainStatus> _autoManagedOnUpload = {
    TrainStatus.maintenance,
    TrainStatus.outOfService,
    TrainStatus.research,
    TrainStatus.night7D,
    TrainStatus.morning7D,
    TrainStatus.afternoon7D,
    TrainStatus.cleaning,
    TrainStatus.okdongStay,
    TrainStatus.okdongReserve,
  };

  static List<Train> applyTodayInfoBulk(List<Train> trains, ParsedCompanyDial parsed) {
    final latestByOkdongSlot = <String, ParsedArrivalRow>{};
    final explicitReserveIds = <String>{};

    for (final arrival in parsed.arrivals) {
      final combined = '${arrival.lane} ${arrival.note}';
      if (combined.contains('예비')) {
        explicitReserveIds.add(arrival.trainId);
      }

      final slot = _extractOkdongSlot(combined);
      if (slot == null) continue;

      final previous = latestByOkdongSlot[slot];
      if (previous == null || _isLaterArrival(arrival, previous)) {
        latestByOkdongSlot[slot] = arrival;
      }
    }

    final okdongStayIds = latestByOkdongSlot.values.map((a) => a.trainId).toSet();
    final reserveIds = parsed.okdongReserveFromG35 != null
        ? <String>{parsed.okdongReserveFromG35!}
        : (explicitReserveIds.isNotEmpty ? explicitReserveIds : _inferReserveIds(trains, parsed));

    return trains.map((train) {
      final dep = parsed.departures[train.id] ?? train.departureMins;
      final next = {...train.statuses}..removeAll(_autoManagedOnUpload);

      if (okdongStayIds.contains(train.id)) {
        next.add(TrainStatus.okdongStay);
      }
      if (reserveIds.contains(train.id)) {
        next.add(TrainStatus.okdongReserve);
        next.add(TrainStatus.okdongStay);
      }

      return train.copyWith(departureMins: dep, statuses: next);
    }).toList();
  }

  static Set<String> _inferReserveIds(List<Train> trains, ParsedCompanyDial parsed) {
    final allIds = trains.map((t) => t.id).toSet();
    final movedIds = <String>{...parsed.departures.keys, ...parsed.arrivals.map((a) => a.trainId)};
    final idleIds = allIds.difference(movedIds).toList()..sort();
    if (idleIds.length == 1) {
      return {idleIds.first};
    }
    return const <String>{};
  }

  static String? _extractOkdongSlot(String text) {
    final match = RegExp(r'옥동\s*([1-4])').firstMatch(text);
    if (match == null) return null;
    final num = match.group(1);
    if (num == null) return null;
    return '옥동$num';
  }

  static bool _isLaterArrival(ParsedArrivalRow a, ParsedArrivalRow b) {
    final aMins = a.arrivalMins == TimeParser.unknown ? -1 : a.arrivalMins;
    final bMins = b.arrivalMins == TimeParser.unknown ? -1 : b.arrivalMins;
    return aMins > bMins;
  }
}

class DoubleParkingEntry {
  final String trainId;
  final int timeMins;

  const DoubleParkingEntry({
    required this.trainId,
    required this.timeMins,
  });
}

class DoubleParkingLineResult {
  final String lineName;
  final List<DoubleParkingEntry> outbound;
  final List<DoubleParkingEntry> inbound;
  final String inboundGuide;
  final List<String> warnings;

  const DoubleParkingLineResult({
    required this.lineName,
    required this.outbound,
    required this.inbound,
    required this.inboundGuide,
    required this.warnings,
  });
}

class OkdongParkingSummary {
  final List<String> normalIds;
  final String? reserveId;
  final List<String> allIds;
  final List<String> warnings;

  const OkdongParkingSummary({
    required this.normalIds,
    required this.reserveId,
    required this.allIds,
    required this.warnings,
  });
}

class DoubleParkingResult {
  final List<DoubleParkingLineResult> lines;
  final String globalInboundGuide;
  final List<String> warnings;
  final OkdongParkingSummary okdongSummary;

  const DoubleParkingResult({
    required this.lines,
    required this.globalInboundGuide,
    required this.warnings,
    required this.okdongSummary,
  });
}

class DoubleParkingEngine {
  static const List<String> monitoredLines = ['M1', 'M2', 'D3', 'D2', 'D1', 'C2', 'C1'];

  static DoubleParkingResult simulate({
    required ParsedCompanyDial todayDial,
    required ParsedCompanyDial tomorrowDial,
  }) {
    final inboundByLine = <String, List<DoubleParkingEntry>>{
      for (final line in monitoredLines) line: <DoubleParkingEntry>[],
    };
    final outboundByLine = <String, List<DoubleParkingEntry>>{
      for (final line in monitoredLines) line: <DoubleParkingEntry>[],
    };
    final globalWarnings = <String>[];

    for (final arrival in todayDial.arrivals) {
      final line = _lineFromText('${arrival.lane} ${arrival.note}');
      if (line == null) continue;
      inboundByLine[line]!.add(
        DoubleParkingEntry(
          trainId: arrival.trainId,
          timeMins: arrival.arrivalMins,
        ),
      );
    }

    for (final departure in tomorrowDial.departureRows) {
      final line = _lineFromText('${departure.lane} ${departure.note}');
      if (line == null) continue;
      outboundByLine[line]!.add(
        DoubleParkingEntry(
          trainId: departure.trainId,
          timeMins: departure.departureMins,
        ),
      );
    }

    for (final line in monitoredLines) {
      inboundByLine[line]!.sort((a, b) => _timeSort(a.timeMins, b.timeMins));
      outboundByLine[line]!.sort((a, b) => _timeSort(a.timeMins, b.timeMins));
    }

    _validateDuplicateTrainUse(outboundByLine, globalWarnings, '명일 출고');
    _validateDuplicateTrainUse(inboundByLine, globalWarnings, '금일 입고');

    final lineResults = monitoredLines.map((line) {
      final inbound = inboundByLine[line]!;
      final outbound = outboundByLine[line]!;
      final warnings = <String>[];

      if (inbound.isNotEmpty && outbound.isNotEmpty) {
        final firstOut = outbound.first.timeMins;
        final lastIn = inbound.last.timeMins;
        if (_isKnownTime(firstOut) && _isKnownTime(lastIn) && firstOut < lastIn) {
          warnings.add('해당 선로에서 명일 첫 출고가 금일 마지막 입고보다 빠릅니다.');
        }
      }

      if (outbound.length > inbound.length + 1) {
        warnings.add('해당 선로의 출고 편성이 입고 대비 많아 초기 배치 점검이 필요합니다.');
      }

      final inboundGuide = inbound.isEmpty
          ? '입고 없음'
          : inbound
              .asMap()
              .entries
              .map((e) => '${e.key + 1}.${e.value.trainId}')
              .join(' -> ');

      return DoubleParkingLineResult(
        lineName: line,
        outbound: outbound,
        inbound: inbound,
        inboundGuide: inboundGuide,
        warnings: warnings,
      );
    }).toList();

    final globalInbound = lineResults.expand((l) => l.inbound).toList()
      ..sort((a, b) => _timeSort(a.timeMins, b.timeMins));
    final globalInboundGuide = globalInbound.isEmpty
        ? '입고 데이터 없음'
        : globalInbound
            .asMap()
            .entries
            .map((e) => '${e.key + 1}.${e.value.trainId}')
            .join(' -> ');

    final okdongSummary = _buildOkdongSummary(todayDial, tomorrowDial);
    globalWarnings.addAll(okdongSummary.warnings);

    return DoubleParkingResult(
      lines: lineResults,
      globalInboundGuide: globalInboundGuide,
      warnings: globalWarnings,
      okdongSummary: okdongSummary,
    );
  }

  static OkdongParkingSummary _buildOkdongSummary(
    ParsedCompanyDial todayDial,
    ParsedCompanyDial tomorrowDial,
  ) {
    final latestBySlot = <String, ParsedArrivalRow>{};
    String? explicitReserveId;
    final warnings = <String>[];

    for (final arrival in todayDial.arrivals) {
      final combined = '${arrival.lane} ${arrival.note}';
      final slot = _extractOkdongSlot(combined);
      if (slot != null) {
        final prev = latestBySlot[slot];
        if (prev == null || _timeSort(arrival.arrivalMins, prev.arrivalMins) > 0) {
          latestBySlot[slot] = arrival;
        }
      }
      if (combined.contains('옥동') && combined.contains('예비')) {
        explicitReserveId ??= arrival.trainId;
      }
    }

    final normalIds = latestBySlot.values.map((e) => e.trainId).toSet().toList()..sort();
    final reserveId = explicitReserveId;
    final allSet = {...normalIds};
    if (reserveId != null && reserveId.isNotEmpty) {
      allSet.add(reserveId);
    }
    final allIds = allSet.toList()..sort();

    if (allIds.length != 5) {
      warnings.add('옥동 주박 5대 규칙 불일치: 현재 ${allIds.length}대입니다.');
    }
    if (reserveId == null || reserveId.isEmpty) {
      warnings.add('금일 옥동 예비차를 명시적으로 찾지 못했습니다.');
    }

    final tomorrowOkdongOut =
        tomorrowDial.departureRows.where((a) => a.lane.contains('옥동')).map((a) => a.trainId).toSet();
    if (reserveId != null && reserveId.isNotEmpty && !tomorrowOkdongOut.contains(reserveId)) {
      warnings.add('금일 옥동 예비차($reserveId)가 명일 옥동 출고에 포함되지 않았습니다.');
    }

    return OkdongParkingSummary(
      normalIds: normalIds,
      reserveId: reserveId,
      allIds: allIds,
      warnings: warnings,
    );
  }

  static void _validateDuplicateTrainUse(
    Map<String, List<DoubleParkingEntry>> byLine,
    List<String> warnings,
    String kind,
  ) {
    final lineByTrain = <String, String>{};
    for (final entry in byLine.entries) {
      for (final item in entry.value) {
        final prev = lineByTrain[item.trainId];
        if (prev != null && prev != entry.key) {
          warnings.add('$kind 중복: ${item.trainId}가 $prev, ${entry.key}에 동시에 배치되었습니다.');
        } else {
          lineByTrain[item.trainId] = entry.key;
        }
      }
    }
  }

  static String? _extractOkdongSlot(String text) {
    final match = RegExp(r'옥동\s*([1-4])').firstMatch(text);
    if (match == null) return null;
    final num = match.group(1);
    return num == null ? null : '옥동$num';
  }

  static bool _isKnownTime(int mins) => mins != TimeParser.unknown;

  static int _timeSort(int a, int b) {
    final aa = a == TimeParser.unknown ? 999999 : a;
    final bb = b == TimeParser.unknown ? 999999 : b;
    return aa.compareTo(bb);
  }

  static String? _lineFromText(String text) {
    final upper = text.toUpperCase().replaceAll(' ', '');
    for (final line in monitoredLines) {
      if (upper.contains(line)) return line;
    }
    return null;
  }
}

class TomorrowDialEngine {
  static const Map<int, TrainStatus> fixedStatusSlots = {
    3: TrainStatus.cleaning,
    5: TrainStatus.night7D,
    7: TrainStatus.morning7D,
    9: TrainStatus.morning7D,
    10: TrainStatus.afternoon7D,
    15: TrainStatus.afternoon7D,
    12: TrainStatus.oneLoop,
    13: TrainStatus.oneLoop,
  };

  static TomorrowDialGenerationResult generateDetailed({
    required List<Train> trains,
    required ParsedCompanyDial todayDial,
    required Set<String> okdongForbidden,
  }) {
    final slots = todayDial.departureRows.take(18).toList()..sort((a, b) => a.slot.compareTo(b.slot));
    final assignmentsBySlot = <int, TomorrowAssignment>{};
    final usedTrainIds = <String>{};
    final logs = <String>[];
    final available = trains.where((t) => !_isExcluded(t)).toList();

    Train? pickForSlot(
      ParsedDepartureRow slot,
      bool Function(Train) predicate, {
      bool highMileage = false,
    }) {
      final isOkdongSlot = _isOkdongLane(slot.lane);
      final pool = available.where((train) {
        if (usedTrainIds.contains(train.id)) return false;
        if (isOkdongSlot && okdongForbidden.contains(train.id)) return false;
        return predicate(train);
      }).toList()
        ..sort((a, b) => _compareTrain(a, b, highMileage: highMileage));
      return firstOrNull(pool);
    }

    bool assignSlot(
      ParsedDepartureRow slot,
      Train train, {
      required String note,
      required String reason,
    }) {
      if (assignmentsBySlot.containsKey(slot.slot)) {
        logs.add('slot ${slot.slot}: skipped($reason), already assigned');
        return false;
      }
      if (usedTrainIds.contains(train.id)) {
        logs.add('slot ${slot.slot}: skipped($reason), train ${train.id} already used');
        return false;
      }

      assignmentsBySlot[slot.slot] = TomorrowAssignment(
        slot: slot.slot,
        trainId: train.id,
        departureMins: slot.departureMins,
        lane: slot.lane,
        note: note,
        reason: reason,
      );
      usedTrainIds.add(train.id);
      logs.add('slot ${slot.slot}: ${train.id} ($reason)');
      return true;
    }

    // Level 1: absolute rules
    final tomorrowReserveSlot = firstOrNull(slots.where((s) => s.slot == 18));
    if (tomorrowReserveSlot != null) {
      final reserve = pickForSlot(
            tomorrowReserveSlot,
            (t) => t.has(TrainStatus.tomorrowReserve),
          ) ??
          pickForSlot(
            tomorrowReserveSlot,
            (t) => t.has(TrainStatus.okdongStay) && !t.has(TrainStatus.okdongReserve),
          );
      if (reserve != null) {
        assignSlot(
          tomorrowReserveSlot,
          reserve,
          note: statusLabel(TrainStatus.tomorrowReserve),
          reason: 'L1 tomorrow reserve fixed',
        );
      } else {
        logs.add('slot 18: no candidate for L1 tomorrow reserve');
      }
    }

    final okdongOutboundSlots = slots.where((s) => _isOkdongLane(s.lane) && s.slot != 18).toList()
      ..sort((a, b) => a.slot.compareTo(b.slot));

    final firstOkdongOutbound = firstOrNull(
      okdongOutboundSlots.where((s) => !assignmentsBySlot.containsKey(s.slot)),
    );
    if (firstOkdongOutbound != null) {
      final reserveOut = pickForSlot(firstOkdongOutbound, (t) => t.has(TrainStatus.okdongReserve));
      if (reserveOut != null) {
        assignSlot(
          firstOkdongOutbound,
          reserveOut,
          note: statusLabel(TrainStatus.okdongReserve),
          reason: 'L1 today reserve must outbound',
        );
      } else {
        logs.add('slot ${firstOkdongOutbound.slot}: no candidate for L1 today reserve outbound');
      }
    }

    // Level 2: okdong structure + forbidden filter
    for (final slot in okdongOutboundSlots) {
      if (assignmentsBySlot.containsKey(slot.slot)) continue;
      final train =
          pickForSlot(slot, (t) => t.has(TrainStatus.okdongStay)) ?? pickForSlot(slot, (t) => true);
      if (train != null) {
        final reason = train.has(TrainStatus.okdongStay)
            ? 'L2 keep okdong structure'
            : 'L2 keep okdong structure (fallback)';
        assignSlot(
          slot,
          train,
          note: statusLabel(TrainStatus.okdongStay),
          reason: reason,
        );
      } else {
        logs.add('slot ${slot.slot}: unassigned at L2 (no okdong candidate)');
      }
    }

    // Level 3: fixed status slots
    for (final slot in slots) {
      if (assignmentsBySlot.containsKey(slot.slot)) continue;
      final fixedStatus = fixedStatusSlots[slot.slot];
      if (fixedStatus == null) continue;

      final train = pickForSlot(
        slot,
        (t) => t.has(fixedStatus),
        highMileage: fixedStatus == TrainStatus.oneLoop,
      );
      if (train != null) {
        assignSlot(
          slot,
          train,
          note: statusLabel(fixedStatus),
          reason: 'L3 fixed slot ${statusLabel(fixedStatus)}',
        );
      } else {
        logs.add('slot ${slot.slot}: no candidate for L3 fixed slot ${statusLabel(fixedStatus)}');
      }
    }

    // Level 4: dial guidance
    for (final slot in slots) {
      if (assignmentsBySlot.containsKey(slot.slot)) continue;
      final train = pickForSlot(slot, (t) => _matchesDialGuide(t, slot));
      if (train != null) {
        final fixedStatus = fixedStatusSlots[slot.slot];
        final note = fixedStatus == null ? '' : '${statusLabel(fixedStatus)} 보강';
        assignSlot(
          slot,
          train,
          note: note,
          reason: 'L4 dial guide',
        );
      }
    }

    // Level 5: deterministic mileage fallback
    for (final slot in slots) {
      if (assignmentsBySlot.containsKey(slot.slot)) continue;
      final fixedStatus = fixedStatusSlots[slot.slot];
      final train = pickForSlot(
        slot,
        (t) => true,
        highMileage: fixedStatus == TrainStatus.oneLoop,
      );
      if (train != null) {
        final note = fixedStatus == null ? '' : '${statusLabel(fixedStatus)} 대체';
        assignSlot(
          slot,
          train,
          note: note,
          reason: 'L5 mileage fallback',
        );
      } else {
        logs.add('slot ${slot.slot}: unassigned at L5 (no available train)');
      }
    }

    final assignments = assignmentsBySlot.values.toList()..sort((a, b) => a.slot.compareTo(b.slot));
    final unassignedSlots =
        slots.where((s) => !assignmentsBySlot.containsKey(s.slot)).map((s) => s.slot).toList()..sort();
    if (unassignedSlots.isNotEmpty) {
      logs.add('unassigned slots: ${unassignedSlots.join(', ')}');
    }

    return TomorrowDialGenerationResult(
      assignments: assignments,
      logs: logs,
      unassignedSlots: unassignedSlots,
    );
  }

  static List<TomorrowAssignment> generate({
    required List<Train> trains,
    required ParsedCompanyDial todayDial,
    required Set<String> okdongForbidden,
  }) {
    final slots = todayDial.departureRows.take(18).toList();
    final assigned = <TomorrowAssignment>[];
    final used = <String>{};
    final available = trains.where((t) => !_isExcluded(t)).toList()
      ..sort((a, b) => a.mileage.compareTo(b.mileage));

    TomorrowAssignment assign(ParsedDepartureRow slot, Train train, String note) {
      used.add(train.id);
      return TomorrowAssignment(
        slot: slot.slot,
        trainId: train.id,
        departureMins: slot.departureMins,
        lane: slot.lane,
        note: note,
      );
    }

    Train? pick(bool Function(Train) test, {bool highMileage = false}) {
      final pool = available.where((t) => !used.contains(t.id) && test(t)).toList();
      pool.sort((a, b) => highMileage ? b.mileage.compareTo(a.mileage) : a.mileage.compareTo(b.mileage));
      return firstOrNull(pool);
    }

    final tomorrowReserveSlot = firstOrNull(slots.where((s) => s.slot == 18));
    if (tomorrowReserveSlot != null) {
      final reserve = pick((t) => t.has(TrainStatus.tomorrowReserve) && !okdongForbidden.contains(t.id)) ??
          pick((t) => t.has(TrainStatus.okdongStay) && !t.has(TrainStatus.okdongReserve) && !okdongForbidden.contains(t.id));
      if (reserve != null) {
        assigned.add(assign(tomorrowReserveSlot, reserve, '명일 옥동 예비'));
      }
    }

    for (final slot in slots.where((s) => _isOkdongLane(s.lane) && s.slot != 18)) {
      final train = pick((t) => t.has(TrainStatus.okdongReserve) && !okdongForbidden.contains(t.id)) ??
          pick((t) => t.has(TrainStatus.okdongStay) && !okdongForbidden.contains(t.id)) ??
          pick((t) => !okdongForbidden.contains(t.id));
      if (train != null) {
        assigned.add(assign(slot, train, train.has(TrainStatus.okdongReserve) ? '금일 옥동 예비 출고' : '옥동 출고'));
      }
    }

    for (final slot in slots.where((s) => !usedSlots(assigned).contains(s.slot))) {
      final fixedStatus = fixedStatusSlots[slot.slot];
      Train? train;
      if (fixedStatus != null) {
        train = pick((t) => t.has(fixedStatus), highMileage: fixedStatus == TrainStatus.oneLoop);
      }
      train ??= pick((t) => _matchesDialGuide(t, slot));
      train ??= pick((t) => true);
      if (train != null) {
        assigned.add(assign(slot, train, fixedStatus == null ? '' : statusLabel(fixedStatus)));
      }
    }

    assigned.sort((a, b) => a.slot.compareTo(b.slot));
    return assigned;
  }

  static Set<int> usedSlots(List<TomorrowAssignment> assigned) {
    return assigned.map((a) => a.slot).toSet();
  }

  static bool _isExcluded(Train train) {
    return train.has(TrainStatus.maintenance) ||
        train.has(TrainStatus.outOfService) ||
        train.has(TrainStatus.research);
  }

  static bool _isOkdongLane(String lane) => lane.contains('옥동');

  static int _compareTrain(
    Train a,
    Train b, {
    required bool highMileage,
  }) {
    final mileageCompare = highMileage ? b.mileage.compareTo(a.mileage) : a.mileage.compareTo(b.mileage);
    if (mileageCompare != 0) return mileageCompare;
    return a.id.compareTo(b.id);
  }

  static bool _matchesDialGuide(Train train, ParsedDepartureRow slot) {
    if (train.has(TrainStatus.longDia)) return slot.note.contains('10') || slot.note.contains('6.5');
    if (train.has(TrainStatus.mediumDia)) return slot.note.contains('5.5') || slot.note.contains('4');
    if (train.has(TrainStatus.shortDia)) return slot.note.contains('2') || slot.note.contains('3');
    return false;
  }
}

class CompanyExcelWriter {
  static Uint8List write({
    required Uint8List templateBytes,
    required List<TomorrowAssignment> assignments,
    DoubleParkingResult? doubleParkingResult,
  }) {
    final xlsx = XlsxSheetDoc.fromBytes(templateBytes, preferredSheetName: '평일');

    final bySlot = {for (final assignment in assignments) assignment.slot: assignment};
    for (final row in xlsx.rows) {
      final rowNum = xlsx.rowNumber(row);
      if (rowNum <= CompanyDialParser.headerRow + 1) continue;

      final slot = CompanyDialParser._parseSlot(xlsx.readCellValue(row, CompanyDialParser.outSlotCol), -1);
      final assignment = bySlot[slot];
      if (assignment == null) continue;

      xlsx.writeCell(row, CompanyDialParser.outTrainCol, assignment.trainId, asText: true);
      xlsx.writeCell(
        row,
        CompanyDialParser.outTimeCol,
        TimeParser.toExcelTime(assignment.departureMins),
        asText: false,
      );
      xlsx.writeCell(row, CompanyDialParser.outLaneCol, assignment.lane, asText: true);
      xlsx.writeCell(row, CompanyDialParser.outNoteCol, assignment.note, asText: true);
    }

    if (doubleParkingResult != null) {
      _writeDoubleParkingBlock(xlsx, doubleParkingResult);
    }

    return xlsx.encode();
  }

  static void _writeDoubleParkingBlock(XlsxSheetDoc xlsx, DoubleParkingResult result) {
    const int startCol = 22; // W
    var row = 4;

    xlsx.writeAt(row, startCol, '이중주차 시뮬레이터', asText: true);
    row++;
    xlsx.writeAt(row, startCol, '전체 입고 순서', asText: true);
    xlsx.writeAt(row, startCol + 1, result.globalInboundGuide, asText: true);
    row++;
    xlsx.writeAt(row, startCol, '옥동 일반4', asText: true);
    xlsx.writeAt(row, startCol + 1, result.okdongSummary.normalIds.join(', '), asText: true);
    row++;
    xlsx.writeAt(row, startCol, '옥동 예비1', asText: true);
    xlsx.writeAt(row, startCol + 1, result.okdongSummary.reserveId ?? '미확인', asText: true);
    row++;
    xlsx.writeAt(row, startCol, '옥동 총합', asText: true);
    xlsx.writeAt(row, startCol + 1, result.okdongSummary.allIds.join(', '), asText: true);
    row++;

    final allWarnings = <String>[
      ...result.warnings,
      ...result.lines.expand((l) => l.warnings),
    ];
    if (allWarnings.isNotEmpty) {
      xlsx.writeAt(row, startCol, '경고', asText: true);
      xlsx.writeAt(row, startCol + 1, allWarnings.join(' | '), asText: true);
      row++;
    }

    row++;
    xlsx.writeAt(row, startCol, '선로', asText: true);
    xlsx.writeAt(row, startCol + 1, '명일 OUT', asText: true);
    xlsx.writeAt(row, startCol + 2, '금일 IN', asText: true);
    xlsx.writeAt(row, startCol + 3, '선로 경고', asText: true);
    row++;

    for (final line in result.lines) {
      xlsx.writeAt(row, startCol, line.lineName, asText: true);
      xlsx.writeAt(row, startCol + 1, _entriesText(line.outbound), asText: true);
      xlsx.writeAt(row, startCol + 2, _entriesText(line.inbound), asText: true);
      xlsx.writeAt(row, startCol + 3, line.warnings.join(' | '), asText: true);
      row++;
    }
  }

  static String _entriesText(List<DoubleParkingEntry> entries) {
    if (entries.isEmpty) return '없음';
    return entries.map((e) => '${e.trainId}(${TimeParser.format(e.timeMins)})').join(', ');
  }
}
