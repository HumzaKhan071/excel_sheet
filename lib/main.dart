import 'dart:convert';
import 'dart:io';
import 'package:flutter/material.dart';
import 'package:open_filex/open_filex.dart';
import 'package:path_provider/path_provider.dart';
import 'package:syncfusion_flutter_xlsio/xlsio.dart';
import 'package:universal_html/html.dart' show AnchorElement;
import 'package:flutter/foundation.dart' show kIsWeb;

void main() {
  runApp(const MyApp());
}

class MyApp extends StatelessWidget {
  const MyApp({super.key});

  // This widget is the root of your application.
  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      title: 'Flutter Demo',
      home: const MyHomePage(),
    );
  }
}

class MyHomePage extends StatelessWidget {
  const MyHomePage({super.key});

  @override
  Widget build(BuildContext context) {
    return Scaffold(
        body: Center(
      child: TextButton(
        onPressed: createExcelSheet,
        child: const Text('Create ExcelSheet'),
      ),
    ));
  }
}

Future<void> createExcelSheet() async {
  final Workbook workbook = Workbook();
  final Worksheet sheet = workbook.worksheets[0];
  sheet.getRangeByName('A1').setNumber(22);
  sheet.getRangeByName('A2').setNumber(44);

//Formula calculation is enabled for the sheet
  sheet.enableSheetCalculations();

//Setting formula in the cell
  sheet.getRangeByName('A3').setFormula('=A1+A2');
  final List<int> bytes = workbook.saveAsStream();
  workbook.dispose();

  if (kIsWeb) {
    AnchorElement(
        href:
            'data:application/octet-stream;charset=utf-161e;base64,${base64.encode(bytes)}')
      ..setAttribute('download', 'ExcelSheet.xlsx')
      ..click();
  } else {
    final String path = (await getApplicationSupportDirectory()).path;
    final String fileName =
        Platform.isWindows ? '$path/ExcelSheet.xlsx' : '$path/ExcelSheet.xlsx';
    final File file = File(fileName);
    await file.writeAsBytes(bytes, flush: true);

    OpenFilex.open(fileName);
  }
}
