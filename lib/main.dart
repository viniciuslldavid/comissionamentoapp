import 'package:flutter/material.dart';
import 'package:excel/excel.dart';
import 'dart:io';
import 'package:path_provider/path_provider.dart';
import 'package:flutter/services.dart' show ByteData, rootBundle;

void main() {
  runApp(MyApp());
}

class MyApp extends StatelessWidget {
  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      home: DataExportScreen(),
    );
  }
}

class DataExportScreen extends StatefulWidget {
  @override
  _DataExportScreenState createState() => _DataExportScreenState();
}

class _DataExportScreenState extends State<DataExportScreen> {
  TextEditingController nameController = TextEditingController();
  TextEditingController ageController = TextEditingController();
  List<int>? templateBytes;
  String selectedSheet = "TE01";

  Future<void> loadTemplateFromAssets() async {
    ByteData data = await rootBundle.load("assets/modelo.xlsx");
    setState(() {
      templateBytes = data.buffer.asUint8List();
    });
  }

  Future<void> exportData() async {
    if (templateBytes == null) {
      ScaffoldMessenger.of(context).showSnackBar(
        SnackBar(content: Text("Erro ao carregar a planilha modelo")),
      );
      return;
    }
    var excel = Excel.decodeBytes(templateBytes!);

    if (!excel.sheets.keys.contains(selectedSheet)) {
      ScaffoldMessenger.of(context).showSnackBar(
        SnackBar(content: Text("A sheet selecionada não existe!")),
      );
      return;
    }

    Sheet sheet = excel[selectedSheet]!;
    sheet.updateCell(CellIndex.indexByString("H45"), TextCellValue(nameController.text));
    sheet.updateCell(CellIndex.indexByString("I45"), TextCellValue(ageController.text));

    var outputFile = await getExternalStorageDirectory();
    String outputPath = "${outputFile!.path}/dados_exportados.xlsx";
    File(outputPath)
      ..createSync(recursive: true)
      ..writeAsBytesSync(excel.encode()!);

    ScaffoldMessenger.of(context).showSnackBar(
      SnackBar(content: Text("Dados exportados para $outputPath")),
    );
  }

  @override
  void initState() {
    super.initState();
    loadTemplateFromAssets();
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(title: Text("Exportação de Dados")),
      body: Padding(
        padding: EdgeInsets.all(16.0),
        child: Column(
          children: [
            TextField(controller: nameController, decoration: InputDecoration(labelText: "Nome")),
            TextField(controller: ageController, decoration: InputDecoration(labelText: "Idade"), keyboardType: TextInputType.number),
            SizedBox(height: 10),
            ElevatedButton(
              onPressed: exportData,
              child: Text("Exportar Dados"),
            ),
          ],
        ),
      ),
    );
  }
}
