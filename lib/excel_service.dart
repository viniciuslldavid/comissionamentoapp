import 'package:excel/excel.dart';
import 'dart:io';
import 'package:path_provider/path_provider.dart';
import 'package:flutter/services.dart' show ByteData, rootBundle;
import 'package:flutter/material.dart';

class ExcelService {
  List<int>? templateBytes;
  final String selectedSheet = "TE01";

  Future<void> loadTemplateFromAssets(BuildContext context) async {
    try {
      ByteData data = await rootBundle.load("assets/spreadsheettemplates/megohmetroca.xlsx");
      templateBytes = data.buffer.asUint8List();
    } catch (e) {
      _showMessage(context, "Erro ao carregar a planilha modelo");
    }
  }

  Future<void> exportData(BuildContext context, Map<String, TextEditingController> controllers) async {
    if (templateBytes == null) {
      _showMessage(context, "Erro ao carregar a planilha modelo");
      return;
    }

    var excel = Excel.decodeBytes(templateBytes!)!;
    if (!excel.sheets.keys.contains(selectedSheet)) {
      _showMessage(context, "A sheet selecionada não existe!");
      return;
    }

    _updateExcelSheet(context, excel, controllers);
    await _saveExcelFile(context, excel);
  }

  void _updateExcelSheet(BuildContext context, Excel excel, Map<String, TextEditingController> controllers) {
    Sheet sheet = excel[selectedSheet]!;

    // Atualizando informações gerais
    sheet.updateCell(CellIndex.indexByString("E5"), TextCellValue(controllers["Cliente"]?.text.trim() ?? ""));
    sheet.updateCell(CellIndex.indexByString("E6"), TextCellValue(controllers["Local"]?.text.trim() ?? ""));

    // Atualizando equipe de execução
    List<String> nomeCells = ["C30", "C31", "C32", "C33"];
    List<String> funcaoCells = ["I30", "I31", "I32", "I33"];
    List<String> empresaCells = ["L30", "L31", "L32", "L33"];

    for (int i = 0; i < 4; i++) {
      sheet.updateCell(CellIndex.indexByString(nomeCells[i]), TextCellValue(controllers["Nome ${i + 1}"]?.text.trim() ?? ""));
      sheet.updateCell(CellIndex.indexByString(funcaoCells[i]), TextCellValue(controllers["Função ${i + 1}"]?.text.trim() ?? ""));
      sheet.updateCell(CellIndex.indexByString(empresaCells[i]), TextCellValue(controllers["Empresa ${i + 1}"]?.text.trim() ?? ""));
    }

    // Atualizando dados dos ensaios
    sheet.updateCell(CellIndex.indexByString("D42"), TextCellValue(controllers["Data de Início"]?.text.trim() ?? ""));
    sheet.updateCell(CellIndex.indexByString("F42"), TextCellValue(controllers["Hora de Início"]?.text.trim() ?? ""));
    sheet.updateCell(CellIndex.indexByString("J42"), TextCellValue(controllers["Data de Término"]?.text.trim() ?? ""));
    sheet.updateCell(CellIndex.indexByString("L42"), TextCellValue(controllers["Hora de Término"]?.text.trim() ?? ""));

    // Atualizando os dados dos INV
    int row = 45;
    for (var inv in ["INV01", "INV02", "INV03", "INV04"]) {
      for (var phase in [
        "Fase R/Terra",
        "Fase S/Terra",
        "Fase T/Terra",
        "Fase R/Fase S",
        "Fase R/Fase T",
        "Fase S/Fase T"
      ]) {
        sheet.updateCell(CellIndex.indexByString("E$row"), TextCellValue(controllers["$inv $phase"]?.text.trim() ?? ""));
        sheet.updateCell(CellIndex.indexByString("I$row"), TextCellValue(controllers["$inv $phase Temp. Amb. (°C)"]?.text.trim() ?? ""));
        sheet.updateCell(CellIndex.indexByString("J$row"), TextCellValue(controllers["$inv $phase UMID. REL (%)"]?.text.trim() ?? ""));
        row++;
      }
    }
  }

  Future<void> _saveExcelFile(BuildContext context, Excel excel) async {
    try {
      if (Platform.isAndroid || Platform.isIOS) {
        var outputFile = await getExternalStorageDirectory();
        String outputPath = "${outputFile!.path}/dados_exportados.xlsx";
        File(outputPath)
          ..createSync(recursive: true)
          ..writeAsBytesSync(excel.encode()!);
        _showMessage(context, "Dados exportados para $outputPath");
      } else {
        _showMessage(context, "Exportação de arquivos não suportada nesta plataforma.");
      }
    } catch (e) {
      _showMessage(context, "Erro ao salvar o arquivo");
    }
  }

  void _showMessage(BuildContext context, String message) {
    ScaffoldMessenger.of(context).showSnackBar(SnackBar(content: Text(message)));
  }
}
