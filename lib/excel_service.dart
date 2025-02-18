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
      _showMessage(context, "Erro ao carregar a planilha modelo: $e");
    }
  }

  Future<void> exportData(BuildContext context, Map<String, TextEditingController> controllers) async {
    if (templateBytes == null) {
      _showMessage(context, "Erro ao carregar a planilha modelo");
      return;
    }

    var excel = Excel.decodeBytes(templateBytes!);
    if (!excel.sheets.containsKey(selectedSheet)) {
      _showMessage(context, "A sheet selecionada não existe!");
      return;
    }

    Map<String, String> formData = {};

    controllers.forEach((key, controller) {
      String value = controller.text.trim();
      if (value.isNotEmpty) {
        formData[key] = value; // Apenas preenche se houver texto
      }
    });

    _updateExcelSheet(context, excel, formData);
    await _saveExcelFile(context, excel);
  }

  void _updateExcelSheet(BuildContext context, Excel excel, Map<String, String> formData) {
    Sheet sheet = excel[selectedSheet]!;

    // Atualizando Cliente e Local apenas se houver valor
    if (formData["Cliente"] != null) {
      sheet.updateCell(CellIndex.indexByString("E5"), TextCellValue(formData["Cliente"]!));
    }
    if (formData["Local"] != null) {
      sheet.updateCell(CellIndex.indexByString("E6"), TextCellValue(formData["Local"]!));
    }

    // Equipe de Execução (somente se houver valores)
    List<String> nomeCells = ["C30", "C31", "C32", "C33"];
    List<String> funcaoCells = ["I30", "I31", "I32", "I33"];
    List<String> empresaCells = ["L30", "L31", "L32", "L33"];

    for (int i = 0; i < 4; i++) {
      if (formData["Nome ${i + 1}"] != null) {
        sheet.updateCell(CellIndex.indexByString(nomeCells[i]), TextCellValue(formData["Nome ${i + 1}"]!));
      }
      if (formData["Função ${i + 1}"] != null) {
        sheet.updateCell(CellIndex.indexByString(funcaoCells[i]), TextCellValue(formData["Função ${i + 1}"]!));
      }
      if (formData["Empresa ${i + 1}"] != null) {
        sheet.updateCell(CellIndex.indexByString(empresaCells[i]), TextCellValue(formData["Empresa ${i + 1}"]!));
      }
    }

    // Atualizando Datas e Horários apenas se houver valores
    if (formData["Data de Início"] != null) {
      sheet.updateCell(CellIndex.indexByString("D42"), TextCellValue(formData["Data de Início"]!));
    }
    if (formData["Hora de Início"] != null) {
      sheet.updateCell(CellIndex.indexByString("F42"), TextCellValue(formData["Hora de Início"]!));
    }
    if (formData["Data de Término"] != null) {
      sheet.updateCell(CellIndex.indexByString("J42"), TextCellValue(formData["Data de Término"]!));
    }
    if (formData["Hora de Término"] != null) {
      sheet.updateCell(CellIndex.indexByString("L42"), TextCellValue(formData["Hora de Término"]!));
    }

    // Atualizando medições apenas se houver valores
    int row = 45;
    for (var inv in ["INV01", "INV02", "INV03", "INV04"]) {
      for (var phase in [
        "Fase R/Terra", "Fase S/Terra", "Fase T/Terra", "Fase R/Fase S", "Fase R/Fase T", "Fase S/Fase T"
      ]) {
        if (formData["$inv $phase"] != null) {
          sheet.updateCell(CellIndex.indexByString("E$row"), TextCellValue(formData["$inv $phase"]!));
        }
        if (formData["$inv $phase Temp. Amb. (°C)"] != null) {
          sheet.updateCell(CellIndex.indexByString("I$row"), TextCellValue(formData["$inv $phase Temp. Amb. (°C)"]!));
        }
        if (formData["$inv $phase UMID. REL (%)"] != null) {
          sheet.updateCell(CellIndex.indexByString("J$row"), TextCellValue(formData["$inv $phase UMID. REL (%)"]!));
        }
        row++;
      }
    }
  }

  Future<void> _saveExcelFile(BuildContext context, Excel excel) async {
    try {
      Directory? outputDir = await getExternalStorageDirectory();
      if (outputDir == null) {
        _showMessage(context, "Erro: Não foi possível acessar o diretório de armazenamento.");
        return;
      }
      String outputPath = "${outputDir.path}/dados_exportados.xlsx";
      File file = File(outputPath);
      file.createSync(recursive: true);
      file.writeAsBytesSync(excel.encode()!);
      _showMessage(context, "Dados exportados com sucesso: $outputPath");
    } catch (e) {
      _showMessage(context, "Erro ao salvar o arquivo: $e");
    }
  }

  void _showMessage(BuildContext context, String message) {
    ScaffoldMessenger.of(context).showSnackBar(SnackBar(content: Text(message)));
  }
}
