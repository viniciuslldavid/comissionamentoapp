import 'package:flutter/material.dart';
import 'excel_service.dart';

class DataExportScreen extends StatefulWidget {
  const DataExportScreen({super.key});

  @override
  _DataExportScreenState createState() => _DataExportScreenState();
}

class _DataExportScreenState extends State<DataExportScreen> {
  final ExcelService excelService = ExcelService();
  final Map<String, TextEditingController> controllers = {};

  final TextEditingController clienteController = TextEditingController();
  final TextEditingController localController = TextEditingController();
  final TextEditingController dataInicioController = TextEditingController();
  final TextEditingController horaInicioController = TextEditingController();
  final TextEditingController dataTerminoController = TextEditingController();
  final TextEditingController horaTerminoController = TextEditingController();

  @override
  void initState() {
    super.initState();
    excelService.loadTemplateFromAssets(context);
  }

  @override
  void dispose() {
    clienteController.dispose();
    localController.dispose();
    dataInicioController.dispose();
    horaInicioController.dispose();
    dataTerminoController.dispose();
    horaTerminoController.dispose();
    super.dispose();
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(title: const Text("Ato. Comissionamento")),
      body: Padding(
        padding: const EdgeInsets.all(16.0),
        child: ListView(
          children: [
            _buildTextField(clienteController, "Cliente"),
            _buildTextField(localController, "Local"),
            _buildTextField(dataInicioController, "Data de Início"),
            _buildTextField(horaInicioController, "Hora de Início"),
            _buildTextField(dataTerminoController, "Data de Término"),
            _buildTextField(horaTerminoController, "Hora de Término"),
          ],
        ),
      ),
      floatingActionButton: FloatingActionButton(
        onPressed: () {
          controllers["Cliente"] = clienteController;
          controllers["Local"] = localController;
          controllers["Data de Início"] = dataInicioController;
          controllers["Hora de Início"] = horaInicioController;
          controllers["Data de Término"] = dataTerminoController;
          controllers["Hora de Término"] = horaTerminoController;

          excelService.exportData(context, controllers);
        },
        child: const Icon(Icons.save),
      ),
    );
  }

  Widget _buildTextField(TextEditingController controller, String label) {
    return Padding(
      padding: const EdgeInsets.symmetric(vertical: 8.0),
      child: TextField(
        controller: controller,
        decoration: InputDecoration(labelText: label, border: OutlineInputBorder()),
      ),
    );
  }
}