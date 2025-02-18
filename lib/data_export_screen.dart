import 'package:flutter/material.dart';
import 'excel_service.dart';

class DataExportScreen extends StatefulWidget {
  const DataExportScreen({super.key});

  @override
  _DataExportScreenState createState() => _DataExportScreenState();
}

class _DataExportScreenState extends State<DataExportScreen> {
  final ExcelService excelService = ExcelService();
  final List<String> invNumbers = ['INV01', 'INV02', 'INV03', 'INV04'];
  final List<String> phases = [
    'Fase R/Terra',
    'Fase S/Terra',
    'Fase T/Terra',
    'Fase R/Fase S',
    'Fase R/Fase T',
    'Fase S/Fase T'
  ];
  final Map<String, TextEditingController> controllers = {};

  final List<TextEditingController> nomeControllers = List.generate(4, (_) => TextEditingController());
  final List<TextEditingController> funcaoControllers = List.generate(4, (_) => TextEditingController());
  final List<TextEditingController> empresaControllers = List.generate(4, (_) => TextEditingController());
  final TextEditingController dataInicioController = TextEditingController();
  final TextEditingController horaInicioController = TextEditingController();
  final TextEditingController dataTerminoController = TextEditingController();
  final TextEditingController horaTerminoController = TextEditingController();
  final TextEditingController clienteController = TextEditingController();
  final TextEditingController localController = TextEditingController();

  @override
  void initState() {
    super.initState();
    excelService.loadTemplateFromAssets(context);
    for (var inv in invNumbers) {
      for (var phase in phases) {
        controllers['$inv $phase'] = TextEditingController();
        controllers['$inv $phase Temp. Amb. (°C)'] = TextEditingController();
        controllers['$inv $phase UMID. REL (%)'] = TextEditingController();
      }
    }
  }

  @override
  void dispose() {
    for (var controller in controllers.values) {
      controller.dispose();
    }
    nomeControllers.forEach((controller) => controller.dispose());
    funcaoControllers.forEach((controller) => controller.dispose());
    empresaControllers.forEach((controller) => controller.dispose());
    dataInicioController.dispose();
    horaInicioController.dispose();
    dataTerminoController.dispose();
    horaTerminoController.dispose();
    clienteController.dispose();
    localController.dispose();
    super.dispose();
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(
        title: const Text("Ato. Comissionamento", style: TextStyle(color: Colors.white)),
        backgroundColor: const Color(0xFF2C2C5C),
      ),
      body: Padding(
        padding: const EdgeInsets.all(16.0),
        child: ListView(
          children: [
            _buildTextField(controller: clienteController, label: "Cliente", inputType: TextInputType.text),
            _buildTextField(controller: localController, label: "Local", inputType: TextInputType.text),
            _buildExecutionTeamCard(),
            _buildTestDataCard(),
            ...invNumbers.map((inv) => _buildInvCard(inv)).toList(),
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

          for (int i = 0; i < 4; i++) {
            controllers["Nome ${i + 1}"] = nomeControllers[i];
            controllers["Função ${i + 1}"] = funcaoControllers[i];
            controllers["Empresa ${i + 1}"] = empresaControllers[i];
          }

          excelService.exportData(context, controllers);
        },
        child: const Icon(Icons.save),
      ),
    );
  }

  Widget _buildExecutionTeamCard() {
    return Card(
      margin: const EdgeInsets.symmetric(vertical: 8.0),
      child: ExpansionTile(
        title: const Text("Equipe de Execução dos Ensaios e Verificações", style: TextStyle(fontWeight: FontWeight.bold)),
        children: List.generate(4, (i) => Column(
          children: [
            _buildTextField(controller: nomeControllers[i], label: "Nome ${i + 1}"),
            _buildTextField(controller: funcaoControllers[i], label: "Função ${i + 1}"),
            _buildTextField(controller: empresaControllers[i], label: "Empresa ${i + 1}"),
          ],
        )),
      ),
    );
  }

  Widget _buildTestDataCard() {
    return Card(
      margin: const EdgeInsets.symmetric(vertical: 8.0),
      child: ExpansionTile(
        title: const Text("Dados dos Ensaios (Obrigatório)", style: TextStyle(fontWeight: FontWeight.bold, color: Colors.red)),
        children: [
          _buildTextField(controller: dataInicioController, label: "Data de Início", inputType: TextInputType.datetime),
          _buildTextField(controller: horaInicioController, label: "Hora de Início", inputType: TextInputType.datetime),
          _buildTextField(controller: dataTerminoController, label: "Data de Término", inputType: TextInputType.datetime),
          _buildTextField(controller: horaTerminoController, label: "Hora de Término", inputType: TextInputType.datetime),
        ],
      ),
    );
  }


  Widget _buildInvCard(String inv) {
    return Card(
      margin: const EdgeInsets.symmetric(vertical: 8.0),
      child: ExpansionTile(
        title: Text(inv, style: const TextStyle(fontWeight: FontWeight.bold)),
        children: phases.map((phase) => _buildPhaseFields(inv, phase)).toList(),
      ),
    );
  }

  Widget _buildPhaseFields(String inv, String phase) {
    return Padding(
      padding: const EdgeInsets.symmetric(horizontal: 16.0, vertical: 8.0),
      child: Column(
        crossAxisAlignment: CrossAxisAlignment.start,
        children: [
          Text(phase, style: const TextStyle(fontWeight: FontWeight.bold)),
          _buildTextField(controller: controllers['$inv $phase']!, label: "Valor"),
          _buildTextField(controller: controllers['$inv $phase Temp. Amb. (°C)']!, label: "Temp. Amb. (°C)"),
          _buildTextField(controller: controllers['$inv $phase UMID. REL (%)']!, label: "UMID. REL (%)"),
        ],
      ),
    );
  }

  Widget _buildTextField({required TextEditingController controller, required String label, TextInputType inputType = TextInputType.text}) {
    return Padding(
      padding: const EdgeInsets.symmetric(vertical: 4.0),
      child: TextField(
        controller: controller,
        decoration: InputDecoration(labelText: label, border: OutlineInputBorder()),
        keyboardType: inputType,
      ),
    );
  }
}



/*import 'package:flutter/material.dart';
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
*/