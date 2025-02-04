import 'package:flutter/material.dart';

import '../data_export_screen.dart';

class InitialScreen extends StatelessWidget {
  const InitialScreen({super.key});

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(title: const Text("Ato. Comissionamento",
        style: TextStyle(
            color: Colors.white
        ),),
          backgroundColor: Color(0xFF2C2C5C)),
      body: Container(
        decoration: BoxDecoration(
          image: DecorationImage(image:
            AssetImage('assets/images/backgroundato.png'),
            fit: BoxFit.cover
          )
        ),
        child: Center(
          child: Column(
            mainAxisAlignment: MainAxisAlignment.center,
            crossAxisAlignment: CrossAxisAlignment.center,
            children: [
              ElevatedButton(
                onPressed: () {
                  Navigator.push(
                    context,
                    MaterialPageRoute(builder: (context) => DataExportScreen()),);
                },
                child: Text("Megohmetro CA"),
                  style: ElevatedButton.styleFrom(
                    backgroundColor: Color(0xFF2C2C5C),
                    foregroundColor: Colors.white,
                    shape: RoundedRectangleBorder(
                      borderRadius: BorderRadius.circular(30.0),
                    ),
                    elevation: 10,
                    padding: EdgeInsets.symmetric(horizontal: 40, vertical: 15),
                    shadowColor: Color(0xFF2C2C5C).withOpacity(0.5),
                ),
              ),
            ],
          ),
        ),
      ),
    );
  }
}
