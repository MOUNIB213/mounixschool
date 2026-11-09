import 'package:flutter/material.dart';
import 'package:flutter_localizations/flutter_localizations.dart';
import 'app_core.dart';

void main() {
  WidgetsFlutterBinding.ensureInitialized();
  runApp(const MounixSchoolApp());
}

class MounixSchoolApp extends StatefulWidget {
  const MounixSchoolApp({super.key});
  @override
  State<MounixSchoolApp> createState() => _MounixSchoolAppState();
}

class _MounixSchoolAppState extends State<MounixSchoolApp> {
  Locale _locale = const Locale('ar');
  ThemeMode _theme = ThemeMode.system;

  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      title: 'MOUNIX SCHOOL',
      debugShowCheckedModeBanner: false,
      locale: _locale,
      supportedLocales: const [Locale('ar'), Locale('en')],
      localizationsDelegates: const [
        GlobalMaterialLocalizations.delegate,
        GlobalWidgetsLocalizations.delegate,
        GlobalCupertinoLocalizations.delegate,
      ],
      theme: ThemeData(
        useMaterial3: true,
        colorSchemeSeed: const Color(0xFF1E3A8A),
        brightness: Brightness.light,
      ),
      darkTheme: ThemeData(
        useMaterial3: true,
        colorSchemeSeed: const Color(0xFF3B82F6),
        brightness: Brightness.dark,
      ),
      themeMode: _theme,
      home: MounixPortalRoot(
        onToggleLang: () => setState(() =>
            _locale = Locale(_locale.languageCode == 'ar' ? 'en' : 'ar')),
        onToggleTheme: () => setState(() =>
            _theme = _theme == ThemeMode.dark ? ThemeMode.light : ThemeMode.dark),
      ),
    );
  }
}
