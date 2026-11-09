import 'package:flutter/material.dart';

class MounixPortalRoot extends StatefulWidget {
  final VoidCallback onToggleLang;
  final VoidCallback onToggleTheme;
  const MounixPortalRoot({
    super.key,
    required this.onToggleLang,
    required this.onToggleTheme,
  });

  @override
  State<MounixPortalRoot> createState() => _MounixPortalRootState();
}

class _MounixPortalRootState extends State<MounixPortalRoot> {
  int _selectedIndex = 0;

  static const List<(String, IconData)> _sections = [
    ('لوحة التحكم / Dashboard', Icons.dashboard_rounded),
    ('التلاميذ / Students', Icons.people_alt_rounded),
    ('ملف التلميذ / Student Dossier', Icons.badge_rounded),
    ('الأساتذة / Teachers', Icons.record_voice_over_rounded),
    ('أولياء الأمور / Parents', Icons.family_restroom_rounded),
    ('المستويات / Levels', Icons.layers_rounded),
    ('المواد / Subjects', Icons.menu_book_rounded),
    ('الأقسام / Classes', Icons.meeting_room_rounded),
    ('الجداول / Schedules', Icons.calendar_month_rounded),
    ('الحضور والغياب / Attendance', Icons.fact_check_rounded),
    ('الاختبارات / Exams', Icons.quiz_rounded),
    ('النتائج والمعدلات / Grades', Icons.grade_rounded),
    ('الإدارة المالية / Payments', Icons.payments_rounded),
    ('الملفات والواجبات / Materials', Icons.folder_copy_rounded),
    ('الإشعارات / Notifications', Icons.notifications_rounded),
    ('التقارير / Reports', Icons.analytics_rounded),
    ('الملف الشخصي / Profile', Icons.person_rounded),
    ('الإعدادات والنسخ الاحتياطي / Settings', Icons.settings_rounded),
  ];

  @override
  Widget build(BuildContext context) {
    final current = _sections[_selectedIndex];
    return Scaffold(
      appBar: AppBar(
        title: Text('MOUNIX SCHOOL — ${current.$1}'),
        actions: [
          IconButton(
            icon: const Icon(Icons.translate_rounded),
            onPressed: widget.onToggleLang,
          ),
          IconButton(
            icon: const Icon(Icons.brightness_6_rounded),
            onPressed: widget.onToggleTheme,
          ),
        ],
      ),
      drawer: Drawer(
        child: ListView(
          children: [
            const DrawerHeader(
              decoration: BoxDecoration(color: Color(0xFF1E3A8A)),
              child: Column(
                crossAxisAlignment: CrossAxisAlignment.start,
                mainAxisAlignment: MainAxisAlignment.end,
                children: [
                  Icon(Icons.school_rounded, size: 42, color: Colors.white),
                  SizedBox(height: 8),
                  Text(
                    'MOUNIX SCHOOL',
                    style: TextStyle(
                      color: Colors.white,
                      fontSize: 20,
                      fontWeight: FontWeight.bold,
                    ),
                  ),
                ],
              ),
            ),
            for (int i = 0; i < _sections.length; i++)
              ListTile(
                selected: i == _selectedIndex,
                leading: Icon(_sections[i].$2),
                title: Text(_sections[i].$1),
                onTap: () {
                  setState(() => _selectedIndex = i);
                  Navigator.pop(context);
                },
              ),
          ],
        ),
      ),
      body: Center(
        child: Card(
          margin: const EdgeInsets.all(24),
          child: Padding(
            padding: const EdgeInsets.all(24),
            child: Column(
              mainAxisSize: MainAxisSize.min,
              children: [
                Icon(current.$2, size: 56, color: const Color(0xFF1E3A8A)),
                const SizedBox(height: 16),
                Text(
                  current.$1,
                  style: Theme.of(context).textTheme.headlineSmall,
                  textAlign: TextAlign.center,
                ),
                const SizedBox(height: 8),
                const Text(
                  'MOUNIX SCHOOL — Connected to Unified FastAPI & SQLite Ecosystem',
                  textAlign: TextAlign.center,
                ),
              ],
            ),
          ),
        ),
      ),
    );
  }
}
