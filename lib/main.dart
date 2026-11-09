import 'dart:convert';
import 'package:flutter/material.dart';
import 'package:flutter_localizations/flutter_localizations.dart';
import 'package:http/http.dart' as http;
import 'package:shared_preferences/shared_preferences.dart';

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
  String _serverUrl = 'https://mounixschool.onrender.com';
  String? _authToken;
  Map<String, dynamic>? _userAccount;

  @override
  void initState() {
    super.initState();
    _loadState();
  }

  Future<void> _loadState() async {
    final prefs = await SharedPreferences.getInstance();
    setState(() {
      _locale = Locale(prefs.getString('mounix_lang') ?? 'ar');
      _serverUrl = prefs.getString('mounix_server') ?? 'https://mounixschool.onrender.com';
      _authToken = prefs.getString('mounix_jwt');
    });
    if (_authToken != null) {
      _verifyToken();
    }
  }

  Future<void> _verifyToken() async {
    try {
      final res = await http.get(
        Uri.parse('$_serverUrl/me'),
        headers: {'Authorization': 'Bearer $_authToken'},
      );
      if (res.statusCode == 200) {
        final data = jsonDecode(utf8.decode(res.bodyBytes));
        setState(() => _userAccount = data['user'] as Map<String, dynamic>);
      } else {
        final prefs = await SharedPreferences.getInstance();
        await prefs.remove('mounix_jwt');
        setState(() {
          _authToken = null;
          _userAccount = null;
        });
      }
    } catch (_) {}
  }

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
      home: _userAccount == null
          ? MounixLoginScreen(
              initialServerUrl: _serverUrl,
              onLoginSuccess: (server, token, user) async {
                final prefs = await SharedPreferences.getInstance();
                await prefs.setString('mounix_server', server);
                await prefs.setString('mounix_jwt', token);
                setState(() {
                  _serverUrl = server;
                  _authToken = token;
                  _userAccount = user;
                });
              },
            )
          : MounixWorkspace(
              serverUrl: _serverUrl,
              token: _authToken!,
              user: _userAccount!,
              onToggleLanguage: () => setState(() =>
                  _locale = Locale(_locale.languageCode == 'ar' ? 'en' : 'ar')),
              onToggleTheme: () => setState(() =>
                  _theme = _theme == ThemeMode.dark ? ThemeMode.light : ThemeMode.dark),
              onLogout: () async {
                final prefs = await SharedPreferences.getInstance();
                await prefs.remove('mounix_jwt');
                setState(() {
                  _authToken = null;
                  _userAccount = null;
                });
              },
            ),
    );
  }
}

class MounixLoginScreen extends StatefulWidget {
  final String initialServerUrl;
  final void Function(String, String, Map<String, dynamic>) onLoginSuccess;
  const MounixLoginScreen({
    super.key,
    required this.initialServerUrl,
    required this.onLoginSuccess,
  });

  @override
  State<MounixLoginScreen> createState() => _MounixLoginScreenState();
}

class _MounixLoginScreenState extends State<MounixLoginScreen> {
  late final _serverCtrl = TextEditingController(text: widget.initialServerUrl);
  final _userCtrl = TextEditingController(text: 'admin');
  final _passCtrl = TextEditingController();
  bool _loading = false;
  String? _errorMessage;

  Future<void> _performLogin() async {
    setState(() {
      _loading = true;
      _errorMessage = null;
    });
    try {
      final srv = _serverCtrl.text.trim().replaceAll(RegExp(r'/$'), '');
      final res = await http.post(
        Uri.parse('$srv/auth/login'),
        headers: {'Content-Type': 'application/json'},
        body: jsonEncode({
          'username': _userCtrl.text.trim(),
          'password': _passCtrl.text,
        }),
      );
      final data = jsonDecode(utf8.decode(res.bodyBytes));
      if (res.statusCode == 200) {
        widget.onLoginSuccess(
          srv,
          data['access_token'].toString(),
          data['user'] as Map<String, dynamic>,
        );
      } else {
        setState(() => _errorMessage =
            data['error']?['message']?.toString() ?? data['detail']?.toString() ?? 'خطأ في المصادقة');
      }
    } catch (e) {
      setState(() => _errorMessage = 'تعذر الاتصال بالخادم: $e');
    } finally {
      if (mounted) setState(() => _loading = false);
    }
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      body: Center(
        child: SingleChildScrollView(
          padding: const EdgeInsets.all(24),
          child: ConstrainedBox(
            constraints: const BoxConstraints(maxWidth: 440),
            child: Card(
              elevation: 4,
              shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(16)),
              child: Padding(
                padding: const EdgeInsets.all(28),
                child: Column(
                  mainAxisSize: MainAxisSize.min,
                  children: [
                    const Icon(Icons.school_rounded, size: 68, color: Color(0xFF1E3A8A)),
                    const SizedBox(height: 12),
                    const Text(
                      'MOUNIX SCHOOL',
                      style: TextStyle(fontSize: 24, fontWeight: FontWeight.bold),
                    ),
                    const SizedBox(height: 4),
                    const Text(
                      'منظومة إدارة المدارس ومراكز الدعم',
                      style: TextStyle(color: Colors.grey, fontSize: 13),
                    ),
                    const SizedBox(height: 24),
                    TextField(
                      controller: _serverCtrl,
                      decoration: const InputDecoration(
                        labelText: 'عنوان خادم الـ API (Server URL)',
                        prefixIcon: Icon(Icons.dns_rounded),
                        border: OutlineInputBorder(),
                      ),
                    ),
                    const SizedBox(height: 14),
                    TextField(
                      controller: _userCtrl,
                      decoration: const InputDecoration(
                        labelText: 'اسم المستخدم (Username)',
                        prefixIcon: Icon(Icons.person_rounded),
                        border: OutlineInputBorder(),
                      ),
                    ),
                    const SizedBox(height: 14),
                    TextField(
                      controller: _passCtrl,
                      obscureText: true,
                      decoration: const InputDecoration(
                        labelText: 'كلمة المرور (Password)',
                        prefixIcon: Icon(Icons.lock_rounded),
                        border: OutlineInputBorder(),
                      ),
                      onSubmitted: (_) => _performLogin(),
                    ),
                    if (_errorMessage != null) ...[
                      const SizedBox(height: 14),
                      Text(
                        _errorMessage!,
                        style: const TextStyle(color: Colors.red, fontSize: 13),
                        textAlign: TextAlign.center,
                      ),
                    ],
                    const SizedBox(height: 22),
                    SizedBox(
                      width: double.infinity,
                      height: 48,
                      child: FilledButton.icon(
                        onPressed: _loading ? null : _performLogin,
                        icon: _loading
                            ? const SizedBox(
                                width: 18,
                                height: 18,
                                child: CircularProgressIndicator(strokeWidth: 2, color: Colors.white),
                              )
                            : const Icon(Icons.login_rounded),
                        label: const Text('تسجيل الدخول', style: TextStyle(fontSize: 16)),
                      ),
                    ),
                  ],
                ),
              ),
            ),
          ),
        ),
      ),
    );
  }
}

class MounixWorkspace extends StatefulWidget {
  final String serverUrl;
  final String token;
  final Map<String, dynamic> user;
  final VoidCallback onToggleLanguage;
  final VoidCallback onToggleTheme;
  final VoidCallback onLogout;
  const MounixWorkspace({
    super.key,
    required this.serverUrl,
    required this.token,
    required this.user,
    required this.onToggleLanguage,
    required this.onToggleTheme,
    required this.onLogout,
  });

  @override
  State<MounixWorkspace> createState() => _MounixWorkspaceState();
}

class _MounixWorkspaceState extends State<MounixWorkspace> {
  int _activeNavIndex = 0;

  static const List<(String, IconData, String)> _endpoints = [
    ('لوحة التحكم / Dashboard', Icons.dashboard_rounded, '/dashboard'),
    ('التلاميذ / Students', Icons.people_alt_rounded, '/students'),
    ('الأساتذة / Teachers', Icons.record_voice_over_rounded, '/teachers'),
    ('أولياء الأمور / Parents', Icons.family_restroom_rounded, '/parents'),
    ('المستويات / Levels', Icons.layers_rounded, '/levels'),
    ('المواد / Subjects', Icons.menu_book_rounded, '/subjects'),
    ('الأقسام / Classes', Icons.meeting_room_rounded, '/classes'),
    ('الجداول / Schedules', Icons.calendar_month_rounded, '/schedules'),
    ('الحضور / Attendance', Icons.fact_check_rounded, '/attendance'),
    ('الاختبارات / Exams', Icons.quiz_rounded, '/exams'),
    ('النتائج / Grades', Icons.grade_rounded, '/grades'),
    ('المدفوعات / Payments', Icons.payments_rounded, '/payments'),
    ('الملفات والواجبات / Materials', Icons.folder_copy_rounded, '/materials'),
    ('الإشعارات / Notifications', Icons.notifications_rounded, '/notifications'),
    ('النسخ الاحتياطي / Backups', Icons.storage_rounded, '/admin/backups'),
  ];

  @override
  Widget build(BuildContext context) {
    final active = _endpoints[_activeNavIndex];
    return Scaffold(
      appBar: AppBar(
        title: Text('MOUNIX SCHOOL — ${active.$1}'),
        actions: [
          IconButton(
            icon: const Icon(Icons.translate_rounded),
            onPressed: widget.onToggleLanguage,
            tooltip: 'تغيير اللغة',
          ),
          IconButton(
            icon: const Icon(Icons.brightness_6_rounded),
            onPressed: widget.onToggleTheme,
            tooltip: 'تغيير المظهر',
          ),
          IconButton(
            icon: const Icon(Icons.logout_rounded),
            onPressed: widget.onLogout,
            tooltip: 'تسجيل الخروج',
          ),
        ],
      ),
      drawer: Drawer(
        child: ListView(
          children: [
            UserAccountsDrawerHeader(
              decoration: const BoxDecoration(color: Color(0xFF1E3A8A)),
              accountName: Text('${widget.user['full_name']}'),
              accountEmail: Text('${widget.user['username']} (${widget.user['role']})'),
              currentAccountPicture: const CircleAvatar(
                backgroundColor: Colors.white,
                child: Icon(Icons.school_rounded, color: Color(0xFF1E3A8A), size: 36),
              ),
            ),
            for (int i = 0; i < _endpoints.length; i++)
              ListTile(
                selected: i == _activeNavIndex,
                leading: Icon(_endpoints[i].$2),
                title: Text(_endpoints[i].$1),
                onTap: () {
                  setState(() => _activeNavIndex = i);
                  Navigator.pop(context);
                },
              ),
          ],
        ),
      ),
      body: MounixResourceView(
        key: ValueKey(active.$3),
        serverUrl: widget.serverUrl,
        token: widget.token,
        endpointPath: active.$3,
        sectionTitle: active.$1,
      ),
    );
  }
}

class MounixResourceView extends StatefulWidget {
  final String serverUrl;
  final String token;
  final String endpointPath;
  final String sectionTitle;
  const MounixResourceView({
    super.key,
    required this.serverUrl,
    required this.token,
    required this.endpointPath,
    required this.sectionTitle,
  });

  @override
  State<MounixResourceView> createState() => _MounixResourceViewState();
}

class _MounixResourceViewState extends State<MounixResourceView> {
  bool _loading = true;
  String? _fetchError;
  Map<String, dynamic>? _responseMap;

  @override
  void initState() {
    super.initState();
    _fetch();
  }

  Future<void> _fetch() async {
    setState(() {
      _loading = true;
      _fetchError = null;
    });
    try {
      final res = await http.get(
        Uri.parse('${widget.serverUrl}${widget.endpointPath}'),
        headers: {'Authorization': 'Bearer ${widget.token}'},
      );
      if (res.statusCode == 200) {
        setState(() => _responseMap = jsonDecode(utf8.decode(res.bodyBytes)));
      } else {
        setState(() => _fetchError = 'خطأ HTTP ${res.statusCode}');
      }
    } catch (e) {
      setState(() => _fetchError = 'تعذر الاتصال بالخادم: $e');
    } finally {
      if (mounted) setState(() => _loading = false);
    }
  }

  @override
  Widget build(BuildContext context) {
    if (_loading) return const Center(child: CircularProgressIndicator());
    if (_fetchError != null) {
      return Center(
        child: Padding(
          padding: const EdgeInsets.all(20),
          child: Column(
            mainAxisSize: MainAxisSize.min,
            children: [
              Text(_fetchError!, style: const TextStyle(color: Colors.red), textAlign: TextAlign.center),
              const SizedBox(height: 12),
              FilledButton.icon(onPressed: _fetch, icon: const Icon(Icons.refresh), label: const Text('إعادة المحاولة')),
            ],
          ),
        ),
      );
    }
    if (widget.endpointPath == '/dashboard') {
      final stats = (_responseMap?['stats'] as Map<String, dynamic>?) ?? {};
      return ListView(
        padding: const EdgeInsets.all(16),
        children: stats.entries
            .map((entry) => Card(
                  margin: const EdgeInsets.only(bottom: 12),
                  child: ListTile(
                    title: Text(entry.key, style: const TextStyle(fontSize: 15)),
                    trailing: Text('${entry.value}',
                        style: const TextStyle(fontSize: 18, fontWeight: FontWeight.bold, color: Color(0xFF1E3A8A))),
                  ),
                ))
            .toList(),
      );
    }
    final items = (_responseMap?['items'] as List?) ?? [];
    if (items.isEmpty) {
      return Center(
        child: Column(
          mainAxisSize: MainAxisSize.min,
          children: [
            const Icon(Icons.inbox_rounded, size: 54, color: Colors.grey),
            const SizedBox(height: 10),
            const Text('لا توجد سجلات حالياً في هذا القسم'),
            const SizedBox(height: 12),
            FilledButton.tonalIcon(onPressed: _fetch, icon: const Icon(Icons.refresh), label: const Text('تحديث')),
          ],
        ),
      );
    }
    return RefreshIndicator(
      onRefresh: _fetch,
      child: ListView.builder(
        padding: const EdgeInsets.all(12),
        itemCount: items.length,
        itemBuilder: (ctx, i) {
          final item = items[i] as Map<String, dynamic>;
          final header = item['full_name'] ??
              item['name_ar'] ??
              item['name'] ??
              item['title'] ??
              item['title_ar'] ??
              item['receipt_number'] ??
              item['filename'] ??
              '#${item['id']}';
          final subheader = item['student_code'] ??
              item['teacher_code'] ??
              item['parent_code'] ??
              item['code'] ??
              item['status'] ??
              item['payment_date'] ??
              '';
          return Card(
            margin: const EdgeInsets.only(bottom: 8),
            child: ListTile(
              leading: CircleAvatar(child: Text('${i + 1}')),
              title: Text('$header', style: const TextStyle(fontWeight: FontWeight.bold)),
              subtitle: subheader.isNotEmpty ? Text('$subheader') : null,
            ),
          );
        },
      ),
    );
  }
}
