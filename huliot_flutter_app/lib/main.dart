import 'dart:convert';
import 'package:flutter/material.dart';
import 'package:http/http.dart' as http;
import 'web_download.dart';

void main() {
  runApp(const HuliotApp());
}

class HuliotApp extends StatelessWidget {
  const HuliotApp({super.key});

  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      title: 'VIEW SUMMARY STUDIO',
      debugShowCheckedModeBanner: false,
      theme: ThemeData(
        useMaterial3: true,
        colorScheme: ColorScheme.fromSeed(
          seedColor: const Color(0xFF0F4C81), // Classic Classic Blue
          primary: const Color(0xFF0F4C81),
          secondary: const Color(0xFF00897B), // Teal
          surface: Colors.white,
        ),
        fontFamily: 'Segoe UI',
      ),
      home: const MainStudioScreen(),
    );
  }
}

class MainStudioScreen extends StatefulWidget {
  const MainStudioScreen({super.key});

  @override
  State<MainStudioScreen> createState() => _MainStudioScreenState();
}

class _MainStudioScreenState extends State<MainStudioScreen> {
  final String apiUrl = '${Uri.base.origin}/api';

  int selectedMode = 0; // 0: Generate, 1: Train, 2: Admin Dashboard
  bool isBackendOnline = false;

  // Auth State ('user' or 'admin', null = Login Screen)
  String? loggedInRole;
  String adminPassword = '';
  bool isAdminLoggingIn = false;
  bool isLoadingDashboard = false;
  Map<String, dynamic>? adminDashboardData;
  final TextEditingController adminPassCtrl = TextEditingController();

  // Generate Mode State
  WebFile? rawFile;
  bool isAnalyzing = false;
  bool isGenerating = false;
  bool isAutoAgentRunning = false;
  bool isBatchGenerating = false;
  bool includeRainwater = false;
  Map<String, dynamic>? analysisResult;
  Map<String, dynamic>? agentLog;

  // Controllers for parameters
  final TextEditingController projectNameCtrl = TextEditingController();
  final TextEditingController toiletsCtrl = TextEditingController(text: '500');
  final TextEditingController kitchensCtrl = TextEditingController(text: '250');
  final TextEditingController discPipeCtrl = TextEditingController(text: '35');
  final TextEditingController discFitCtrl = TextEditingController(text: '35');
  final TextEditingController discMftCtrl = TextEditingController(text: '35');
  final TextEditingController discLubCtrl = TextEditingController(text: '35');
  final TextEditingController discClampCtrl = TextEditingController(text: '35');

  // Train Mode State
  WebFile? trainFile;
  bool isTraining = false;
  Map<String, dynamic>? trainingResult;

  @override
  void initState() {
    super.initState();
    checkBackendStatus();
  }

  Future<void> checkBackendStatus() async {
    try {
      final res = await http.get(Uri.parse('$apiUrl/status')).timeout(const Duration(seconds: 4));
      if (res.statusCode == 200) {
        setState(() => isBackendOnline = true);
      }
    } catch (_) {
      setState(() => isBackendOnline = false);
    }
  }

  // --- AUTH & ADMIN FLOW ---
  void loginAsUser() {
    setState(() {
      loggedInRole = 'user';
      adminPassword = '';
      selectedMode = 0;
    });
  }

  Future<void> loginAsAdmin() async {
    final pwd = adminPassCtrl.text.trim();
    if (pwd.isEmpty) {
      showSnack('कृपया Admin Password दर्ज करें', isError: true);
      return;
    }
    setState(() => isAdminLoggingIn = true);
    try {
      final res = await http.post(
        Uri.parse('$apiUrl/admin/login'),
        body: {'password': pwd},
      );
      if (res.statusCode == 200) {
        setState(() {
          loggedInRole = 'admin';
          adminPassword = pwd;
          selectedMode = 2;
          adminPassCtrl.clear();
        });
        await fetchAdminDashboard();
      } else {
        showSnack('❌ गलत Admin Password', isError: true);
      }
    } catch (e) {
      showSnack('Admin Login Error: $e', isError: true);
    } finally {
      setState(() => isAdminLoggingIn = false);
    }
  }

  void logout() {
    setState(() {
      loggedInRole = null;
      adminPassword = '';
      adminDashboardData = null;
      selectedMode = 0;
    });
  }

  Future<void> fetchAdminDashboard() async {
    if (loggedInRole != 'admin') return;
    setState(() => isLoadingDashboard = true);
    try {
      final res = await http.get(
        Uri.parse('$apiUrl/admin/dashboard'),
        headers: {'x-admin-password': adminPassword},
      );
      if (res.statusCode == 200) {
        setState(() {
          adminDashboardData = jsonDecode(res.body) as Map<String, dynamic>;
        });
      }
    } catch (e) {
      showSnack('Dashboard Error: $e', isError: true);
    } finally {
      setState(() => isLoadingDashboard = false);
    }
  }

  Future<void> downloadKnowledgeBackup() async {
    if (loggedInRole != 'admin') return;
    try {
      final res = await http.get(
        Uri.parse('$apiUrl/admin/download-knowledge'),
        headers: {'x-admin-password': adminPassword},
      );
      if (res.statusCode == 200) {
        downloadBytes(res.bodyBytes, 'knowledge_backup.json', mimeType: 'application/json');
        showSnack('✅ Knowledge Backup Downloaded!');
      } else {
        showSnack('Backup Error: ${res.body}', isError: true);
      }
    } catch (e) {
      showSnack('Backup Error: $e', isError: true);
    }
  }

  // --- GENERATE FLOW ---
  void resetGenerateState() {
    setState(() {
      rawFile = null;
      analysisResult = null;
      agentLog = null;
      isAnalyzing = false;
      isGenerating = false;
      isAutoAgentRunning = false;
      includeRainwater = false;
      projectNameCtrl.clear();
      toiletsCtrl.clear();
      kitchensCtrl.clear();
      discPipeCtrl.text = '35';
      discFitCtrl.text = '35';
      discMftCtrl.text = '35';
      discLubCtrl.text = '35';
      discClampCtrl.text = '35';
    });
  }

  void resetTrainState() {
    setState(() {
      trainFile = null;
      trainingResult = null;
      isTraining = false;
    });
  }

  Future<void> selectRawFile() async {
    final picked = await pickExcelFile();
    if (picked != null) {
      setState(() {
        rawFile = picked;
        analysisResult = null;
        agentLog = null;
      });
      await analyzeSelectedFile();
    }
  }

  Future<void> analyzeSelectedFile() async {
    if (rawFile == null) return;
    setState(() => isAnalyzing = true);

    try {
      final uri = Uri.parse('$apiUrl/analyze');
      final request = http.MultipartRequest('POST', uri);
      request.files.add(http.MultipartFile.fromBytes(
        'file',
        rawFile!.bytes,
        filename: rawFile!.name,
      ));

      final streamedRes = await request.send();
      final res = await http.Response.fromStream(streamedRes);

      if (res.statusCode == 200) {
        final data = jsonDecode(res.body);
        setState(() {
          analysisResult = data;
          includeRainwater = (data['has_rainwater'] == true);
          projectNameCtrl.text = data['project_name'] ?? 'Huliot Project';
          toiletsCtrl.text = (data['estimated_toilets'] ?? 500).toString();
          kitchensCtrl.text = (data['estimated_kitchens'] ?? 250).toString();
          final disc = ((data['baseline_discount'] ?? 0.35) * 100).round().toString();
          final learned = data['learned_discounts'] as Map<String, dynamic>?;
          discPipeCtrl.text = learned != null && learned['P'] != null
              ? ((learned['P'] as num) * 100).round().toString()
              : disc;
          discFitCtrl.text = learned != null && learned['F'] != null
              ? ((learned['F'] as num) * 100).round().toString()
              : disc;
          discMftCtrl.text = learned != null && learned['M'] != null
              ? ((learned['M'] as num) * 100).round().toString()
              : disc;
          discLubCtrl.text = learned != null && learned['L'] != null
              ? ((learned['L'] as num) * 100).round().toString()
              : disc;
          discClampCtrl.text = learned != null && learned['C'] != null
              ? ((learned['C'] as num) * 100).round().toString()
              : disc;
        });
      } else {
        showSnack('Analysis Error: ${res.body}', isError: true);
      }
    } catch (e) {
      showSnack('Server connection error: $e', isError: true);
    } finally {
      setState(() => isAnalyzing = false);
    }
  }

  Future<void> generateSummarySheet() async {
    if (rawFile == null) return;
    setState(() => isGenerating = true);

    try {
      final uri = Uri.parse('$apiUrl/generate');
      final request = http.MultipartRequest('POST', uri);
      request.files.add(http.MultipartFile.fromBytes(
        'file',
        rawFile!.bytes,
        filename: rawFile!.name,
      ));

      request.fields['project_name'] = projectNameCtrl.text;
      request.fields['num_toilets'] = toiletsCtrl.text;
      request.fields['num_kitchens'] = kitchensCtrl.text;
      request.fields['discount_p'] = ((double.tryParse(discPipeCtrl.text) ?? 35) / 100).toString();
      request.fields['discount_f'] = ((double.tryParse(discFitCtrl.text) ?? 35) / 100).toString();
      request.fields['discount_m'] = ((double.tryParse(discMftCtrl.text) ?? 35) / 100).toString();
      request.fields['discount_l'] = ((double.tryParse(discLubCtrl.text) ?? 35) / 100).toString();
      request.fields['discount_c'] = ((double.tryParse(discClampCtrl.text) ?? 35) / 100).toString();
      request.fields['include_rainwater'] = includeRainwater ? 'true' : 'false';

      final streamedRes = await request.send();
      final res = await http.Response.fromStream(streamedRes);

      if (res.statusCode == 200) {
        final outName = rawFile!.name.replaceAll('.xlsx', '_with_Summary.xlsx');
        downloadBytes(res.bodyBytes, outName);
        final verified = res.headers['x-summary-verified'] != 'false';
        final warnRaw = res.headers['x-summary-warnings'];
        if (!verified && warnRaw != null && warnRaw != '[]') {
          showSnack('⚠️ Downloaded $outName (Warnings: $warnRaw)', isError: true);
        } else {
          showSnack('✅ Verified (0 Formula Errors) & Downloaded: $outName');
        }
      } else {
        showSnack('Generation Error: ${res.body}', isError: true);
      }
    } catch (e) {
      showSnack('Generation error: $e', isError: true);
    } finally {
      setState(() => isGenerating = false);
    }
  }

  Future<void> runAutoAgent() async {
    if (rawFile == null) return;
    setState(() {
      isAutoAgentRunning = true;
      agentLog = null;
    });

    try {
      final uri = Uri.parse('$apiUrl/auto-agent');
      final request = http.MultipartRequest('POST', uri);
      request.files.add(http.MultipartFile.fromBytes(
        'file',
        rawFile!.bytes,
        filename: rawFile!.name,
      ));

      if (projectNameCtrl.text.trim().isNotEmpty) {
        request.fields['project_name'] = projectNameCtrl.text.trim();
      }
      if (toiletsCtrl.text.trim().isNotEmpty) {
        request.fields['num_toilets'] = toiletsCtrl.text.trim();
      }
      if (kitchensCtrl.text.trim().isNotEmpty) {
        request.fields['num_kitchens'] = kitchensCtrl.text.trim();
      }
      request.fields['discount_p'] = ((double.tryParse(discPipeCtrl.text) ?? 35) / 100).toString();
      request.fields['discount_f'] = ((double.tryParse(discFitCtrl.text) ?? 35) / 100).toString();
      request.fields['discount_m'] = ((double.tryParse(discMftCtrl.text) ?? 35) / 100).toString();
      request.fields['discount_l'] = ((double.tryParse(discLubCtrl.text) ?? 35) / 100).toString();
      request.fields['discount_c'] = ((double.tryParse(discClampCtrl.text) ?? 35) / 100).toString();
      request.fields['include_rainwater'] = includeRainwater ? 'true' : 'false';

      final streamedRes = await request.send();
      final res = await http.Response.fromStream(streamedRes);

      if (res.statusCode == 200) {
        final outName = rawFile!.name.replaceAll('.xlsx', '_with_Summary.xlsx');
        downloadBytes(res.bodyBytes, outName);
        final logHeader = res.headers['x-agent-log'];
        if (logHeader != null && logHeader.isNotEmpty) {
          setState(() {
            agentLog = jsonDecode(logHeader) as Map<String, dynamic>;
          });
        }
        showSnack('🤖 Auto Agent Complete: $outName');
      } else {
        showSnack('Auto Agent Error: ${res.body}', isError: true);
      }
    } catch (e) {
      showSnack('Auto Agent error: $e', isError: true);
    } finally {
      setState(() => isAutoAgentRunning = false);
    }
  }

  Future<void> generateBatchSummaries() async {
    final pickedList = await pickMultipleExcelFiles();
    if (pickedList.isEmpty) return;
    setState(() => isBatchGenerating = true);

    try {
      final uri = Uri.parse('$apiUrl/generate-batch');
      final request = http.MultipartRequest('POST', uri);
      request.fields['include_rainwater'] = includeRainwater ? 'true' : 'false';
      for (final wf in pickedList) {
        request.files.add(http.MultipartFile.fromBytes('files', wf.bytes, filename: wf.name));
      }

      final streamedRes = await request.send();
      final res = await http.Response.fromStream(streamedRes);

      if (res.statusCode == 200) {
        downloadBytes(res.bodyBytes, 'BOQ_Summaries_Batch.zip', mimeType: 'application/zip');
        showSnack('✅ Batch Complete (${pickedList.length} files): BOQ_Summaries_Batch.zip Downloaded!');
      } else {
        showSnack('Batch Error: ${res.body}', isError: true);
      }
    } catch (e) {
      showSnack('Batch error: $e', isError: true);
    } finally {
      setState(() => isBatchGenerating = false);
    }
  }

  // --- TRAIN FLOW ---
  Future<void> selectTrainFile() async {
    final picked = await pickExcelFile();
    if (picked != null) {
      setState(() {
        trainFile = picked;
        trainingResult = null;
      });
      await trainOnSelectedFile();
    }
  }

  Future<void> trainOnSelectedFile() async {
    if (trainFile == null) return;
    setState(() => isTraining = true);

    try {
      final uri = Uri.parse('$apiUrl/train');
      final request = http.MultipartRequest('POST', uri);
      request.files.add(http.MultipartFile.fromBytes(
        'file',
        trainFile!.bytes,
        filename: trainFile!.name,
      ));

      final streamedRes = await request.send();
      final res = await http.Response.fromStream(streamedRes);

      if (res.statusCode == 200) {
        final data = jsonDecode(res.body);
        setState(() {
          trainingResult = data;
        });
        showSnack('🧠 नया फॉर्मेट सफलतापूर्वक सीखा गया और सुरक्षित कर लिया गया!');
      } else {
        showSnack('Training Error: ${res.body}', isError: true);
      }
    } catch (e) {
      showSnack('Error connecting to backend: $e', isError: true);
    } finally {
      setState(() => isTraining = false);
    }
  }

  void showSnack(String msg, {bool isError = false}) {
    ScaffoldMessenger.of(context).showSnackBar(SnackBar(
      content: Text(msg),
      backgroundColor: isError ? const Color(0xFFC62828) : const Color(0xFF2E7D32),
      duration: const Duration(seconds: 4),
      behavior: SnackBarBehavior.floating,
      shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(8)),
    ));
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      backgroundColor: loggedInRole == null ? const Color(0xFF0B1120) : const Color(0xFFF8FAFC),
      appBar: loggedInRole == null
          ? null
          : PreferredSize(
        preferredSize: const Size.fromHeight(68),
        child: Container(
          decoration: const BoxDecoration(
            color: Colors.white,
            border: Border(bottom: BorderSide(color: Color(0xFFE2E8F0))),
          ),
          padding: const EdgeInsets.symmetric(horizontal: 24, vertical: 12),
          child: Row(
            children: [
              Column(
                crossAxisAlignment: CrossAxisAlignment.start,
                mainAxisAlignment: MainAxisAlignment.center,
                children: const [
                  Text(
                    'VIEW SUMMARY STUDIO',
                    style: TextStyle(color: Color(0xFF0F172A), fontWeight: FontWeight.bold, fontSize: 18, letterSpacing: 0.5),
                  ),
                  Text(
                    'Automated Commercial & MEP Estimation Engine',
                    style: TextStyle(color: Color(0xFF64748B), fontSize: 11),
                  ),
                ],
              ),
              const Spacer(),

              // Clean Mode Selector (Segmented)
              if (loggedInRole != null) ...[
                Container(
                  decoration: BoxDecoration(
                    color: const Color(0xFFF1F5F9),
                    borderRadius: BorderRadius.circular(8),
                  ),
                  padding: const EdgeInsets.all(4),
                  child: Row(
                    children: [
                      _buildModeButton(
                        index: 0,
                        icon: Icons.flash_on,
                        label: 'Generate Summary',
                      ),
                      const SizedBox(width: 4),
                      _buildModeButton(
                        index: 1,
                        icon: Icons.school,
                        label: 'Train New Format',
                      ),
                      if (loggedInRole == 'admin') ...[
                        const SizedBox(width: 4),
                        _buildModeButton(
                          index: 2,
                          icon: Icons.admin_panel_settings,
                          label: 'Admin Dashboard',
                        ),
                      ],
                    ],
                  ),
                ),
                const SizedBox(width: 12),
                Container(
                  padding: const EdgeInsets.symmetric(horizontal: 10, vertical: 5),
                  decoration: BoxDecoration(
                    color: loggedInRole == 'admin' ? const Color(0xFFFEF3C7) : const Color(0xFFEFF6FF),
                    borderRadius: BorderRadius.circular(16),
                  ),
                  child: Text(
                    loggedInRole == 'admin' ? 'Admin' : 'User',
                    style: TextStyle(
                      fontSize: 11,
                      fontWeight: FontWeight.bold,
                      color: loggedInRole == 'admin' ? const Color(0xFF92400E) : const Color(0xFF1D4ED8),
                    ),
                  ),
                ),
                const SizedBox(width: 8),
                IconButton(
                  tooltip: 'Logout',
                  onPressed: logout,
                  icon: const Icon(Icons.logout, size: 18, color: Color(0xFF64748B)),
                ),
                const SizedBox(width: 6),
              ],

              // Status Chip
              Container(
                padding: const EdgeInsets.symmetric(horizontal: 10, vertical: 5),
                decoration: BoxDecoration(
                  color: isBackendOnline ? const Color(0xFFE8F5E9) : const Color(0xFFFFEBEE),
                  borderRadius: BorderRadius.circular(16),
                ),
                child: Row(
                  children: [
                    CircleAvatar(radius: 4, backgroundColor: isBackendOnline ? Colors.green[700] : Colors.red[700]),
                    const SizedBox(width: 6),
                    Text(
                      isBackendOnline ? 'System Ready' : 'Connecting...',
                      style: TextStyle(color: isBackendOnline ? Colors.green[900] : Colors.red[900], fontSize: 11, fontWeight: FontWeight.w600),
                    ),
                  ],
                ),
              ),
            ],
          ),
        ),
      ),
      body: loggedInRole == null
          ? _buildLoginView()
          : (selectedMode == 0
              ? _buildGenerateView()
              : (selectedMode == 1
                  ? _buildTrainView()
                  : (loggedInRole == 'admin' ? _buildAdminDashboardView() : _buildGenerateView()))),
    );
  }

  Widget _buildModeButton({required int index, required IconData icon, required String label}) {
    final isSelected = selectedMode == index;
    return InkWell(
      onTap: () {
        setState(() => selectedMode = index);
        if (index == 2 && loggedInRole == 'admin') {
          fetchAdminDashboard();
        }
      },
      borderRadius: BorderRadius.circular(6),
      child: Container(
        padding: const EdgeInsets.symmetric(horizontal: 14, vertical: 6),
        decoration: BoxDecoration(
          color: isSelected ? Colors.white : Colors.transparent,
          borderRadius: BorderRadius.circular(6),
          boxShadow: isSelected
              ? [BoxElevation(color: Colors.black.withOpacity(0.05), blurRadius: 4, offset: const Offset(0, 1))]
              : null,
        ),
        child: Row(
          children: [
            Icon(icon, size: 16, color: isSelected ? const Color(0xFF0F4C81) : const Color(0xFF64748B)),
            const SizedBox(width: 6),
            Text(
              label,
              style: TextStyle(
                color: isSelected ? const Color(0xFF0F4C81) : const Color(0xFF64748B),
                fontSize: 13,
                fontWeight: isSelected ? FontWeight.bold : FontWeight.w500,
              ),
            ),
          ],
        ),
      ),
    );
  }

  // --- VIEW 1: GENERATE (MAIN FOR TEAM) ---
  Widget _buildGenerateView() {
    return SingleChildScrollView(
      padding: const EdgeInsets.symmetric(horizontal: 24, vertical: 28),
      child: Center(
        child: Container(
          constraints: const BoxConstraints(maxWidth: 880),
          child: Column(
            crossAxisAlignment: CrossAxisAlignment.stretch,
            children: [
              // Hero Upload Box
              Container(
                decoration: BoxDecoration(
                  color: Colors.white,
                  borderRadius: BorderRadius.circular(14),
                  border: Border.all(color: const Color(0xFFCBD5E1), width: 1.5),
                  boxShadow: [
                    BoxShadow(color: Colors.black.withOpacity(0.02), blurRadius: 8, offset: const Offset(0, 2)),
                  ],
                ),
                padding: const EdgeInsets.symmetric(horizontal: 24, vertical: 32),
                child: Stack(
                  children: [
                    if (rawFile != null || analysisResult != null)
                      Positioned(
                        top: 0,
                        right: 0,
                        child: OutlinedButton.icon(
                          onPressed: resetGenerateState,
                          icon: const Icon(Icons.refresh, size: 16),
                          label: const Text('Reset', style: TextStyle(fontSize: 12.5, fontWeight: FontWeight.w600)),
                          style: OutlinedButton.styleFrom(
                            foregroundColor: const Color(0xFF475569),
                            side: const BorderSide(color: Color(0xFFCBD5E1)),
                            padding: const EdgeInsets.symmetric(horizontal: 12, vertical: 8),
                            shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(8)),
                          ),
                        ),
                      ),
                    SizedBox(
                      width: double.infinity,
                      child: Column(
                        children: [
                          Container(
                            padding: const EdgeInsets.all(16),
                            decoration: BoxDecoration(
                              color: const Color(0xFFEFF6FF),
                              shape: BoxShape.circle,
                            ),
                            child: const Icon(Icons.file_present, size: 40, color: Color(0xFF0F4C81)),
                          ),
                          const SizedBox(height: 16),
                          Text(
                            rawFile == null ? 'रॉ Excel BOQ फाइल यहाँ अपलोड करें' : rawFile!.name,
                            style: const TextStyle(fontSize: 17, fontWeight: FontWeight.bold, color: Color(0xFF0F172A)),
                          ),
                          const SizedBox(height: 6),
                          const Text(
                            'सिस्टम शीट के अंदरूनी शाफ्ट, टॉयलेट, किचन और डिस्काउंट्स को खुद समझकर लाइव फॉर्मूलों के साथ समरी बनाएगा',
                            textAlign: TextAlign.center,
                            style: TextStyle(color: Color(0xFF64748B), fontSize: 13),
                          ),
                          const SizedBox(height: 18),
                          Wrap(
                            spacing: 12,
                            runSpacing: 10,
                            alignment: WrapAlignment.center,
                            children: [
                              ElevatedButton.icon(
                                onPressed: selectRawFile,
                                icon: const Icon(Icons.upload_file, size: 18),
                                label: Text(rawFile == null ? 'Select Raw BOQ File (.xlsx)' : 'Choose Another File'),
                                style: ElevatedButton.styleFrom(
                                  backgroundColor: const Color(0xFF0F4C81),
                                  foregroundColor: Colors.white,
                                  padding: const EdgeInsets.symmetric(horizontal: 24, vertical: 13),
                                  shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(8)),
                                  elevation: 0,
                                ),
                              ),
                              OutlinedButton.icon(
                                onPressed: isBatchGenerating ? null : generateBatchSummaries,
                                icon: isBatchGenerating
                                    ? const SizedBox(width: 16, height: 16, child: CircularProgressIndicator(strokeWidth: 2))
                                    : const Icon(Icons.folder_zip_outlined, size: 18),
                                label: Text(isBatchGenerating ? 'Processing Batch...' : 'Batch Mode (Multiple .xlsx → ZIP)'),
                                style: OutlinedButton.styleFrom(
                                  foregroundColor: const Color(0xFF0F4C81),
                                  side: const BorderSide(color: Color(0xFF0F4C81)),
                                  padding: const EdgeInsets.symmetric(horizontal: 20, vertical: 13),
                                  shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(8)),
                                ),
                              ),
                            ],
                          ),
                        ],
                      ),
                    ),
                  ],
                ),
              ),

              if (isAnalyzing) ...[
                const SizedBox(height: 28),
                const Center(child: CircularProgressIndicator(strokeWidth: 3)),
                const SizedBox(height: 10),
                const Center(
                  child: Text(
                    'Reading shaft topology, building composition & discount structure...',
                    style: TextStyle(color: Color(0xFF64748B), fontSize: 13),
                  ),
                ),
              ],

              if (analysisResult != null) ...[
                const SizedBox(height: 24),
                // Detailed Parameters Card
                Container(
                  decoration: BoxDecoration(
                    color: Colors.white,
                    borderRadius: BorderRadius.circular(14),
                    border: Border.all(color: const Color(0xFFE2E8F0)),
                    boxShadow: [
                      BoxShadow(color: Colors.black.withOpacity(0.02), blurRadius: 10, offset: const Offset(0, 4)),
                    ],
                  ),
                  padding: const EdgeInsets.all(24),
                  child: Column(
                    crossAxisAlignment: CrossAxisAlignment.start,
                    children: [
                      // Header Row
                      Row(
                        children: [
                          const Icon(Icons.verified, color: Color(0xFF059669), size: 22),
                          const SizedBox(width: 8),
                          const Text(
                            'प्रोजेक्ट पैरामीटर्स (Auto-Detected & Editable)',
                            style: TextStyle(fontSize: 16, fontWeight: FontWeight.bold, color: Color(0xFF0F172A)),
                          ),
                          const Spacer(),
                          if (analysisResult!['knowledge_match'] == true) ...[
                            Container(
                              padding: const EdgeInsets.symmetric(horizontal: 10, vertical: 4),
                              decoration: BoxDecoration(
                                color: const Color(0xFFEFF6FF),
                                border: Border.all(color: const Color(0xFFBFDBFE)),
                                borderRadius: BorderRadius.circular(6),
                              ),
                              child: const Text(
                                'Loaded from Training Data ✓',
                                style: TextStyle(color: Color(0xFF1D4ED8), fontSize: 12, fontWeight: FontWeight.bold),
                              ),
                            ),
                            const SizedBox(width: 8),
                          ],
                          Container(
                            padding: const EdgeInsets.symmetric(horizontal: 10, vertical: 4),
                            decoration: BoxDecoration(
                              color: const Color(0xFFF0FDF4),
                              border: Border.all(color: const Color(0xFFBBF7D0)),
                              borderRadius: BorderRadius.circular(6),
                            ),
                            child: Text(
                              analysisResult!['is_combined_single_sheet'] == true
                                  ? 'Single Combined Sheet'
                                  : (analysisResult!['has_diversion'] == true ? 'Residential + Diversion' : 'Multi-Sheet Standard'),
                              style: const TextStyle(color: Color(0xFF166534), fontSize: 12, fontWeight: FontWeight.bold),
                            ),
                          ),
                        ],
                      ),
                      const SizedBox(height: 12),

                      // Format Confidence Score Banner
                      Builder(builder: (context) {
                        final int conf = (analysisResult!['format_confidence'] as num?)?.toInt() ?? 0;
                        final String fmtLabel = (analysisResult!['detected_format_label'] ?? 'Standard Format').toString();
                        final bool isNew = analysisResult!['is_new_format'] == true || conf < 70;
                        final String? matchedProj = analysisResult!['matched_format_project']?.toString();
                        return Container(
                          width: double.infinity,
                          padding: const EdgeInsets.symmetric(horizontal: 14, vertical: 10),
                          decoration: BoxDecoration(
                            color: isNew ? const Color(0xFFFFF7ED) : const Color(0xFFECFDF5),
                            borderRadius: BorderRadius.circular(8),
                            border: Border.all(
                              color: isNew ? const Color(0xFFFDBA74) : const Color(0xFFA7F3D0),
                            ),
                          ),
                          child: Row(
                            children: [
                              Icon(
                                isNew ? Icons.new_releases_outlined : Icons.verified_user_outlined,
                                size: 18,
                                color: isNew ? const Color(0xFFC2410C) : const Color(0xFF047857),
                              ),
                              const SizedBox(width: 8),
                              Expanded(
                                child: Text(
                                  isNew
                                      ? 'नया Format लगता है (Confidence $conf%) — कृपया Generate से पहले पैरामीटर्स जांच लें'
                                      : 'Detected: $fmtLabel (Confidence $conf%)${matchedProj != null && matchedProj.isNotEmpty ? '  •  Match: $matchedProj' : ''}',
                                  style: TextStyle(
                                    fontSize: 13,
                                    fontWeight: FontWeight.bold,
                                    color: isNew ? const Color(0xFF9A3412) : const Color(0xFF065F46),
                                  ),
                                ),
                              ),
                            ],
                          ),
                        );
                      }),
                      const SizedBox(height: 18),

                      // Project Name & Quantities
                      Row(
                        children: [
                          Expanded(
                            flex: 2,
                            child: TextField(
                              controller: projectNameCtrl,
                              decoration: InputDecoration(
                                labelText: 'Project Name',
                                prefixIcon: const Icon(Icons.business, size: 18),
                                border: OutlineInputBorder(borderRadius: BorderRadius.circular(8)),
                                contentPadding: const EdgeInsets.symmetric(horizontal: 12, vertical: 12),
                              ),
                            ),
                          ),
                          const SizedBox(width: 14),
                          Expanded(
                            child: TextField(
                              controller: toiletsCtrl,
                              keyboardType: TextInputType.number,
                              decoration: InputDecoration(
                                labelText: 'No. of Toilets',
                                prefixIcon: const Icon(Icons.bathtub_outlined, size: 18),
                                border: OutlineInputBorder(borderRadius: BorderRadius.circular(8)),
                                contentPadding: const EdgeInsets.symmetric(horizontal: 12, vertical: 12),
                              ),
                            ),
                          ),
                          const SizedBox(width: 14),
                          Expanded(
                            child: TextField(
                              controller: kitchensCtrl,
                              keyboardType: TextInputType.number,
                              decoration: InputDecoration(
                                labelText: 'No. of Kitchens / Flats',
                                prefixIcon: const Icon(Icons.kitchen_outlined, size: 18),
                                border: OutlineInputBorder(borderRadius: BorderRadius.circular(8)),
                                contentPadding: const EdgeInsets.symmetric(horizontal: 12, vertical: 12),
                              ),
                            ),
                          ),
                        ],
                      ),
                      const SizedBox(height: 20),

                      // Discount Structure Row
                      const Text(
                        'Discount Structure (%):',
                        style: TextStyle(fontWeight: FontWeight.bold, fontSize: 13, color: Color(0xFF334155)),
                      ),
                      const SizedBox(height: 8),
                      Row(
                        children: [
                          Expanded(child: _buildDiscountField(controller: discPipeCtrl, label: 'Pipe (P) %')),
                          const SizedBox(width: 10),
                          Expanded(child: _buildDiscountField(controller: discFitCtrl, label: 'Fittings (F) %')),
                          const SizedBox(width: 10),
                          Expanded(child: _buildDiscountField(controller: discMftCtrl, label: 'MFT (M) %')),
                          const SizedBox(width: 10),
                          Expanded(child: _buildDiscountField(controller: discLubCtrl, label: 'Lubricant (L) %')),
                          const SizedBox(width: 10),
                          Expanded(child: _buildDiscountField(controller: discClampCtrl, label: 'Clamps (C) %')),
                        ],
                      ),
                      const SizedBox(height: 18),

                      // Live Detection Preview (Modules, Towers & Warnings)
                      _buildLivePreviewPanel(),
                      if (analysisResult!['has_rainwater'] == true) ...[
                        const SizedBox(height: 12),
                        CheckboxListTile(
                          value: includeRainwater,
                          onChanged: (v) => setState(() => includeRainwater = v ?? false),
                          title: const Text(
                            'Include Rainwater BOQ (RWP) in Summary Sheet',
                            style: TextStyle(fontSize: 13.5, fontWeight: FontWeight.w600, color: Color(0xFF0F172A)),
                          ),
                          controlAffinity: ListTileControlAffinity.leading,
                          contentPadding: EdgeInsets.zero,
                          dense: true,
                        ),
                      ],
                      const SizedBox(height: 18),

                      // Generate Buttons (Manual + Auto Agent)
                      Row(
                        children: [
                          Expanded(
                            child: ElevatedButton.icon(
                              onPressed: (isGenerating || isAutoAgentRunning) ? null : generateSummarySheet,
                              icon: isGenerating
                                  ? const SizedBox(width: 18, height: 18, child: CircularProgressIndicator(color: Colors.white, strokeWidth: 2))
                                  : const Icon(Icons.bolt, size: 20),
                              label: Text(
                                isGenerating ? 'Generating Live Summary Sheet...' : '⚡ Generate Summary Sheet & Download',
                                style: const TextStyle(fontSize: 14.5, fontWeight: FontWeight.bold),
                              ),
                              style: ElevatedButton.styleFrom(
                                backgroundColor: const Color(0xFF00897B),
                                foregroundColor: Colors.white,
                                padding: const EdgeInsets.symmetric(vertical: 16),
                                shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(8)),
                                elevation: 0,
                              ),
                            ),
                          ),
                          const SizedBox(width: 12),
                          Expanded(
                            child: ElevatedButton.icon(
                              onPressed: (isGenerating || isAutoAgentRunning) ? null : runAutoAgent,
                              icon: isAutoAgentRunning
                                  ? const SizedBox(width: 18, height: 18, child: CircularProgressIndicator(color: Colors.white, strokeWidth: 2))
                                  : const Icon(Icons.smart_toy_outlined, size: 20),
                              label: Text(
                                isAutoAgentRunning ? 'Auto Agent Running...' : '🤖 Auto Agent से चलाओ',
                                style: const TextStyle(fontSize: 14.5, fontWeight: FontWeight.bold),
                              ),
                              style: ElevatedButton.styleFrom(
                                backgroundColor: const Color(0xFF4F46E5),
                                foregroundColor: Colors.white,
                                padding: const EdgeInsets.symmetric(vertical: 16),
                                shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(8)),
                                elevation: 0,
                              ),
                            ),
                          ),
                        ],
                      ),

                      if (agentLog != null) ...[
                        const SizedBox(height: 16),
                        Container(
                          width: double.infinity,
                          padding: const EdgeInsets.all(14),
                          decoration: BoxDecoration(
                            color: const Color(0xFFEEF2FF),
                            borderRadius: BorderRadius.circular(10),
                            border: Border.all(color: const Color(0xFFC7D2FE)),
                          ),
                          child: Column(
                            crossAxisAlignment: CrossAxisAlignment.start,
                            children: [
                              Row(
                                children: [
                                  const Icon(Icons.smart_toy, size: 18, color: Color(0xFF4338CA)),
                                  const SizedBox(width: 8),
                                  const Text(
                                    'Agent Log',
                                    style: TextStyle(fontSize: 14, fontWeight: FontWeight.bold, color: Color(0xFF312E81)),
                                  ),
                                  if (agentLog!['total_time_s'] != null) ...[
                                    const SizedBox(width: 8),
                                    Container(
                                      padding: const EdgeInsets.symmetric(horizontal: 7, vertical: 2),
                                      decoration: BoxDecoration(
                                        color: Colors.white,
                                        borderRadius: BorderRadius.circular(6),
                                        border: Border.all(color: const Color(0xFFC7D2FE)),
                                      ),
                                      child: Text(
                                        '⏱ ${agentLog!['total_time_s']}s',
                                        style: const TextStyle(fontSize: 11, fontWeight: FontWeight.bold, color: Color(0xFF4338CA)),
                                      ),
                                    ),
                                  ],
                                  const Spacer(),
                                  Container(
                                    padding: const EdgeInsets.symmetric(horizontal: 8, vertical: 2),
                                    decoration: BoxDecoration(
                                      color: agentLog!['trained'] == true ? const Color(0xFFFEF3C7) : const Color(0xFFDCFCE7),
                                      borderRadius: BorderRadius.circular(6),
                                    ),
                                    child: Text(
                                      agentLog!['trained'] == true ? 'Trained + Generated' : 'Direct Generate',
                                      style: TextStyle(
                                        fontSize: 11,
                                        fontWeight: FontWeight.bold,
                                        color: agentLog!['trained'] == true ? const Color(0xFF92400E) : const Color(0xFF166534),
                                      ),
                                    ),
                                  ),
                                ],
                              ),
                              const SizedBox(height: 8),
                              Text(
                                agentLog!['summary_line']?.toString() ?? '',
                                style: const TextStyle(fontSize: 13.5, fontWeight: FontWeight.bold, color: Color(0xFF1E1B4B)),
                              ),
                              if (agentLog!['steps'] is List) ...[
                                const SizedBox(height: 12),
                                ...List.generate((agentLog!['steps'] as List).length, (idx) {
                                  final steps = agentLog!['steps'] as List;
                                  final st = steps[idx] as Map<String, dynamic>;
                                  final bool isLast = idx == steps.length - 1;
                                  final String status = (st['status'] ?? 'done').toString();
                                  final Color nodeColor = status == 'warning'
                                      ? const Color(0xFFD97706)
                                      : (status == 'trained'
                                          ? const Color(0xFF7E22CE)
                                          : (status == 'skipped' ? const Color(0xFF2563EB) : const Color(0xFF059669)));
                                  return IntrinsicHeight(
                                    child: Row(
                                      crossAxisAlignment: CrossAxisAlignment.start,
                                      children: [
                                        SizedBox(
                                          width: 24,
                                          child: Column(
                                            children: [
                                              Container(
                                                width: 18,
                                                height: 18,
                                                decoration: BoxDecoration(
                                                  color: nodeColor,
                                                  shape: BoxShape.circle,
                                                ),
                                                child: Center(
                                                  child: Text(
                                                    '${idx + 1}',
                                                    style: const TextStyle(color: Colors.white, fontSize: 10, fontWeight: FontWeight.bold),
                                                  ),
                                                ),
                                              ),
                                              if (!isLast)
                                                Expanded(
                                                  child: Container(
                                                    width: 2,
                                                    color: const Color(0xFFA5B4FC),
                                                  ),
                                                ),
                                            ],
                                          ),
                                        ),
                                        const SizedBox(width: 8),
                                        Expanded(
                                          child: Padding(
                                            padding: EdgeInsets.only(bottom: isLast ? 2 : 10),
                                            child: Container(
                                              padding: const EdgeInsets.symmetric(horizontal: 10, vertical: 7),
                                              decoration: BoxDecoration(
                                                color: Colors.white,
                                                borderRadius: BorderRadius.circular(8),
                                                border: Border.all(color: const Color(0xFFE0E7FF)),
                                              ),
                                              child: Row(
                                                children: [
                                                  Expanded(
                                                    child: Column(
                                                      crossAxisAlignment: CrossAxisAlignment.start,
                                                      children: [
                                                        Text(
                                                          st['step']?.toString() ?? '',
                                                          style: const TextStyle(fontSize: 12.5, fontWeight: FontWeight.bold, color: Color(0xFF1E1B4B)),
                                                        ),
                                                        const SizedBox(height: 2),
                                                        Text(
                                                          st['detail']?.toString() ?? '',
                                                          style: const TextStyle(fontSize: 11.5, color: Color(0xFF475569)),
                                                        ),
                                                      ],
                                                    ),
                                                  ),
                                                  Container(
                                                    padding: const EdgeInsets.symmetric(horizontal: 7, vertical: 3),
                                                    decoration: BoxDecoration(
                                                      color: const Color(0xFFF1F5F9),
                                                      borderRadius: BorderRadius.circular(6),
                                                    ),
                                                    child: Text(
                                                      '${st['time_s'] ?? 0}s',
                                                      style: const TextStyle(fontSize: 11, fontWeight: FontWeight.bold, color: Color(0xFF334155)),
                                                    ),
                                                  ),
                                                ],
                                              ),
                                            ),
                                          ),
                                        ),
                                      ],
                                    ),
                                  );
                                }),
                              ],
                              const SizedBox(height: 8),
                              Text('• मैंने क्या देखा: ${agentLog!['saw'] ?? ''}', style: const TextStyle(fontSize: 12, color: Color(0xFF3730A3))),
                              const SizedBox(height: 2),
                              Text('• मैंने क्या Decide किया: ${agentLog!['decided'] ?? ''}', style: const TextStyle(fontSize: 12, color: Color(0xFF3730A3))),
                              const SizedBox(height: 2),
                              Text('• मैंने क्या किया: ${agentLog!['did'] ?? ''}', style: const TextStyle(fontSize: 12, color: Color(0xFF3730A3))),
                            ],
                          ),
                        ),
                      ],
                    ],
                  ),
                ),
              ],
            ],
          ),
        ),
      ),
    );
  }

  Widget _buildLivePreviewPanel() {
    if (analysisResult == null) return const SizedBox.shrink();
    final sheetCls = (analysisResult!['sheet_classification'] as Map<String, dynamic>?) ?? {};
    final int typicalCnt = (sheetCls['typical_count'] as num?)?.toInt() ?? 0;
    final int divCnt = (sheetCls['diversion_count'] as num?)?.toInt() ?? 0;
    final int rainCnt = (sheetCls['rainwater_count'] as num?)?.toInt() ?? 0;
    final List<dynamic> towers = (analysisResult!['towers'] as List<dynamic>?) ?? [];
    final List<dynamic> warnings = (analysisResult!['warnings'] as List<dynamic>?) ?? [];
    final List<dynamic> opts = (analysisResult!['detected_options'] as List<dynamic>?) ?? [];

    return Container(
      padding: const EdgeInsets.all(14),
      decoration: BoxDecoration(
        color: const Color(0xFFF8FAFC),
        borderRadius: BorderRadius.circular(10),
        border: Border.all(color: const Color(0xFFE2E8F0)),
      ),
      child: Column(
        crossAxisAlignment: CrossAxisAlignment.start,
        children: [
          Wrap(
            spacing: 8,
            runSpacing: 6,
            children: [
              _buildStatusPill(
                label: 'Typical ($typicalCnt)',
                ok: typicalCnt > 0,
              ),
              _buildStatusPill(
                label: divCnt > 0 ? 'Diversion ($divCnt)' : 'Diversion: None',
                ok: divCnt > 0,
                neutral: divCnt == 0,
              ),
              if (rainCnt > 0)
                _buildStatusPill(
                  label: 'Rainwater ($rainCnt)',
                  ok: true,
                ),
              if (opts.isNotEmpty)
                Container(
                  padding: const EdgeInsets.symmetric(horizontal: 10, vertical: 4),
                  decoration: BoxDecoration(
                    color: Colors.white,
                    borderRadius: BorderRadius.circular(6),
                    border: Border.all(color: const Color(0xFFCBD5E1)),
                  ),
                  child: Text(
                    'Options: ${opts.join(', ')}',
                    style: const TextStyle(fontSize: 12, color: Color(0xFF334155), fontWeight: FontWeight.w500),
                  ),
                ),
            ],
          ),
          if (towers.isNotEmpty) ...[
            const SizedBox(height: 10),
            Text(
              'Detected Towers (${towers.length}):',
              style: const TextStyle(fontSize: 12, fontWeight: FontWeight.bold, color: Color(0xFF334155)),
            ),
            const SizedBox(height: 6),
            Wrap(
              spacing: 8,
              runSpacing: 6,
              children: towers.map((t) {
                final m = t as Map<String, dynamic>;
                return Container(
                  padding: const EdgeInsets.symmetric(horizontal: 10, vertical: 5),
                  decoration: BoxDecoration(
                    color: Colors.white,
                    borderRadius: BorderRadius.circular(6),
                    border: Border.all(color: const Color(0xFFCBD5E1)),
                  ),
                  child: Text(
                    '${m['name']}  •  ${m['floors']} Flr  |  T:${m['toilets']}  K:${m['kitchens']}',
                    style: const TextStyle(fontSize: 11.5, color: Color(0xFF0F172A)),
                  ),
                );
              }).toList(),
            ),
          ],
          if (warnings.isNotEmpty) ...[
            const SizedBox(height: 10),
            Container(
              width: double.infinity,
              padding: const EdgeInsets.all(10),
              decoration: BoxDecoration(
                color: const Color(0xFFFFFBEB),
                borderRadius: BorderRadius.circular(8),
                border: Border.all(color: const Color(0xFFFDE68A)),
              ),
              child: Column(
                crossAxisAlignment: CrossAxisAlignment.start,
                children: warnings.map((w) => Padding(
                  padding: const EdgeInsets.symmetric(vertical: 2),
                  child: Row(
                    crossAxisAlignment: CrossAxisAlignment.start,
                    children: [
                      const Icon(Icons.warning_amber_rounded, size: 16, color: Color(0xFFD97706)),
                      const SizedBox(width: 6),
                      Expanded(
                        child: Text(
                          w.toString(),
                          style: const TextStyle(fontSize: 12, color: Color(0xFF92400E), fontWeight: FontWeight.w500),
                        ),
                      ),
                    ],
                  ),
                )).toList(),
              ),
            ),
          ],
        ],
      ),
    );
  }

  Widget _buildStatusPill({required String label, required bool ok, bool neutral = false, bool warning = false}) {
    final bg = warning
        ? const Color(0xFFFFFBEB)
        : (neutral ? const Color(0xFFF1F5F9) : (ok ? const Color(0xFFF0FDF4) : const Color(0xFFFEF2F2)));
    final border = warning
        ? const Color(0xFFFDE68A)
        : (neutral ? const Color(0xFFCBD5E1) : (ok ? const Color(0xFFBBF7D0) : const Color(0xFFFECACA)));
    final fg = warning
        ? const Color(0xFFB45309)
        : (neutral ? const Color(0xFF64748B) : (ok ? const Color(0xFF166534) : const Color(0xFFB91C1C)));
    return Container(
      padding: const EdgeInsets.symmetric(horizontal: 10, vertical: 4),
      decoration: BoxDecoration(
        color: bg,
        borderRadius: BorderRadius.circular(6),
        border: Border.all(color: border),
      ),
      child: Text(
        '${ok && !neutral && !warning ? "✅ " : ""}$label',
        style: TextStyle(fontSize: 12, fontWeight: FontWeight.w600, color: fg),
      ),
    );
  }

  Widget _buildDiscountField({required TextEditingController controller, required String label}) {
    return TextField(
      controller: controller,
      keyboardType: TextInputType.number,
      textAlign: TextAlign.center,
      decoration: InputDecoration(
        labelText: label,
        border: OutlineInputBorder(borderRadius: BorderRadius.circular(8)),
        contentPadding: const EdgeInsets.symmetric(horizontal: 8, vertical: 10),
      ),
    );
  }

  // --- VIEW 2: TRAIN (FOR NEW FORMATS ONLY) ---
  Widget _buildTrainView() {
    return SingleChildScrollView(
      padding: const EdgeInsets.symmetric(horizontal: 24, vertical: 28),
      child: Center(
        child: Container(
          constraints: const BoxConstraints(maxWidth: 780),
          child: Column(
            crossAxisAlignment: CrossAxisAlignment.stretch,
            children: [
              Container(
                decoration: BoxDecoration(
                  color: Colors.white,
                  borderRadius: BorderRadius.circular(14),
                  border: Border.all(color: const Color(0xFFCBD5E1), width: 1.5),
                  boxShadow: [
                    BoxShadow(color: Colors.black.withOpacity(0.02), blurRadius: 8, offset: const Offset(0, 2)),
                  ],
                ),
                padding: const EdgeInsets.symmetric(horizontal: 24, vertical: 32),
                child: Stack(
                  children: [
                    if (trainFile != null || trainingResult != null)
                      Positioned(
                        top: 0,
                        right: 0,
                        child: OutlinedButton.icon(
                          onPressed: resetTrainState,
                          icon: const Icon(Icons.refresh, size: 16),
                          label: const Text('Reset', style: TextStyle(fontSize: 12.5, fontWeight: FontWeight.w600)),
                          style: OutlinedButton.styleFrom(
                            foregroundColor: const Color(0xFF475569),
                            side: const BorderSide(color: Color(0xFFCBD5E1)),
                            padding: const EdgeInsets.symmetric(horizontal: 12, vertical: 8),
                            shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(8)),
                          ),
                        ),
                      ),
                    SizedBox(
                      width: double.infinity,
                      child: Column(
                        children: [
                          Container(
                            padding: const EdgeInsets.all(16),
                            decoration: const BoxDecoration(
                              color: Color(0xFFF3E8FF),
                              shape: BoxShape.circle,
                            ),
                            child: const Icon(Icons.school, size: 40, color: Color(0xFF7E22CE)),
                          ),
                          const SizedBox(height: 16),
                          const Text(
                            'सिस्टम को नया BOQ फॉर्मेट सिखाएं (Summary या Raw BOQ)',
                            style: TextStyle(fontSize: 17, fontWeight: FontWeight.bold, color: Color(0xFF0F172A)),
                          ),
                          const SizedBox(height: 6),
                          const Text(
                            'यहाँ आप पहले से बनी Summary वाली फाइल या बिना Summary वाली Raw BOQ — दोनों अपलोड कर सकते हैं। सिस्टम उसका पूरा लेआउट, टावर्स और लॉजिक समझकर Knowledge Base में सेव कर लेगा।',
                            textAlign: TextAlign.center,
                            style: TextStyle(color: Color(0xFF64748B), fontSize: 13),
                          ),
                          const SizedBox(height: 18),
                          ElevatedButton.icon(
                            onPressed: selectTrainFile,
                            icon: const Icon(Icons.upload_file, size: 18),
                            label: Text(trainFile == null ? 'Select BOQ File (.xlsx - Raw or with Summary)' : 'Selected: ${trainFile!.name}'),
                            style: ElevatedButton.styleFrom(
                              backgroundColor: const Color(0xFF7E22CE),
                              foregroundColor: Colors.white,
                              padding: const EdgeInsets.symmetric(horizontal: 26, vertical: 13),
                              shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(8)),
                              elevation: 0,
                            ),
                          ),
                        ],
                      ),
                    ),
                  ],
                ),
              ),

              if (isTraining) ...[
                const SizedBox(height: 24),
                const Center(child: CircularProgressIndicator(strokeWidth: 3)),
                const SizedBox(height: 8),
                const Center(child: Text('Extracting rules, formulas & saving into brain...')),
              ],

              if (trainingResult != null) ...[
                const SizedBox(height: 24),
                Builder(builder: (context) {
                  final learned = (trainingResult!['learned'] as Map<String, dynamic>?) ?? {};
                  final tType = trainingResult!['training_type'] ?? learned['training_type'] ?? 'full_summary';
                  final isRaw = tType == 'raw_layout';
                  final List<dynamic> towerNames = (learned['tower_names'] as List<dynamic>?) ?? [];
                  final List<dynamic> warnings = (trainingResult!['warnings'] as List<dynamic>?) ?? [];
                  return Container(
                    decoration: BoxDecoration(
                      color: const Color(0xFFFDF4FF),
                      borderRadius: BorderRadius.circular(12),
                      border: Border.all(color: const Color(0xFFF0ABFC)),
                    ),
                    padding: const EdgeInsets.all(20),
                    child: Column(
                      crossAxisAlignment: CrossAxisAlignment.start,
                      children: [
                        Row(
                          children: [
                            const Icon(Icons.check_circle, color: Color(0xFF7E22CE), size: 20),
                            const SizedBox(width: 8),
                            Expanded(
                              child: Text(
                                isRaw ? 'Raw BOQ लेआउट सफलतापूर्वक सीखा गया!' : 'Full Summary फ़ॉर्मेट सफलतापूर्वक सीखा गया!',
                                style: const TextStyle(fontWeight: FontWeight.bold, fontSize: 15, color: Color(0xFF7E22CE)),
                              ),
                            ),
                            Container(
                              padding: const EdgeInsets.symmetric(horizontal: 8, vertical: 3),
                              decoration: BoxDecoration(
                                color: isRaw ? const Color(0xFFFEF3C7) : const Color(0xFFDCFCE7),
                                borderRadius: BorderRadius.circular(6),
                              ),
                              child: Text(
                                isRaw ? 'Raw Layout ✓' : 'Full Summary ✓',
                                style: TextStyle(
                                  fontSize: 11,
                                  fontWeight: FontWeight.bold,
                                  color: isRaw ? const Color(0xFF92400E) : const Color(0xFF166534),
                                ),
                              ),
                            ),
                          ],
                        ),
                        const SizedBox(height: 10),
                        Text('प्रोजेक्ट: ${learned['project_title'] ?? ''}', style: const TextStyle(fontWeight: FontWeight.bold)),
                        const SizedBox(height: 4),
                        Text('फॉर्मेट / Archetype: ${learned['archetype'] ?? 'standard'}  •  Towers: ${learned['tower_count'] ?? 1} ${towerNames.isNotEmpty ? '(${towerNames.join(', ')})' : ''}'),
                        const SizedBox(height: 4),
                        Text('पहचाने गए ऑप्शंस: ${(learned['options'] as List?)?.map((e) => e['option']).join(', ') ?? ''}'),
                        const SizedBox(height: 4),
                        Builder(builder: (context) {
                          final qty = (learned['quantities'] as Map<String, dynamic>?) ?? {};
                          final k = qty['kitchens'];
                          return Text(
                            'संख्या: Toilets: ${qty['toilets'] ?? '-'}  •  Kitchens: ${k ?? 0}${k == 0 ? ' (No Kitchen / Commercial)' : ''}  •  Floors: ${learned['floors_count'] ?? '-'}',
                            style: const TextStyle(fontSize: 12.5, color: Color(0xFF475569), fontWeight: FontWeight.w600),
                          );
                        }),
                        if (trainingResult!['preserved_existing_summary'] == true) ...[
                          const SizedBox(height: 8),
                          const Text(
                            'नोट: इस प्रोजेक्ट की पहले से मौजूद Verified Summary सुरक्षित रखी गई है।',
                            style: TextStyle(fontSize: 12, color: Color(0xFF1D4ED8), fontWeight: FontWeight.w600),
                          ),
                        ],
                        if (warnings.isNotEmpty) ...[
                          const SizedBox(height: 10),
                          ...warnings.map((w) => Padding(
                                padding: const EdgeInsets.only(top: 3),
                                child: Text('⚠️ $w', style: const TextStyle(fontSize: 12, color: Color(0xFFB45309))),
                              )),
                        ],
                      ],
                    ),
                  );
                }),
              ],
            ],
          ),
        ),
      ),
    );
  }

  // --- VIEW 0: LOGIN SCREEN (DARK LAMP THEME) ---
  Future<void> _showAdminPasswordDialog() async {
    adminPassCtrl.clear();
    String? dialogError;
    await showDialog(
      context: context,
      barrierDismissible: true,
      builder: (ctx) {
        return StatefulBuilder(builder: (ctx, setDlgState) {
          Future<void> submit() async {
            final pwd = adminPassCtrl.text.trim();
            if (pwd.isEmpty) {
              setDlgState(() => dialogError = 'कृपया Admin Password दर्ज करें');
              return;
            }
            setDlgState(() {
              dialogError = null;
              isAdminLoggingIn = true;
            });
            try {
              final res = await http.post(
                Uri.parse('$apiUrl/admin/login'),
                body: {'password': pwd},
              );
              if (res.statusCode == 200) {
                if (ctx.mounted) Navigator.of(ctx).pop();
                setState(() {
                  loggedInRole = 'admin';
                  adminPassword = pwd;
                  selectedMode = 2;
                  adminPassCtrl.clear();
                });
                await fetchAdminDashboard();
              } else {
                String msg = '❌ गलत Admin Password';
                try {
                  final errData = jsonDecode(res.body);
                  if (errData is Map && errData['detail'] != null) {
                    msg = '❌ ${errData['detail']}';
                  }
                } catch (_) {}
                setDlgState(() => dialogError = msg);
              }
            } catch (e) {
              setDlgState(() => dialogError = 'Server Error: $e');
            } finally {
              if (ctx.mounted) setDlgState(() => isAdminLoggingIn = false);
            }
          }

          return Dialog(
            backgroundColor: const Color(0xFF1E293B),
            shape: RoundedRectangleBorder(
              borderRadius: BorderRadius.circular(16),
              side: const BorderSide(color: Color(0xFF334155)),
            ),
            child: Container(
              width: 380,
              padding: const EdgeInsets.all(24),
              child: Column(
                mainAxisSize: MainAxisSize.min,
                children: [
                  Container(
                    padding: const EdgeInsets.all(12),
                    decoration: BoxDecoration(
                      color: const Color(0xFFF59E0B).withOpacity(0.15),
                      shape: BoxShape.circle,
                    ),
                    child: const Icon(Icons.lock_outline, color: Color(0xFFFBBF24), size: 28),
                  ),
                  const SizedBox(height: 14),
                  const Text(
                    'Admin Authentication',
                    style: TextStyle(fontSize: 18, fontWeight: FontWeight.bold, color: Colors.white),
                  ),
                  const SizedBox(height: 4),
                  const Text(
                    'Admin Dashboard खोलने के लिए Password डालें',
                    style: TextStyle(fontSize: 12.5, color: Color(0xFF94A3B8)),
                  ),
                  const SizedBox(height: 18),
                  TextField(
                    controller: adminPassCtrl,
                    obscureText: true,
                    autofocus: true,
                    style: const TextStyle(color: Colors.white),
                    onSubmitted: (_) => submit(),
                    decoration: InputDecoration(
                      labelText: 'Admin Password',
                      labelStyle: const TextStyle(color: Color(0xFF94A3B8)),
                      prefixIcon: const Icon(Icons.key, color: Color(0xFFFBBF24), size: 18),
                      filled: true,
                      fillColor: const Color(0xFF0F172A),
                      enabledBorder: OutlineInputBorder(
                        borderRadius: BorderRadius.circular(10),
                        borderSide: const BorderSide(color: Color(0xFF334155)),
                      ),
                      focusedBorder: OutlineInputBorder(
                        borderRadius: BorderRadius.circular(10),
                        borderSide: const BorderSide(color: Color(0xFFFBBF24), width: 1.5),
                      ),
                    ),
                  ),
                  if (dialogError != null) ...[
                    const SizedBox(height: 10),
                    Text(dialogError!, style: const TextStyle(color: Color(0xFFFCA5A5), fontSize: 12.5, fontWeight: FontWeight.w600)),
                  ],
                  const SizedBox(height: 20),
                  Row(
                    children: [
                      Expanded(
                        child: OutlinedButton(
                          onPressed: () => Navigator.of(ctx).pop(),
                          style: OutlinedButton.styleFrom(
                            foregroundColor: const Color(0xFFCBD5E1),
                            side: const BorderSide(color: Color(0xFF475569)),
                            padding: const EdgeInsets.symmetric(vertical: 13),
                            shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(8)),
                          ),
                          child: const Text('Cancel'),
                        ),
                      ),
                      const SizedBox(width: 12),
                      Expanded(
                        child: ElevatedButton(
                          onPressed: isAdminLoggingIn ? null : submit,
                          style: ElevatedButton.styleFrom(
                            backgroundColor: const Color(0xFFF59E0B),
                            foregroundColor: const Color(0xFF0F172A),
                            padding: const EdgeInsets.symmetric(vertical: 13),
                            shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(8)),
                          ),
                          child: isAdminLoggingIn
                              ? const SizedBox(width: 16, height: 16, child: CircularProgressIndicator(strokeWidth: 2))
                              : const Text('Unlock Admin', style: TextStyle(fontWeight: FontWeight.bold)),
                        ),
                      ),
                    ],
                  ),
                ],
              ),
            ),
          );
        });
      },
    );
  }

  Widget _buildLoginView() {
    return Container(
      width: double.infinity,
      height: double.infinity,
      decoration: const BoxDecoration(
        gradient: LinearGradient(
          begin: Alignment.topCenter,
          end: Alignment.bottomCenter,
          colors: [Color(0xFF070B14), Color(0xFF0F172A), Color(0xFF111827)],
        ),
      ),
      child: Stack(
        alignment: Alignment.topCenter,
        children: [
          // Warm Lamp Ambient Glow Cone at Top Center
          Positioned(
            top: 0,
            child: Column(
              children: [
                // Lamp Cord
                Container(
                  width: 3,
                  height: 36,
                  decoration: const BoxDecoration(
                    gradient: LinearGradient(
                      begin: Alignment.topCenter,
                      end: Alignment.bottomCenter,
                      colors: [Color(0xFF334155), Color(0xFF94A3B8)],
                    ),
                  ),
                ),
                // Lamp Shade
                Container(
                  width: 96,
                  height: 38,
                  decoration: BoxDecoration(
                    gradient: const LinearGradient(
                      begin: Alignment.topCenter,
                      end: Alignment.bottomCenter,
                      colors: [Color(0xFF334155), Color(0xFF1E293B)],
                    ),
                    borderRadius: const BorderRadius.only(
                      topLeft: Radius.circular(48),
                      topRight: Radius.circular(48),
                      bottomLeft: Radius.circular(10),
                      bottomRight: Radius.circular(10),
                    ),
                    border: Border.all(color: const Color(0xFF475569)),
                    boxShadow: [
                      BoxShadow(
                        color: const Color(0xFFF59E0B).withOpacity(0.35),
                        blurRadius: 30,
                        spreadRadius: 4,
                        offset: const Offset(0, 12),
                      ),
                    ],
                  ),
                  child: Align(
                    alignment: Alignment.bottomCenter,
                    child: Container(
                      width: 34,
                      height: 6,
                      margin: const EdgeInsets.only(bottom: 2),
                      decoration: BoxDecoration(
                        color: const Color(0xFFFDE68A),
                        borderRadius: BorderRadius.circular(6),
                        boxShadow: [
                          BoxShadow(
                            color: const Color(0xFFFBBF24).withOpacity(0.9),
                            blurRadius: 18,
                            spreadRadius: 4,
                          ),
                        ],
                      ),
                    ),
                  ),
                ),
                // Warm Radial Light Beam
                Container(
                  width: 520,
                  height: 220,
                  decoration: BoxDecoration(
                    gradient: RadialGradient(
                      center: Alignment.topCenter,
                      radius: 0.9,
                      colors: [
                        const Color(0xFFF59E0B).withOpacity(0.20),
                        const Color(0xFFF59E0B).withOpacity(0.05),
                        Colors.transparent,
                      ],
                    ),
                  ),
                ),
              ],
            ),
          ),

          // Center Single Welcome Window & 2 Tabs (User | Admin)
          Center(
            child: SingleChildScrollView(
              padding: const EdgeInsets.fromLTRB(24, 96, 24, 32),
              child: Container(
                constraints: const BoxConstraints(maxWidth: 460),
                padding: const EdgeInsets.symmetric(horizontal: 32, vertical: 34),
                decoration: BoxDecoration(
                  color: const Color(0xFF1E293B).withOpacity(0.90),
                  borderRadius: BorderRadius.circular(22),
                  border: Border.all(color: const Color(0xFF334155), width: 1.2),
                  boxShadow: [
                    BoxShadow(
                      color: Colors.black.withOpacity(0.50),
                      blurRadius: 34,
                      offset: const Offset(0, 14),
                    ),
                    BoxShadow(
                      color: const Color(0xFFF59E0B).withOpacity(0.09),
                      blurRadius: 42,
                      offset: const Offset(0, -6),
                    ),
                  ],
                ),
                child: Column(
                  mainAxisSize: MainAxisSize.min,
                  children: [
                    Container(
                      padding: const EdgeInsets.symmetric(horizontal: 12, vertical: 5),
                      decoration: BoxDecoration(
                        color: const Color(0xFF0F172A),
                        borderRadius: BorderRadius.circular(20),
                        border: Border.all(color: const Color(0xFF334155)),
                      ),
                      child: Row(
                        mainAxisSize: MainAxisSize.min,
                        children: [
                          CircleAvatar(
                            radius: 4,
                            backgroundColor: isBackendOnline ? const Color(0xFF34D399) : const Color(0xFFF87171),
                          ),
                          const SizedBox(width: 7),
                          Text(
                            isBackendOnline ? 'System Online' : 'Connecting...',
                            style: const TextStyle(color: Color(0xFFCBD5E1), fontSize: 11.5, fontWeight: FontWeight.w600),
                          ),
                        ],
                      ),
                    ),
                    const SizedBox(height: 20),
                    const Text(
                      'Welcome to',
                      style: TextStyle(fontSize: 14, color: Color(0xFFFBBF24), fontWeight: FontWeight.w600, letterSpacing: 1.2),
                    ),
                    const SizedBox(height: 6),
                    const Text(
                      'VIEW SUMMARY STUDIO',
                      textAlign: TextAlign.center,
                      style: TextStyle(
                        fontSize: 26,
                        fontWeight: FontWeight.w900,
                        color: Colors.white,
                        letterSpacing: 1.0,
                      ),
                    ),
                    const SizedBox(height: 6),
                    const Text(
                      'Automated Commercial & MEP Estimation Engine',
                      textAlign: TextAlign.center,
                      style: TextStyle(fontSize: 12.5, color: Color(0xFF94A3B8)),
                    ),
                    const SizedBox(height: 28),

                    // Single Segmented Tab Bar at bottom of Window: Tab 1 (User) | Tab 2 (Admin)
                    Container(
                      padding: const EdgeInsets.all(6),
                      decoration: BoxDecoration(
                        color: const Color(0xFF0F172A),
                        borderRadius: BorderRadius.circular(14),
                        border: Border.all(color: const Color(0xFF334155)),
                      ),
                      child: Row(
                        children: [
                          Expanded(
                            child: InkWell(
                              onTap: loginAsUser,
                              borderRadius: BorderRadius.circular(10),
                              child: Container(
                                padding: const EdgeInsets.symmetric(vertical: 14),
                                decoration: BoxDecoration(
                                  gradient: const LinearGradient(
                                    colors: [Color(0xFF0284C7), Color(0xFF0369A1)],
                                  ),
                                  borderRadius: BorderRadius.circular(10),
                                ),
                                child: const Row(
                                  mainAxisAlignment: MainAxisAlignment.center,
                                  children: [
                                    Icon(Icons.person_outline, size: 18, color: Colors.white),
                                    SizedBox(width: 8),
                                    Text(
                                      'User',
                                      style: TextStyle(color: Colors.white, fontSize: 14.5, fontWeight: FontWeight.bold),
                                    ),
                                  ],
                                ),
                              ),
                            ),
                          ),
                          const SizedBox(width: 8),
                          Expanded(
                            child: InkWell(
                              onTap: _showAdminPasswordDialog,
                              borderRadius: BorderRadius.circular(10),
                              child: Container(
                                padding: const EdgeInsets.symmetric(vertical: 14),
                                decoration: BoxDecoration(
                                  color: const Color(0xFF1E293B),
                                  borderRadius: BorderRadius.circular(10),
                                  border: Border.all(color: const Color(0xFFF59E0B).withOpacity(0.6)),
                                ),
                                child: const Row(
                                  mainAxisAlignment: MainAxisAlignment.center,
                                  children: [
                                    Icon(Icons.lock_outline, size: 17, color: Color(0xFFFBBF24)),
                                    SizedBox(width: 8),
                                    Text(
                                      'Admin',
                                      style: TextStyle(color: Color(0xFFFBBF24), fontSize: 14.5, fontWeight: FontWeight.bold),
                                    ),
                                  ],
                                ),
                              ),
                            ),
                          ),
                        ],
                      ),
                    ),
                  ],
                ),
              ),
            ),
          ),
        ],
      ),
    );
  }

  // --- VIEW 3: ADMIN DASHBOARD (ADMIN ONLY) ---
  Widget _buildAdminDashboardView() {
    final totalFormats = adminDashboardData?['total_formats'] ?? '-';
    final trainedToday = adminDashboardData?['trained_today'] ?? '-';
    final lastTrained = adminDashboardData?['last_trained_at'] ?? '-';
    final List<dynamic> recent10 = (adminDashboardData?['recent_10'] as List<dynamic>?) ?? [];

    return SingleChildScrollView(
      padding: const EdgeInsets.symmetric(horizontal: 24, vertical: 28),
      child: Center(
        child: Container(
          constraints: const BoxConstraints(maxWidth: 920),
          child: Column(
            crossAxisAlignment: CrossAxisAlignment.stretch,
            children: [
              Row(
                children: [
                  const Icon(Icons.admin_panel_settings, color: Color(0xFFB45309), size: 26),
                  const SizedBox(width: 10),
                  const Text(
                    'Admin Dashboard',
                    style: TextStyle(fontSize: 20, fontWeight: FontWeight.bold, color: Color(0xFF0F172A)),
                  ),
                  const Spacer(),
                  OutlinedButton.icon(
                    onPressed: isLoadingDashboard ? null : fetchAdminDashboard,
                    icon: const Icon(Icons.refresh, size: 17),
                    label: const Text('Refresh'),
                  ),
                  const SizedBox(width: 10),
                  ElevatedButton.icon(
                    onPressed: downloadKnowledgeBackup,
                    icon: const Icon(Icons.download, size: 18),
                    label: const Text('Download Knowledge Backup'),
                    style: ElevatedButton.styleFrom(
                      backgroundColor: const Color(0xFF0F4C81),
                      foregroundColor: Colors.white,
                      padding: const EdgeInsets.symmetric(horizontal: 18, vertical: 12),
                      shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(8)),
                    ),
                  ),
                ],
              ),
              const SizedBox(height: 20),
              Row(
                children: [
                  Expanded(child: _buildStatCard('Total Formats', '$totalFormats', Icons.library_books, const Color(0xFF0F4C81))),
                  const SizedBox(width: 14),
                  Expanded(child: _buildStatCard('आज नई Train हुईं', '$trainedToday', Icons.today, const Color(0xFF059669))),
                  const SizedBox(width: 14),
                  Expanded(child: _buildStatCard('Last Train कब हुई', '$lastTrained', Icons.schedule, const Color(0xFF7E22CE))),
                ],
              ),
              const SizedBox(height: 24),
              Container(
                padding: const EdgeInsets.all(20),
                decoration: BoxDecoration(
                  color: Colors.white,
                  borderRadius: BorderRadius.circular(12),
                  border: Border.all(color: const Color(0xFFE2E8F0)),
                ),
                child: Column(
                  crossAxisAlignment: CrossAxisAlignment.start,
                  children: [
                    const Text(
                      'आखिरी 10 Train की हुई फाइलें',
                      style: TextStyle(fontSize: 15, fontWeight: FontWeight.bold, color: Color(0xFF0F172A)),
                    ),
                    const SizedBox(height: 12),
                    if (isLoadingDashboard)
                      const Center(child: Padding(padding: EdgeInsets.all(20), child: CircularProgressIndicator()))
                    else if (recent10.isEmpty)
                      const Text('कोई डेटा उपलब्ध नहीं है।', style: TextStyle(color: Color(0xFF64748B)))
                    else
                      ...recent10.asMap().entries.map((entry) {
                        final idx = entry.key + 1;
                        final item = entry.value as Map<String, dynamic>;
                        return Container(
                          padding: const EdgeInsets.symmetric(vertical: 10, horizontal: 12),
                          margin: const EdgeInsets.only(bottom: 8),
                          decoration: BoxDecoration(
                            color: const Color(0xFFF8FAFC),
                            borderRadius: BorderRadius.circular(8),
                            border: Border.all(color: const Color(0xFFE2E8F0)),
                          ),
                          child: Row(
                            children: [
                              CircleAvatar(
                                radius: 13,
                                backgroundColor: const Color(0xFF0F4C81),
                                child: Text('$idx', style: const TextStyle(color: Colors.white, fontSize: 11, fontWeight: FontWeight.bold)),
                              ),
                              const SizedBox(width: 12),
                              Expanded(
                                child: Column(
                                  crossAxisAlignment: CrossAxisAlignment.start,
                                  children: [
                                    Text(
                                      item['filename']?.toString() ?? '',
                                      style: const TextStyle(fontSize: 13, fontWeight: FontWeight.bold, color: Color(0xFF0F172A)),
                                    ),
                                    Text(
                                      '${item['project_title'] ?? ''}  •  ${item['archetype'] ?? ''}',
                                      style: const TextStyle(fontSize: 11.5, color: Color(0xFF64748B)),
                                    ),
                                  ],
                                ),
                              ),
                              Container(
                                padding: const EdgeInsets.symmetric(horizontal: 8, vertical: 3),
                                decoration: BoxDecoration(
                                  color: item['training_type'] == 'raw_layout'
                                      ? const Color(0xFFFEF3C7)
                                      : (item['training_type'] == 'generated_summary'
                                          ? const Color(0xFFEFF6FF)
                                          : const Color(0xFFDCFCE7)),
                                  borderRadius: BorderRadius.circular(6),
                                ),
                                child: Text(
                                  item['training_type']?.toString() ?? '',
                                  style: const TextStyle(fontSize: 11, fontWeight: FontWeight.bold),
                                ),
                              ),
                              const SizedBox(width: 10),
                              Text(
                                item['trained_at']?.toString() ?? 'Legacy',
                                style: const TextStyle(fontSize: 11.5, color: Color(0xFF475569)),
                              ),
                            ],
                          ),
                        );
                      }),
                  ],
                ),
              ),
            ],
          ),
        ),
      ),
    );
  }

  Widget _buildStatCard(String label, String value, IconData icon, Color color) {
    return Container(
      padding: const EdgeInsets.all(18),
      decoration: BoxDecoration(
        color: Colors.white,
        borderRadius: BorderRadius.circular(12),
        border: Border.all(color: const Color(0xFFE2E8F0)),
      ),
      child: Row(
        children: [
          Container(
            padding: const EdgeInsets.all(10),
            decoration: BoxDecoration(color: color.withOpacity(0.1), borderRadius: BorderRadius.circular(10)),
            child: Icon(icon, color: color, size: 22),
          ),
          const SizedBox(width: 12),
          Expanded(
            child: Column(
              crossAxisAlignment: CrossAxisAlignment.start,
              children: [
                Text(label, style: const TextStyle(fontSize: 12, color: Color(0xFF64748B), fontWeight: FontWeight.w600)),
                const SizedBox(height: 4),
                Text(value, style: const TextStyle(fontSize: 17, fontWeight: FontWeight.bold, color: Color(0xFF0F172A)), overflow: TextOverflow.ellipsis),
              ],
            ),
          ),
        ],
      ),
    );
  }
}

class BoxElevation extends BoxShadow {
  const BoxElevation({required super.color, required super.blurRadius, required super.offset});
}
