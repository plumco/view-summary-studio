import 'dart:async';
import 'dart:html' as html;
import 'dart:typed_data';

class WebFile {
  final String name;
  final Uint8List bytes;

  WebFile({required this.name, required this.bytes});
}

Future<WebFile?> pickExcelFile() async {
  final completer = Completer<WebFile?>();
  final uploadInput = html.FileUploadInputElement();
  uploadInput.accept = '.xlsx';
  uploadInput.click();

  uploadInput.onChange.listen((e) {
    final files = uploadInput.files;
    if (files != null && files.isNotEmpty) {
      final file = files[0];
      final reader = html.FileReader();
      reader.onLoadEnd.listen((e) {
        final result = reader.result;
        if (result is Uint8List) {
          completer.complete(WebFile(name: file.name, bytes: result));
        } else if (result is ByteBuffer) {
          completer.complete(WebFile(name: file.name, bytes: result.asUint8List()));
        } else {
          completer.complete(null);
        }
      });
      reader.readAsArrayBuffer(file);
    } else {
      completer.complete(null);
    }
  });

  return completer.future;
}

Future<List<WebFile>> pickMultipleExcelFiles() async {
  final completer = Completer<List<WebFile>>();
  final uploadInput = html.FileUploadInputElement();
  uploadInput.accept = '.xlsx';
  uploadInput.multiple = true;
  uploadInput.click();

  uploadInput.onChange.listen((e) async {
    final files = uploadInput.files;
    if (files == null || files.isEmpty) {
      completer.complete([]);
      return;
    }
    final List<WebFile> resultList = [];
    for (final file in files) {
      final fileCompleter = Completer<WebFile?>();
      final reader = html.FileReader();
      reader.onLoadEnd.listen((e) {
        final result = reader.result;
        if (result is Uint8List) {
          fileCompleter.complete(WebFile(name: file.name, bytes: result));
        } else if (result is ByteBuffer) {
          fileCompleter.complete(WebFile(name: file.name, bytes: result.asUint8List()));
        } else {
          fileCompleter.complete(null);
        }
      });
      reader.readAsArrayBuffer(file);
      final wf = await fileCompleter.future;
      if (wf != null) resultList.add(wf);
    }
    completer.complete(resultList);
  });

  return completer.future;
}

void downloadBytes(List<int> bytes, String filename, {String mimeType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}) {
  final blob = html.Blob([bytes], mimeType);
  final url = html.Url.createObjectUrlFromBlob(blob);
  html.AnchorElement(href: url)
    ..setAttribute('download', filename)
    ..click();
  html.Url.revokeObjectUrl(url);
}
