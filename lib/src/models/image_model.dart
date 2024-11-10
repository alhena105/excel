import 'dart:typed_data';
import 'package:image/image.dart' as img;
import 'package:uuid/uuid.dart';
import 'package:archive/archive.dart';

class ExcelImage {
  final String id;
  final String name;
  final String extension;
  final Uint8List imageBytes;
  final int width;
  final int height;
  final String contentType;

  ExcelImage._({
    required this.id,
    required this.name,
    required this.extension,
    required this.imageBytes,
    required this.width,
    required this.height,
    required this.contentType,
  });

  static ExcelImage from(Uint8List bytes, {String? name}) {
    final image = img.decodeImage(bytes);
    if (image == null) throw Exception('Invalid image data');

    final extension = _getImageExtension(bytes);
    final contentType = _getContentType(extension);

    return ExcelImage._(
      id: 'rId${const Uuid().v4()}',
      name: name ?? 'Image_${DateTime.now().millisecondsSinceEpoch}$extension',
      extension: extension,
      imageBytes: bytes,
      width: image.width,
      height: image.height,
      contentType: contentType,
    );
  }

  static String _getImageExtension(Uint8List bytes) {
    if (bytes.length >= 2) {
      if (bytes[0] == 0xFF && bytes[1] == 0xD8) return '.jpg';
      if (bytes[0] == 0x89 && bytes[1] == 0x50) return '.png';
    }
    throw Exception(
        'Unsupported image format. Only JPG and PNG are supported.');
  }

  static String _getContentType(String extension) {
    switch (extension) {
      case '.jpg':
        return 'image/jpeg';
      case '.png':
        return 'image/png';
      default:
        throw Exception('Unsupported image format');
    }
  }

  String get mediaPath => 'xl/media/$name';

  ArchiveFile toArchiveFile() {
    return ArchiveFile(
      mediaPath,
      imageBytes.length,
      imageBytes,
    );
  }
}
