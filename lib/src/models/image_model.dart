import 'dart:typed_data';
import 'package:image/image.dart' as img;

import 'package:archive/archive.dart';

class ExcelImage {
  final String id;
  final int nvPrId;
  final String name;
  final String extension;
  final Uint8List imageBytes;
  final int originalWidth;
  final int originalHeight;
  final String contentType;
  int width;
  int height;
  int offsetX;
  int offsetY;

  static int _imageCounter = 2;

  static int _pixelsToEmu(int pixels) => pixels * 9525;

  static int _emuToPixels(int emu) => (emu / 9525).round();

  ExcelImage._({
    required this.name,
    required this.extension,
    required this.imageBytes,
    required int width,
    required int height,
    required this.contentType,
  })  : this.id = 'rId${_imageCounter - 1}',
        this.nvPrId = _imageCounter++,
        this.originalWidth = width,
        this.originalHeight = height,
        this.width = _pixelsToEmu((width * 0.5).round()),
        this.height = _pixelsToEmu((height * 0.5).round()),
        this.offsetX = 0,
        this.offsetY = 0;

  static ExcelImage from(
    Uint8List bytes, {
    String? name,
    int? widthInPixels,
    int? heightInPixels,
    int offsetXInPixels = 0,
    int offsetYInPixels = 0,
  }) {
    final image = img.decodeImage(bytes);
    if (image == null) throw Exception('Invalid image data');

    final extension = _getImageExtension(bytes);
    final contentType = _getContentType(extension);

    final excelImage = ExcelImage._(
      name: name ?? 'Image_${DateTime.now().millisecondsSinceEpoch}$extension',
      extension: extension,
      imageBytes: bytes,
      width: image.width,
      height: image.height,
      contentType: contentType,
    );

    if (widthInPixels != null) {
      excelImage.width = _pixelsToEmu(widthInPixels);
    }
    if (heightInPixels != null) {
      excelImage.height = _pixelsToEmu(heightInPixels);
    }
    excelImage.offsetX = _pixelsToEmu(offsetXInPixels);
    excelImage.offsetY = _pixelsToEmu(offsetYInPixels);

    return excelImage;
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
