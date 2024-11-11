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

  static int _pixelsToEmu(int pixels) {
    // 1 inch = 96 pixels = 914400 EMU
    return (pixels * 914400 / 96).round();
  }

  static int _emuToPixels(int emu) => (emu / 9525).round();

  ExcelImage._({
    required this.name,
    required this.extension,
    required this.imageBytes,
    required int width,
    required int height,
    required this.contentType,
    String? reuseRid,
  })  : this.id = reuseRid ?? 'rId${_imageCounter}',
        this.nvPrId = reuseRid != null ? _imageCounter : _imageCounter++,
        this.originalWidth = width,
        this.originalHeight = height,
        this.width = _pixelsToEmu((width * 0.5).round()),
        this.height = _pixelsToEmu((height * 0.5).round()),
        this.offsetX = 0,
        this.offsetY = 0;

  factory ExcelImage.from(
    Uint8List imageBytes, {
    String? name,
    int? widthInPixels,
    int? heightInPixels,
    int offsetXInPixels = 0,
    int offsetYInPixels = 0,
    String? reuseRid,
  }) {
    final image = img.decodeImage(imageBytes);
    if (image == null) throw Exception('Invalid image data');

    final extension = '.jpg';
    final imageName =
        name ?? 'Image_${DateTime.now().millisecondsSinceEpoch}$extension';

    final instance = ExcelImage._(
      name: imageName,
      extension: extension,
      imageBytes: imageBytes,
      width: image.width,
      height: image.height,
      contentType: 'image/jpeg',
      reuseRid: reuseRid,
    );

    if (widthInPixels != null) instance.width = _pixelsToEmu(widthInPixels);
    if (heightInPixels != null) instance.height = _pixelsToEmu(heightInPixels);
    instance.offsetX = _pixelsToEmu(offsetXInPixels);
    instance.offsetY = _pixelsToEmu(offsetYInPixels);

    return instance;
  }

  String get mediaPath => 'xl/media/$name';

  ArchiveFile toArchiveFile() {
    return ArchiveFile(
      mediaPath,
      imageBytes.length,
      imageBytes,
    );
  }

  void fitToCell(double cellWidth, double cellHeight) {
    // 셀 패딩 고려 (Excel 기본값: 약 5%)
    final padding = 0.05;
    final availableWidth = cellWidth * (1 - 2 * padding);
    final availableHeight = cellHeight * (1 - 2 * padding);

    final cellWidthEmu = _pixelsToEmu(availableWidth.round());
    final cellHeightEmu = _pixelsToEmu(availableHeight.round());

    final ratio = originalWidth / originalHeight;

    if (ratio > availableWidth / availableHeight) {
      width = cellWidthEmu;
      height = (cellWidthEmu / ratio).round();
    } else {
      height = cellHeightEmu;
      width = (cellHeightEmu * ratio).round();
    }

    // 패딩을 고려한 오프셋 계산
    offsetX = _pixelsToEmu((cellWidth * padding).round());
    offsetY = _pixelsToEmu((cellHeight * padding).round());
  }
}
