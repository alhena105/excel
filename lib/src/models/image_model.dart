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

  static int _pixelsToEmu(double pixels) {
    // 1 inch = 96 pixels = 914400 EMU
    return (pixels * 9525).round();
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
        this.width = _pixelsToEmu(width.toDouble()),
        this.height = _pixelsToEmu(height.toDouble()),
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
      width: widthInPixels ?? image.width,
      height: heightInPixels ?? image.height,
      contentType: 'image/jpeg',
      reuseRid: reuseRid,
    );

    if (widthInPixels != null)
      instance.width = _pixelsToEmu(widthInPixels.toDouble());
    if (heightInPixels != null)
      instance.height = _pixelsToEmu(heightInPixels.toDouble());
    instance.offsetX = _pixelsToEmu(offsetXInPixels.toDouble());
    instance.offsetY = _pixelsToEmu(offsetYInPixels.toDouble());

    print('width: ${instance.width}');
    print('height: ${instance.height}');
    print('offsetX: ${instance.offsetX}');
    print('offsetY: ${instance.offsetY}');

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

    final ratio = originalWidth / originalHeight;
    final cellRatio = availableWidth / availableHeight;

    double targetWidth, targetHeight;

    if (ratio > cellRatio) {
      targetWidth = availableWidth;
      targetHeight = availableWidth / ratio;
    } else {
      targetHeight = availableHeight;
      targetWidth = availableHeight * ratio;
    }

    width = _pixelsToEmu(targetWidth);
    height = _pixelsToEmu(targetHeight);

    // 가운데 정렬을 위한 오프셋 계산 (EMU 단위로 직접 지정)
    offsetX = ((cellWidth * 9525) - width) ~/ 2; // EMU 단위로 직접 계산
    offsetY = ((cellHeight * 9525) - height) ~/ 2; // EMU 단위로 직접 계산
  }
}
