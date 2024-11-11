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
  int columnSpan;
  int rowSpan;
  static int _imageCounter = 2;

  static int _pixelsToEmu(double pixels) {
    return (pixels * (914400 / 96)).round();
  }

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
        this.offsetY = 0,
        this.columnSpan = 1,
        this.rowSpan = 1;

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

    final finalWidth = widthInPixels ?? image.width;
    final finalHeight = heightInPixels ??
        (widthInPixels != null
            ? (image.height * widthInPixels / image.width).round()
            : image.height);

    final instance = ExcelImage._(
      name: imageName,
      extension: extension,
      imageBytes: imageBytes,
      width: finalWidth,
      height: finalHeight,
      contentType: 'image/jpeg',
      reuseRid: reuseRid,
    );

    instance.offsetX = _pixelsToEmu(offsetXInPixels.toDouble());
    instance.offsetY = _pixelsToEmu(offsetYInPixels.toDouble());

    return instance;
  }

  void calculateCellSpan(double cellWidth, double cellHeight) {
    final imageWidthPx = width / _pixelsToEmu(1.0);
    final imageHeightPx = height / _pixelsToEmu(1.0);

    columnSpan = (imageWidthPx / cellWidth).ceil();
    rowSpan = (imageHeightPx / cellHeight).ceil();

    print('Cell span calculations:');
    print('  Image size (px): ${imageWidthPx}x${imageHeightPx}');
    print('  Cell size (px): ${cellWidth}x${cellHeight}');
    print('  Required spans: ${columnSpan}x${rowSpan}');
    print('  Final size (EMU): ${width}x${height}');
    print('  Offset (EMU): ${offsetX}x${offsetY}');
  }

  String get mediaPath => 'xl/media/$name';

  ArchiveFile toArchiveFile() {
    return ArchiveFile(mediaPath, imageBytes.length, imageBytes);
  }
}
