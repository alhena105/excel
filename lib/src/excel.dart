part of excel;

Excel _newExcel(Archive archive) {
  // Lookup at file format
  String? format;

  var mimetype = archive.findFile('mimetype');
  if (mimetype == null) {
    var xl = archive.findFile('xl/workbook.xml');
    if (xl != null) {
      format = _spreasheetXlsx;
    }
  }

  switch (format) {
    case _spreasheetXlsx:
      return Excel._(archive);
    default:
      throw UnsupportedError(
          'Excel format unsupported. Only .xlsx files are supported');
  }
}

/// Decode a excel file.
class Excel {
  bool _styleChanges = false;
  bool _mergeChanges = false;
  bool _rtlChanges = false;

  Archive _archive;

  final Map<String, XmlNode> _sheets = {};
  final Map<String, XmlDocument> _xmlFiles = {};
  final Map<String, String> _xmlSheetId = {};
  final Map<String, Map<String, int>> _cellStyleReferenced = {};
  final Map<String, Sheet> _sheetMap = {};

  List<CellStyle> _cellStyleList = [];
  List<String> _patternFill = [];
  final List<String> _mergeChangeLook = [];
  final List<String> _rtlChangeLook = [];
  List<_FontStyle> _fontStyleList = [];
  final List<int> _numFmtIds = [];
  final NumFormatMaintainer _numFormats = NumFormatMaintainer();
  List<_BorderSet> _borderSetList = [];

  _SharedStringsMaintainer _sharedStrings = _SharedStringsMaintainer._();

  String _stylesTarget = '';
  String _sharedStringsTarget = '';
  get _absSharedStringsTarget {
    if (_sharedStringsTarget.isNotEmpty && _sharedStringsTarget[0] == "/") {
      return _sharedStringsTarget.substring(1);
    }
    return "xl/${_sharedStringsTarget}";
  }

  String? _defaultSheet;
  late Parser parser;

  Excel._(this._archive) {
    parser = Parser._(this);
    parser._startParsing();
  }

  factory Excel.createExcel() {
    return Excel.decodeBytes(Base64Decoder().convert(_newSheet));
  }

  factory Excel.decodeBytes(List<int> data) {
    final Archive archive;
    try {
      archive = ZipDecoder().decodeBytes(data);
    } catch (e) {
      throw UnsupportedError(
          'Excel format unsupported. Only .xlsx files are supported');
    }
    return _newExcel(archive);
  }

  factory Excel.decodeBuffer(InputStream input) {
    return _newExcel(ZipDecoder().decodeBuffer(input));
  }

  ///
  ///It will return `tables` as map in order to mimic the previous versions reading the data.
  ///
  Map<String, Sheet> get tables {
    if (this._sheetMap.isEmpty) {
      _damagedExcel(text: "Corrupted Excel file.");
    }
    return Map<String, Sheet>.from(this._sheetMap);
  }

  ///
  ///It will return the SheetObject of `sheet`.
  ///
  ///If the `sheet` does not exist then it will create `sheet` with `New Sheet Object`
  ///
  Sheet operator [](String sheet) {
    _availSheet(sheet);
    return _sheetMap[sheet]!;
  }

  ///
  ///Returns the `Map<String, Sheet>`
  ///
  ///where `key` is the `Sheet Name` and the `value` is the `Sheet Object`
  ///
  Map<String, Sheet> get sheets {
    return Map<String, Sheet>.from(_sheetMap);
  }

  ///
  ///If `sheet` does not exist then it will be automatically created with contents of `sheetObject`
  ///
  ///Newly created sheet with name = `sheet` will have seperate reference and will not be linked to sheetObject.
  ///
  operator []=(String sheet, Sheet sheetObject) {
    _availSheet(sheet);

    _sheetMap[sheet] = Sheet._clone(this, sheet, sheetObject);
  }

  ///
  ///`sheet2Object` will be linked with `sheet1`.
  ///
  ///If `sheet1` does not exist then it will be automatically created.
  ///
  ///Important Note: After linkage the operations performed on `sheet1`, will also get performed on `sheet2Object` and `vica-versa`.
  ///
  void link(String sheet1, Sheet existingSheetObject) {
    if (_sheetMap[existingSheetObject.sheetName] != null) {
      _availSheet(sheet1);

      _sheetMap[sheet1] = _sheetMap[existingSheetObject.sheetName]!;

      if (_cellStyleReferenced[existingSheetObject.sheetName] != null) {
        _cellStyleReferenced[sheet1] = Map<String, int>.from(
            _cellStyleReferenced[existingSheetObject.sheetName]!);
      }
    }
  }

  ///
  ///If `sheet` is linked with any other sheet's object then it's link will be broke
  ///
  void unLink(String sheet) {
    if (_sheetMap[sheet] != null) {
      ///
      /// copying the sheet into itself thus resulting in breaking the linkage as Sheet._clone() will provide new reference;
      copy(sheet, sheet);
    }
  }

  ///
  ///Copies the content of `fromSheet` into `toSheet`.
  ///
  ///In order to successfully copy: `fromSheet` should exist in `excel.tables.keys`.
  ///
  ///If `toSheet` does not exist then it will be automatically created.
  ///
  void copy(String fromSheet, String toSheet) {
    _availSheet(toSheet);

    if (_sheetMap[fromSheet] != null) {
      this[toSheet] = this[fromSheet];
    }
    if (_cellStyleReferenced[fromSheet] != null) {
      _cellStyleReferenced[toSheet] =
          Map<String, int>.from(_cellStyleReferenced[fromSheet]!);
    }
  }

  ///
  ///Changes the name from `oldSheetName` to `newSheetName`.
  ///
  ///In order to rename : `oldSheetName` should exist in `excel.tables.keys` and `newSheetName` must not exist.
  ///
  void rename(String oldSheetName, String newSheetName) {
    if (_sheetMap[oldSheetName] != null && _sheetMap[newSheetName] == null) {
      ///
      /// rename from _defaultSheet var also
      if (_defaultSheet == oldSheetName) {
        _defaultSheet = newSheetName;
      }

      copy(oldSheetName, newSheetName);

      ///
      /// delete the `oldSheetName` as sheet with `newSheetName` is having cloned `SheetObject of oldSheetName` with new reference,
      delete(oldSheetName);
    }
  }

  ///
  ///If `sheet` exist in `excel.tables.keys` and `excel.tables.keys.length >= 2` then it will be `deleted`.
  ///
  void delete(String sheet) {
    ///
    /// remove the sheet `name` or `key` from the below locations if they exist.

    ///
    /// If it is not the last sheet then `delete` otherwise `return`;
    if (_sheetMap.length <= 1) {
      return;
    }

    ///
    ///remove from _defaultSheet var also
    if (_defaultSheet == sheet) {
      _defaultSheet = null;
    }

    ///
    /// remove the `Sheet Object` from `_sheetMap`.
    if (_sheetMap[sheet] != null) {
      _sheetMap.remove(sheet);
    }

    ///
    /// remove from `_mergeChangeLook`.
    if (_mergeChangeLook.contains(sheet)) {
      _mergeChangeLook.remove(sheet);
    }

    ///
    /// remove from `_rtlChangeLook`.
    if (_rtlChangeLook.contains(sheet)) {
      _rtlChangeLook.remove(sheet);
    }

    ///
    /// remove from `_xmlSheetId`.
    if (_xmlSheetId[sheet] != null) {
      String sheetId1 =
              "worksheets" + _xmlSheetId[sheet]!.split('worksheets')[1],
          sheetId2 = _xmlSheetId[sheet]!;

      _xmlFiles['xl/_rels/workbook.xml.rels']
          ?.rootElement
          .children
          .removeWhere((_sheetName) {
        return _sheetName.getAttribute('Target') != null &&
            _sheetName.getAttribute('Target') == sheetId1;
      });

      _xmlFiles['[Content_Types].xml']
          ?.rootElement
          .children
          .removeWhere((_sheetName) {
        return _sheetName.getAttribute('PartName') != null &&
            _sheetName.getAttribute('PartName') == '/' + sheetId2;
      });

      ///
      /// Also remove from the _xmlFiles list as we might want to create this sheet again from new starting.
      if (_xmlFiles[_xmlSheetId[sheet]] != null) {
        _xmlFiles.remove(_xmlSheetId[sheet]);
      }

      ///
      /// Maybe overkill and unsafe to do this, but works for now especially
      /// delete or renaming default sheet name (`Sheet1`),
      /// another safer method preferred
      _archive = _cloneArchive(
        _archive,
        _xmlFiles.map((k, v) {
          final encode = utf8.encode(v.toString());
          final value = ArchiveFile(k, encode.length, encode);
          return MapEntry(k, value);
        }),
        excludedFile: _xmlSheetId[sheet],
      );

      _xmlSheetId.remove(sheet);
    }

    ///
    /// remove from key = `sheet` from `_sheets`
    if (_sheets[sheet] != null) {
      ///
      /// Remove from `xl/workbook.xml`
      ///
      _xmlFiles['xl/workbook.xml']
          ?.findAllElements('sheets')
          .first
          .children
          .removeWhere((element) {
        return element.getAttribute('name') != null &&
            element.getAttribute('name').toString() == sheet;
      });

      _sheets.remove(sheet);
    }

    ///
    /// remove the cellStlye Referencing as it would be useless to have cellStyleReferenced saved
    if (_cellStyleReferenced[sheet] != null) {
      _cellStyleReferenced.remove(sheet);
    }
  }

  ///
  ///It will start setting the edited values of `sheets` into the `files` and then `exports the file`.
  ///
  List<int>? encode() {
    Save s = Save._(this, parser);
    return s._save();
  }

  /// Starts Saving the file.
  /// `On Web`
  /// ```
  /// // Call function save() to download the file
  /// var bytes = excel.save(fileName: "My_Excel_File_Name.xlsx");
  ///
  ///
  /// ```
  /// `On Android / iOS`
  ///
  /// For getting directory on Android or iOS, Use: [path_provider](https://pub.dev/packages/path_provider)
  /// ```
  /// // Call function save() to download the file
  /// var fileBytes = excel.save();
  /// var directory = await getApplicationDocumentsDirectory();
  ///
  /// File(join("$directory/output_file_name.xlsx"))
  ///   ..createSync(recursive: true)
  ///   ..writeAsBytesSync(fileBytes);
  ///
  ///```
  List<int>? save({String fileName = 'FlutterExcel.xlsx'}) {
    Save s = Save._(this, parser);
    var onValue = s._save();
    return helper.SavingHelper.saveFile(onValue, fileName);
  }

  ///
  ///returns the name of the `defaultSheet` (the sheet which opens firstly when xlsx file is opened in `excel based software`).
  ///
  String? getDefaultSheet() {
    if (_defaultSheet != null) {
      return _defaultSheet;
    } else {
      String? re = _getDefaultSheet();
      return re;
    }
  }

  ///
  ///Internal function which returns the defaultSheet-Name by reading from `workbook.xml`
  ///
  String? _getDefaultSheet() {
    Iterable<XmlElement>? elements =
        _xmlFiles['xl/workbook.xml']?.findAllElements('sheet');
    XmlElement? _sheet;
    if (elements?.isNotEmpty ?? false) {
      _sheet = elements?.first;
    }

    if (_sheet != null) {
      var defaultSheet = _sheet.getAttribute('name');
      if (defaultSheet != null) {
        return defaultSheet;
      } else {
        _damagedExcel(
            text: 'Excel sheet corrupted!! Try creating new excel file.');
      }
    }
    return null;
  }

  ///
  ///It returns `true` if the passed `sheetName` is successfully set to `default opening sheet` otherwise returns `false`.
  ///
  bool setDefaultSheet(String sheetName) {
    if (_sheetMap[sheetName] != null) {
      _defaultSheet = sheetName;
      return true;
    }
    return false;
  }

  ///
  ///Inserts an empty `column` in sheet at position = `columnIndex`.
  ///
  ///If `columnIndex == null` or `columnIndex < 0` if will not execute
  ///
  ///If the `sheet` does not exists then it will be created automatically.
  ///
  void insertColumn(String sheet, int columnIndex) {
    if (columnIndex < 0) {
      return;
    }
    _availSheet(sheet);
    _sheetMap[sheet]!.insertColumn(columnIndex);
  }

  ///
  ///If `sheet` exists and `columnIndex < maxColumns` then it removes column at index = `columnIndex`
  ///
  void removeColumn(String sheet, int columnIndex) {
    if (columnIndex >= 0 && _sheetMap[sheet] != null) {
      _sheetMap[sheet]!.removeColumn(columnIndex);
    }
  }

  ///
  ///Inserts an empty row in `sheet` at position = `rowIndex`.
  ///
  ///If `rowIndex == null` or `rowIndex < 0` if will not execute
  ///
  ///If the `sheet` does not exists then it will be created automatically.
  ///
  void insertRow(String sheet, int rowIndex) {
    if (rowIndex < 0) {
      return;
    }
    _availSheet(sheet);
    _sheetMap[sheet]!.insertRow(rowIndex);
  }

  ///
  ///If `sheet` exists and `rowIndex < maxRows` then it removes row at index = `rowIndex`
  ///
  void removeRow(String sheet, int rowIndex) {
    if (rowIndex >= 0 && _sheetMap[sheet] != null) {
      _sheetMap[sheet]!.removeRow(rowIndex);
    }
  }

  ///
  ///Appends [row] iterables just post the last filled index in the [sheet]
  ///
  ///If `sheet` does not exist then it will be automatically created.
  ///
  void appendRow(String sheet, List<CellValue?> row) {
    if (row.isEmpty) {
      return;
    }
    _availSheet(sheet);
    int targetRow = _sheetMap[sheet]!.maxRows;
    insertRowIterables(sheet, row, targetRow);
  }

  ///
  ///If `sheet` does not exist then it will be automatically created.
  ///
  ///Adds the [row] iterables in the given rowIndex = [rowIndex] in [sheet]
  ///
  ///[startingColumn] tells from where we should start putting the [row] iterables
  ///
  ///[overwriteMergedCells] when set to [true] will over-write mergedCell and does not jumps to next unqiue cell.
  ///
  ///[overwriteMergedCells] when set to [false] puts the cell value to next unique cell available by putting the value in merged cells only once and jumps to next unique cell.
  ///
  void insertRowIterables(String sheet, List<CellValue?> row, int rowIndex,
      {int startingColumn = 0, bool overwriteMergedCells = true}) {
    if (rowIndex < 0) {
      return;
    }
    _availSheet(sheet);
    _sheetMap['$sheet']!.insertRowIterables(row, rowIndex,
        startingColumn: startingColumn,
        overwriteMergedCells: overwriteMergedCells);
  }

  ///
  ///Returns the `count` of replaced `source` with `target`
  ///
  ///`source` is Pattern which allows you to pass your custom `RegExp` or a `String` providing more control over it.
  ///
  ///optional argument `first` is used to replace the number of first earlier occurrences
  ///
  ///If `first` is set to `3` then it will replace only first `3 occurrences` of the `source` with `target`.
  ///
  ///       excel.findAndReplace('MySheetName', 'sad', 'happy', first: 3);
  ///
  ///       or
  ///
  ///       var mySheet = excel['mySheetName'];
  ///       mySheet.findAndReplace('MySheetName', 'sad', 'happy', first: 3);
  ///
  ///In the above example it will replace all the occurences of `sad` with `happy` in the cells
  ///
  ///Other `options` are used to `narrow down` the `starting and ending ranges of cells`.
  ///
  int findAndReplace(String sheet, Pattern source, dynamic target,
      {int first = -1,
      int startingRow = -1,
      int endingRow = -1,
      int startingColumn = -1,
      int endingColumn = -1}) {
    int replaceCount = 0;
    if (_sheetMap[sheet] == null) return replaceCount;

    _sheetMap['$sheet']!.findAndReplace(
      source,
      target,
      first: first,
      startingRow: startingRow,
      endingRow: endingRow,
      startingColumn: startingColumn,
      endingColumn: endingColumn,
    );

    return replaceCount;
  }

  ///
  ///Make `sheet` available if it does not exist in `_sheetMap`
  ///
  void _availSheet(String sheet) {
    if (_sheetMap[sheet] == null) {
      _sheetMap[sheet] = Sheet._(this, sheet);
    }
  }

  ///
  ///Updates the contents of `sheet` of the `cellIndex: CellIndex.indexByColumnRow(0, 0);` where indexing starts from 0
  ///
  ///----or---- by `cellIndex: CellIndex.indexByString("A3");`.
  ///
  ///Styling of cell can be done by passing the CellStyle object to `cellStyle`.
  ///
  ///If `sheet` does not exist then it will be automatically created.
  ///
  void updateCell(String sheet, CellIndex cellIndex, CellValue? value,
      {CellStyle? cellStyle}) {
    _availSheet(sheet);

    _sheetMap[sheet]!.updateCell(cellIndex, value, cellStyle: cellStyle);
  }

  ///
  ///Merges the cells starting from `start` to `end`.
  ///
  ///If `custom value` is not defined then it will look for the very first available value in range `start` to `end` by searching row-wise from left to right.
  ///
  ///If `sheet` does not exist then it will be automatically created.
  ///
  void merge(String sheet, CellIndex start, CellIndex end,
      {CellValue? customValue}) {
    _availSheet(sheet);
    _sheetMap[sheet]!.merge(start, end, customValue: customValue);
  }

  ///
  ///returns an Iterable of `cell-Id` for the previously merged cell-Ids.
  ///
  List<String> getMergedCells(String sheet) {
    return List<String>.from(
        _sheetMap[sheet] != null ? _sheetMap[sheet]!.spannedItems : <String>[]);
  }

  ///
  ///unMerge the merged cells.
  ///
  ///       var sheet = 'DesiredSheet';
  ///       List<String> spannedCells = excel.getMergedCells(sheet);
  ///       var cellToUnMerge = "A1:A2";
  ///       excel.unMerge(sheet, cellToUnMerge);
  ///
  void unMerge(String sheet, String unmergeCells) {
    if (_sheetMap[sheet] != null) {
      _sheetMap[sheet]!.unMerge(unmergeCells);
    }
  }

  ///
  ///Internal function taking care of adding the `sheetName` to the `mergeChangeLook` List
  ///So that merging function will be only called on `sheetNames of mergeChangeLook`
  ///
  set _mergeChangeLookup(String value) {
    if (!_mergeChangeLook.contains(value)) {
      _mergeChangeLook.add(value);
      //_mergeChanges = true;
    }
  }

  set _rtlChangeLookup(String value) {
    if (!_rtlChangeLook.contains(value)) {
      _rtlChangeLook.add(value);
      _rtlChanges = true;
    }
  }
}

extension ImageExtension on Excel {
  void addImage({
    required String sheet,
    required Uint8List imageBytes,
    required CellIndex cellIndex,
    String? name,
  }) {
    final image = ExcelImage.from(imageBytes, name: name);

    // 1. media 폴더에 이미지 추가
    _archive.addFile(image.toArchiveFile());

    // 2. drawing.xml 파일 생성
    _createDrawingFile(sheet, image, cellIndex);

    // 3. drawing relationships 파일 생성
    _createDrawingRels(sheet, image);

    // 4. worksheet relationships 업데이트
    _updateWorksheetRels(sheet);

    // 5. Content Types에 이미지 타입 추가
    _updateContentTypes(image);
  }

  void _createDrawingRels(String sheet, ExcelImage image) {
    final relsContent =
        '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="${image.id}" 
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" 
                Target="../media/${image.name}"/>
</Relationships>''';

    _archive.addFile(ArchiveFile(
      'xl/drawings/_rels/drawing1.xml.rels',
      relsContent.length,
      utf8.encode(relsContent),
    ));
  }

  void _updateContentTypes(ExcelImage image) {
    final types =
        _xmlFiles['[Content_Types].xml']!.findAllElements('Types').first;

    // 이미지 확장자에 대한 ContentType이 없으면 추가
    if (!types.children.any((node) =>
        node is XmlElement &&
        node.getAttribute('Extension') == image.extension.substring(1))) {
      types.children.add(XmlElement(
        XmlName('Default'),
        <XmlAttribute>[
          XmlAttribute(XmlName('Extension'), image.extension.substring(1)),
          XmlAttribute(XmlName('ContentType'), image.contentType),
        ],
      ));
    }
  }

  void _createDrawingFile(String sheet, ExcelImage image, CellIndex cellIndex) {
    final drawingXml =
        '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <xdr:oneCellAnchor>
    <xdr:from>
      <xdr:col>${cellIndex.columnIndex}</xdr:col>
      <xdr:colOff>0</xdr:colOff>
      <xdr:row>${cellIndex.rowIndex}</xdr:row>
      <xdr:rowOff>0</xdr:rowOff>
    </xdr:from>
    <xdr:ext cx="${image.width * 9525}" cy="${image.height * 9525}"/>
    <xdr:pic>
      <xdr:nvPicPr>
        <xdr:cNvPr id="1" name="${image.name}"/>
        <xdr:cNvPicPr>
          <a:picLocks noChangeAspect="1"/>
        </xdr:cNvPicPr>
      </xdr:nvPicPr>
      <xdr:blipFill>
        <a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="${image.id}">
        </a:blip>
        <a:stretch>
          <a:fillRect/>
        </a:stretch>
      </xdr:blipFill>
      <xdr:spPr>
        <a:xfrm>
          <a:off x="0" y="0"/>
          <a:ext cx="${image.width * 9525}" cy="${image.height * 9525}"/>
        </a:xfrm>
        <a:prstGeom prst="rect">
          <a:avLst/>
        </a:prstGeom>
      </xdr:spPr>
    </xdr:pic>
    <xdr:clientData/>
  </xdr:oneCellAnchor>
</xdr:wsDr>''';

    _archive.addFile(ArchiveFile(
      'xl/drawings/drawing1.xml',
      drawingXml.length,
      utf8.encode(drawingXml),
    ));
  }

  void _updateWorksheetRels(String sheet) {
    // worksheet rels 파일 경로 생성
    final sheetRelsPath =
        'xl/worksheets/_rels/${_xmlSheetId[sheet]!.split('/').last}.rels';

    // 기존 rels 파일이 있는지 확인
    var relsFile = _archive.findFile(sheetRelsPath);

    if (relsFile == null) {
      // 새로운 rels 파일 생성
      final relsContent =
          '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" 
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" 
                Target="../drawings/drawing1.xml"/>
</Relationships>''';

      _archive.addFile(ArchiveFile(
        sheetRelsPath,
        relsContent.length,
        utf8.encode(relsContent),
      ));

      // worksheet XML 업데이트
      final worksheet = _xmlFiles[_xmlSheetId[sheet]]!;
      final sheetData = worksheet.findAllElements('sheetData').first;

      // drawing 요소 추가
      final drawingElement = XmlElement(
        XmlName('drawing'),
        [XmlAttribute(XmlName('r:id'), 'rId1')],
      );

      // 기존 drawing 요소가 있는지 확인하고 없으면 추가
      if (!worksheet.findAllElements('drawing').any((element) => true)) {
        // dimension 요소 다음, sheetData 이전에 drawing 요소 추가
        final dimensionElement = worksheet.findAllElements('dimension').first;
        final dimensionIndex = worksheet.children.indexOf(dimensionElement);
        worksheet.children.insert(dimensionIndex + 1, drawingElement);
      }

      // 파일 업데이트
      final sheetId = _xmlSheetId[sheet];
      if (sheetId != null) {
        _xmlFiles[sheetId] = worksheet;
      }
    }
  }
}
