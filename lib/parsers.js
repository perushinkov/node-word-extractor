

const compressedCharMapping = {};
compressedCharMapping[0x82] = 0x201A;
compressedCharMapping[0x83] = 0x0192;
compressedCharMapping[0x84] = 0x201E;
compressedCharMapping[0x85] = 0x2026;
compressedCharMapping[0x86] = 0x2020;
compressedCharMapping[0x87] = 0x2021;
compressedCharMapping[0x88] = 0x02C6;
compressedCharMapping[0x89] = 0x2030;
compressedCharMapping[0x8A] = 0x0160;
compressedCharMapping[0x8B] = 0x2039;
compressedCharMapping[0x8C] = 0x0152;
compressedCharMapping[0x91] = 0x2018;
compressedCharMapping[0x92] = 0x2019;
compressedCharMapping[0x93] = 0x201C;
compressedCharMapping[0x94] = 0x201D;
compressedCharMapping[0x95] = 0x2022;
compressedCharMapping[0x96] = 0x2013;
compressedCharMapping[0x97] = 0x2014;
compressedCharMapping[0x98] = 0x02DC;
compressedCharMapping[0x99] = 0x2122;
compressedCharMapping[0x9A] = 0x0161;
compressedCharMapping[0x9B] = 0x203A;
compressedCharMapping[0x9C] = 0x0153;
compressedCharMapping[0x9F] = 0x0178;

function assertEquals(a, b, errorMessage) {
  if (a !== b) {
    throw Error(errorMessage);
  }
}


// function _error(msg) {
//   return {error: msg};
// }

// const baseOffsetInBytes = 0;
// fib.base.wIdent = buffer.readUInt16LE(baseOffsetInBytes);
// fib.base.nFib = buffer.readUInt16LE(baseOffsetInBytes + 2);
// if (fib.base.wIdent === 0xa5ec) {
//   return _error('fib.base.wIdent must be 0xa5ec');
// }

function makeProcessor() {
  const streams = {
    word: null,
    table: null,
    data: null,
  };
  const wordDocument = {
    base: {//32 bytes
      wIdent: null, // (2 bytes): MAGIC number: 0xA5EC.
      nFib: null, // (2 bytes)
      // An unsigned integer that specifies the version number of the file format used.
      // Superseded by FibRgCswNew.nFibNew if it is present. T
      unused_4: [],
      pnNext: null, // 2 bytes that should be 0. Else warnings for AutoText
      flags: { // 2 bytes
        fDot: null, // 0x0001 Set if this document is a template
        fComplex: null, // 0x0004 If complex fast-saved format
        fHasPic: null, // 0x0008 If contains at least one image
        fWhichTblStm: null,// 0x0200 When 0, this fib refers to the table stream named ―0Table‖, when 1, this fib refers to the table stream named ―1Table‖.
        fFarEast: null
      },
      unused_20: null, // 20 unused bytes
    },
    csw: null, // 0x000E, number of 16-bit values in figRgW (Always 14)
    fibRgW: { // 28 bytes
      unused_26: null,
      lidFE: null, // unsigned 16 bit. Language related, so can be ignored
    },
    cslw: null, // 0x0016, number of 32-bit values in figRgW (Always 22)
    fibRgLw: { // 88 bytes
      cbMac: 0x0, // Count of meaningful bytes in WordDocument stream. Ignore bytes after cbMac-th byte.
      // unused 8
      ccpText: 0x0, // count of CPs in the "main_document"
      ccpFtn: 0x0, // count of CPs in the "footnotes_subdocument"
      ccpHdd: 0x0, // count of CPs in the "header_subdocument"
      // unused 4
      ccpAtn: 0x0, // count of CPs in the "comment_subdocument"
      ccpEdn: 0x0, // count of CPs in the "endnote_subdocument"
      ccpTxbx: 0x0, // count of CPs in the "textbox_subdocument_of_the_main_document"
      ccpHdrTxbx: 0x0, // count of CPs in the "textbox subdocument of the header"
      // unused 4
      pnChpFirst: 0x0, // The page number of the lowest numbered page in the document that records CHPX FKP information
      cpnBteChp: 0x0, // Count of CHPX FKPs recorded in file.
      // unused 4
      pnPapFirst: 0x0, // The page number of the lowest numbered page in the document that records PAPX FKP information
      cpnBtePap: 0x0, // Count of PAPX FKPs recorded in file
      // unused 4
      pnLvcFirst: 0x0, // The page number of the lowest numbered page in the document that records LVC FKP information
      cpnBteLvc: 0x0, // Count of LVC FKPS recorded in file
      // unused 8
    },
    cbRgFcLcb: 0x0, // 2 bytes (The number of 64-bit values in fibRgFcLcbBlob)
    // Value of nFib => cbRgFcLcb
    // 0x00C1 0x005D
    // 0x00D9 0x006C
    // 0x0101 0x0088
    // 0x010C 0x00A4
    // 0x0112 0x00B7
    fibRgFcLcbBlob: {
      fcStshfOrig: 0x0,  // file offset of original STSH in table stream
      lcbStshfOrig: 0x0, // sizeof(original STSH) in bytes
      fcStshf: 0x0,  // file offset of STSH in table stream
      lcbStshf: 0x0, // sizeof(STSH) in bytes
      fcPlcffndRef: 0x0,  // Offset in table stream of footnote reference PLCF of FRD structures. CPs in     PLC are relative to main document text stream and give location of footnote references.
      lcbPlcffndRef: 0x0, // Footnote reference size in bytes
      fcPlcffndTxt: 0x0,  // Offset in table stream of footnote text PLC. CPs in PLC are relative to footnote subdocument text stream and give location of beginnings of footnote text for corresponding references recorded in plcffndRef. No structure is stored in this plc. There will just be n+1 FC entries in this PLC when there are n footnotes
      lcbPlcffndTxt: 0x0, // Count of bytes of footnote text PLC. == 0 if no footnotes defined in document
      fcPlcfandRef: 0x0,  // IGNORING Annotations
      lcbPlcfandRef: 0x0, // IGNORING Annotations
      fcPlcfandTxt: 0x0,  // IGNORING Annotations
      lcbPlcfandTxt: 0x0, // IGNORING Annotations
      fcPlcfsed: 0x0,  // Offset in table stream of section descriptor SED PLC. CPs in PLC are relative to main document.
      lcbPlcfsed: 0x0, // Count of bytes of section descriptor PLC
      fcPlcfhdd: 0x0,  // IGNORING Headers
      lcbPlcfhdd: 0x0, // IGNORING Headers
      fcPlcfbteChpx: 0x0,   // Offset in table stream of character property bin table.PLC. FCs in PLC are file offsets in the main stream. Describes text of main document and all subdocuments.
      lcbPlcfbteChpx: 0x0,  // Count of bytes of character property bin table PLC
      fcPlcfbtePapx: 0x0,   // Offset in table stream of paragraph property bin table PLC. FCs in PLC are file offsets in the main stream. Describes text of main document and all subdocuments.
      lcbPlcfbtePapx: 0x0,  // Count of bytes of paragraph property bin table PLC
      fcSttbfffn: 0x0,  // Offset in table stream of font information STTBF. The sttbfffn is a STTBF where is string is actually an FFN structure. The nth entry in the STTBF describes the font that will be displayed when the chp.ftc for text is equal to n. See the FFN file structure definition.
      lcbSttbfffn: 0x0, // Count of bytes in sttbfffn
      fcClx: 0x0,  // Offset in table stream of beginning of information for complex files. Consists of an encoding of all of the prms quoted by the document followed by the plcpcd (piece table) for the document.
      lcbClx: 0x0, // Count of bytes of complex file information == 0 if file is non-complex.
      fcPlcfLst: 0x0,  // List format              // TODO: determine if needed
      lcbPlcfLst: 0x0, // List format              // TODO: determine if needed
      fcPlfLfo: 0x0,   // List format override              // TODO: determine if needed
      lcbPlfLfo: 0x0,  // List format override              // TODO: determine if needed
      fcPlcfbteLvc: 0x0, // Offset in table stream of character(list level, prolly typo) property bin table.PLC. FCs in PLC are file offsets. Describes text of main document and all subdocuments.
      lcbPlcfbteLvc: 0x0, // Byte count
      fcPlcfpgp: 0x0,  // Paragraph Group Properties This is undocumented HTML DIV (paragraph group) formatting information
      lcbPlcfpgp: 0x0, //
    },
    cswNew: 0x0, // 2 bytes. The number of 16-bit values in fibRgCswNew
    // nFib to cswNew mapping
    // 0x00C1 0
    // 0x00D9 0x0002
    // 0x0101 0x0002
    // 0x010C 0x0002
    // 0x0112 0x0005
    fibRgCswNew: { // variable
      nFibNew: null, // 0x00D9 || 0x0101 || 0x010C || 0x0112, 2 bytes. Must be one of these 4 versions.
      // The actual nFib, moved here because some readers assumed they couldn't read any format with nFib > some constant
    },
  };

  const tableX = {
    cursor: 0,
    clx: {
      prcs: [],
      pcdt: {
        cps: [],
        pcds: []
      }
    }
  };

  let currentPrc;
  let currentSprm;
  let currentSprmOperand;
  let currentPcd;
  let currentFcCompressed;
  let currentPrm;
  let unicode16Chars;

  function readFIB() {
    // Must be 0xA5EC
    wordDocument.base.wIdent = streams.word.readUInt16LE(0);
    assertEquals(wordDocument.base.wIdent, 0xA5EC, 'Magic number incorrect!');
    wordDocument.base.nFib = streams.word.readUInt16LE(2);
    wordDocument.base.pnNext = streams.word.readUInt16LE(8);

    const flags = streams.word.readUInt16LE(10);
    wordDocument.base.flags.fDot = flags & 0x0001;
    wordDocument.base.flags.fComplex = flags & 0x0004;
    wordDocument.base.flags.fHasPic = flags & 0x0008;
    wordDocument.base.flags.fWhichTblStm = flags & 0x0200;
    wordDocument.base.flags.fFarEast = flags & 0x4000;

    wordDocument.csw = streams.word.readUInt16LE(32);
    assertEquals(wordDocument.csw, 14, 'csw must be 14');
    wordDocument.fibRgW.lidFE = streams.word.readUInt16LE(60);

    wordDocument.cslw = streams.word.readUInt16LE(62);
    assertEquals(wordDocument.cslw, 22, 'cslw must be 22');

    wordDocument.fibRgLw.cbMac = streams.word.readUInt32LE(64);
    wordDocument.fibRgLw.ccpText = streams.word.readUInt32LE(76);
    wordDocument.fibRgLw.ccpFtn = streams.word.readUInt32LE(80);
    wordDocument.fibRgLw.ccpHdd = streams.word.readUInt32LE(84);
    wordDocument.fibRgLw.ccpAtn = streams.word.readUInt32LE(92);
    wordDocument.fibRgLw.ccpEdn = streams.word.readUInt32LE(96);
    wordDocument.fibRgLw.ccpTxbx = streams.word.readUInt32LE(100);
    wordDocument.fibRgLw.ccpHdrTxbx = streams.word.readUInt32LE(104);

    wordDocument.fibRgLw.pnChpFirst = streams.word.readUInt32LE(112);
    wordDocument.fibRgLw.cpnBteChp = streams.word.readUInt32LE(116);
    wordDocument.fibRgLw.pnPapFirst = streams.word.readUInt32LE(124);
    wordDocument.fibRgLw.cpnBtePap = streams.word.readUInt32LE(128);
    wordDocument.fibRgLw.pnLvcFirst = streams.word.readUInt32LE(136);
    wordDocument.fibRgLw.cpnBteLvc = streams.word.readUInt32LE(140);

    wordDocument.cbRgFcLcb = streams.word.readUInt16LE(152);

    wordDocument.fibRgFcLcbBlob.fcStshfOrig = streams.word.readUInt32LE(154);
    wordDocument.fibRgFcLcbBlob.lcbStshfOrig = streams.word.readUInt32LE(158);

    wordDocument.fibRgFcLcbBlob.fcStshf = streams.word.readUInt32LE(162);
    wordDocument.fibRgFcLcbBlob.lcbStshf = streams.word.readUInt32LE(166);

    wordDocument.fibRgFcLcbBlob.fcPlcffndRef = streams.word.readUInt32LE(170);
    wordDocument.fibRgFcLcbBlob.lcbPlcffndRef = streams.word.readUInt32LE(174);

    wordDocument.fibRgFcLcbBlob.fcPlcffndTxt = streams.word.readUInt32LE(178);
    wordDocument.fibRgFcLcbBlob.lcbPlcffndTxt = streams.word.readUInt32LE(182);

    wordDocument.fibRgFcLcbBlob.fcPlcfsed = streams.word.readUInt32LE(202);
    wordDocument.fibRgFcLcbBlob.lcbPlcfsed = streams.word.readUInt32LE(206);

    wordDocument.fibRgFcLcbBlob.fcPlcfbteChpx = streams.word.readUInt32LE(250);
    wordDocument.fibRgFcLcbBlob.lcbPlcfbteChpx = streams.word.readUInt32LE(254);

    wordDocument.fibRgFcLcbBlob.fcPlcfbtePapx = streams.word.readUInt32LE(258);
    wordDocument.fibRgFcLcbBlob.lcbPlcfbtePapx = streams.word.readUInt32LE(262);

    wordDocument.fibRgFcLcbBlob.fcSttbfffn = streams.word.readUInt32LE(274);
    wordDocument.fibRgFcLcbBlob.lcbSttbfffn = streams.word.readUInt32LE(278);

    wordDocument.fibRgFcLcbBlob.fcClx = streams.word.readUInt32LE(418);
    wordDocument.fibRgFcLcbBlob.lcbClx = streams.word.readUInt32LE(422);

    wordDocument.fibRgFcLcbBlob.fcPlcfbteLvc = streams.word.readUInt32LE(842);
    wordDocument.fibRgFcLcbBlob.lcbPlcfbteLvc = streams.word.readUInt32LE(846);

    wordDocument.cswNew = streams.word.readUInt16LE(1466);
    wordDocument.fibRgCswNew.nFibNew = streams.word.readUInt16LE(1468);
  }

  function readSprm() {
    const sprm = streams.table.readUInt16LE(tableX.cursor);
    tableX.cursor += 2;
    currentSprm = {
      ispmd: sprm & 0x01FF,
      fSpec: (sprm / 512) & 0x0001,
      sgc: (sprm / 1024) & 0x0007,
      spra: sprm / 8192
    };
  }

  function readSprmOperand() {
    let byteLength;
    switch (currentSprm.spra) {
      case 0:
      case 1:
        byteLength = 1;
        break;
      case 2:
      case 4:
      case 5:
        byteLength = 2;
        break;
      case 3:
        byteLength = 4;
        break;
      case 7:
        byteLength = 3;
        break;
      case 6:
        if (currentSprm.ispmd === 0x08) { // sprmTDefTable
          byteLength = streams.table.readUint16LE(tableX.cursor) + 1;
        } else if (currentSprm.ispmd === 0x15) { // sprmPChgTabs
          const byteCount = streams.table.readUInt8(tableX.cursor);
          if (byteCount < 255) {
            byteLength = byteCount + 1;
          } else { // some of the brightest examples of a wonderfully designed binary format here
            byteLength = 1;
            byteLength += streams.table.readUInt8(tableX.cursor + byteLength) * 4 + 1;
            byteLength += streams.table.readUInt8(tableX.cursor + byteLength) * 3 + 1;
          }
        } else {
          byteLength = streams.table.readUInt8(tableX.cursor) + 1;
        }
        return;
      default:
        throw Error('unexpected spra value!');
    }
    currentSprmOperand = streams.table.slice(tableX.cursor, byteLength);
    tableX.cursor += byteLength;
  }

  function readPrl() {
    readSprm();
    readSprmOperand();
    currentPrc.prls.push({
      sprm: currentSprm,
      operand: currentSprmOperand
    });
  }

  function readPrcData() {
    const prlEntriesCount = streams.data.readUInt16LE(tableX.cursor);
    tableX.cursor += 2;
    currentPrc.prls = [];

    for (let currentPrl = 0; currentPrl < prlEntriesCount; currentPrl++) {
      readPrl();
    }
  }

  function readPrc() {
    if (streams.table.readUInt8(tableX.cursor) !== 0x01) {
      return false;
    }
    tableX.cursor++;
    currentPrc = {};
    readPrcData();
    return true;
  }

  function readFcCompressed() {
    const fcCompressed = streams.table.readUInt32LE(tableX.cursor);
    tableX.cursor += 4;

    currentFcCompressed = {
      fc: fcCompressed & 0x3FFFFFFF,
      fCompressed: fcCompressed & 0x40000000,
      r1: fcCompressed & 0x80000000
    };
    if (currentFcCompressed.r1 !== 0) {
      throw Error('Each fcCompressed must have r1 of 0.');
    }
  }

  function readPrm() {
    const prm = streams.table.readUInt16LE(tableX.cursor);
    tableX.cursor += 2;
    currentPrm = {
      fComplex: prm & 0x0001,
      isprm_0: prm & 0xFE,
      value_0: prm & 0xFF00,
      value_1: prm & 0xFFFE,
    };
  }

  function readPcd() {
    const flags = streams.table.readUInt16LE(tableX.cursor);
    tableX.cursor += 2;

    readFcCompressed();
    readPrm();

    currentPcd = {
      fNoParaLast: flags & 0x0001, // if true, text must not contain \p
      fcCompressed: currentFcCompressed,
      prm: currentPrm
    };
  }

  function readPcdt() {
    if (streams.table.readUInt8(tableX.cursor) !== 0x02) {
      throw Error('Pcdt must follow after prc* in clx');
    }
    tableX.cursor++;

    tableX.clx.pcdt.byteCount = streams.table.readUInt32LE(tableX.cursor);
    tableX.cursor += 4;

    const plcPcdSize = tableX.clx.pcdt.byteCount;
    const pcdSize = 8;
    const pcdCount = (plcPcdSize - 4) / (4 + pcdSize);

    // plcPcd
    tableX.clx.pcdt.cps = [];
    for (let cpIndex = 0; cpIndex < pcdCount + 1; cpIndex++) {
      const cp = streams.table.readUInt32LE(tableX.cursor);
      tableX.cursor += 4;
      tableX.clx.pcdt.cps.push(cp);
    }

    tableX.clx.pcdt.pcds = [];
    for (let pcdIndex = 0; pcdIndex < pcdCount; pcdIndex++) {
      readPcd();
      tableX.clx.pcdt.pcds.push(currentPcd);
    }
  }

  function readCLX() {
    tableX.cursor = wordDocument.fibRgFcLcbBlob.fcClx;
    while (readPrc()) {
      tableX.clx.prcs.push(currentPrc);
    }
    readPcdt();
  }

  function compressedCharToUnicode(uint8Char) {
    const mapping = compressedCharMapping[uint8Char];
    return mapping ? mapping : uint8Char;
  }

  function extractPcdText(pcd, cpStart, cpEnd) {
    const fcCompressed = pcd.fcCompressed;
    const fCompressed = fcCompressed.fCompressed;

    if (fCompressed === 0) {
      // Unicode

      // This works?
      // streams.word.slice(tableXLocation, (cpEnd - cpStart) * 2).toString('utf16le ');
      for (let currentCp = 0; currentCp < cpEnd - cpStart; currentCp++) {
        const u16Char = streams.word.readUInt16LE(currentCp * 2 + fcCompressed.fc);
        unicode16Chars.push(u16Char);
      }
    } else {
      for (let currentCp = 0; currentCp < cpEnd - cpStart; currentCp++) {
        const u8Char = streams.word.readUInt16LE(currentCp + fcCompressed.fc/2);
        unicode16Chars.push(compressedCharToUnicode(u8Char));
      }
    }
    // TODO:
  }
  // This method gets all the text and converts it from
  // mixed representation to unicode

  function retrieveText() {
    unicode16Chars = [];
    const pcdt = tableX.clx.pcdt;

    for (let pieceIndex = 0; pieceIndex < pcdt.cps.length - 1; pieceIndex++) {
      extractPcdText(pcdt.pcds[pieceIndex], pcdt.cps[pieceIndex], pcdt.cps[pieceIndex + 1]);
    }
  }

  function readPlcBtePapx() {
    // TODO:
  }
  return function(docStream, table0Stream, table1Stream, dataStream) {
    streams.word = docStream;
    streams.data = dataStream;
    readFIB();
    streams.table = wordDocument.base.flags.fWhichTblStm ? table1Stream : table0Stream;
    readCLX();

    retrieveText();
    readPlcBtePapx();

    return {
      table: wordDocument.base.flags.fWhichTblStm + 'Table',
      complex: wordDocument.base.flags.fComplex,
      lcbClx: wordDocument.fibRgFcLcbBlob.lcbClx,
      version: {
        OLD: wordDocument.base.nFib,
        NEW: wordDocument.fibRgCswNew.nFibNew
      },
      actualData: tableX
    };
  };
}

function processDocStreams(streamBuffers) {
  const wordDocumentStream = streamBuffers[0];
  const table0Stream = streamBuffers[1];
  const table1Stream = streamBuffers[2];
  const dataStream = streamBuffers[3];
  return makeProcessor()(wordDocumentStream, table0Stream, table1Stream, dataStream);
}

module.exports = {
  processDocStreams: processDocStreams,
};

