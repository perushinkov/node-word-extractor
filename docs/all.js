const def = {
  WordDocument: {
    fib: {
      base: {//32 bytes
        wIdent: 0xA5EC, // (2 bytes): MAGIC number: 0xA5EC.
        nFib: 0X0000, // (2 bytes)
        // An unsigned integer that specifies the version number of the file format used.
        // Superseded by FibRgCswNew.nFibNew if it is present. T
        unused_4: [],
        pnNext: 0, // 2 bytes that should be 0. Else warnings for AutoText
        fDot: 0, // if doc is template (must be 0)
        flags: { // 2 bytes
          fWhichTblStm: 0b1000000000, // 10th bit
        },
        unused_20: [] // 20 unused bytes
      },
      csw: 0x000E, // number of 16-bit values in figRgW (Always 14)
      fibRgW: { // 28 bytes
        unused_26: [],
        lidFE: 0, // unsigned 16 bit. Language related, so can be ignored
      },
      cslw: 0x0016, // number of 32-bit values in figRgW (Always 22)
      fibRgLw: { // 88 bytes
        cbMac: 0x0, // Count of meaningful bytes in WordDocument stream. Ignore bytes after cbMac-th byte.
        unused_8: 0x0,
        ccpText: 0x0, // count of CPs in the "main_document"
        ccpFtn: 0x0, // count of CPs in the "footnotes_subdocument"
        ccpHdd: 0x0, // count of CPs in the "header_subdocument"
        unused_4: 0x0,
        ccpAtn: 0x0, // count of CPs in the "comment_subdocument"
        ccpEdn: 0x0, // count of CPs in the "endnote_subdocument"
        ccpTxbx: 0x0, // count of CPs in the "textbox_subdocument_of_the_main_document"
        ccpHdrTxbx: 0x0, // count of CPs in the "textbox subdocument of the header"
        unused_44: 0x0,
      },
      cbRgFcLcb: {}, // 2 bytes (The number of 64-bit values in fibRgFcLcbBlob)
      // Value of nFib => cbRgFcLcb
      // 0x00C1 0x005D
      // 0x00D9 0x006C
      // 0x0101 0x0088
      // 0x010C 0x00A4
      // 0x0112 0x00B7
      // TODO: Fill in figRgFcLcbBlob data only if a need arises`
      fibRgFcLcbBlob_c1: {  // Version 97: 93 q-words
      },
      fibRgFcLcbBlob_d9_extension: {  // Version 2000: 15 extra q-words => 108 total
      },
      fibRgFcLcbBlob_101: { // Version 2002: 0x88 or 28 extra q-words => 136 total
      },
      fibRgFcLcbBlob_10c: { // Version 2003: 0xa4 or 28 extra q-words => 164 total
      },
      fibRgFcLcbBlob_112: { // Version 2007: 0xb7 or 19 extra q-words => 183 total
      },
      cswNew: 0x0, // 2 bytes. The number of 16-bit values in fibRgCswNew
      // nFib to cswNew mapping
      // 0x00C1 0
      // 0x00D9 0x0002
      // 0x0101 0x0002
      // 0x010C 0x0002
      // 0x0112 0x0005
      fibRgCswNew: { // variable
        nFibNew: 0x00D9 || 0x0101 || 0x010C || 0x0112, // 2 bytes. Must be one of these 4 versions.
        unused_2: 0x0, // 2 bytes
        // ONLY IF nFibNew is 0x0112
        unused_6: 0x0
      },
    }
  },
};

const tableName = def.WordDocument.fib.base.flags.fWhichTblStm ? '1Table' : '0Table';
def[tableName] = {};

const nFib = def.WordDocument.fib.cswNew === 0 ?
  def.WordDocument.fib.base.nFib :
  def.WordDocument.fib.fibRgCswNew.nFibNew;

// A CP (Character Position) is a 32bit index serving as a zero-based index
module.exports = {
  parsedOle: def,
  nFib: nFib
};
