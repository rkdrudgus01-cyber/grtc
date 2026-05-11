window.grtcCanReadLegacyXls = function () {
  return typeof window.XLSX !== 'undefined' && !!window.XLSX.read;
};

window.grtcReadXlsRowsFromBase64 = async function (base64) {
  if (!window.grtcCanReadLegacyXls()) {
    throw new Error('SheetJS parser is not loaded.');
  }

  const workbook = window.XLSX.read(base64, {
    type: 'base64',
    cellDates: false,
  });

  const sheetName = workbook.SheetNames && workbook.SheetNames[0];
  if (!sheetName) {
    return [];
  }

  const sheet = workbook.Sheets[sheetName];
  return window.XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    raw: false,
    defval: '',
  });
};
