(function () {
  "use strict";

  function normalizeCell(value) {
    if (value === undefined || value === null) return "";
    if (value instanceof Date) return value.toISOString();
    return String(value).trim();
  }

  function selectSheet(workbook) {
    for (const name of workbook.SheetNames || []) {
      const sheet = workbook.Sheets[name];
      if (sheet && sheet["!ref"]) return sheet;
    }
    const firstName = workbook.SheetNames && workbook.SheetNames[0];
    return firstName ? workbook.Sheets[firstName] : null;
  }

  window.grtcCanReadLegacyXls = function () {
    return !!(window.XLSX && window.XLSX.read && window.XLSX.utils);
  };

  window.grtcReadXlsRowsFromBase64 = function (base64) {
    return new Promise(function (resolve, reject) {
      setTimeout(function () {
        try {
          if (!window.grtcCanReadLegacyXls()) {
            throw new Error("XLS parser is not loaded");
          }
          const workbook = window.XLSX.read(base64, {
            type: "base64",
            raw: false,
            cellDates: false,
          });
          const sheet = selectSheet(workbook);
          if (!sheet) {
            throw new Error("No worksheet found");
          }
          const rows = window.XLSX.utils.sheet_to_json(sheet, {
            header: 1,
            raw: false,
            defval: "",
            blankrows: false,
          });
          resolve(rows.map(function (row) {
            return row.map(normalizeCell);
          }));
        } catch (error) {
          reject(error && error.message ? error.message : String(error));
        }
      }, 0);
    });
  };
})();
