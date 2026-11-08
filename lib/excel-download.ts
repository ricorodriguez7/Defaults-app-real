import * as XLSX from 'xlsx';

export const MAX_EXCEL_BYTES = 4 * 1024 * 1024;
export const MAX_EXCEL_FILE_MESSAGE = 'Excel files must be under 4 MB. Split the file and try again.';

export function downloadWorkbook(workbook: XLSX.WorkBook, fileName: string) {
  const data = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
  if (data.byteLength > MAX_EXCEL_BYTES) {
    throw new Error('This export exceeds 4 MB. Apply filters or export fewer records at a time.');
  }

  const url = URL.createObjectURL(new Blob([data], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  }));
  const link = document.createElement('a');
  link.href = url;
  link.download = fileName;
  link.click();
  URL.revokeObjectURL(url);
}
