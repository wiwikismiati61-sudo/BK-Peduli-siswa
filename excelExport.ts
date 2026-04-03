import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import { CaseRecord } from './types';

const LOGO_URL = 'https://iili.io/KDFk4fI.png';

export async function exportToExcel(data: CaseRecord[], title: string = 'LAPORAN BIMBINGAN KONSELING SISWA') {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Rekapitulasi');

  // Set column widths
  worksheet.columns = [
    { header: 'NO', key: 'no', width: 5 },
    { header: 'NAMA SISWA', key: 'nama_siswa', width: 30 },
    { header: 'KELAS', key: 'kelas', width: 10 },
    { header: 'TANGGAL', key: 'tanggal', width: 15 },
    { header: 'KATEGORI', key: 'kategori', width: 20 },
    { header: 'KRONOLOGI', key: 'kronologi', width: 40 },
    { header: 'TINDAK LANJUT', key: 'tindak_lanjut', width: 30 },
    { header: 'STATUS', key: 'status', width: 12 },
  ];

  // 1. Header Section
  // Institutional Header
  worksheet.mergeCells('C1:H1');
  worksheet.getCell('C1').value = 'PEMERINTAH KOTA PASURUAN';
  worksheet.getCell('C1').font = { size: 14, bold: true };
  worksheet.getCell('C1').alignment = { horizontal: 'center' };

  worksheet.mergeCells('C2:H2');
  worksheet.getCell('C2').value = 'SMP NEGERI 7';
  worksheet.getCell('C2').font = { size: 18, bold: true };
  worksheet.getCell('C2').alignment = { horizontal: 'center' };

  worksheet.mergeCells('C3:H3');
  worksheet.getCell('C3').value = 'Jalan Simpang Slamet Riadi Nomor 2, Kota Pasuruan, Jawa Timur, 67139';
  worksheet.getCell('C3').font = { size: 10 };
  worksheet.getCell('C3').alignment = { horizontal: 'center' };

  worksheet.mergeCells('C4:H4');
  worksheet.getCell('C4').value = 'Telepon (0343) 426845';
  worksheet.getCell('C4').font = { size: 10 };
  worksheet.getCell('C4').alignment = { horizontal: 'center' };

  worksheet.mergeCells('C5:H5');
  worksheet.getCell('C5').value = 'Pos-el smp7pas@yahoo.co.id , Laman www.smpn7pasuruan.sch.id';
  worksheet.getCell('C5').font = { size: 10, italic: true };
  worksheet.getCell('C5').alignment = { horizontal: 'center' };

  // Logo
  try {
    const response = await fetch(LOGO_URL);
    const buffer = await response.arrayBuffer();
    const logoId = workbook.addImage({
      buffer: buffer,
      extension: 'png',
    });
    worksheet.addImage(logoId, {
      tl: { col: 0.5, row: 0.2 },
      ext: { width: 80, height: 80 }
    });
  } catch (error) {
    console.error('Failed to load logo', error);
  }

  // Blue divider line
  worksheet.mergeCells('A6:H6');
  const dividerCell = worksheet.getCell('A6');
  dividerCell.border = {
    bottom: { style: 'thick', color: { argb: 'FF3B82F6' } }
  };

  // Report Title
  worksheet.mergeCells('A8:H8');
  const titleCell = worksheet.getCell('A8');
  titleCell.value = title.toUpperCase();
  titleCell.font = { size: 20, bold: true, color: { argb: 'FF336699' } };
  titleCell.alignment = { horizontal: 'center' };

  // 2. Table Header
  const headerRowIndex = 10;
  const headerRow = worksheet.getRow(headerRowIndex);
  headerRow.values = ['NO', 'NAMA SISWA', 'KELAS', 'TANGGAL', 'KATEGORI', 'KRONOLOGI', 'TINDAK LANJUT', 'STATUS'];
  
  headerRow.eachCell((cell) => {
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF336699' }
    };
    cell.font = { color: { argb: 'FFFFFFFF' }, bold: true };
    cell.alignment = { horizontal: 'center', vertical: 'middle' };
    cell.border = {
      top: { style: 'thin' },
      left: { style: 'thin' },
      bottom: { style: 'thin' },
      right: { style: 'thin' }
    };
  });
  headerRow.height = 25;

  // 3. Data Rows
  data.forEach((item, index) => {
    const rowIndex = headerRowIndex + 1 + index;
    const row = worksheet.getRow(rowIndex);
    row.values = [
      index + 1,
      item.nama_siswa,
      item.kelas,
      item.tanggal,
      item.kategori_kasus,
      item.kronologi,
      item.tindak_lanjut,
      item.status
    ];

    row.eachCell((cell) => {
      cell.alignment = { vertical: 'middle', wrapText: true };
      cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
      };
      // Alternating row colors
      if (index % 2 === 1) {
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FFF0F7FF' }
        };
      }
    });
  });

  // 4. Footer Section
  const footerStartRow = headerRowIndex + data.length + 3;
  const today = new Date();
  const dateStr = `Pasuruan, ${today.toLocaleDateString('id-ID', { day: 'numeric', month: 'long', year: 'numeric' })}`;

  worksheet.getCell(`G${footerStartRow}`).value = dateStr;
  worksheet.getCell(`G${footerStartRow}`).alignment = { horizontal: 'center' };

  worksheet.getCell(`B${footerStartRow + 1}`).value = 'Mengetahui';
  worksheet.getCell(`B${footerStartRow + 1}`).alignment = { horizontal: 'center' };
  worksheet.getCell(`G${footerStartRow + 1}`).value = 'Guru BK';
  worksheet.getCell(`G${footerStartRow + 1}`).alignment = { horizontal: 'center' };

  worksheet.getCell(`B${footerStartRow + 2}`).value = 'Kepala Sekolah';
  worksheet.getCell(`B${footerStartRow + 2}`).alignment = { horizontal: 'center' };

  // Signature names (placeholders or from context if available)
  // Based on screenshot: NUR FADILAH, S.Pd and WIWIK ISMIATI, S.Pd
  worksheet.getCell(`B${footerStartRow + 6}`).value = 'NUR FADILAH, S.Pd';
  worksheet.getCell(`B${footerStartRow + 6}`).font = { bold: true, underline: true };
  worksheet.getCell(`B${footerStartRow + 6}`).alignment = { horizontal: 'center' };
  worksheet.getCell(`B${footerStartRow + 7}`).value = 'NIP. 19860410 201001 2 030';
  worksheet.getCell(`B${footerStartRow + 7}`).alignment = { horizontal: 'center' };

  worksheet.getCell(`G${footerStartRow + 6}`).value = 'WIWIK ISMIATI, S.Pd';
  worksheet.getCell(`G${footerStartRow + 6}`).font = { bold: true, underline: true };
  worksheet.getCell(`G${footerStartRow + 6}`).alignment = { horizontal: 'center' };
  worksheet.getCell(`G${footerStartRow + 7}`).value = 'NIP. 19831116 200904 2 003';
  worksheet.getCell(`G${footerStartRow + 7}`).alignment = { horizontal: 'center' };

  // Generate and save
  const buffer = await workbook.xlsx.writeBuffer();
  saveAs(new Blob([buffer]), `Rekap_BK_Siswa_${today.toISOString().split('T')[0]}.xlsx`);
}

export async function exportPivotToExcel(students: string[], pivotData: any, categories: string[]) {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Rekapitulasi Pivot');

  // Header Section (Same as above)
  worksheet.mergeCells('C1:H1');
  worksheet.getCell('C1').value = 'PEMERINTAH KOTA PASURUAN';
  worksheet.getCell('C1').font = { size: 14, bold: true };
  worksheet.getCell('C1').alignment = { horizontal: 'center' };

  worksheet.mergeCells('C2:H2');
  worksheet.getCell('C2').value = 'SMP NEGERI 7';
  worksheet.getCell('C2').font = { size: 18, bold: true };
  worksheet.getCell('C2').alignment = { horizontal: 'center' };

  worksheet.mergeCells('C3:H3');
  worksheet.getCell('C3').value = 'Jalan Simpang Slamet Riadi Nomor 2, Kota Pasuruan, Jawa Timur, 67139';
  worksheet.getCell('C3').font = { size: 10 };
  worksheet.getCell('C3').alignment = { horizontal: 'center' };

  worksheet.mergeCells('C4:H4');
  worksheet.getCell('C4').value = 'Telepon (0343) 426845';
  worksheet.getCell('C4').font = { size: 10 };
  worksheet.getCell('C4').alignment = { horizontal: 'center' };

  worksheet.mergeCells('C5:H5');
  worksheet.getCell('C5').value = 'Pos-el smp7pas@yahoo.co.id , Laman www.smpn7pasuruan.sch.id';
  worksheet.getCell('C5').font = { size: 10, italic: true };
  worksheet.getCell('C5').alignment = { horizontal: 'center' };

  try {
    const response = await fetch(LOGO_URL);
    const buffer = await response.arrayBuffer();
    const logoId = workbook.addImage({
      buffer: buffer,
      extension: 'png',
    });
    worksheet.addImage(logoId, {
      tl: { col: 0.5, row: 0.2 },
      ext: { width: 80, height: 80 }
    });
  } catch (error) {}

  worksheet.mergeCells('A6:H6');
  worksheet.getCell('A6').border = { bottom: { style: 'thick', color: { argb: 'FF3B82F6' } } };

  worksheet.mergeCells('A8:H8');
  const titleCell = worksheet.getCell('A8');
  titleCell.value = 'REKAPITULASI PENANGANAN SISWA';
  titleCell.font = { size: 20, bold: true, color: { argb: 'FF336699' } };
  titleCell.alignment = { horizontal: 'center' };

  // Table Header
  const headerRowIndex = 10;
  const headerRow = worksheet.getRow(headerRowIndex);
  const headers = ['NO', 'NAMA SISWA', ...categories, 'TOTAL'];
  headerRow.values = headers;
  
  headerRow.eachCell((cell) => {
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF336699' } };
    cell.font = { color: { argb: 'FFFFFFFF' }, bold: true };
    cell.alignment = { horizontal: 'center', vertical: 'middle' };
    cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
  });

  // Data Rows
  students.forEach((student, index) => {
    const rowIndex = headerRowIndex + 1 + index;
    const row = worksheet.getRow(rowIndex);
    const rowValues = [
      index + 1,
      student,
      ...categories.map(c => pivotData[student][c] || 0),
      Object.values(pivotData[student]).reduce((a: any, b: any) => a + b, 0)
    ];
    row.values = rowValues;
    row.eachCell((cell) => {
      cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
      if (index % 2 === 1) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF0F7FF' } };
    });
  });

  // Footer (Same as above)
  const footerStartRow = headerRowIndex + students.length + 3;
  const today = new Date();
  const dateStr = `Pasuruan, ${today.toLocaleDateString('id-ID', { day: 'numeric', month: 'long', year: 'numeric' })}`;
  worksheet.getCell(`G${footerStartRow}`).value = dateStr;
  worksheet.getCell(`G${footerStartRow}`).alignment = { horizontal: 'center' };
  worksheet.getCell(`B${footerStartRow + 1}`).value = 'Mengetahui';
  worksheet.getCell(`B${footerStartRow + 1}`).alignment = { horizontal: 'center' };
  worksheet.getCell(`G${footerStartRow + 1}`).value = 'Guru BK';
  worksheet.getCell(`G${footerStartRow + 1}`).alignment = { horizontal: 'center' };
  worksheet.getCell(`B${footerStartRow + 2}`).value = 'Kepala Sekolah';
  worksheet.getCell(`B${footerStartRow + 2}`).alignment = { horizontal: 'center' };
  worksheet.getCell(`B${footerStartRow + 6}`).value = 'NUR FADILAH, S.Pd';
  worksheet.getCell(`B${footerStartRow + 6}`).font = { bold: true, underline: true };
  worksheet.getCell(`B${footerStartRow + 6}`).alignment = { horizontal: 'center' };
  worksheet.getCell(`B${footerStartRow + 7}`).value = 'NIP. 19860410 201001 2 030';
  worksheet.getCell(`B${footerStartRow + 7}`).alignment = { horizontal: 'center' };
  worksheet.getCell(`G${footerStartRow + 6}`).value = 'WIWIK ISMIATI, S.Pd';
  worksheet.getCell(`G${footerStartRow + 6}`).font = { bold: true, underline: true };
  worksheet.getCell(`G${footerStartRow + 6}`).alignment = { horizontal: 'center' };
  worksheet.getCell(`G${footerStartRow + 7}`).value = 'NIP. 19831116 200904 2 003';
  worksheet.getCell(`G${footerStartRow + 7}`).alignment = { horizontal: 'center' };

  const buffer = await workbook.xlsx.writeBuffer();
  saveAs(new Blob([buffer]), `Rekap_Penanganan_Siswa_${today.toISOString().split('T')[0]}.xlsx`);
}
