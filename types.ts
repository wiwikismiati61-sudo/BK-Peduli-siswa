
export interface Attachment {
  name: string;
  type: string;
  data: string; // base64
}

export enum CaseStatus {
  BARU = 'Baru',
  PROSES = 'Proses',
  SELESAI = 'Selesai'
}

export interface CaseRecord {
  id?: number;
  tanggal: string;
  kategori_kasus: string;
  kelas: string;
  nama_siswa: string;
  guru_kelas: string;
  guru_bk: string;
  kronologi: string;
  tindak_lanjut: string;
  status: CaseStatus;
  lampiran: Attachment[];
  created_at: number;
}

export interface Student {
  id?: number;
  Nama: string;
  Kelas: string;
}

export interface Teacher {
  id?: number;
  Nama: string;
}

export interface DatabaseState {
  siswa: Student[];
  wali_kelas: Teacher[];
  guru_bk: Teacher[];
  kasus: CaseRecord[];
}

export type Page = 'dashboard' | 'master' | 'input' | 'laporan' | 'rekap';
