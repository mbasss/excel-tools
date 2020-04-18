<?php
include('koneksi.php');
require '../vendor/autoload.php';
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

$spreadsheet = new Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();
$sheet->setCellValue('A1', 'No');
$sheet->setCellValue('B1', 'Id Desa');
$sheet->setCellValue('C1', 'Id Warga');
$sheet->setCellValue('D1', 'Id Bangunan');
$sheet->setCellValue('E1', 'Nomor KK');
$sheet->setCellValue('F1', 'Nomor KTP');
$sheet->setCellValue('G1', 'Nomor HP');
$sheet->setCellValue('H1', 'Nama Warga');
$sheet->setCellValue('I1', 'Jenis Kelamin');
$sheet->setCellValue('J1', 'Tempat Lahir');
$sheet->setCellValue('K1', 'Tanggal Lahir');
$sheet->setCellValue('L1', 'Hub Keluarga');
$sheet->setCellValue('M1', 'Status Nikah');
$sheet->setCellValue('N1', 'Kelengkapan Dokumen');
$sheet->setCellValue('O1', 'Tercantum Di KK Ini');
$sheet->setCellValue('P1', 'Status Hamil');
$sheet->setCellValue('Q1', 'Tempat Periksa Kehamilan');
$sheet->setCellValue('R1', 'Jenis Kontrasepsi');
$sheet->setCellValue('S1', 'jenis Cacat');
$sheet->setCellValue('T1', 'Penyakit Kronis');
$sheet->setCellValue('U1', 'Keberadaan Sekarang');
$sheet->setCellValue('V1', 'Partisipasi Sekolah');
$sheet->setCellValue('W1', 'Nama Sekolah');
$sheet->setCellValue('X1', 'Jenjang Sekolah Sekarang');
$sheet->setCellValue('Y1', 'Ijazah Tertinggi');
$sheet->setCellValue('Z1', 'Status Kerja');
$sheet->setCellValue('AA1', 'Lapangan Usaha');
$sheet->setCellValue('AB1', 'Keahlian');
$sheet->setCellValue('AC1', 'Penghasilan/Bulan');
$sheet->setCellValue('AD1', 'Kategori Sosial');
$sheet->setCellValue('AE1', 'Masalah Kesejahteraan');
$sheet->setCellValue('AF1', 'Gangguan Lingkungan');
$sheet->setCellValue('AG1', 'Bantuan Yang Diterima');
$sheet->setCellValue('AH1', 'Afiliasi Kelompok');
$sheet->setCellValue('AI1', 'Golongan Darah');
$sheet->setCellValue('AJ1', 'Agama');
$sheet->setCellValue('AK1', 'Tanggal Pendataan');
$sheet->setCellValue('AL1', 'Nama Petugas');
$sheet->setCellValue('AM1', 'Foto Diri');
$sheet->setCellValue('AN1', 'Foto KTP');
$sheet->setCellValue('AO1', 'Foto KK');
$sheet->setCellValue('AP1', 'Peran di Desa');

$query = mysqli_query($koneksi, "select * from warga");
$i = 2;
$no = 1;
while ($row = mysqli_fetch_array($query)) {
    $sheet->setCellValue('A' . $i, $no++);
    $sheet->setCellValue('B' . $i, $row['id_desa']);
    $sheet->setCellValue('C' . $i, $row['id_warga']);
    $sheet->setCellValue('D' . $i, $row['id_bangunan']);
    $sheet->setCellValue('E' . $i, $row['nomor_kk']);
    $sheet->setCellValue('F' . $i, $row['nomor_ktp']);
    $sheet->setCellValue('G' . $i, $row['nomor_hp']);
    $sheet->setCellValue('H' . $i, $row['nama_warga']);
    $sheet->setCellValue('I' . $i, $row['jenis_kelamin']);
    $sheet->setCellValue('J' . $i, $row['tempat_lahir']);
    $sheet->setCellValue('K' . $i, $row['tanggal_lahir']);
    $sheet->setCellValue('L' . $i, $row['hub_keluarga']);
    $sheet->setCellValue('M' . $i, $row['status_nikah']);
    $sheet->setCellValue('N' . $i, $row['kelengkapan_dokumen']);
    $sheet->setCellValue('O' . $i, $row['tercantum_di_kk_ini']);
    $sheet->setCellValue('P' . $i, $row['status_hamil']);
    $sheet->setCellValue('Q' . $i, $row['periksa_kehamilan_di']);
    $sheet->setCellValue('R' . $i, $row['jenis_kontrasepsi']);
    $sheet->setCellValue('S' . $i, $row['jenis_cacat']);
    $sheet->setCellValue('T' . $i, $row['penyakit_kronis']);
    $sheet->setCellValue('U' . $i, $row['keberadaan_sekarang']);
    $sheet->setCellValue('V' . $i, $row['partisipasi_sekolah']);
    $sheet->setCellValue('W' . $i, $row['nama_sekolah']);
    $sheet->setCellValue('X' . $i, $row['jenjang_sekolah_sekarang']);
    $sheet->setCellValue('Y' . $i, $row['ijazah_tertinggi']);
    $sheet->setCellValue('Z' . $i, $row['status_kerja']);
    $sheet->setCellValue('AA' . $i, $row['lap_usaha']);
    $sheet->setCellValue('AB' . $i, $row['keahlian_dimiliki']);
    $sheet->setCellValue('AC' . $i, $row['penghasilan_perbulan']);
    $sheet->setCellValue('AD' . $i, $row['kategori_sosial']);
    $sheet->setCellValue('AE' . $i, $row['masalah_kesejahteraan']);
    $sheet->setCellValue('AF' . $i, $row['gangguan_lingkungan']);
    $sheet->setCellValue('AG' . $i, $row['bantuan_yang_diterima']);
    $sheet->setCellValue('AH' . $i, $row['afiliasi_kelompok']);
    $sheet->setCellValue('AI' . $i, $row['gol_darah']);
    $sheet->setCellValue('AJ' . $i, $row['agama']);
    $sheet->setCellValue('AK' . $i, $row['tgl_pendataan']);
    $sheet->setCellValue('AL' . $i, $row['nama_petugas']);
    $sheet->setCellValue('AM' . $i, $row['foto_diri']);
    $sheet->setCellValue('AN' . $i, $row['foto_ktp']);
    $sheet->setCellValue('AO' . $i, $row['foto_kk']);
    $sheet->setCellValue('AP' . $i, $row['peran_di_desa']);
    $i++;
}

$styleArray = [
    'borders' => [
        'allBorders' => [
            'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
        ],
    ],
];
$i = $i - 1;
$sheet->getStyle('A1:AP' . $i)->applyFromArray($styleArray);


$writer = new Xlsx($spreadsheet);
$writer->save('../Data Export/Data Warga.xlsx');
echo "<script>alert('Data berhasil di Export!');history.go(-1);</script>";
// header("Location: ../index.php");
