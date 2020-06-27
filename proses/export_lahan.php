<?php
include('koneksi.php');
require '../vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

$spreadsheet = new Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();
$sheet->setCellValue('A1', 'No');
$sheet->setCellValue('B1', 'Id Desa');
$sheet->setCellValue('C1', 'Id Lahan');
$sheet->setCellValue('D1', 'Asal Pemilik');
$sheet->setCellValue('E1', 'Nama Pemilik');
$sheet->setCellValue('F1', 'Alamat Pemilik');
$sheet->setCellValue('G1', 'Id Pemilik');
$sheet->setCellValue('H1', 'Id Penanggungjawab');
$sheet->setCellValue('I1', 'afiliasi_kelompok');
$sheet->setCellValue('J1', 'Jenis Lahan');
$sheet->setCellValue('K1', 'Status Lahan');
$sheet->setCellValue('L1', 'Fungsi Lahan');
$sheet->setCellValue('M1', 'Kelengkapan Dokumen');
$sheet->setCellValue('N1', 'Kondisi Tanah');
$sheet->setCellValue('O1', 'Luas Tanaman Pertahun');
$sheet->setCellValue('P1', 'Nilai Produksi Pertahun');
$sheet->setCellValue('Q1', 'Biaya Pemupukan Pertahun');
$sheet->setCellValue('R1', 'Biaya Bibit Pertahun');
$sheet->setCellValue('S1', 'Biaya Obat Pertahun');
$sheet->setCellValue('T1', 'Biaa Lain Pertahun');
$sheet->setCellValue('U1', 'Sarana Irigasi');
$sheet->setCellValue('V1', 'Pjg Irigrasi Primer');
$sheet->setCellValue('W1', 'Pjg Irigrasi Sekunder');
$sheet->setCellValue('X1', 'Pjg Irigrasi Tersier');
$sheet->setCellValue('Y1', 'Jml Pintu Sadap');
$sheet->setCellValue('Z1', 'Jm Pintu Air');
$sheet->setCellValue('AA1', 'Fasilitas Pendukung');
$sheet->setCellValue('AB1', 'Jenis Fas Umum');
$sheet->setCellValue('AC1', 'Transportasi Terparkir');
$sheet->setCellValue('AD1', 'Jenis Irigari');
$sheet->setCellValue('AE1', 'Produk Dihasilkan');
$sheet->setCellValue('AF1', 'Jenis Ternak');
$sheet->setCellValue('AG1', 'Lahan Gembala');
$sheet->setCellValue('AH1', 'Jumlah Populasi');
$sheet->setCellValue('AI1', 'Omzet Pertahun');
$sheet->setCellValue('AJ1', 'Modal Pertahun');
$sheet->setCellValue('AK1', 'Jml Pekerja');
$sheet->setCellValue('AL1', 'Pemasaran');
$sheet->setCellValue('AM1', 'Luas Lahan');
$sheet->setCellValue('AN1', 'Kondisi Hutan');
$sheet->setCellValue('AO1', 'Gangguan Dirasakan');
$sheet->setCellValue('AP1', 'Dampak Ke Lingkungan');
$sheet->setCellValue('AQ1', 'Foto Lahan');
$sheet->setCellValue('AR1', 'Foto Fasilitas');
$sheet->setCellValue('AS1', 'Foto Produk');
$sheet->setCellValue('AT1', 'Koordinat');
$sheet->setCellValue('AU1', 'Nama Petugas');
$sheet->setCellValue('AV1', 'Tgl Pendaftaran');


$query = mysqli_query($koneksi, "select * from lahan");
$i = 2;
$no = 1;
while ($row = mysqli_fetch_array($query)) {
    $sheet->setCellValue('A' . $i, $no++);
    $sheet->setCellValue('B' . $i, $row['id_desa']);
    $sheet->setCellValue('C' . $i, $row['id_lahan']);
    $sheet->setCellValue('D' . $i, $row['asal_pemilik']);
    $sheet->setCellValue('E' . $i, $row['nama_pemilik']);
    $sheet->setCellValue('F' . $i, $row['alamat_pemilik']);
    $sheet->setCellValue('G' . $i, $row['id_pemilik']);
    $sheet->setCellValue('H' . $i, $row['id_penanggungjawab']);
    $sheet->setCellValue('I' . $i, $row['afiliasi_kelompok']);
    $sheet->setCellValue('J' . $i, $row['jenis_lahan']);
    $sheet->setCellValue('K' . $i, $row['status_lahan']);
    $sheet->setCellValue('L' . $i, $row['fungsi_lahan']);
    $sheet->setCellValue('M' . $i, $row['kelengkapan_dokumen']);
    $sheet->setCellValue('N' . $i, $row['kondisi_tanah']);
    $sheet->setCellValue('O' . $i, $row['luas_tanaman_pertahun']);
    $sheet->setCellValue('P' . $i, $row['nilai_produksi_pertahun']);
    $sheet->setCellValue('Q' . $i, $row['biaya_pemupukan_pertahun']);
    $sheet->setCellValue('R' . $i, $row['biaya_bibit_pertahun']);
    $sheet->setCellValue('S' . $i, $row['biaya_obat_pertahun']);
    $sheet->setCellValue('T' . $i, $row['biaya_lain_pertahun']);
    $sheet->setCellValue('U' . $i, $row['sarana_irigasi']);
    $sheet->setCellValue('V' . $i, $row['pjg_irigasi_primer']);
    $sheet->setCellValue('W' . $i, $row['pjg_irigasi_sekunder']);
    $sheet->setCellValue('X' . $i, $row['pjg_irigasi_tersier']);
    $sheet->setCellValue('Y' . $i, $row['jml_pintu_sadap']);
    $sheet->setCellValue('Z' . $i, $row['jm_pintu_air']);
    $sheet->setCellValue('AA' . $i, $row['fasilitas_pendukung']);
    $sheet->setCellValue('AB' . $i, $row['jenis_fas_umum']);
    $sheet->setCellValue('AC' . $i, $row['transportasi_terparkir']);
    $sheet->setCellValue('AD' . $i, $row['jenis_irigasi']);
    $sheet->setCellValue('AE' . $i, $row['produk_dihasilkan']);
    $sheet->setCellValue('AF' . $i, $row['jenis_ternak']);
    $sheet->setCellValue('AG' . $i, $row['lahan_gembala']);
    $sheet->setCellValue('AH' . $i, $row['jumlah_populasi']);
    $sheet->setCellValue('AI' . $i, $row['omzet_pertahun']);
    $sheet->setCellValue('AJ' . $i, $row['modal_pertahun']);
    $sheet->setCellValue('AK' . $i, $row['jml_pekerja']);
    $sheet->setCellValue('AL' . $i, $row['pemasaran']);
    $sheet->setCellValue('AM' . $i, $row['luas_lahan']);
    $sheet->setCellValue('AN' . $i, $row['kondisi_hutan']);
    $sheet->setCellValue('AO' . $i, $row['gangguan_dirasakan']);
    $sheet->setCellValue('AP' . $i, $row['dampak_ke_lingkungan']);
    $sheet->setCellValue('AQ' . $i, $row['foto_lahan']);
    $sheet->setCellValue('AR' . $i, $row['foto_fasilitas']);
    $sheet->setCellValue('AS' . $i, $row['foto_produk']);
    $sheet->setCellValue('AT' . $i, $row['koordinat']);
    $sheet->setCellValue('AU' . $i, $row['nama_petugas']);
    $sheet->setCellValue('AV' . $i, $row['tgl_pendataan']);
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
$sheet->getStyle('A1:AX' . $i)->applyFromArray($styleArray);

$files = glob('../Data Export/*'); // get all file names
foreach ($files as $file) { // iterate files
    if (is_file($file))
        unlink($file); // delete file
}

$file_name = "Data Warga " . date("(Y-m-d h-i-s)") . ".xlsx";
$writer = new Xlsx($spreadsheet);
$writer->save('../Data Export/' . $file_name);
$url = "http://localhost/excel/Data Export/" . $file_name;
header("Location: $url");
// echo "<script>alert(' Data berhasil di Ex p ort!');history.go(-1);</script>";
