<?php
include('koneksi.php');
require '../vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

$spreadsheet = new Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();
$sheet->setCellValue('A1', 'No');
$sheet->setCellValue('B1', 'Id Desa');
$sheet->setCellValue('C1', 'Id Bangunan');
$sheet->setCellValue('D1', 'Asal Pemilik');
$sheet->setCellValue('E1', 'Fungsi Bangunan');
$sheet->setCellValue('F1', 'Layanan Kesehatan');
$sheet->setCellValue('G1', 'Jml Dokter');
$sheet->setCellValue('H1', 'Jml Dokter');
$sheet->setCellValue('I1', 'Jml Bidan');
$sheet->setCellValue('J1', 'Jml Perawat');
$sheet->setCellValue('K1', 'Jenis Lembaga Keamanan');
$sheet->setCellValue('L1', 'Jml Linmas');
$sheet->setCellValue('M1', 'Jml Satgas Linmas');
$sheet->setCellValue('N1', 'Jml Satpam Swakarsa');
$sheet->setCellValue('O1', 'Nama Induk Satpam Swakarsa');
$sheet->setCellValue('P1', 'Kepemilikan Satpam Swakarsa');
$sheet->setCellValue('Q1', 'Jml Mtra TNI');
$sheet->setCellValue('R1', 'Jml Kegiatan TNI');
$sheet->setCellValue('S1', 'Jml Mtra Polri');
$sheet->setCellValue('T1', 'Jml Kegiatan Polri');
$sheet->setCellValue('U1', 'Kelengkapan Lbg Adat');
$sheet->setCellValue('V1', 'Jenis Lbg Pendidikan');
$sheet->setCellValue('W1', 'Status Lbg Pendidikan');
$sheet->setCellValue('X1', 'Jml Pengajar');
$sheet->setCellValue('Y1', 'Jml SIswa');
$sheet->setCellValue('Z1', 'Jenis Usaha');
$sheet->setCellValue('AA1', 'Jenis Angkutan');
$sheet->setCellValue('AB1', 'Kapasitas Angkut Orang');
$sheet->setCellValue('AC1', 'Kapasitas Angkut Barang');
$sheet->setCellValue('AD1', 'Trayek');
$sheet->setCellValue('AE1', 'Nama Pemilik');
$sheet->setCellValue('AF1', 'Alamat Pemilik');
$sheet->setCellValue('AG1', 'Id Kelompok');
$sheet->setCellValue('AH1', 'Jenis Kepemilikan');
$sheet->setCellValue('AI1', 'Id Fasilitas Pendukung');
$sheet->setCellValue('AJ1', 'Omzet Pertahun');
$sheet->setCellValue('AK1', 'Modal Pertahun');
$sheet->setCellValue('AL1', 'Jml Pekerja');
$sheet->setCellValue('AM1', 'Produk Dihasilkan');
$sheet->setCellValue('AN1', 'Pemasaran');
$sheet->setCellValue('AO1', 'Kelengkapan Izin');
$sheet->setCellValue('AP1', 'Gangguan Dari Lingkungan');
$sheet->setCellValue('AQ1', 'Dampak Ke Lingkungan');
$sheet->setCellValue('AR1', 'Id Warga');
$sheet->setCellValue('AS1', 'Dusun');
$sheet->setCellValue('AT1', 'RT');
$sheet->setCellValue('AU1', 'RW');
$sheet->setCellValue('AV1', 'Jalan');
$sheet->setCellValue('AW1', 'Nomor Rumah');
$sheet->setCellValue('AX1', 'Jml Anggota Rumah');
$sheet->setCellValue('AY1', 'Jml KK Dirumah');
$sheet->setCellValue('AZ1', 'Status Rumah');
$sheet->setCellValue('BA1', 'Status Lahan Tinggal');
$sheet->setCellValue('BB1', 'Luas Lantai');
$sheet->setCellValue('BC1', 'Jenis Lantai Terluas');
$sheet->setCellValue('BD1', 'Jenis Dinding Terluas');
$sheet->setCellValue('BE1', 'Kondisi Dinding Terluas');
$sheet->setCellValue('BF1', 'Jenis Atap Terluas');
$sheet->setCellValue('BG1', 'Kondisi Atap Terluas');
$sheet->setCellValue('BH1', 'Jml Kamar Tidur');
$sheet->setCellValue('BI1', 'Sumber Air Minum');
$sheet->setCellValue('BJ1', 'Id Pelanggan Air');
$sheet->setCellValue('BK1', 'Cara Memperoleh Air Minum');
$sheet->setCellValue('BL1', 'Sumber Penerangan Utama');
$sheet->setCellValue('BM1', 'Daya Listrik Terpasang');
$sheet->setCellValue('BN1', 'Nomor Rek Listrik');
$sheet->setCellValue('BO1', 'Energi Untuk Memasak');
$sheet->setCellValue('BP1', 'Id Gas Pipa');
$sheet->setCellValue('BQ1', 'Fasilitas BAB');
$sheet->setCellValue('BR1', 'Jenis Kloset');
$sheet->setCellValue('BS1', 'Pembuangan Tinja');
$sheet->setCellValue('BT1', 'Perlengkapan Rumah');
$sheet->setCellValue('BU1', 'Luas Lahan');
$sheet->setCellValue('BV1', 'Foto Pemilik');
$sheet->setCellValue('BW1', 'Foto Depan');
$sheet->setCellValue('BX1', 'Foto IMB');
$sheet->setCellValue('BY1', 'Foto PBB');
$sheet->setCellValue('BZ1', 'Koordinat');
$sheet->setCellValue('CA1', 'Tgl Pendataan');
$sheet->setCellValue('CB1', 'Nama Petugas');
$sheet->setCellValue('CC1', 'Tgl Pemeriksaan');
$sheet->setCellValue('CD1', 'Nama Pemeriksa');
$sheet->setCellValue('CE1', 'Hasil Verivali');
$sheet->setCellValue('CF1', 'Ttd Verivali');
$sheet->setCellValue('CG1', 'Ttd Pemeriksa');
$sheet->setCellValue('CH1', 'Keluhan');
$sheet->setCellValue('CI1', 'Cataan Tambahan');


$query = mysqli_query($koneksi, "select * from bangunan");
$i = 2;
$no = 1;
while ($row = mysqli_fetch_array($query)) {
    $sheet->setCellValue('A' . $i, $no++);
    $sheet->setCellValue('B' . $i, $row['id_desa']);
    $sheet->setCellValue('C' . $i, $row['id_bangunan']);
    $sheet->setCellValue('D' . $i, $row['asal_pemilik']);
    $sheet->setCellValue('E' . $i, $row['fungsi_bangunan']);
    $sheet->setCellValue('F' . $i, $row['layanan_kesehatan']);
    $sheet->setCellValue('G' . $i, $row['jml_dokter']);
    $sheet->setCellValue('H' . $i, $row['jml_bidan']);
    $sheet->setCellValue('I' . $i, $row['jml_perawat']);
    $sheet->setCellValue('J' . $i, $row['jenis_lbg_keamanan']);
    $sheet->setCellValue('K' . $i, $row['jml_linmas']);
    $sheet->setCellValue('L' . $i, $row['jml_satgas_linmas']);
    $sheet->setCellValue('M' . $i, $row['jml_satpam_swakarsa']);
    $sheet->setCellValue('N' . $i, $row['namainduk_satpam_swakarsa']);
    $sheet->setCellValue('O' . $i, $row['kepemilikan_satpam_swakarsa']);
    $sheet->setCellValue('P' . $i, $row['jml_mitra_tni']);
    $sheet->setCellValue('Q' . $i, $row['jml_kegiatan_tni']);
    $sheet->setCellValue('R' . $i, $row['jml_mitra_polri']);
    $sheet->setCellValue('S' . $i, $row['jml_kegiatan_polri']);
    $sheet->setCellValue('T' . $i, $row['kelengkapan_lbg_adat']);
    $sheet->setCellValue('U' . $i, $row['jenis_lbg_pendidikan']);
    $sheet->setCellValue('V' . $i, $row['status_lbg_pendidikan']);
    $sheet->setCellValue('W' . $i, $row['jml_pengajar']);
    $sheet->setCellValue('X' . $i, $row['jml_siswa']);
    $sheet->setCellValue('Y' . $i, $row['jenis_usaha']);
    $sheet->setCellValue('Z' . $i, $row['jenis_angkutan']);
    $sheet->setCellValue('AA' . $i, $row['kapasitas_angkut_orang']);
    $sheet->setCellValue('AB' . $i, $row['kapasitas_angkut_barang']);
    $sheet->setCellValue('AC' . $i, $row['trayek']);
    $sheet->setCellValue('AD' . $i, $row['nama_pemilik']);
    $sheet->setCellValue('AE' . $i, $row['alamat_pemilik']);
    $sheet->setCellValue('AF' . $i, $row['id_kelompok']);
    $sheet->setCellValue('AG' . $i, $row['jenis_kepemilikan']);
    $sheet->setCellValue('AH' . $i, $row['id_fasilitas_pendukung']);
    $sheet->setCellValue('AI' . $i, $row['omzet_pertahun']);
    $sheet->setCellValue('AJ' . $i, $row['modal_pertahun']);
    $sheet->setCellValue('AK' . $i, $row['jml_pekerja']);
    $sheet->setCellValue('AL' . $i, $row['produk_dihasilkan']);
    $sheet->setCellValue('AM' . $i, $row['pemasaran']);
    $sheet->setCellValue('AN' . $i, $row['kelengkapan_izin']);
    $sheet->setCellValue('AO' . $i, $row['gangguan_dari_lingkungan']);
    $sheet->setCellValue('AP' . $i, $row['dampak_ke_lingkungan']);
    $sheet->setCellValue('AQ' . $i, $row['id_warga']);
    $sheet->setCellValue('AR' . $i, $row['dusun']);
    $sheet->setCellValue('AS' . $i, $row['rt']);
    $sheet->setCellValue('AT' . $i, $row['rw']);
    $sheet->setCellValue('AU' . $i, $row['jalan']);
    $sheet->setCellValue('AV' . $i, $row['nomor_rumah']);
    $sheet->setCellValue('AW' . $i, $row['jml_anggota_rumah']);
    $sheet->setCellValue('AX' . $i, $row['jumlah_kk_dirumah']);
    $sheet->setCellValue('AY' . $i, $row['status_rumah']);
    $sheet->setCellValue('AZ' . $i, $row['status_lahan_tinggal']);
    $sheet->setCellValue('BA' . $i, $row['luas_lantai']);
    $sheet->setCellValue('BB' . $i, $row['jenis_lantai_terluas']);
    $sheet->setCellValue('BC' . $i, $row['jenis_dinding_terluas']);
    $sheet->setCellValue('BD' . $i, $row['kondisi_dinding_terluas']);
    $sheet->setCellValue('BE' . $i, $row['jenis_atap_terluas']);
    $sheet->setCellValue('BF' . $i, $row['kondisi_atap_terluas']);
    $sheet->setCellValue('BG' . $i, $row['jml_kamar_tidur']);
    $sheet->setCellValue('BH' . $i, $row['sumber_air_minum']);
    $sheet->setCellValue('BI' . $i, $row['id_pelanggan_air']);
    $sheet->setCellValue('BJ' . $i, $row['cara_memperoleh_airminum']);
    $sheet->setCellValue('BK' . $i, $row['sumber_peneranga_utama']);
    $sheet->setCellValue('BL' . $i, $row['daya_listrik_terpasang']);
    $sheet->setCellValue('BM' . $i, $row['nomor_rek_listrik']);
    $sheet->setCellValue('BN' . $i, $row['energi_untuk_memasak']);
    $sheet->setCellValue('BO' . $i, $row['id_gas_pipa']);
    $sheet->setCellValue('BP' . $i, $row['fasilitas_bab']);
    $sheet->setCellValue('BQ' . $i, $row['jenis_kloset']);
    $sheet->setCellValue('BR' . $i, $row['pembuangan_tinja']);
    $sheet->setCellValue('BS' . $i, $row['perlengkapan_rumah']);
    $sheet->setCellValue('BT' . $i, $row['luas_lahan']);
    $sheet->setCellValue('BU' . $i, $row['foto_pemilik']);
    $sheet->setCellValue('BV' . $i, $row['foto_depan']);
    $sheet->setCellValue('BW' . $i, $row['foto_imb']);
    $sheet->setCellValue('BX' . $i, $row['foto_pbb']);
    $sheet->setCellValue('BY' . $i, $row['foto_lampu_jalan']);
    $sheet->setCellValue('BZ' . $i, $row['koordinat']);
    $sheet->setCellValue('CA' . $i, $row['tgl_pendataan']);
    $sheet->setCellValue('CB' . $i, $row['nama_petugas']);
    $sheet->setCellValue('CC' . $i, $row['tgl_pemeriksaan']);
    $sheet->setCellValue('CD' . $i, $row['nama_pemeriksa']);
    $sheet->setCellValue('CE' . $i, $row['hasil_verivali']);
    $sheet->setCellValue('CF' . $i, $row['ttd_petugas_verivali']);
    $sheet->setCellValue('CG' . $i, $row['ttd_pemeriksa']);
    $sheet->setCellValue('CH' . $i, $row['keluhan']);
    $sheet->setCellValue('CI' . $i, $row['catatan_tambahan']);
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
$sheet->getStyle('A1:CI' . $i)->applyFromArray($styleArray);

$files = glob('../Data Export/*'); // get all file names
foreach ($files as $file) { // iterate files
    if (is_file($file))
        unlink($file); // delete file
}

$file_name = "Data Bangunan " . date("(Y-m-d h-i-s)") . ".xlsx";
$writer = new Xlsx($spreadsheet);
$writer->save('../Data Export/' . $file_name);
$url = $site_url . "Data Export/" . $file_name;
header("Location: $url");
// echo "<script>alert(' Data berhasil di Ex p ort!');history.go(-1);</script>";
