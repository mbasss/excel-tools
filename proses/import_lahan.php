<?php
include('koneksi.php');
require '../vendor/autoload.php';

echo "<a href='../index.php'>Home</a> </br>";

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Reader\Csv;
use PhpOffice\PhpSpreadsheet\Reader\Xlsx;

$file_mimes = array('application/octet-stream', 'application/vnd.ms-excel', 'application/x-csv', 'text/x-csv', 'text/csv', 'application/csv', 'application/excel', 'application/vnd.msexcel', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');

if (isset($_FILES['berkas_excel']['name']) && in_array($_FILES['berkas_excel']['type'], $file_mimes)) {

    $arr_file = explode('.', $_FILES['berkas_excel']['name']);
    $extension = end($arr_file);

    if ('csv' == $extension) {
        $reader = new \PhpOffice\PhpSpreadsheet\Reader\Csv();
    } else {
        $reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
    }

    $spreadsheet = $reader->load($_FILES['berkas_excel']['tmp_name']);
    $no = 0;
    $sheetData = $spreadsheet->getActiveSheet()->toArray();
    for ($i = 1; $i < count($sheetData); $i++) {
        $no++;
        $id_desa                        = $sheetData[$i]['1'];
        $id_lahan                       = $sheetData[$i]['2'];
        $asal_pemilik                   = $sheetData[$i]['3'];
        $nama_pemilik                   = addslashes($sheetData[$i]['4']);
        $alamat_pemilik                 = $sheetData[$i]['5'];
        $id_pemilik                     = $sheetData[$i]['6'];
        $id_penanggungjawab             = addslashes($sheetData[$i]['7']);
        $afiliasi_kelompok              = $sheetData[$i]['8'];
        $jenis_lahan                    = addslashes($sheetData[$i]['9']);
        $status_lahan                   = $sheetData[$i]['10'];
        $fungsi_lahan                   = $sheetData[$i]['11'];
        $kelengkapan_dokumen            = $sheetData[$i]['12'];
        $kondisi_tanah                  = $sheetData[$i]['13'];
        $luas_tanaman_pertahun          = $sheetData[$i]['14'];
        $nilai_produksi_pertahun        = $sheetData[$i]['15'];
        $biaya_pemupukan_pertahun       = $sheetData[$i]['16'];
        $biaya_bibit_pertahun           = $sheetData[$i]['17'];
        $biaya_obat_pertahun            = $sheetData[$i]['18'];
        $biaya_lain_pertahun            = $sheetData[$i]['19'];
        $sarana_irigasi                 = $sheetData[$i]['20'];
        $pjg_irigasi_primer             = $sheetData[$i]['21'];
        $pjg_irigasi_sekunder           = $sheetData[$i]['22'];
        $pjg_irigasi_tersier            = $sheetData[$i]['23'];
        $jml_pintu_sadap                = $sheetData[$i]['24'];
        $jm_pintu_air                   = $sheetData[$i]['25'];
        $fasilitas_pendukung            = $sheetData[$i]['26'];
        $jenis_fas_umum                 = $sheetData[$i]['27'];
        $transportasi_terparkir         = $sheetData[$i]['28'];
        $jenis_irigasi                  = $sheetData[$i]['29'];
        $produk_dihasilkan              = $sheetData[$i]['30'];
        $jenis_ternak                   = $sheetData[$i]['31'];
        $lahan_gembala                  = $sheetData[$i]['32'];
        $jumlah_populasi                = $sheetData[$i]['33'];
        $omzet_pertahun                 = $sheetData[$i]['34'];
        $modal_pertahun                 = $sheetData[$i]['35'];
        $jml_pekerja                    = $sheetData[$i]['36'];
        $pemasaran                      = $sheetData[$i]['37'];
        $luas_lahan                     = $sheetData[$i]['38'];
        $kondisi_hutan                  = $sheetData[$i]['39'];
        $gangguan_dirasakan             = $sheetData[$i]['40'];
        $dampak_ke_lingkungan           = $sheetData[$i]['41'];
        $foto_lahan                     = $sheetData[$i]['42'];
        $foto_fasilitas                 = $sheetData[$i]['43'];
        $foto_produk                    = $sheetData[$i]['44'];
        $koordinat                      = $sheetData[$i]['45'];
        $nama_petugas                   = $sheetData[$i]['46'];
        $tgl_pendataan                  = $sheetData[$i]['47'];
        echo $no . ". " .  $id_pemilik . "<br>";

        mysqli_query($koneksi, "insert into warga values ('$id_desa','$id_lahan','$asal_pemilik','$nama_pemilik','$alamat_pemilik','$id_pemilik','$id_penanggungjawab','$afiliasi_kelompok','$jenis_lahan','$status_lahan','$fungsi_lahan','$kelengkapan_dokumen','$kondisi_tanah','$luas_tanaman_pertahun','$nilai_produksi_pertahun','$biaya_pemupukan_pertahun','$biaya_bibit_pertahun','$biaya_obat_pertahun', `$biaya_lain_pertahun`, `$sarana_irigasi`, `$pjg_irigasi_primer`, `$pjg_irigasi_sekunder`, `$pjg_irigasi_tersier`, `$jml_pintu_sadap`, `$jm_pintu_air`, `$fasilitas_pendukung`, `$jenis_fas_umum`, `$transportasi_terparkir`, `$jenis_irigasi`, `$produk_dihasilkan`, `$jenis_ternak`, `$lahan_gembala`, `$jumlah_populasi`, `$omzet_pertahun`, `$modal_pertahun`, `$jml_pekerja`, `$pemasaran`, `$luas_lahan`, `$kondisi_hutan`, `$gangguan_dirasakan`, `$dampak_ke_lingkungan`, `$foto_lahan`, `$foto_fasilitas`, `$foto_produk`, `$koordinat`, `$nama_petugas`, `$tgl_pendataan`)");
    }

    echo "<a href='../index.php'>Home</a>";
    // header("Location: ../index.php");
}
