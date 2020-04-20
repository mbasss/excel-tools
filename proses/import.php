<?php
include('koneksi.php');
require '../vendor/autoload.php';

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
        $id_desa                    = $sheetData[$i]['1'];
        $id_warga                   = $sheetData[$i]['2'];
        $id_bangunan                = $sheetData[$i]['3'];
        $nomor_kk                   = $sheetData[$i]['4'];
        $nomor_ktp                  = $sheetData[$i]['5'];
        $nomor_hp                   = $sheetData[$i]['6'];
        $nama_warga                 = addslashes($sheetData[$i]['7']);
        $jenis_kelamin              = $sheetData[$i]['8'];
        $tempat_lahir               = addslashes($sheetData[$i]['9']);
        $tanggal_lahir              = $sheetData[$i]['10'];
        $hub_keluarga               = $sheetData[$i]['11'];
        $status_nikah               = $sheetData[$i]['12'];
        $kelengkapan_dokumen        = $sheetData[$i]['13'];
        $tercantum_di_kk_ini        = $sheetData[$i]['14'];
        $status_hamil               = $sheetData[$i]['15'];
        $periksa_kehamilan_di       = $sheetData[$i]['16'];
        $jenis_kontrasepsi          = $sheetData[$i]['17'];
        $jenis_cacat                = $sheetData[$i]['18'];
        $penyakit_kronis            = $sheetData[$i]['19'];
        $keberadaan_sekarang        = $sheetData[$i]['20'];
        $partisipasi_sekolah        = $sheetData[$i]['21'];
        $nama_sekolah               = $sheetData[$i]['22'];
        $jenjang_sekolah_sekarang   = $sheetData[$i]['23'];
        $ijazah_tertinggi           = $sheetData[$i]['24'];
        $status_kerja               = $sheetData[$i]['25'];
        $lap_usaha                  = $sheetData[$i]['26'];
        $keahlian_dimiliki          = $sheetData[$i]['27'];
        $penghasilan_perbulan       = $sheetData[$i]['28'];
        $kategori_sosial            = $sheetData[$i]['29'];
        $masalah_kesejahteraan      = $sheetData[$i]['30'];
        $gangguan_lingkungan        = $sheetData[$i]['31'];
        $bantuan_yang_diterima      = $sheetData[$i]['32'];
        $afiliasi_kelompok          = $sheetData[$i]['33'];
        $gol_darah                  = $sheetData[$i]['34'];
        $agama                      = $sheetData[$i]['35'];
        $tgl_pendataan              = $sheetData[$i]['36'];
        $nama_petugas               = $sheetData[$i]['37'];
        $foto_diri                  = $sheetData[$i]['38'];
        $foto_ktp                   = $sheetData[$i]['39'];
        $foto_kk                    = $sheetData[$i]['40'];
        $peran_di_desa              = $sheetData[$i]['41'];
        echo $no . ". " .  $nama_warga . "<br>";

        mysqli_query($koneksi, "insert into warga values ('$id_desa','$id_warga','$id_bangunan','$nomor_kk','$nomor_ktp','$nomor_hp','$no','$nama_warga','$jenis_kelamin','$tempat_lahir','$tanggal_lahir','$hub_keluarga','$status_nikah','$kelengkapan_dokumen','$tercantum_di_kk_ini','$status_hamil','$periksa_kehamilan_di','$jenis_kontrasepsi','$jenis_cacat','$penyakit_kronis','$keberadaan_sekarang','$partisipasi_sekolah','$nama_sekolah','$jenjang_sekolah_sekarang','$ijazah_tertinggi','$status_kerja','$lap_usaha','$keahlian_dimiliki','$penghasilan_perbulan','$kategori_sosial','$masalah_kesejahteraan','$gangguan_lingkungan','$bantuan_yang_diterima','$afiliasi_kelompok','$gol_darah','$agama','$tgl_pendataan','$nama_petugas','$foto_diri','$foto_ktp','$foto_kk','$peran_di_desa')");
    }

    // echo "<script>alert('Data berhasil di Import!');history.go(-1);</script>";
    // header("Location: ../index.php");
}
