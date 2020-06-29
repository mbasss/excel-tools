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
        $id_bangunan                    = $sheetData[$i]['2'];
        $asal_pemilik                   = $sheetData[$i]['3'];
        $fungsi_bangunan                = $sheetData[$i]['4'];
        $layanan_kesehatan              = $sheetData[$i]['5'];
        $jml_dokter                     = $sheetData[$i]['6'];
        $jml_bidan                      = $sheetData[$i]['7'];
        $jml_perawat                    = $sheetData[$i]['8'];
        $jenis_lbg_keamanan             = $sheetData[$i]['9'];
        $jml_linmas                     = $sheetData[$i]['10'];
        $jml_satgas_linmas              = $sheetData[$i]['11'];
        $jml_satpam_swakarsa            = $sheetData[$i]['12'];
        $namainduk_satpam_swakarsa      = $sheetData[$i]['13'];
        $kepemilikan_satpam_swakarsa    = $sheetData[$i]['14'];
        $jml_mitra_tni                  = $sheetData[$i]['15'];
        $jml_kegiatan_tni               = $sheetData[$i]['16'];
        $jml_mitra_polri                = $sheetData[$i]['17'];
        $jml_kegiatan_polri             = $sheetData[$i]['18'];
        $kelengkapan_lbg_adat           = $sheetData[$i]['19'];
        $jenis_lbg_pendidikan           = $sheetData[$i]['20'];
        $status_lbg_pendidikan          = $sheetData[$i]['21'];
        $jml_pengajar                   = $sheetData[$i]['22'];
        $jml_siswa                      = $sheetData[$i]['23'];
        $jenis_usaha                    = $sheetData[$i]['24'];
        $jenis_angkutan                 = $sheetData[$i]['25'];
        $kapasitas_angkut_orang         = $sheetData[$i]['26'];
        $kapasitas_angkut_barang        = $sheetData[$i]['27'];
        $trayek                         = $sheetData[$i]['28'];
        $nama_pemilik                   = $sheetData[$i]['29'];
        $alamat_pemilik                 = $sheetData[$i]['30'];
        $id_kelompok                    = $sheetData[$i]['31'];
        $jenis_kepemilikan              = $sheetData[$i]['32'];
        $id_fasilitas_pendukung         = $sheetData[$i]['33'];
        $omzet_pertahun                 = $sheetData[$i]['34'];
        $modal_pertahun                 = $sheetData[$i]['35'];
        $jml_pekerja                    = $sheetData[$i]['36'];
        $produk_dihasilkan              = $sheetData[$i]['37'];
        $pemasaran                      = $sheetData[$i]['38'];
        $kelengkapan_izin               = $sheetData[$i]['39'];
        $gangguan_dari_lingkungan       = $sheetData[$i]['40'];
        $dampak_ke_lingkungan           = $sheetData[$i]['41'];
        $id_warga                       = $sheetData[$i]['42'];
        $dusun                          = $sheetData[$i]['43'];
        $rt                             = $sheetData[$i]['44'];
        $rw                             = $sheetData[$i]['45'];
        $jalan                          = $sheetData[$i]['46'];
        $nomor_rumah                    = $sheetData[$i]['47'];
        $jml_anggota_rumah              = $sheetData[$i]['48'];
        $jumlah_kk_dirumah              = $sheetData[$i]['49'];
        $status_rumah                   = $sheetData[$i]['50'];
        $status_lahan_tinggal           = $sheetData[$i]['51'];
        $luas_lantai                    = $sheetData[$i]['52'];
        $jenis_lantai_terluas           = $sheetData[$i]['53'];
        $jenis_dinding_terluas          = $sheetData[$i]['54'];
        $kondisi_dinding_terluas        = $sheetData[$i]['55'];
        $jenis_atap_terluas             = $sheetData[$i]['56'];
        $kondisi_atap_terluas           = $sheetData[$i]['57'];
        $jml_kamar_tidur                = $sheetData[$i]['58'];
        $sumber_air_minum               = $sheetData[$i]['59'];
        $id_pelanggan_air               = $sheetData[$i]['60'];
        $cara_memperoleh_airminum       = $sheetData[$i]['61'];
        $sumber_peneranga_utama         = $sheetData[$i]['62'];
        $daya_listrik_terpasang         = $sheetData[$i]['63'];
        $nomor_rek_listrik              = $sheetData[$i]['64'];
        $energi_untuk_memasak           = $sheetData[$i]['65'];
        $id_gas_pipa                    = $sheetData[$i]['66'];
        $fasilitas_bab                  = $sheetData[$i]['67'];
        $jenis_kloset                   = $sheetData[$i]['68'];
        $pembuangan_tinja               = $sheetData[$i]['69'];
        $perlengkapan_rumah             = $sheetData[$i]['70'];
        $luas_lahan                     = $sheetData[$i]['71'];
        $foto_pemilik                   = $sheetData[$i]['72'];
        $foto_depan                     = $sheetData[$i]['73'];
        $foto_imb                       = $sheetData[$i]['74'];
        $foto_pbb                       = $sheetData[$i]['75'];
        $foto_lampu_jalan               = $sheetData[$i]['76'];
        $koordinat                      = $sheetData[$i]['77'];
        $tgl_pendataan                  = $sheetData[$i]['78'];
        $nama_petugas                   = $sheetData[$i]['79'];
        $tgl_pemeriksaan                = $sheetData[$i]['80'];
        $nama_pemeriksa                 = $sheetData[$i]['81'];
        $hasil_verivali                 = $sheetData[$i]['82'];
        $ttd_petugas_verivali           = $sheetData[$i]['83'];
        $ttd_pemeriksa                  = $sheetData[$i]['84'];
        $keluhan                        = $sheetData[$i]['85'];
        $catatan_tambahan               = $sheetData[$i]['86'];
        echo $no . ". " .  $id_bangunan . "<br>";

        mysqli_query($koneksi, "insert into bangunan values ('$id_desa','$id_bangunan','$asal_pemilik','$fungsi_bangunan','$layanan_kesehatan','$jml_dokter','$jml_bidan','$jml_perawat','$jenis_lbg_keamanan','$jml_linmas','$jml_satgas_linmas','$jml_satpam_swakarsa','$namainduk_satpam_swakarsa','$kepemilikan_satpam_swakarsa','$jml_mitra_tni','$jml_kegiatan_tni','$jml_mitra_polri','$jml_kegiatan_polri','$kelengkapan_lbg_adat','$jenis_lbg_pendidikan','$status_lbg_pendidikan','$jml_pengajar','$jml_siswa','$jenis_usaha','$jenis_angkutan','$kapasitas_angkut_orang','$kapasitas_angkut_barang','$trayek','$nama_pemilik','$alamat_pemilik','$id_kelompok','$jenis_kepemilikan','$id_fasilitas_pendukung','$omzet_pertahun','$modal_pertahun','$jml_pekerja','$produk_dihasilkan','$pemasaran','$kelengkapan_izin','$gangguan_dari_lingkungan','$dampak_ke_lingkungan','$id_warga','$dusun','$rt','$rw','$jalan','$nomor_rumah','$jml_anggota_rumah','$jumlah_kk_dirumah','$status_rumah','$status_lahan_tinggal','$luas_lantai','$jenis_lantai_terluas','$jenis_dinding_terluas','$kondisi_dinding_terluas','$jenis_atap_terluas','$kondisi_atap_terluas','$jml_kamar_tidur','$sumber_air_minum','$id_pelanggan_air','$cara_memperoleh_airminum','$sumber_peneranga_utama','$daya_listrik_terpasang','$nomor_rek_listrik','$energi_untuk_memasak','$id_gas_pipa','$fasilitas_bab','$jenis_kloset','$pembuangan_tinja','$perlengkapan_rumah','$luas_lahan','$foto_pemilik','$foto_depan','$foto_imb','$foto_pbb','$foto_lampu_jalan','$koordinat','$tgl_pendataan','$nama_petugas','$tgl_pemeriksaan','$nama_pemeriksa','$hasil_verivali','$ttd_petugas_verivali','$ttd_pemeriksa','$keluhan','$catatan_tambahan')");
    }

    echo "<a href='../index.php'>Home</a>";
    // header("Location: ../index.php");
}
