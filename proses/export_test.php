<?php
include('koneksi.php');
require '../vendor/autoload.php';
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

$spreadsheet = new Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();
$sheet->setCellValue('A1', 'No');
$sheet->setCellValue('B1', 'Nama Test');
$sheet->setCellValue('C1', 'Keterangan Test');
$sheet->setCellValue('D1', 'Tanggal Test');

$query = mysqli_query($koneksi, "select * from test_excel");
$i = 2;
$no = 1;
while ($row = mysqli_fetch_array($query)) {
    $sheet->setCellValue('A' . $i, $no++);
    $sheet->setCellValue('B' . $i, $row['nama_test']);
    $sheet->setCellValue('C' . $i, $row['keterangan_test']);
    $sheet->setCellValue('D' . $i, $row['tanggal_test']);
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
$sheet->getStyle('A1:D' . $i)->applyFromArray($styleArray);


$writer = new Xlsx($spreadsheet);
$writer->save('../Data Export/Data Test.xlsx');
echo "<script>alert('Data berhasil di Export!');history.go(-1);</script>";
// header("Location: ../index.php");
