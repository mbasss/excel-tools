<?php

$sheetimport1 = '$sheetData[$i][';
$sheetimport2 = "'];";

$sheet1 = '$sheet->setCellValue(';
$sheet2 = ' . $i, $row[';
$sheet3 = ' ]);';
$petik = "'";



include('koneksi.php');

$i = 1;
$query = mysqli_query($koneksi, "DESC bangunan");
while ($row = mysqli_fetch_row($query)) {
    // echo $sheet1 . $petik . "pb" . $petik . $sheet2 . $row[0] . $petik . $sheet3 . "<BR>";
    // echo "$" . $row[0] . "<BR>";
    // echo "'$" . $row[0] . "',";
    // echo  $sheetimport1 . $petik . $i++ . $sheetimport2 . "<BR>";
}




// function excelColumnRange($lower, $upper)
// {
//     ++$upper;
//     for ($i = $lower; $i !== $upper; ++$i) {
//         yield $i;
//     }
// }

// foreach (excelColumnRange('A', 'CI') as $value) {
//     echo $value . "1<BR>";
// }





//  gAK kEPAKE


// function createColumnsArray($end_column, $first_letters = '')
// {
//     $columns = array();
//     $length = strlen($end_column);
//     $letters = range('A', 'Z');

//     // Iterate over 26 letters.
//     foreach ($letters as $letter) {
//         // Paste the $first_letters before the next.
//         $column = $first_letters . $letter;

//         // Add the column to the final array.
//         $columns[] = $column;

//         // If it was the end column that was added, return the columns.
//         if ($column == $end_column)
//             return $columns;
//     }

//     // Add the column children.
//     foreach ($columns as $column) {
//         // Don't itterate if the $end_column was already set in a previous itteration.
//         // Stop iterating if you've reached the maximum character length.
//         if (!in_array($end_column, $columns) && strlen($column) < $length) {
//             $new_columns = createColumnsArray($end_column, $column);
//             // Merge the new columns which were created with the final columns array.
//             $columns = array_merge($columns, $new_columns);
//         }
//     }

//     return $columns;
// }
// echo "<pre>";
// print_r(createColumnsArray('BZ'));