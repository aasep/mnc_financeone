<?php
require_once 'Classes/PHPExcel.php';
require_once 'Classes/PHPExcel/IOFactory.php';

// Create new PHPExcel object
$objPHPExcel = new PHPExcel();

$objPHPExcel->setActiveSheetIndex(0);

$styleArrayFontBold = array('font' => array('bold'  => true,'color' => array('rgb' => ''),'size'  => 11,'name'  => 'Calibri'));
$styleArrayAlignment1 = array('alignment' => array(
            'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
        ));
$styleArrayAlignment2 = array('alignment' => array(
            'horizontal' => PHPExcel_Style_Alignment::VERTICAL_CENTER,
        ));

$styleArrayColorFont = array(
    'font'  => array(
        'bold'  => true,
        'color' => array('rgb' => 'FFFFFF'),
        'size'  => 11,
        'name'  => 'Calibri'
    ));

$styleArrayBorder1 = array('borders' => array('allborders' => array('style' => PHPExcel_Style_Border::BORDER_THIN,'color' => array('argb' => ''),),),);
$styleArrayBorder2 = array('borders' => array('outline' => array('style' => PHPExcel_Style_Border::BORDER_THIN,'color' => array('argb' => ''),),),);
$styleArrayBorder3 = array('borders' => array('bottom' => array('style' => PHPExcel_Style_Border::BORDER_THIN,'color' => array('argb' => ''),),),);


$objPHPExcel->getActiveSheet()->getStyle('A7:G8')->applyFromArray($styleArrayBorder1);

$objPHPExcel->getActiveSheet()->getStyle('A11:G109')->applyFromArray($styleArrayBorder1);

$objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth(5);
$objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth(5);
$objPHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth(10);
$objPHPExcel->getActiveSheet()->getColumnDimension('D')->setWidth(60);
$objPHPExcel->getActiveSheet()->getColumnDimension('E')->setWidth(15);
$objPHPExcel->getActiveSheet()->getColumnDimension('F')->setWidth(25);
$objPHPExcel->getActiveSheet()->getColumnDimension('G')->setWidth(25);


$objPHPExcel->setActiveSheetIndex(0)->mergeCells('A1:G1');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('A2:G2');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('A3:G3');

$objPHPExcel->setActiveSheetIndex(0)->mergeCells('B7:D8');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('F7:G7');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('A7:A8');
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('E7:E8');


for ($i=11; $i <=109 ; $i++) { 
$objPHPExcel->setActiveSheetIndex(0)->mergeCells("B$i:D$i");
}


$objPHPExcel->getActiveSheet()->getStyle('A1:G3')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('B7')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('A7:G8')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('A7:G8')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);

$objPHPExcel->getActiveSheet()->getStyle('A1:G3')->applyFromArray($styleArrayFontBold);
$objPHPExcel->getActiveSheet()->getStyle('A7:G8')->applyFromArray($styleArrayFontBold);
$objPHPExcel->getActiveSheet()->getStyle('A10')->applyFromArray($styleArrayFontBold);

$objPHPExcel->getActiveSheet()->setCellValue('A1', "LAPORAN POSISI KEUANGAN");
$objPHPExcel->getActiveSheet()->setCellValue('A2', 'PT BANK MNC INTERNASIONAL Tbk.');
$objPHPExcel->getActiveSheet()->setCellValue('A3', "Per Tanggal1 dan Tanggal2 ");

$objPHPExcel->getActiveSheet()->setCellValue('A7', 'No');
$objPHPExcel->getActiveSheet()->setCellValue('B7', 'POS - POS');
$objPHPExcel->getActiveSheet()->setCellValue('E7', 'Sandi LBU');
$objPHPExcel->getActiveSheet()->setCellValue('F7', 'BANK');
$objPHPExcel->getActiveSheet()->setCellValue('F8', 'Tanggal1');
$objPHPExcel->getActiveSheet()->setCellValue('G8', 'Tanggal2');

$objPHPExcel->getActiveSheet()->setCellValue('A10', 'ASET');

$objPHPExcel->getActiveSheet()->setCellValue('A11', '1.');
$objPHPExcel->getActiveSheet()->setCellValue('B11', 'Kas');
$objPHPExcel->getActiveSheet()->setCellValue('E11', 100);
$objPHPExcel->getActiveSheet()->setCellValue('A12', '2.');
$objPHPExcel->getActiveSheet()->setCellValue('B12', 'Penempatan pada Bank Indonesia');
$objPHPExcel->getActiveSheet()->setCellValue('E12', 120);
$objPHPExcel->getActiveSheet()->setCellValue('A13', '3.');
$objPHPExcel->getActiveSheet()->setCellValue('B13', 'Penempatan pada bank lain');
$objPHPExcel->getActiveSheet()->setCellValue('E13', 130);
$objPHPExcel->getActiveSheet()->setCellValue('A14', '4.');
$objPHPExcel->getActiveSheet()->setCellValue('B14', 'Tagihan spot dan derivatif');
$objPHPExcel->getActiveSheet()->setCellValue('E14', 135);
$objPHPExcel->getActiveSheet()->setCellValue('A15', '5.');
$objPHPExcel->getActiveSheet()->setCellValue('B15', 'Surat berharga');
$objPHPExcel->getActiveSheet()->setCellValue('E15', "");

$objPHPExcel->getActiveSheet()->setCellValue('A16', '');
$objPHPExcel->getActiveSheet()->setCellValue('B16', 'a. Diukur pada nilai wajar melalui laporan laba/rugi');
$objPHPExcel->getActiveSheet()->setCellValue('E16', "");
$objPHPExcel->getActiveSheet()->setCellValue('A17', '');
$objPHPExcel->getActiveSheet()->setCellValue('B17', 'b. Tersedia untuk dijual');
$objPHPExcel->getActiveSheet()->setCellValue('E17', 143);
$objPHPExcel->getActiveSheet()->setCellValue('A18', '');
$objPHPExcel->getActiveSheet()->setCellValue('B18', 'c. Dimiliki hingga jatuh tempo');
$objPHPExcel->getActiveSheet()->setCellValue('E18', 144);
$objPHPExcel->getActiveSheet()->setCellValue('A19', '');
$objPHPExcel->getActiveSheet()->setCellValue('B19', 'd. Pinjaman yang diberikan dan piutang');
$objPHPExcel->getActiveSheet()->setCellValue('E19', 145);

$objPHPExcel->getActiveSheet()->setCellValue('A20', '6.');
$objPHPExcel->getActiveSheet()->setCellValue('B20', 'Surat berharga yang dijual dengan janji dibeli kembali (repo)');
$objPHPExcel->getActiveSheet()->setCellValue('E20', 160);
$objPHPExcel->getActiveSheet()->setCellValue('A21', '7.');
$objPHPExcel->getActiveSheet()->setCellValue('B21', 'Tagihan atas surat berharga yang dibeli dengan janji dijual kembali (reverse repo)');
$objPHPExcel->getActiveSheet()->setCellValue('E21', 164);
$objPHPExcel->getActiveSheet()->setCellValue('A22', '8.');
$objPHPExcel->getActiveSheet()->setCellValue('B22', 'Tagihan akseptasi');
$objPHPExcel->getActiveSheet()->setCellValue('E22', 166);
$objPHPExcel->getActiveSheet()->setCellValue('A23', '9.');
$objPHPExcel->getActiveSheet()->setCellValue('B23', 'Kredit ');
$objPHPExcel->getActiveSheet()->setCellValue('E23', "");
$objPHPExcel->getActiveSheet()->setCellValue('A24', '');
$objPHPExcel->getActiveSheet()->setCellValue('B24', 'a. Diukur pada nilai wajar melalui laporan laba/rugi');
$objPHPExcel->getActiveSheet()->setCellValue('E24', "");
$objPHPExcel->getActiveSheet()->setCellValue('A25', '');
$objPHPExcel->getActiveSheet()->setCellValue('B25', 'b. Tersedia untuk dijual');
$objPHPExcel->getActiveSheet()->setCellValue('E25', 172);
$objPHPExcel->getActiveSheet()->setCellValue('A26', '');
$objPHPExcel->getActiveSheet()->setCellValue('B26', 'c. Dimiliki hingga jatuh tempo');
$objPHPExcel->getActiveSheet()->setCellValue('E26', 173);
$objPHPExcel->getActiveSheet()->setCellValue('A27', '');
$objPHPExcel->getActiveSheet()->setCellValue('B27', 'd. Pinjaman yang diberikan dan piutang');
$objPHPExcel->getActiveSheet()->setCellValue('E27', 175);

$objPHPExcel->getActiveSheet()->setCellValue('A28', '10.');
$objPHPExcel->getActiveSheet()->setCellValue('B28', 'Pembiayaan syariah ¹⁾');
$objPHPExcel->getActiveSheet()->setCellValue('E28', 174);
$objPHPExcel->getActiveSheet()->setCellValue('A29', '11.');
$objPHPExcel->getActiveSheet()->setCellValue('B29', 'Penyertaan ');
$objPHPExcel->getActiveSheet()->setCellValue('E29', 200);
$objPHPExcel->getActiveSheet()->setCellValue('A30', '12.');
$objPHPExcel->getActiveSheet()->setCellValue('B30', 'Cadangan kerugian penurunan nilai aset keuangan -/-');
$objPHPExcel->getActiveSheet()->setCellValue('E30', "");

$objPHPExcel->getActiveSheet()->setCellValue('A31', '');
$objPHPExcel->getActiveSheet()->setCellValue('B31', 'a. Surat Berharga');
$objPHPExcel->getActiveSheet()->setCellValue('E31', 201);
$objPHPExcel->getActiveSheet()->setCellValue('A32', '');
$objPHPExcel->getActiveSheet()->setCellValue('B32', 'b. Kredit');
$objPHPExcel->getActiveSheet()->setCellValue('E32', 202);
$objPHPExcel->getActiveSheet()->setCellValue('A33', '');
$objPHPExcel->getActiveSheet()->setCellValue('B33', 'c. Lainnya');
$objPHPExcel->getActiveSheet()->setCellValue('E33', 206);

$objPHPExcel->getActiveSheet()->setCellValue('A34', '13.');
$objPHPExcel->getActiveSheet()->setCellValue('B34', 'Aset Tidak Berwujud');
$objPHPExcel->getActiveSheet()->setCellValue('E34', 212);
$objPHPExcel->getActiveSheet()->setCellValue('A35', '');
$objPHPExcel->getActiveSheet()->setCellValue('B35', 'Akumulasi amortisasi aset tidak berwujud -/-');
$objPHPExcel->getActiveSheet()->setCellValue('E35', 213);

$objPHPExcel->getActiveSheet()->setCellValue('A36', '14.');
$objPHPExcel->getActiveSheet()->setCellValue('B36', 'Aset tetap dan inventaris');
$objPHPExcel->getActiveSheet()->setCellValue('E36', 214);
$objPHPExcel->getActiveSheet()->setCellValue('A37', '');
$objPHPExcel->getActiveSheet()->setCellValue('B37', 'Akumulasi penyusutan aset tetap dan inventaris -/-');
$objPHPExcel->getActiveSheet()->setCellValue('E37', 215);
$objPHPExcel->getActiveSheet()->setCellValue('A38', '15.');
$objPHPExcel->getActiveSheet()->setCellValue('B38', 'Aset non produktif');
$objPHPExcel->getActiveSheet()->setCellValue('E38', "");

$objPHPExcel->getActiveSheet()->setCellValue('A39', '');
$objPHPExcel->getActiveSheet()->setCellValue('B39', 'a. Properti terbengkalai');
$objPHPExcel->getActiveSheet()->setCellValue('E39', 217);
$objPHPExcel->getActiveSheet()->setCellValue('A40', '');
$objPHPExcel->getActiveSheet()->setCellValue('B40', 'b. Aset yang diambil alih');
$objPHPExcel->getActiveSheet()->setCellValue('E40', 218);
$objPHPExcel->getActiveSheet()->setCellValue('A41', '');
$objPHPExcel->getActiveSheet()->setCellValue('B41', 'c. Rekening tunda');
$objPHPExcel->getActiveSheet()->setCellValue('E41', 219);
$objPHPExcel->getActiveSheet()->setCellValue('A42', '');
$objPHPExcel->getActiveSheet()->setCellValue('B42', 'd. Aset antarkantor ²⁾');
$objPHPExcel->getActiveSheet()->setCellValue('E42', "");


$objPHPExcel->getActiveSheet()->setCellValue('B43', "i. Melakukan kegiatan operasional di Indonesia ");
$objPHPExcel->getActiveSheet()->setCellValue('E43', 223);
$objPHPExcel->getActiveSheet()->setCellValue('B44', "ii. Melakukan kegiatan operasional di luar Indonesia ");
$objPHPExcel->getActiveSheet()->setCellValue('E44', 224);

$objPHPExcel->getActiveSheet()->setCellValue('A45', '16.');
$objPHPExcel->getActiveSheet()->setCellValue('B45', 'Cadangan kerugian penurunan nilai aset non keuangan -/-');
$objPHPExcel->getActiveSheet()->setCellValue('E45', "Form 21");
$objPHPExcel->getActiveSheet()->setCellValue('A46', '17.');
$objPHPExcel->getActiveSheet()->setCellValue('B46', 'Sewa pembiayaan ¹⁾');
$objPHPExcel->getActiveSheet()->setCellValue('E46', 227);
$objPHPExcel->getActiveSheet()->setCellValue('A47', '18.');
$objPHPExcel->getActiveSheet()->setCellValue('B47', 'Aset pajak tangguhan ');
$objPHPExcel->getActiveSheet()->setCellValue('E47', 228);
$objPHPExcel->getActiveSheet()->setCellValue('A48', '19.');
$objPHPExcel->getActiveSheet()->setCellValue('B48', 'Aset lainnya');
$objPHPExcel->getActiveSheet()->setCellValue('E48', 230);

$objPHPExcel->getActiveSheet()->setCellValue('B50', "TOTAL ASET");
$objPHPExcel->getActiveSheet()->setCellValue('E50', 290);


$objPHPExcel->getActiveSheet()->setCellValue('A51', "LIABILITAS DAN EKUITAS");
$objPHPExcel->getActiveSheet()->setCellValue('B52', "LIABILITAS");

$objPHPExcel->getActiveSheet()->setCellValue('A53', "1.");
$objPHPExcel->getActiveSheet()->setCellValue('B53', "Giro");
$objPHPExcel->getActiveSheet()->setCellValue('E53', 300);
$objPHPExcel->getActiveSheet()->setCellValue('A54', "2.");
$objPHPExcel->getActiveSheet()->setCellValue('B54', "Tabungan");
$objPHPExcel->getActiveSheet()->setCellValue('E54', 320);
$objPHPExcel->getActiveSheet()->setCellValue('A55', "3.");
$objPHPExcel->getActiveSheet()->setCellValue('B55', "Simpanan berjangka");
$objPHPExcel->getActiveSheet()->setCellValue('E55', 330);
$objPHPExcel->getActiveSheet()->setCellValue('A56', "4.");
$objPHPExcel->getActiveSheet()->setCellValue('B56', "Dana investasi revenue sharing ¹⁾");
$objPHPExcel->getActiveSheet()->setCellValue('E56', "");
$objPHPExcel->getActiveSheet()->setCellValue('A57', "5.");
$objPHPExcel->getActiveSheet()->setCellValue('B57', "Pinjaman dari Bank Indonesia");
$objPHPExcel->getActiveSheet()->setCellValue('E57', 340);
$objPHPExcel->getActiveSheet()->setCellValue('A58', "6.");
$objPHPExcel->getActiveSheet()->setCellValue('B58', "Pinjaman dari bank lain");
$objPHPExcel->getActiveSheet()->setCellValue('E58', 350);
$objPHPExcel->getActiveSheet()->setCellValue('A59', "7.");
$objPHPExcel->getActiveSheet()->setCellValue('B59', "Liabilitas spot dan derivatif");
$objPHPExcel->getActiveSheet()->setCellValue('E59', 351);
$objPHPExcel->getActiveSheet()->setCellValue('A60', "8.");
$objPHPExcel->getActiveSheet()->setCellValue('B60', "Utang atas surat berharga yang dijual dengan janji dibeli kembali (repo)");
$objPHPExcel->getActiveSheet()->setCellValue('E60', 352);
$objPHPExcel->getActiveSheet()->setCellValue('A61', "9.");
$objPHPExcel->getActiveSheet()->setCellValue('B61', "Utang akseptasi");
$objPHPExcel->getActiveSheet()->setCellValue('E61', 353);
$objPHPExcel->getActiveSheet()->setCellValue('A62', "10.");
$objPHPExcel->getActiveSheet()->setCellValue('B62', "Surat berharga yang diterbitkan");
$objPHPExcel->getActiveSheet()->setCellValue('E62', 355);
$objPHPExcel->getActiveSheet()->setCellValue('A63', "11.");
$objPHPExcel->getActiveSheet()->setCellValue('B63', "Pinjaman yang diterima");
$objPHPExcel->getActiveSheet()->setCellValue('E63', 360);
$objPHPExcel->getActiveSheet()->setCellValue('A64', "12.");
$objPHPExcel->getActiveSheet()->setCellValue('B64', "Setoran jaminan");
$objPHPExcel->getActiveSheet()->setCellValue('E64', "");
$objPHPExcel->getActiveSheet()->setCellValue('A65', "13.");
$objPHPExcel->getActiveSheet()->setCellValue('B65', "Liabilitas antar kantor ²⁾");
$objPHPExcel->getActiveSheet()->setCellValue('E65', "");
$objPHPExcel->getActiveSheet()->setCellValue('B66', "a. Melakukan kegiatan operasional di Indonesia");
$objPHPExcel->getActiveSheet()->setCellValue('E66', 393);
$objPHPExcel->getActiveSheet()->setCellValue('B67', "b. Melakukan kegiatan operasional di luar Indonesia");
$objPHPExcel->getActiveSheet()->setCellValue('E67', 394);
$objPHPExcel->getActiveSheet()->setCellValue('A68', "14.");
$objPHPExcel->getActiveSheet()->setCellValue('B68', "Liabilitas pajak tangguhan");
$objPHPExcel->getActiveSheet()->setCellValue('E68', 396);
$objPHPExcel->getActiveSheet()->setCellValue('A69', "15.");
$objPHPExcel->getActiveSheet()->setCellValue('B69', "Liabilitas lainnya");
$objPHPExcel->getActiveSheet()->setCellValue('E69', 400);
$objPHPExcel->getActiveSheet()->setCellValue('A70', "16.");
$objPHPExcel->getActiveSheet()->setCellValue('B70', "Dana investasi profit sharing ¹⁾");
$objPHPExcel->getActiveSheet()->setCellValue('E70', 401);

$objPHPExcel->getActiveSheet()->setCellValue('B71', "TOTAL LIABILITAS");
$objPHPExcel->getActiveSheet()->setCellValue('E71', 401);


$objPHPExcel->getActiveSheet()->setCellValue('B73', "EKUITAS");
$objPHPExcel->getActiveSheet()->setCellValue('A74', "17");
$objPHPExcel->getActiveSheet()->setCellValue('B74', "Modal disetor");
$objPHPExcel->getActiveSheet()->setCellValue('E74', "");
$objPHPExcel->getActiveSheet()->setCellValue('B75', "a. Modal Dasar");
$objPHPExcel->getActiveSheet()->setCellValue('E75', "421");
$objPHPExcel->getActiveSheet()->setCellValue('B76', "b. Modal yang belum disetor -/-");
$objPHPExcel->getActiveSheet()->setCellValue('E76', "422");
$objPHPExcel->getActiveSheet()->setCellValue('B77', "c. Saham yang dibeli kembali (treasury stock) -/-");
$objPHPExcel->getActiveSheet()->setCellValue('E77', "423");
$objPHPExcel->getActiveSheet()->setCellValue('A78', "18");
$objPHPExcel->getActiveSheet()->setCellValue('B78', "Tambahan Modal disetor");
$objPHPExcel->getActiveSheet()->setCellValue('E78', "");
$objPHPExcel->getActiveSheet()->setCellValue('B79', "a. Agio");
$objPHPExcel->getActiveSheet()->setCellValue('E79', "431");
$objPHPExcel->getActiveSheet()->setCellValue('B80', "b. Disagio -/-");
$objPHPExcel->getActiveSheet()->setCellValue('E80', "432");
$objPHPExcel->getActiveSheet()->setCellValue('B81', "c. Modal sumbangan");
$objPHPExcel->getActiveSheet()->setCellValue('E81', "433");
$objPHPExcel->getActiveSheet()->setCellValue('B82', "d. Dana setoran modal");
$objPHPExcel->getActiveSheet()->setCellValue('E82', "");
$objPHPExcel->getActiveSheet()->setCellValue('B83', "e. Lainnya");
$objPHPExcel->getActiveSheet()->setCellValue('E83', "");
$objPHPExcel->getActiveSheet()->setCellValue('A84', "19");
$objPHPExcel->getActiveSheet()->setCellValue('B84', "Pendapatan (kerugian) komprehensif lainnya");
$objPHPExcel->getActiveSheet()->setCellValue('E84', "");
$objPHPExcel->getActiveSheet()->setCellValue('B85', "a. Penyesuaian akibat penjabaran laporan keuangan dfalam mata uang asing");
$objPHPExcel->getActiveSheet()->setCellValue('E85', "436-437");
$objPHPExcel->getActiveSheet()->setCellValue('B86', "b. Keuntungan (kerugian) dari perubahan nilai aset keuangan dalam kelompok tersedia untuk dijual");
$objPHPExcel->getActiveSheet()->setCellValue('E86', "");
$objPHPExcel->getActiveSheet()->setCellValue('B87', "c. Bagian efektif lindung nilai arus kas");
$objPHPExcel->getActiveSheet()->setCellValue('E87', "");
$objPHPExcel->getActiveSheet()->setCellValue('B88', "d. Keuntungan revaluasi aset tetap");
$objPHPExcel->getActiveSheet()->setCellValue('E88', "");
$objPHPExcel->getActiveSheet()->setCellValue('B89', "e. Bagian pendapatan komprehensif lain dari entitas asosi");
$objPHPExcel->getActiveSheet()->setCellValue('E89', "");
$objPHPExcel->getActiveSheet()->setCellValue('B90', "f. Keuntungan (kerugian) aktuarial program imbalan pasti");
$objPHPExcel->getActiveSheet()->setCellValue('E90', "");
$objPHPExcel->getActiveSheet()->setCellValue('B91', "g. Pajak penghasilan terkait dengan penghasilan komprehensif lain");
$objPHPExcel->getActiveSheet()->setCellValue('E91', "");
$objPHPExcel->getActiveSheet()->setCellValue('B92', "h. Lainnya");
$objPHPExcel->getActiveSheet()->setCellValue('E92', "440-445");

$objPHPExcel->getActiveSheet()->setCellValue('A93', "20");
$objPHPExcel->getActiveSheet()->setCellValue('B93', "Selisih kuasi reorganisasi ³⁾");
$objPHPExcel->getActiveSheet()->setCellValue('E93', "");
$objPHPExcel->getActiveSheet()->setCellValue('A94', "21");
$objPHPExcel->getActiveSheet()->setCellValue('B94', "Selisih restrukturisasi entitas sepengendali");
$objPHPExcel->getActiveSheet()->setCellValue('E94', "457");
$objPHPExcel->getActiveSheet()->setCellValue('A95', "22");
$objPHPExcel->getActiveSheet()->setCellValue('B95', "Ekuitas lainnya");
$objPHPExcel->getActiveSheet()->setCellValue('E95', "");
$objPHPExcel->getActiveSheet()->setCellValue('A96', "23");
$objPHPExcel->getActiveSheet()->setCellValue('B96', "Cadangan");
$objPHPExcel->getActiveSheet()->setCellValue('E96', "");
$objPHPExcel->getActiveSheet()->setCellValue('B97', "a. Cadangan Umum");
$objPHPExcel->getActiveSheet()->setCellValue('E97', "451");
$objPHPExcel->getActiveSheet()->setCellValue('B98', "b. Cadangan Tujuan");
$objPHPExcel->getActiveSheet()->setCellValue('E98', "452");
$objPHPExcel->getActiveSheet()->setCellValue('A99', "24");
$objPHPExcel->getActiveSheet()->setCellValue('B99', "laba (Rugi)");
$objPHPExcel->getActiveSheet()->setCellValue('E99', "");
$objPHPExcel->getActiveSheet()->setCellValue('B100', "a. Tahun-tahun Lalu");
$objPHPExcel->getActiveSheet()->setCellValue('E100', "461-462");
$objPHPExcel->getActiveSheet()->setCellValue('B101', "b. Tahun Berjalan");
$objPHPExcel->getActiveSheet()->setCellValue('E101', "465-466");
$objPHPExcel->getActiveSheet()->setCellValue('B102', "TOTAL EKUITAS YANG DAPAT DIATRIBUSIKAN ");
$objPHPExcel->getActiveSheet()->setCellValue('E102', "");
$objPHPExcel->getActiveSheet()->setCellValue('B103', "KEPADA PEMILIK");

$objPHPExcel->getActiveSheet()->setCellValue('A105', "25");
$objPHPExcel->getActiveSheet()->setCellValue('B105', "Kepentingan non pengendali 6)");
$objPHPExcel->getActiveSheet()->setCellValue('E105', "");

$objPHPExcel->getActiveSheet()->setCellValue('B107', "TOTAL EKUITAS");
$objPHPExcel->getActiveSheet()->setCellValue('E107', "");


$objPHPExcel->getActiveSheet()->setCellValue('B109', "TOTAL LIABILITAS DAN EKUITAS");
$objPHPExcel->getActiveSheet()->setCellValue('E109', 490);

// SHEET 1 (BS)


$objPHPExcel->getActiveSheet()->setTitle('BS');




//SHEET 2 (PL)
$objPHPExcel->createSheet();
$objPHPExcel->setActiveSheetIndex(1);

$objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth(5);
$objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth(90);
$objPHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth(15);
$objPHPExcel->getActiveSheet()->getColumnDimension('D')->setWidth(15);


$objPHPExcel->setActiveSheetIndex(1)->mergeCells('A1:D1');
$objPHPExcel->setActiveSheetIndex(1)->mergeCells('A2:D2');
$objPHPExcel->setActiveSheetIndex(1)->mergeCells('A3:D3');
$objPHPExcel->setActiveSheetIndex(1)->mergeCells('C7:D7');

$objPHPExcel->getActiveSheet()->getStyle('A1:D3')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('C7:D8')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

$objPHPExcel->getActiveSheet()->getStyle('A7:D103')->applyFromArray($styleArrayBorder1);

$objPHPExcel->getActiveSheet()->getStyle('A1:A3')->applyFromArray($styleArrayFontBold);
$objPHPExcel->getActiveSheet()->getStyle('A7:D8')->applyFromArray($styleArrayFontBold);
$objPHPExcel->getActiveSheet()->getStyle('A9:D9')->applyFromArray($styleArrayFontBold);
$objPHPExcel->getActiveSheet()->getStyle('A10:D10')->applyFromArray($styleArrayFontBold);
$objPHPExcel->getActiveSheet()->getStyle('A18:D18')->applyFromArray($styleArrayFontBold);
$objPHPExcel->getActiveSheet()->getStyle('A11:D11')->applyFromArray($styleArrayFontBold);
$objPHPExcel->getActiveSheet()->getStyle('A14:D14')->applyFromArray($styleArrayFontBold);
$objPHPExcel->getActiveSheet()->getStyle('A19:D19')->applyFromArray($styleArrayFontBold);
$objPHPExcel->getActiveSheet()->getStyle('A36:D36')->applyFromArray($styleArrayFontBold);

$objPHPExcel->getActiveSheet()->setCellValue('A1', "LAPORAN LABA RUGI KOMPREHENSIF");
$objPHPExcel->getActiveSheet()->setCellValue('A2', 'PT BANK MNC INTERNASIONAL Tbk.');
$objPHPExcel->getActiveSheet()->setCellValue('A3', "Per Tanggal 30 September 2015 dan 2014 ");



$objPHPExcel->getActiveSheet()->setCellValue('A7', "No");
$objPHPExcel->getActiveSheet()->setCellValue('B7', "POS - POS");
$objPHPExcel->getActiveSheet()->setCellValue('C7', "BANK");

$objPHPExcel->getActiveSheet()->setCellValue('C8', "30-Sep-15");
$objPHPExcel->getActiveSheet()->setCellValue('D8', "30-Sep-14");

$objPHPExcel->getActiveSheet()->setCellValue('B9', "PENDAPATAN DAN BEBAN OPERASIONAL");
$objPHPExcel->getActiveSheet()->setCellValue('A10', "A");
$objPHPExcel->getActiveSheet()->setCellValue('B10', "Pendapatan dan Beban Bunga");
$objPHPExcel->getActiveSheet()->setCellValue('A11', "1.");
$objPHPExcel->getActiveSheet()->setCellValue('B11', "Pendapatan Bunga");
$objPHPExcel->getActiveSheet()->setCellValue('B12', "a. Rupiah");
$objPHPExcel->getActiveSheet()->setCellValue('B13', "b. Valuta Asing ");
$objPHPExcel->getActiveSheet()->setCellValue('A14', "2.");
$objPHPExcel->getActiveSheet()->setCellValue('B14', "Beban Bunga");
$objPHPExcel->getActiveSheet()->setCellValue('B15', "a. Rupiah");
$objPHPExcel->getActiveSheet()->setCellValue('B16', "b. Valuta Asing");
$objPHPExcel->getActiveSheet()->setCellValue('B17', "Pendapatan (Beban) Bunga Bersih");

$objPHPExcel->getActiveSheet()->setCellValue('A18', "B.");
$objPHPExcel->getActiveSheet()->setCellValue('B18', "Pendapatan dan Beban Operasional selain Bunga");
$objPHPExcel->getActiveSheet()->setCellValue('A19', "1.");
$objPHPExcel->getActiveSheet()->setCellValue('B19', "Pendapatan Operasional Selain Bunga");
$objPHPExcel->getActiveSheet()->setCellValue('B20', "a. Peningkatan nilai wajar");
$objPHPExcel->getActiveSheet()->setCellValue('B21', "i  . Surat Berharga");
$objPHPExcel->getActiveSheet()->setCellValue('B22', "ii . Kredit");
$objPHPExcel->getActiveSheet()->setCellValue('B23', "iii. Spot dan derivatif");
$objPHPExcel->getActiveSheet()->setCellValue('B24', "iv . Aset keuangan lainnya");
$objPHPExcel->getActiveSheet()->setCellValue('B25', "b. Penurunan nilai wajar liabilitas keuangan");
$objPHPExcel->getActiveSheet()->setCellValue('B26', "c. Keuntungan penjualan aset keuangan");
$objPHPExcel->getActiveSheet()->setCellValue('B27', "i  . Surat Berharga");
$objPHPExcel->getActiveSheet()->setCellValue('B28', "ii . Kredit");
$objPHPExcel->getActiveSheet()->setCellValue('B29', "iii. Aset keuangan lainnya");
$objPHPExcel->getActiveSheet()->setCellValue('B30', "d. Keuntungan transaksi spot dan derivatif (realised)");
$objPHPExcel->getActiveSheet()->setCellValue('B31', "e. Deviden ");
$objPHPExcel->getActiveSheet()->setCellValue('B32', "f. Keuntungan dari Penyertaan dengan equity Method");
$objPHPExcel->getActiveSheet()->setCellValue('B33', "g. Komisi/provisi/fee dan administrasi");
$objPHPExcel->getActiveSheet()->setCellValue('B34', "h. Pemulihan atas cadangan kerugian penurunan nilai");
$objPHPExcel->getActiveSheet()->setCellValue('B35', "i. Pendapatan lainnya");

$objPHPExcel->getActiveSheet()->setCellValue('A36', "2.");
$objPHPExcel->getActiveSheet()->setCellValue('B36', "Beban Operasional Selain Bunga");
$objPHPExcel->getActiveSheet()->setCellValue('B37', "a. Penurunan nilai wajar aset keuangan (mark to market) ");
$objPHPExcel->getActiveSheet()->setCellValue('B38', "i  . Surat Berharga");
$objPHPExcel->getActiveSheet()->setCellValue('B39', "ii . Kredit");
$objPHPExcel->getActiveSheet()->setCellValue('B40', "iii. Spot dan derivatif");
$objPHPExcel->getActiveSheet()->setCellValue('B41', "iv . Aset keuangan lainnya");
$objPHPExcel->getActiveSheet()->setCellValue('B42', "b. Peningkatan nilai wajar liabilitas keuangan (mart to market)");
$objPHPExcel->getActiveSheet()->setCellValue('B43', "c. Kerugian penjualan aset keuangan");
$objPHPExcel->getActiveSheet()->setCellValue('B44', "i  . Surat Berharga");
$objPHPExcel->getActiveSheet()->setCellValue('B45', "ii . Kredit");
$objPHPExcel->getActiveSheet()->setCellValue('B46', "iii. Aset keuangan lainnya");

$objPHPExcel->getActiveSheet()->setCellValue('B47', "d. Kerugian transaksi spot dan derivatif (realised)");
$objPHPExcel->getActiveSheet()->setCellValue('B48', "e. Kerugian penurunan nilai aset keuangan (impairment)");
$objPHPExcel->getActiveSheet()->setCellValue('B49', "i  . Surat Berharga");
$objPHPExcel->getActiveSheet()->setCellValue('B50', "ii . Kredit");
$objPHPExcel->getActiveSheet()->setCellValue('B51', "iii. Pembiayaan syariah");
$objPHPExcel->getActiveSheet()->setCellValue('B52', "iv . Aset keuangan lainnya");

$objPHPExcel->getActiveSheet()->setCellValue('B53', "f. Kerugian terkait risiko operasional *)");
$objPHPExcel->getActiveSheet()->setCellValue('B54', "g. Kerugian dari penyertaan dengan equity method");
$objPHPExcel->getActiveSheet()->setCellValue('B55', "h. Komisi/provisi/fee dan administrasi");
$objPHPExcel->getActiveSheet()->setCellValue('B56', "i. Kerugian penurunan nilai aset lainnya (non keuangan)");
$objPHPExcel->getActiveSheet()->setCellValue('B57', "j. Beban tenaga kerja");
$objPHPExcel->getActiveSheet()->setCellValue('B58', "k. Beban promosi");
$objPHPExcel->getActiveSheet()->setCellValue('B59', "l. Beban lainnya");

$objPHPExcel->getActiveSheet()->setCellValue('B60', "Pendapatan (Beban) Operasional Selain Bunga Bersih");
$objPHPExcel->getActiveSheet()->setCellValue('B62', "LABA (RUGI) OPERASIONAL");

$objPHPExcel->getActiveSheet()->setCellValue('A63', "PENDAPATAN DAN BEBAN NON OPERASIONAL");
$objPHPExcel->getActiveSheet()->setCellValue('A64', "1.");
$objPHPExcel->getActiveSheet()->setCellValue('B64', "Keuntungan (kerugian) penjualan aset tetap dan inventaris");
$objPHPExcel->getActiveSheet()->setCellValue('A65', "2.");
$objPHPExcel->getActiveSheet()->setCellValue('B65', "Keuntungan (kerugian) penjabaran transaksi valuta asing");
$objPHPExcel->getActiveSheet()->setCellValue('A66', "3.");
$objPHPExcel->getActiveSheet()->setCellValue('B66', "Pendapatan (beban) non operasional lainnya");
$objPHPExcel->getActiveSheet()->setCellValue('B68', "LABA (RUGI) OPERASIONAL");
$objPHPExcel->getActiveSheet()->setCellValue('B70', "LABA (RUGI) TAHUN BERJALAN");
$objPHPExcel->getActiveSheet()->setCellValue('B71', "Pajak Penghasilan");
$objPHPExcel->getActiveSheet()->setCellValue('B72', "a. Taksiran Pajak Tahun Berjalan");
$objPHPExcel->getActiveSheet()->setCellValue('B73', "b. Pendapatan (beban) pajak tangguhan");
$objPHPExcel->getActiveSheet()->setCellValue('B75', "LABA (RUGI) BERSIH");

$objPHPExcel->getActiveSheet()->setCellValue('A77', "PENGHASILAN KOMPREHENSIF LAIN");
$objPHPExcel->getActiveSheet()->setCellValue('A78', "1.");
$objPHPExcel->getActiveSheet()->setCellValue('B78', "Pos-pos yang tidak akan direklasifikasi ke laba rugi");
$objPHPExcel->getActiveSheet()->setCellValue('B79', "a. Keuntungan revaluasi aset tetap");
$objPHPExcel->getActiveSheet()->setCellValue('B80', "b. Keuntungan (kerugian) aktuarial program imbalan pasti");
$objPHPExcel->getActiveSheet()->setCellValue('B81', "c. Bagian pendapatan komprehensif lain dari entitas asosiasi");
$objPHPExcel->getActiveSheet()->setCellValue('B82', "d. Lainnya");
$objPHPExcel->getActiveSheet()->setCellValue('B83', "e. Pajak penghasilan terkait pos-pos yang tidak akan direklasifikasi ke laba rugi");
$objPHPExcel->getActiveSheet()->setCellValue('A84', "2.");
$objPHPExcel->getActiveSheet()->setCellValue('B84', "Pos-pos yang akan direklasifikasi ke laba rugi");
$objPHPExcel->getActiveSheet()->setCellValue('B85', "a. Penyesuaian akibat penjabaran laporan keuangan dalam mata uang asing");
$objPHPExcel->getActiveSheet()->setCellValue('B86', "b. Keuntungan (kerugian) dari perubahan nilai aset keuangan dalam kelompok tersedia untuk dijual");
$objPHPExcel->getActiveSheet()->setCellValue('B87', "c. Bagian efektif dari lindung nilai arus kas");
$objPHPExcel->getActiveSheet()->setCellValue('B88', "d. Lainnya");
$objPHPExcel->getActiveSheet()->setCellValue('B89', "e. Pajak penghasilan terkait pos-pos yang akan direklasifikasi ke laba rugi");

$objPHPExcel->getActiveSheet()->setCellValue('B90', "PENGHASILAN KOMPREHENSIF LAIN TAHUN BERJALAN - NET PAJAK PENGHASILAN TERKAIT");
$objPHPExcel->getActiveSheet()->setCellValue('B91', "TOTAL LABA (RUGI) KOMPREHENSIF TAHUN BERJALAN");
$objPHPExcel->getActiveSheet()->setCellValue('B92', "Laba yang dapat diatribusikan kepada :");
$objPHPExcel->getActiveSheet()->setCellValue('B93', "PEMILIK");
$objPHPExcel->getActiveSheet()->setCellValue('B94', "KEPENTINGAN NNON PENGENDALI");
$objPHPExcel->getActiveSheet()->setCellValue('B95', "TOTAL LABA TAHUN BERJALAN ");

$objPHPExcel->getActiveSheet()->setCellValue('B96', "Total Penghasilan Komprehensif Lain yang dapat diatribusikan kepada :");
$objPHPExcel->getActiveSheet()->setCellValue('B97', "PEMILIK");
$objPHPExcel->getActiveSheet()->setCellValue('B98', "KEPENTINGAN NNON PENGENDALI");
$objPHPExcel->getActiveSheet()->setCellValue('B99', "TOTAL PENGHASILAN KOMPREHENSIF LAIN TAHUN BERJALAN ");

$objPHPExcel->getActiveSheet()->setCellValue('B101', "TRANSFER LABA (RUGI) KE KANTOR PUSAT ");
$objPHPExcel->getActiveSheet()->setCellValue('B102', "DIVIDEN ");
$objPHPExcel->getActiveSheet()->setCellValue('B103', "LABA (RUGI) BERSIH PER SAHAM");


$objPHPExcel->getActiveSheet()->setTitle('PL');

###################################################################################3
//SHEET 3 (Rek Adm)
$objPHPExcel->createSheet();
$objPHPExcel->setActiveSheetIndex(2);


$objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth(5);
$objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth(90);
$objPHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth(15);
$objPHPExcel->getActiveSheet()->getColumnDimension('D')->setWidth(15);


$objPHPExcel->setActiveSheetIndex(2)->mergeCells('A1:D1');
$objPHPExcel->setActiveSheetIndex(2)->mergeCells('A2:D2');
$objPHPExcel->setActiveSheetIndex(2)->mergeCells('A3:D3');
$objPHPExcel->setActiveSheetIndex(2)->mergeCells('C7:D7');

$objPHPExcel->getActiveSheet()->getStyle('A1:D3')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('C7:D8')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

$objPHPExcel->getActiveSheet()->getStyle('A7:D54')->applyFromArray($styleArrayBorder1);


$objPHPExcel->getActiveSheet()->setCellValue('A1', "LAPORAN KOMITMEN & KONTINJENSI");
$objPHPExcel->getActiveSheet()->setCellValue('A2', 'PT BANK MNC INTERNASIONAL Tbk.');
$objPHPExcel->getActiveSheet()->setCellValue('A3', "Per 30 September 2015 dan 31 Desember 2014 ");



$objPHPExcel->getActiveSheet()->setCellValue('A7', "No");
$objPHPExcel->getActiveSheet()->setCellValue('B7', "POS - POS");
$objPHPExcel->getActiveSheet()->setCellValue('C7', "BANK");

$objPHPExcel->getActiveSheet()->setCellValue('C8', "30-Sep-15");
$objPHPExcel->getActiveSheet()->setCellValue('D8', "30-Sep-14");

$objPHPExcel->getActiveSheet()->setCellValue('A9', "I");
$objPHPExcel->getActiveSheet()->setCellValue('A9', "TAGIHAN KOMITMEN");
$objPHPExcel->getActiveSheet()->setCellValue('B10', "1. Fasilitas pinjaman yang belum ditarik ");
$objPHPExcel->getActiveSheet()->setCellValue('B11', "a. Rupiah ");
$objPHPExcel->getActiveSheet()->setCellValue('B12', "b. Valuta Asing ");
$objPHPExcel->getActiveSheet()->setCellValue('B13', "2. Posisi pembelian spot dan derivatif yang masih berjalan  ");
$objPHPExcel->getActiveSheet()->setCellValue('B14', "3. Lainnya ");
	
$objPHPExcel->getActiveSheet()->setCellValue('A15', "II");
$objPHPExcel->getActiveSheet()->setCellValue('A15', "KEWAJIBAN KOMITMEN");
$objPHPExcel->getActiveSheet()->setCellValue('B16', "1. Fasilitas kredit kepada nasabah yang belum ditarik ");
$objPHPExcel->getActiveSheet()->setCellValue('B17', "a. BUMN ");
$objPHPExcel->getActiveSheet()->setCellValue('B18', "i  . Committed ");
$objPHPExcel->getActiveSheet()->setCellValue('B19', "- Rupiah");
$objPHPExcel->getActiveSheet()->setCellValue('B20', "- Valuta Asing ");
$objPHPExcel->getActiveSheet()->setCellValue('B21', "ii . Uncommitted ");
$objPHPExcel->getActiveSheet()->setCellValue('B22', "- Rupiah ");
$objPHPExcel->getActiveSheet()->setCellValue('B23', "- Valuta Asing ");
$objPHPExcel->getActiveSheet()->setCellValue('B24', "b. Lainnya ");
$objPHPExcel->getActiveSheet()->setCellValue('B25', "i  . Committed ");
$objPHPExcel->getActiveSheet()->setCellValue('B26', "ii . Uncommitted");

$objPHPExcel->getActiveSheet()->setCellValue('B27', "2. Fasilitas kredit kepada bank lain yang belum ditarik ");
$objPHPExcel->getActiveSheet()->setCellValue('B28', "a. Committed ");
$objPHPExcel->getActiveSheet()->setCellValue('B29', "i  . Rupiah ");
$objPHPExcel->getActiveSheet()->setCellValue('B30', "ii . Valuta Asing ");
$objPHPExcel->getActiveSheet()->setCellValue('B31', "- Valuta Asing ");
$objPHPExcel->getActiveSheet()->setCellValue('B32', "b . Uncommitted ");
$objPHPExcel->getActiveSheet()->setCellValue('B33', "i  . Rupiah ");
$objPHPExcel->getActiveSheet()->setCellValue('B34', "ii . Valuta Asing ");

$objPHPExcel->getActiveSheet()->setCellValue('B35', "3. Irrevocable L/C yang masih berjalan ");
$objPHPExcel->getActiveSheet()->setCellValue('B36', "a . L/C luar negeri ");
$objPHPExcel->getActiveSheet()->setCellValue('B37', "b . L/C dalam negeri");
$objPHPExcel->getActiveSheet()->setCellValue('B38', "4 . Posisi penjualan spot dan derivatif yang masih berjalan  ");
$objPHPExcel->getActiveSheet()->setCellValue('B39', "5 . Lainnya");

$objPHPExcel->getActiveSheet()->setCellValue('A41', "III");
$objPHPExcel->getActiveSheet()->setCellValue('A41', "TAGIHAN KONTINJENSI");
$objPHPExcel->getActiveSheet()->setCellValue('B42', "1. Garansi yang diterima  ");
$objPHPExcel->getActiveSheet()->setCellValue('B43', "a. Rupiah ");
$objPHPExcel->getActiveSheet()->setCellValue('B44', "b. Valuta Asing ");
$objPHPExcel->getActiveSheet()->setCellValue('B45', "2. Pendapatan bunga dalam penyelesaian");
$objPHPExcel->getActiveSheet()->setCellValue('B46', "a. Bunga Kredit yang diberikan ");
$objPHPExcel->getActiveSheet()->setCellValue('B47', "b. Bunga Lainnya");
$objPHPExcel->getActiveSheet()->setCellValue('B48', "3. Lainnya ");


$objPHPExcel->getActiveSheet()->setCellValue('A50', "IV");
$objPHPExcel->getActiveSheet()->setCellValue('B50', "KEWAJIBAN KONTINJENSI");
$objPHPExcel->getActiveSheet()->setCellValue('B51', "1. Garansi yang diterima  ");
$objPHPExcel->getActiveSheet()->setCellValue('B52', "a. Rupiah ");
$objPHPExcel->getActiveSheet()->setCellValue('B53', "b. Valuta Asing ");
$objPHPExcel->getActiveSheet()->setCellValue('B54', "2. Lainnya ");



$objPHPExcel->getActiveSheet()->setTitle('Rek Adm');





//SHEET 4 (Rasio)
$objPHPExcel->createSheet();
$objPHPExcel->setActiveSheetIndex(3);

$objPHPExcel->getActiveSheet()->setTitle('Rasio');





$objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth(5);
$objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth(103);
$objPHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth(15);
$objPHPExcel->getActiveSheet()->getColumnDimension('D')->setWidth(15);


$objPHPExcel->setActiveSheetIndex(3)->mergeCells('A1:D1');
$objPHPExcel->setActiveSheetIndex(3)->mergeCells('A2:D2');
$objPHPExcel->setActiveSheetIndex(3)->mergeCells('A3:D3');
$objPHPExcel->setActiveSheetIndex(3)->mergeCells('C7:D7');

$objPHPExcel->getActiveSheet()->getStyle('A1:D3')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('C7:D8')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

$objPHPExcel->getActiveSheet()->getStyle('A7:D32')->applyFromArray($styleArrayBorder1);


$objPHPExcel->getActiveSheet()->setCellValue('A1', "Rasio");
$objPHPExcel->getActiveSheet()->setCellValue('A2', 'PT BANK MNC INTERNASIONAL Tbk.');
$objPHPExcel->getActiveSheet()->setCellValue('A3', "Per 30 September 2015 dan 2014 ");


$objPHPExcel->getActiveSheet()->setCellValue('A7', "No");
$objPHPExcel->getActiveSheet()->setCellValue('B7', "Rasio");
$objPHPExcel->getActiveSheet()->setCellValue('C7', "BANK");

$objPHPExcel->getActiveSheet()->setCellValue('C8', "30-Sep-15");
$objPHPExcel->getActiveSheet()->setCellValue('D8', "30-Sep-14");

$objPHPExcel->getActiveSheet()->setCellValue('A9', "Rasio Kinerja");
$objPHPExcel->getActiveSheet()->setCellValue('A11', "1");
$objPHPExcel->getActiveSheet()->setCellValue('B11', "Kewajiban Penyediaan Modal Minimum (KPMM) ");
$objPHPExcel->getActiveSheet()->setCellValue('A12', "2");
$objPHPExcel->getActiveSheet()->setCellValue('B12', "Aset produktif bermasalah dan aset non produktif bermasalah terhadap total aset produktif dan aset non produktif");
$objPHPExcel->getActiveSheet()->setCellValue('A13', "3");
$objPHPExcel->getActiveSheet()->setCellValue('B13', "Aset produktif bermasalah terhadap total aset produktif ");
$objPHPExcel->getActiveSheet()->setCellValue('A14', "4");
$objPHPExcel->getActiveSheet()->setCellValue('B14', "Cadangan kerugian penurunan nilai (CKPN) aset keuangan  terhadap aset produktif ");
$objPHPExcel->getActiveSheet()->setCellValue('A15', "5");
$objPHPExcel->getActiveSheet()->setCellValue('B15', "NPL gross ");
$objPHPExcel->getActiveSheet()->setCellValue('A16', "6");
$objPHPExcel->getActiveSheet()->setCellValue('B16', "NPL net ");
$objPHPExcel->getActiveSheet()->setCellValue('A17', "7");
$objPHPExcel->getActiveSheet()->setCellValue('B17', "Return on Asset (ROA) ");
$objPHPExcel->getActiveSheet()->setCellValue('A18', "8");
$objPHPExcel->getActiveSheet()->setCellValue('B18', "Return on Equity (ROE)
 ");
$objPHPExcel->getActiveSheet()->setCellValue('A19', "9");
$objPHPExcel->getActiveSheet()->setCellValue('B19', "Net Interest Margin (NIM) ");
$objPHPExcel->getActiveSheet()->setCellValue('A20', "10");
$objPHPExcel->getActiveSheet()->setCellValue('B20', "Biaya Operasional terhadap Pendapatan Operasional (BOPO) ");
$objPHPExcel->getActiveSheet()->setCellValue('A21', "11");
$objPHPExcel->getActiveSheet()->setCellValue('B21', "Loan to Deposit Ratio (LDR) ");

$objPHPExcel->getActiveSheet()->setCellValue('A22', "Kepatuhan (Compliance)");
$objPHPExcel->getActiveSheet()->setCellValue('A23', "1. ");
$objPHPExcel->getActiveSheet()->setCellValue('B23', "a.   Persentase pelanggaran BMPK");
$objPHPExcel->getActiveSheet()->setCellValue('B24', "i .  Pihak terkait");
$objPHPExcel->getActiveSheet()->setCellValue('B25', "ii.  Pihak tidak terkait");
$objPHPExcel->getActiveSheet()->setCellValue('B26', "b.   Persentase pelampauan BMPK");
$objPHPExcel->getActiveSheet()->setCellValue('B27', "i .  Pihak terkait");
$objPHPExcel->getActiveSheet()->setCellValue('B28', "ii.  Pihak tidak terkait");
$objPHPExcel->getActiveSheet()->setCellValue('A29', "2. ");
$objPHPExcel->getActiveSheet()->setCellValue('B29', "Giro Wajib Minimum (GWM) ");
$objPHPExcel->getActiveSheet()->setCellValue('B30', "a.   GWM Utama Rupiah");
$objPHPExcel->getActiveSheet()->setCellValue('B31', "b.   GWM Valuta asing");
$objPHPExcel->getActiveSheet()->setCellValue('A32', "3. ");
$objPHPExcel->getActiveSheet()->setCellValue('B32', "Posisi Devisa Neto (PDN) secara keseluruhan ");



####################################################################################
//SHEET 5 (Derivatif)
$objPHPExcel->createSheet();
$objPHPExcel->setActiveSheetIndex(4);



$objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth(5);
$objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth(30);
$objPHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth(15);
$objPHPExcel->getActiveSheet()->getColumnDimension('D')->setWidth(15);
$objPHPExcel->getActiveSheet()->getColumnDimension('E')->setWidth(15);
$objPHPExcel->getActiveSheet()->getColumnDimension('F')->setWidth(15);
$objPHPExcel->getActiveSheet()->getColumnDimension('G')->setWidth(15);


$objPHPExcel->setActiveSheetIndex(4)->mergeCells('A1:G1');
$objPHPExcel->setActiveSheetIndex(4)->mergeCells('A2:G2');
$objPHPExcel->setActiveSheetIndex(4)->mergeCells('A3:G3');

$objPHPExcel->setActiveSheetIndex(4)->mergeCells('A7:A9');
$objPHPExcel->setActiveSheetIndex(4)->mergeCells('B7:B9');

$objPHPExcel->setActiveSheetIndex(4)->mergeCells('C7:G7');

$objPHPExcel->setActiveSheetIndex(4)->mergeCells('C8:C9');
$objPHPExcel->setActiveSheetIndex(4)->mergeCells('D8:E8');
$objPHPExcel->setActiveSheetIndex(4)->mergeCells('F8:G8');

$objPHPExcel->setActiveSheetIndex(4)->mergeCells('B11:G11');
$objPHPExcel->setActiveSheetIndex(4)->mergeCells('B26:G26');
$objPHPExcel->setActiveSheetIndex(4)->mergeCells('B39:G39');



$objPHPExcel->getActiveSheet()->getStyle('A1:G3')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('A7:G9')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('A7:G9')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::VERTICAL_CENTER);

$objPHPExcel->getActiveSheet()->getStyle('A7:A40')->applyFromArray($styleArrayBorder2);
$objPHPExcel->getActiveSheet()->getStyle('A7:A9')->applyFromArray($styleArrayBorder2);
$objPHPExcel->getActiveSheet()->getStyle('B7:G40')->applyFromArray($styleArrayBorder1);

$objPHPExcel->getActiveSheet()->getStyle('A7:G9')->applyFromArray($styleArrayFontBold);

$objPHPExcel->getActiveSheet()->getStyle('A11:G11')->applyFromArray($styleArrayFontBold);
$objPHPExcel->getActiveSheet()->getStyle('A26:G26')->applyFromArray($styleArrayFontBold);
$objPHPExcel->getActiveSheet()->getStyle('A39:G39')->applyFromArray($styleArrayFontBold);
$objPHPExcel->getActiveSheet()->getStyle('A40:G40')->applyFromArray($styleArrayFontBold);

$objPHPExcel->getActiveSheet()->setCellValue('A1', "LAPORAN TRANSAKSI SPOT DAN DERIVATIF");
$objPHPExcel->getActiveSheet()->setCellValue('A2', 'PT BANK MNC INTERNASIONAL Tbk.');
$objPHPExcel->getActiveSheet()->setCellValue('A3', "Per 30 September 2015");

$objPHPExcel->getActiveSheet()->setCellValue('A7', "No");
$objPHPExcel->getActiveSheet()->setCellValue('B7', "Transaksi");
$objPHPExcel->getActiveSheet()->setCellValue('C7', "BANK");
$objPHPExcel->getActiveSheet()->setCellValue('C8', "Nilai Notional");
$objPHPExcel->getActiveSheet()->setCellValue('D8', "Tujuan");
$objPHPExcel->getActiveSheet()->setCellValue('F8', "Tagihan dan Liabilitas Derivatif");
$objPHPExcel->getActiveSheet()->setCellValue('D9', "Trading");
$objPHPExcel->getActiveSheet()->setCellValue('E9', "Hedging");
$objPHPExcel->getActiveSheet()->setCellValue('F9', "Tagihan");
$objPHPExcel->getActiveSheet()->setCellValue('G9', "Liabilitas");

$objPHPExcel->getActiveSheet()->setCellValue('A11', "A.");
$objPHPExcel->getActiveSheet()->setCellValue('B11', "Terkait dengan Nilai Tukar");
$objPHPExcel->getActiveSheet()->setCellValue('A12', "1");
$objPHPExcel->getActiveSheet()->setCellValue('B12', "Spot");
$objPHPExcel->getActiveSheet()->setCellValue('A14', "2");
$objPHPExcel->getActiveSheet()->setCellValue('B14', "Forward");
$objPHPExcel->getActiveSheet()->setCellValue('A16', "3");
$objPHPExcel->getActiveSheet()->setCellValue('B16', "Option");
$objPHPExcel->getActiveSheet()->setCellValue('B17', "a. Jual");
$objPHPExcel->getActiveSheet()->setCellValue('B18', "b. Beli");
$objPHPExcel->getActiveSheet()->setCellValue('A20', "4");
$objPHPExcel->getActiveSheet()->setCellValue('B20', "Future");
$objPHPExcel->getActiveSheet()->setCellValue('A22', "5");
$objPHPExcel->getActiveSheet()->setCellValue('B22', "Swap");
$objPHPExcel->getActiveSheet()->setCellValue('A24', "6");
$objPHPExcel->getActiveSheet()->setCellValue('B24', "Lainnya");

$objPHPExcel->getActiveSheet()->setCellValue('A26', "B.");
$objPHPExcel->getActiveSheet()->setCellValue('B26', "Terkait dengan Suku Bunga");
$objPHPExcel->getActiveSheet()->setCellValue('A27', "1");
$objPHPExcel->getActiveSheet()->setCellValue('B27', "Forward");
$objPHPExcel->getActiveSheet()->setCellValue('A29', "2");
$objPHPExcel->getActiveSheet()->setCellValue('B29', "Option");
$objPHPExcel->getActiveSheet()->setCellValue('B30', "a. Jual");
$objPHPExcel->getActiveSheet()->setCellValue('B31', "b. Beli");
$objPHPExcel->getActiveSheet()->setCellValue('A33', "3");
$objPHPExcel->getActiveSheet()->setCellValue('B33', "Future");
$objPHPExcel->getActiveSheet()->setCellValue('A35', "4");
$objPHPExcel->getActiveSheet()->setCellValue('B35', "Swap");
$objPHPExcel->getActiveSheet()->setCellValue('A37', "5");
$objPHPExcel->getActiveSheet()->setCellValue('B37', "Lainnya");
$objPHPExcel->getActiveSheet()->setCellValue('A39', "C.");
$objPHPExcel->getActiveSheet()->setCellValue('B39', "Lainnya");
$objPHPExcel->getActiveSheet()->setCellValue('B40', "JUMLAH");

$objPHPExcel->getActiveSheet()->setTitle('Derivatif');


//SHEET 6 (KA dan CKPN)
$objPHPExcel->createSheet();
$objPHPExcel->setActiveSheetIndex(5);

$objPHPExcel->getActiveSheet()->setTitle('KA dan CKPN');

//SHEET 7 (Pengurus dan Pmlk)
$objPHPExcel->createSheet();
$objPHPExcel->setActiveSheetIndex(6);

$objPHPExcel->getActiveSheet()->setTitle('Pengurus dan Pmlk');


//SHEET 8 (Arus Kas)
$objPHPExcel->createSheet();
$objPHPExcel->setActiveSheetIndex(7);

$objPHPExcel->getActiveSheet()->setTitle('Arus Kas');



// Redirect output to a client’s web browser (Excel5)
header('Content-Type: application/vnd.ms-excel');
header('Content-Disposition: attachment;filename="name_of_file.xls"');
header('Cache-Control: max-age=0');
$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
$objWriter->save('php://output');